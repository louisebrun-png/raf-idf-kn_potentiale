import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import xml.etree.ElementTree as ET
import unicodedata
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURATION DE LA PAGE — CABINET POTENTIALE
# ==============================================================================
st.set_page_config(
    page_title="Potentiale — RAF IDF",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header {
        background-color: #1E293B;
        color: #FFFFFF;
        padding: 1.2rem 1.5rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        border-left: 6px solid #0EA5E9;
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
    .main-header p { margin: 0.3rem 0 0 0; font-size: 0.95rem; color: #CBD5E1; }
    .kpi-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }
    .kpi-title { font-size: 0.85rem; color: #475569; font-weight: 600; text-transform: uppercase; }
    .kpi-val { font-size: 1.4rem; color: #0F172A; font-weight: 700; margin-top: 0.2rem; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1>CABINET POTENTIALE — Rapprochement Automatisé des Achats (RAF IDF)</h1>
        <p>Architecture SI & Contrôle Mensuel des Flux Récurrents, Consommations et Frais Ponctuels</p>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# FONCTIONS UTILITAIRES DE LECTURE & NORMALISATION
# ==============================================================================
def read_xlsx_robust(uploaded_file):
    """Lit un fichier Excel même si le XML contient des filtres complexes d'export ERP."""
    uploaded_file.seek(0)
    try:
        return pd.read_excel(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        with zipfile.ZipFile(uploaded_file, 'r') as z:
            shared = []
            if 'xl/sharedStrings.xml' in z.namelist():
                root_ss = ET.fromstring(z.read('xl/sharedStrings.xml'))
                ns_ss = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in root_ss.findall('s:si', ns_ss):
                    texts = [t.text or '' for t in si.findall('.//s:t', ns_ss)]
                    shared.append(''.join(texts))
            sheet_xml = [n for n in z.namelist() if n.startswith('xl/worksheets/sheet')][0]
            root_sh = ET.fromstring(z.read(sheet_xml))
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

            def col_idx(ref):
                col_str = ''.join(c for c in ref if c.isalpha())
                idx = 0
                for c in col_str:
                    idx = idx * 26 + (ord(c) - ord('A') + 1)
                return idx - 1

            rows_data = []
            max_col = 0
            for row in root_sh.findall('.//s:sheetData/s:row', ns):
                r_dict = {}
                for c in row.findall('s:c', ns):
                    ref = c.attrib.get('r', '')
                    c_i = col_idx(ref)
                    if c_i > max_col:
                        max_col = c_i
                    t = c.attrib.get('t', '')
                    v_el = c.find('s:v', ns)
                    is_el = c.find('s:is', ns)
                    val = None
                    if t == 's' and v_el is not None and v_el.text is not None:
                        val = shared[int(v_el.text)]
                    elif t == 'inlineStr' and is_el is not None:
                        val = ''.join(x.text or '' for x in is_el.findall('.//s:t', ns))
                    elif v_el is not None and v_el.text is not None:
                        val = v_el.text
                    r_dict[c_i] = val
                rows_data.append(r_dict)
        matrix = [[r_dict.get(i, None) for i in range(max_col + 1)] for r_dict in rows_data]
        return pd.DataFrame(matrix[1:], columns=matrix[0])

def read_csv_robust(uploaded_file):
    uploaded_file.seek(0)
    try:
        df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
        if df.shape[1] == 1:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=',', encoding='utf-8')
        return df
    except Exception:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, sep=';', encoding='latin1')

def norm_txt(s):
    s = str(s or '').lower().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]', '', s)

def get_erp_tech_id(r):
    lib = str(r.get('Libellé RFC', ''))
    if '//' in lib:
        p1 = lib.split('//', 1)[1].split('/', 1)[0].strip()
        if p1:
            return p1
    mat = str(r.get('Matricule', '')).strip()
    if mat and mat.lower() != 'nan':
        return mat
    return ''

def depot_to_famille(dep):
    s = str(dep)
    if 'Ope' in s:
        return 'Opérateur'
    if 'MOB' in s:
        return 'Mobile'
    if 'Heberg' in s:
        return 'Hébergement'
    if 'FAS' in s:
        return 'FAS'
    return 'Opérateur'

def detect_annexe_family(df_csv):
    """Identifie automatiquement la famille d'une annexe CSV fournisseur."""
    libs = " ".join(df_csv['Libellé produit'].dropna().astype(str).head(40).tolist()).upper()
    refs = " ".join(df_csv['Réf. produit'].dropna().astype(str).head(40).tolist()).upper()
    if '[MES]' in libs or 'OUVERTURE DE LIGNE' in libs or 'FAS ' in libs or 'PORTABILIT' in libs:
        return 'FAS'
    if 'VMWARE' in libs or 'PROXMOX' in libs or 'VHDD' in libs or 'VRAM' in libs or 'VCPU' in libs or 'VDOM' in libs:
        return 'Hébergement'
    if 'CALL CONNECT' in libs or 'CALL PRO' in libs or 'MVNO' in libs:
        return 'Mobile'
    return 'Opérateur'

# ==============================================================================
# BARRE LATÉRALE — CHARGEMENT DES SOURCES MENSUELLES
# ==============================================================================
with st.sidebar:
    st.header("1. Référentiels ERP")
    file_sim = st.file_uploader("Simulation d'Achat ERP (.xlsx)", type=['xlsx'])
    file_ssc = st.file_uploader("Extraction Parc SSC Actif (.xlsx)", type=['xlsx'])

    st.header("2. Annexes Fournisseur (CSV)")
    files_csv = st.file_uploader(
        "Déposez les 4 annexes CSV du mois (Opérateur, Mobile, Hébergement, FAS)",
        type=['csv'],
        accept_multiple_files=True
    )
    st.info("💡 Aucune donnée n'est enregistrée sur le serveur : tout est traité en mémoire vive.")

# ==============================================================================
# MOTEUR DE RAPPROCHEMENT MULTI-PASSES
# ==============================================================================
if file_sim and file_ssc and files_csv:
    with st.spinner("Analyse et rapprochement multi-passes en cours..."):
        df_sim = read_xlsx_robust(file_sim)
        df_ssc = read_xlsx_robust(file_ssc)

        # 1. Lecture dynamique des dépôts dans le Parc SSC (sans altérer la simulation)
        ssc_active = df_ssc[~df_ssc['Statut RFC'].isin(['Fermée', 'Résiliée.Dans les délais', 'Résiliée.En attente'])].copy()
        ssc_dedup = ssc_active.drop_duplicates(subset=['Code SSC', 'Libellé RFC'])
        depot_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Abonnement Code dépôt'].to_dict()
        charge_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Nom chargé factu SSC'].to_dict()

        ssc_active['k_norm'] = (
            ssc_active['Code SSC'].astype(str).str.strip() + "||" +
            ssc_active['Libellé RFC'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        )
        map_norm = ssc_active.drop_duplicates('k_norm').set_index('k_norm')['Abonnement Code dépôt'].to_dict()

        sim_valid = df_sim[pd.to_numeric(df_sim['Montant HT'], errors='coerce').notna()].copy()
        sim_valid['Montant HT'] = pd.to_numeric(sim_valid['Montant HT'], errors='coerce')
        sim_valid['Quantite'] = pd.to_numeric(sim_valid['Quantite'], errors='coerce').fillna(1.0)
        sim_valid['Prix Unitaire'] = pd.to_numeric(sim_valid['Prix Unitaire'], errors='coerce').fillna(0.0)

        sim_valid['Depot_SSC'] = sim_valid.apply(
            lambda r: depot_map.get((str(r['Code SSC']).strip(), str(r['Libellé RFC']).strip()), None),
            axis=1
        )
        miss = sim_valid['Depot_SSC'].isna()
        sim_valid.loc[miss, 'Depot_SSC'] = sim_valid.loc[miss].apply(
            lambda r: map_norm.get(str(r['Code SSC']).strip() + "||" + " ".join(str(r['Libellé RFC']).split()), 'Non affecté'),
            axis=1
        )
        sim_valid['Depot_SSC'] = sim_valid['Depot_SSC'].fillna('Non affecté')
        sim_valid['Charge_Factu'] = sim_valid.apply(
            lambda r: charge_map.get((str(r['Code SSC']).strip(), str(r['Libellé RFC']).strip()), ''),
            axis=1
        ).fillna('Import Initial')

        sim_valid['Famille'] = sim_valid['Depot_SSC'].apply(depot_to_famille)
        sim_valid['Tech_ID'] = sim_valid.apply(get_erp_tech_id, axis=1)
        sim_valid['Art_Num'] = sim_valid['Code Article'].astype(str).str.extract(r'(\d+)')[0].fillna('')
        sim_valid['Prod_Lib'] = sim_valid['Libellé RFC'].apply(lambda s: str(s).split('//')[0].strip())
        sim_valid['Client_Norm'] = sim_valid['Client'].apply(norm_txt)
        sim_valid['Prod_Norm'] = sim_valid['Prod_Lib'].apply(norm_txt)

        # 2. Chargement et classification des annexes CSV fournisseur
        f_list = []
        for f_csv in files_csv:
            df_c = read_csv_robust(f_csv)
            fam_det = detect_annexe_family(df_c)
            df_c['Famille'] = fam_det
            df_c['Source_CSV'] = f_csv.name
            f_list.append(df_c)

        f_all = pd.concat(f_list, ignore_index=True)
        f_all['Total H.T.'] = pd.to_numeric(f_all['Total H.T.'], errors='coerce').fillna(0.0)
        f_all['P.U. HT'] = pd.to_numeric(f_all['P.U. HT'], errors='coerce').fillna(0.0)
        f_all['Qté'] = pd.to_numeric(f_all['Qté'], errors='coerce').fillna(1.0)

        f_fas = f_all[f_all['Famille'] == 'FAS'].copy()
        f_rec = f_all[f_all['Famille'] != 'FAS'].copy()
        f_conso = f_rec[f_rec['Réf. produit'].astype(str).str.contains('CONSO', case=False, na=False)].copy()
        f_abo = f_rec[~f_rec['Réf. produit'].astype(str).str.contains('CONSO', case=False, na=False)].copy()

        f_abo['Tech_ID'] = f_abo['Matricule fournisseur'].fillna('').astype(str).str.replace('.0', '', regex=False).str.strip()
        f_abo['Art_Num'] = f_abo['Réf. produit'].astype(str).str.strip()
        f_abo['Client_Norm'] = f_abo['Client final'].apply(norm_txt)
        f_abo['Prod_Norm'] = f_abo['Libellé produit'].apply(norm_txt)

        # 3. Exécution du rapprochement en 3 passes
        sim_valid['Statut_RAF'] = 'Non rapproché (Présent ERP / Absent Fournisseur)'
        sim_valid['Pass_RAF'] = ''
        sim_valid['Fourn_HT'] = 0.0
        sim_valid['Ecart_HT'] = sim_valid['Montant HT']

        f_abo['Statut_RAF'] = 'Non rapproché (Présent Fournisseur / Absent ERP)'
        f_abo['Pass_RAF'] = ''

        for fam in ['Opérateur', 'Mobile', 'Hébergement']:
            s_idx = sim_valid[sim_valid['Famille'] == fam].index
            f_idx = f_abo[f_abo['Famille'] == fam].index

            # Passe 1 : Identifiant Technique + Montant HT exact (<= 0.05 €)
            f_by_id = {}
            for fi in f_idx:
                tid = f_abo.at[fi, 'Tech_ID']
                if tid:
                    f_by_id.setdefault(tid, []).append(fi)

            for si in s_idx:
                tid = sim_valid.at[si, 'Tech_ID']
                if not tid or tid not in f_by_id:
                    continue
                s_ht = float(sim_valid.at[si, 'Montant HT'])
                for fi in f_by_id[tid]:
                    if f_abo.at[fi, 'Pass_RAF'] == '' and abs(s_ht - float(f_abo.at[fi, 'Total H.T.'])) <= 0.05:
                        f_ht = float(f_abo.at[fi, 'Total H.T.'])
                        sim_valid.at[si, 'Statut_RAF'] = '1. Parfait (Identifiant Technique + Prix)'
                        sim_valid.at[si, 'Pass_RAF'] = 'P1'
                        sim_valid.at[si, 'Fourn_HT'] = f_ht
                        sim_valid.at[si, 'Ecart_HT'] = s_ht - f_ht
                        f_abo.at[fi, 'Statut_RAF'] = '1. Parfait (Identifiant Technique + Prix)'
                        f_abo.at[fi, 'Pass_RAF'] = 'P1'
                        break

            # Passe 2 : Client + Code Article Numérique (ex: KN5541 <-> 5541)
            s_rem = [si for si in s_idx if sim_valid.at[si, 'Pass_RAF'] == '' and sim_valid.at[si, 'Art_Num'] not in ('', '2912')]
            f_rem_by_ca = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    ca = (f_abo.at[fi, 'Client_Norm'], f_abo.at[fi, 'Art_Num'])
                    f_rem_by_ca.setdefault(ca, []).append(fi)

            s_by_ca = {}
            for si in s_rem:
                ca = (sim_valid.at[si, 'Client_Norm'], sim_valid.at[si, 'Art_Num'])
                s_by_ca.setdefault(ca, []).append(si)

            for ca, sis in s_by_ca.items():
                if ca not in f_rem_by_ca:
                    continue
                fis = [fi for fi in f_rem_by_ca[ca] if f_abo.at[fi, 'Pass_RAF'] == '']
                if not fis:
                    continue
                sum_s = sum(float(sim_valid.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                label = '2. Client + Code Article (Exact)' if is_exact else '2b. Client + Code Article (Écart Qté/Prorata)'
                for si in sis:
                    s_ht = float(sim_valid.at[si, 'Montant HT'])
                    alloc_f = (sum_f * (s_ht / sum_s)) if sum_s > 0 else 0.0
                    sim_valid.at[si, 'Statut_RAF'] = label
                    sim_valid.at[si, 'Pass_RAF'] = 'P2'
                    sim_valid.at[si, 'Fourn_HT'] = alloc_f
                    sim_valid.at[si, 'Ecart_HT'] = s_ht - alloc_f
                for fi in fis:
                    f_abo.at[fi, 'Statut_RAF'] = label
                    f_abo.at[fi, 'Pass_RAF'] = 'P2'

            # Passe 3 : Client + Libellé Produit (pour lignes reprise KN2912 sans identifiant unitaire)
            s_rem3 = [si for si in s_idx if sim_valid.at[si, 'Pass_RAF'] == '']
            f_rem_by_cp = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    cp = (f_abo.at[fi, 'Client_Norm'], f_abo.at[fi, 'Prod_Norm'])
                    f_rem_by_cp.setdefault(cp, []).append(fi)
            s_by_cp = {}
            for si in s_rem3:
                cp = (sim_valid.at[si, 'Client_Norm'], sim_valid.at[si, 'Prod_Norm'])
                s_by_cp.setdefault(cp, []).append(si)

            for cp, sis in s_by_cp.items():
                if cp not in f_rem_by_cp:
                    continue
                fis = [fi for fi in f_rem_by_cp[cp] if f_abo.at[fi, 'Pass_RAF'] == '']
                if not fis:
                    continue
                sum_s = sum(float(sim_valid.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in fis)
                if abs(sum_s - sum_f) <= 0.10:
                    for si in sis:
                        sim_valid.at[si, 'Statut_RAF'] = '3. Client + Libellé Produit'
                        sim_valid.at[si, 'Pass_RAF'] = 'P3'
                        sim_valid.at[si, 'Fourn_HT'] = float(sim_valid.at[si, 'Montant HT'])
                        sim_valid.at[si, 'Ecart_HT'] = 0.0
                    for fi in fis:
                        f_abo.at[fi, 'Statut_RAF'] = '3. Client + Libellé Produit'
                        f_abo.at[fi, 'Pass_RAF'] = 'P3'

    # ==========================================================================
    # AFFICHAGE DES INDICATEURS & ONGLETS DÉCISIONNELS
    # ==========================================================================
    tot_erp_rec = sim_valid[sim_valid['Famille'] != 'FAS']['Montant HT'].sum()
    tot_fourn_abo = f_abo['Total H.T.'].sum()
    tot_conso = f_conso['Total H.T.'].sum()
    tot_fas = f_fas['Total H.T.'].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='kpi-card'><div class='kpi-title'>Simulation ERP (Récurrent)</div><div class='kpi-val'>{tot_erp_rec:,.2f} €</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-card'><div class='kpi-title'>Facturé Fournisseur (Abo)</div><div class='kpi-val'>{tot_fourn_abo:,.2f} €</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='kpi-card'><div class='kpi-title'>Consommations Variables</div><div class='kpi-val'>{tot_conso:,.2f} €</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='kpi-card'><div class='kpi-title'>Frais Ponctuels (FAS)</div><div class='kpi-val'>{tot_fas:,.2f} €</div></div>", unsafe_allow_html=True)

    st.markdown("###")

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Synthèse par Dépôt",
        "➕ Présent Fournisseur / Absent ERP (RFC manquante)",
        "➖ Présent ERP / Absent Fournisseur (Résiliation / Date)",
        "📦 FAS & Consommations"
    ])

    with tab1:
        rows_s = []
        for fam in ['Opérateur', 'Mobile', 'Hébergement']:
            s_f = sim_valid[sim_valid['Famille'] == fam]
            f_f = f_abo[f_abo['Famille'] == fam]
            c_f = f_conso[f_conso['Famille'] == fam]
            m_s = s_f[s_f['Pass_RAF'] != '']
            u_s = s_f[s_f['Pass_RAF'] == '']
            u_f = f_f[f_f['Pass_RAF'] == '']
            rows_s.append({
                'Famille / Dépôt': fam,
                'Simulation ERP HT': round(s_f['Montant HT'].sum(), 2),
                'Facturé Fournisseur Abo HT': round(f_f['Total H.T.'].sum(), 2),
                'Écart Global Abo HT': round(s_f['Montant HT'].sum() - f_f['Total H.T.'].sum(), 2),
                'Montant Rapproché HT': round(m_s['Montant HT'].sum(), 2),
                'Reste ERP (Résilié/Date) HT': round(u_s['Montant HT'].sum(), 2),
                'Reste Fournisseur (RFC Manquante) HT': round(u_f['Total H.T.'].sum(), 2),
                'Consommations Pied Facture HT': round(c_f['Total H.T.'].sum(), 2)
            })
        df_synthese = pd.DataFrame(rows_s)
        st.dataframe(df_synthese, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("Lignes facturées par le fournisseur absentes de l'ERP (Client par Client & Article)")
        f_unm = f_abo[f_abo['Pass_RAF'] == ''][[
            'Famille', 'Client final', 'Site', 'Tech_ID', 'Réf. produit',
            'Libellé produit', 'Date début', 'Qté', 'P.U. HT', 'Total H.T.', 'Code offre'
        ]].sort_values(by=['Famille', 'Client final', 'Total H.T.'], ascending=[True, True, False])
        st.dataframe(f_unm, use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("Lignes simulées dans l'ERP non retrouvées chez le fournisseur (Résiliations / Mouvements)")
        s_unm = sim_valid[(sim_valid['Pass_RAF'] == '') & (sim_valid['Famille'] != 'FAS')][[
            'Depot_SSC', 'Charge_Factu', 'Code client', 'Client', 'Code SSC',
            'Code Article', 'Libellé RFC', 'Tech_ID', 'Quantite', 'Prix Unitaire', 'Montant HT'
        ]].sort_values(by=['Depot_SSC', 'Client', 'Montant HT'], ascending=[True, True, False])
        st.dataframe(s_unm, use_container_width=True, hide_index=True)

    with tab4:
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader(f"Frais de Mise en Service — FAS ({tot_fas:,.2f} € HT)")
            st.dataframe(f_fas[['Client final', 'Libellé produit', 'Date début', 'Qté', 'Total H.T.']], use_container_width=True, hide_index=True)
        with col_b:
            st.subheader(f"Consommations Variables ({tot_conso:,.2f} € HT)")
            st.dataframe(f_conso[['Famille', 'Client final', 'Libellé produit', 'Qté', 'Total H.T.']], use_container_width=True, hide_index=True)

    # Export Excel Complet
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_synthese.to_excel(writer, sheet_name='1_Synthese_RAF', index=False)
        f_unm.to_excel(writer, sheet_name='2_Fourn_Absent_ERP', index=False)
        s_unm.to_excel(writer, sheet_name='3_ERP_Absent_Fourn', index=False)
        sim_valid[sim_valid['Pass_RAF'] != ''][['Depot_SSC', 'Statut_RAF', 'Client', 'Code SSC', 'Code Article', 'Libellé RFC', 'Tech_ID', 'Montant HT', 'Fourn_HT', 'Ecart_HT']].to_excel(writer, sheet_name='4_Audit_Rapproche', index=False)
        f_fas.to_excel(writer, sheet_name='5_FAS_Ponctuels', index=False)
        f_conso.to_excel(writer, sheet_name='6_Consommations', index=False)

    st.download_button(
        label="📥 Télécharger le Classeur Complet de Rapprochement & Écarts (.xlsx)",
        data=output.getvalue(),
        file_name="Potentiale_RAF_IDF_Rapprochement_Mensuel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Chargez la **Simulation d'Achat ERP**, le **Parc SSC Actif** et vos **4 annexes CSV fournisseurs** dans le menu de gauche pour lancer l'analyse.")

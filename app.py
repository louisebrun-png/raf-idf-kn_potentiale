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
    page_title="Potentiale — Outil Métier RAF",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# FONCTIONS ROBUSTES DE LECTURE & NORMALISATION
# ==============================================================================
def read_xlsx_robust(uploaded_file):
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

def normalize_depot_simple(dep):
    s = str(dep or '').strip()
    if 'Ope' in s:
        return 'AF-KN-Ope'
    if 'MOB' in s:
        return 'AF-KN-MOB'
    if 'Heberg' in s:
        return 'AF-KN-Heberg'
    if 'FAS' in s:
        return 'AF-KN-FAS'
    return 'AF-KN-Ope'

def classify_line_depot(text):
    t = str(text or '').upper()
    if any(k in t for k in ['PROXMOX', 'VMWARE', 'VDOM', 'VHDD', 'VRAM', 'VCPU', 'VBACKUP', 'VEEAM', 'SPLA', 'SQL', 'CLOUD', 'VLAN', 'DATA CENTER', 'HEBERG', 'HÉBERG', 'ISI FAX']):
        return 'AF-KN-Heberg'
    if any(k in t for k in ['CONNECT PRO', 'APN', 'M2M', 'BACKUP ONLY', 'SECOURS']):
        return 'AF-KN-Ope'
    if any(k in t for k in ['CALL CONNECT', 'CALL PRO', 'MOBILE', '5G', 'MVNO']):
        return 'AF-KN-MOB'
    return 'AF-KN-Ope'

def detect_annexe_meta(df_csv):
    libs = " ".join(df_csv['Libellé produit'].dropna().astype(str).head(40).tolist()).upper()
    if '[MES]' in libs or 'OUVERTURE DE LIGNE' in libs or 'FAS ' in libs or 'PORTABILIT' in libs:
        return 'AF-KN-FAS', 'Facture FAS (Frais de mise en service)'
    if 'VMWARE' in libs or 'PROXMOX' in libs or 'VHDD' in libs or 'VRAM' in libs or 'VCPU' in libs or 'VDOM' in libs:
        return 'AF-KN-Heberg', 'Facture Hébergement'
    if 'CALL CONNECT' in libs or 'CALL PRO' in libs or 'MVNO' in libs:
        return 'AF-KN-MOB', 'Facture Mobile'
    return 'AF-KN-Ope', 'Facture Opérateur'

def format_ecart_signe(val):
    if abs(val) < 0.005:
        return "0,00 €"
    return f"{val:+,.2f} €".replace(",", " ").replace(".", ",")

def build_single_tab_excel(df_export):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rapprochement_Abonnements_HT"
    ws.views.sheetView[0].showGridLines = True

    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10, color="1E293B")
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")

    fill_hdr = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_ok = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_warn = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_err = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    thin = Side(style='thin', color='CBD5E1')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, col_name in enumerate(df_export.columns, start=1):
        c = ws.cell(row=1, column=ci, value=col_name)
        c.font = font_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
    ws.row_dimensions[1].height = 26

    for ri, row in enumerate(df_export.itertuples(index=False), start=2):
        ws.row_dimensions[ri].height = 20
        diag_val = str(row[-1])
        for ci, val in enumerate(row, start=1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = font_body
            c.border = bdr
            if ri % 2 == 1:
                c.fill = fill_zebra

            col_name = df_export.columns[ci - 1]
            if col_name == 'Diagnostic':
                c.font = font_bold
                if diag_val == 'Conforme':
                    c.fill = fill_ok
                elif 'Manquant' in diag_val:
                    c.fill = fill_err
                else:
                    c.fill = fill_warn

            if isinstance(val, (int, float)) and not isinstance(val, bool):
                if 'HT' in col_name or 'P.U.' in col_name or 'Écart' in col_name:
                    c.number_format = '#,##0.00 €'
                else:
                    c.number_format = '#,##0.##'
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    ws.auto_filter.ref = f"A1:{get_column_letter(len(df_export.columns))}{len(df_export) + 1}"
    ws.freeze_panes = "A2"

    for ci, col_name in enumerate(df_export.columns, start=1):
        max_len = max([len(str(col_name))] + [len(str(x)) for x in df_export.iloc[:, ci - 1].head(150)])
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 4, 14), 46)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ==============================================================================
# BARRE LATÉRALE — IMPORTATION DES FICHIERS UNIQUEMENT
# ==============================================================================
with st.sidebar:
    st.markdown("### 📁 Importation des Fichiers")
    file_sim = st.file_uploader("1. Simulation Achat Artis (.xlsx)", type=['xlsx'])
    file_ssc = st.file_uploader("2. Listing des SSC Artis (.xlsx) — Requis TVA & Dépôts", type=['xlsx'])
    files_pdf = st.file_uploader("3. Factures PDF Fournisseur (Optionnel)", type=['pdf'], accept_multiple_files=True)
    files_csv = st.file_uploader("4. Annexes Fournisseur (.csv)", type=['csv'], accept_multiple_files=True)

# ==============================================================================
# EN-TÊTE PRINCIPAL — IDENTITÉ POTENTIALE
# ==============================================================================
st.markdown("# 💼 Potentiale — Outil Métier RAF (Rapprochement Abonnements Fournisseurs)")
st.markdown(
    "**Fournisseur : Koesio Networks** | Urbanisation SI : Audit Fiscal TVA [S 20% vs AE 0%], "
    "Lecture Dynamique des Dépôts & Numéros RFC, Contrôle ADV Consolidé"
)

with st.expander("📌 Correspondance Dépôts ERP = Factures Fournisseur (Règle d'Urbanisation Potentiale)", expanded=True):
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.info("**`AF-KN-Ope`**\n\n= **Facture Abonnements Opérateur**\n\n*(Liens, SIP, SDA, Centrex)*")
    col_m2.info("**`AF-KN-MOB`**\n\n= **Facture Forfaits Mobiles**\n\n*(Lignes mobiles & options 5G)*")
    col_m3.info("**`AF-KN-Heberg`**\n\n= **Facture Abonnements Hébergement**\n\n*(VM, Cloud, VDOM, Sauvegarde)*")
    col_m4.warning("**`Flux Ponctuels (FAS)`**\n\n= **Facture Frais de Mise en Service**\n\n*(Pointage BL hors récurrent)*")

st.markdown("---")

# ==============================================================================
# TRAITEMENT MÉTIER AUTOMATISÉ
# ==============================================================================
SEUIL_ALERTE_HT = 500.0

if file_sim and file_ssc and files_csv:
    with st.spinner("Exécution de l'audit fiscal, temporel et financier en cours..."):
        df_sim = read_xlsx_robust(file_sim)
        df_ssc = read_xlsx_robust(file_ssc)

        # --- 1. AUDIT FISCAL TVA SUR LE PARC SSC ACTIF ---
        ssc_active = df_ssc[~df_ssc['Statut RFC'].isin(['Fermée', 'Résiliée.Dans les délais', 'Résiliée.En attente'])].copy()
        tva_col = 'Coût ABONNEMENT Gestion TVA' if 'Coût ABONNEMENT Gestion TVA' in ssc_active.columns else None
        taux_col = 'Coût ABONNEMENT Taux TVA' if 'Coût ABONNEMENT Taux TVA' in ssc_active.columns else None

        if tva_col:
            ssc_tva = ssc_active[ssc_active['Type RFC'] == 'Abonnement fournisseur'].copy()
            ssc_tva['Gestion_TVA'] = ssc_tva[tva_col].fillna('Non renseigné').astype(str)
            ssc_tva['Taux_TVA'] = ssc_tva[taux_col].fillna('Non renseigné').astype(str) if taux_col else '20%'
            ssc_tva['Assiette_HT'] = (
                pd.to_numeric(ssc_tva['Coût ABONNEMENT facturé'], errors='coerce').fillna(0.0) *
                pd.to_numeric(ssc_tva['Abonnement Nombre'], errors='coerce').fillna(1.0)
            )
            anomalies_tva = ssc_tva[
                ssc_tva['Gestion_TVA'].str.contains('AE|AUTO|EXO|0', case=False, na=False) |
                ssc_tva['Taux_TVA'].str.contains('0.00|0 %', case=False, na=False)
            ].copy()
        else:
            ssc_tva = pd.DataFrame()
            anomalies_tva = pd.DataFrame()

        # --- 2. LECTURE DYNAMIQUE DES DÉPÔTS & NUMÉROS RFC DEPUIS SSC ---
        ssc_dedup = ssc_active.drop_duplicates(subset=['Code SSC', 'Libellé RFC'])
        depot_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Abonnement Code dépôt'].to_dict()
        rfc_code_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Code RFC'].to_dict() if 'Code RFC' in ssc_dedup.columns else {}
        charge_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Nom chargé factu SSC'].to_dict() if 'Nom chargé factu SSC' in ssc_dedup.columns else {}

        ssc_active['k_norm'] = (
            ssc_active['Code SSC'].astype(str).str.strip() + "||" +
            ssc_active['Libellé RFC'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        )
        map_norm = ssc_active.drop_duplicates('k_norm').set_index('k_norm')['Abonnement Code dépôt'].to_dict()
        rfc_norm_map = ssc_active.drop_duplicates('k_norm').set_index('k_norm')['Code RFC'].to_dict() if 'Code RFC' in ssc_active.columns else {}

        sim_valid = df_sim[pd.to_numeric(df_sim['Montant HT'], errors='coerce').notna()].copy()
        sim_valid['Montant HT'] = pd.to_numeric(sim_valid['Montant HT'], errors='coerce')
        sim_valid['Quantite'] = pd.to_numeric(sim_valid['Quantite'], errors='coerce').fillna(1.0)
        sim_valid['Prix Unitaire'] = pd.to_numeric(sim_valid['Prix Unitaire'], errors='coerce').fillna(0.0)

        sim_valid['Depot_Brut'] = sim_valid.apply(
            lambda r: depot_map.get((str(r['Code SSC']).strip(), str(r['Libellé RFC']).strip()), None),
            axis=1
        )
        sim_valid['Numero_RFC'] = sim_valid.apply(
            lambda r: rfc_code_map.get((str(r['Code SSC']).strip(), str(r['Libellé RFC']).strip()), None),
            axis=1
        )
        sim_valid['Charge_Factu'] = sim_valid.apply(
            lambda r: charge_map.get((str(r['Code SSC']).strip(), str(r['Libellé RFC']).strip()), ''),
            axis=1
        ).fillna('')

        miss = sim_valid['Depot_Brut'].isna()
        sim_valid.loc[miss, 'Depot_Brut'] = sim_valid.loc[miss].apply(
            lambda r: map_norm.get(str(r['Code SSC']).strip() + "||" + " ".join(str(r['Libellé RFC']).split()), 'AF-KN-Ope'),
            axis=1
        )
        miss_rfc = sim_valid['Numero_RFC'].isna()
        sim_valid.loc[miss_rfc, 'Numero_RFC'] = sim_valid.loc[miss_rfc].apply(
            lambda r: rfc_norm_map.get(str(r['Code SSC']).strip() + "||" + " ".join(str(r['Libellé RFC']).split()), ''),
            axis=1
        )
        sim_valid['Numero_RFC'] = sim_valid['Numero_RFC'].fillna('')
        sim_valid['Depot_Cible'] = sim_valid['Depot_Brut'].apply(normalize_depot_simple)

        # Détection automatique du mois d'étude
        sim_valid['Deb_Per'] = pd.to_datetime(sim_valid['Début période facturation'], errors='coerce')
        valid_dates = sim_valid['Deb_Per'].dropna()
        if not valid_dates.empty:
            mode_date = valid_dates.dt.to_period('M').mode()[0]
            target_start = mode_date.to_timestamp()
            mois_detecte = f"{target_start.month:02d}/{target_start.year}"
        else:
            target_start = pd.Timestamp.today().replace(day=1)
            mois_detecte = f"{target_start.month:02d}/{target_start.year}"

        def check_temporalite(dt):
            if pd.isna(dt):
                return 'Date manquante'
            if dt.year == target_start.year and dt.month == target_start.month:
                return 'Cycle Conforme'
            if dt < target_start:
                return 'Régularisation Échue'
            return 'Facturation Anticipée'

        sim_valid['Statut_Temporel'] = sim_valid['Deb_Per'].apply(check_temporalite)
        nb_conformes = (sim_valid['Statut_Temporel'] == 'Cycle Conforme').sum()
        nb_echues = (sim_valid['Statut_Temporel'] == 'Régularisation Échue').sum()
        nb_anticipees = (sim_valid['Statut_Temporel'] == 'Facturation Anticipée').sum()
        nb_retards = df_sim['Anomalie'].notna().sum() if 'Anomalie' in df_sim.columns else 0

        # --- 3. CHARGEMENT DES ANNEXES CSV FOURNISSEUR ---
        f_list = []
        for f_csv in files_csv:
            df_c = read_csv_robust(f_csv)
            dep_det, fact_label = detect_annexe_meta(df_c)
            df_c['Depot_Fourn'] = dep_det
            df_c['Facture_Label'] = fact_label
            f_list.append(df_c)

        f_all = pd.concat(f_list, ignore_index=True)
        f_all['Total H.T.'] = pd.to_numeric(f_all['Total H.T.'], errors='coerce').fillna(0.0)
        f_all['P.U. HT'] = pd.to_numeric(f_all['P.U. HT'], errors='coerce').fillna(0.0)
        f_all['Qté'] = pd.to_numeric(f_all['Qté'], errors='coerce').fillna(1.0)

        f_fas = f_all[f_all['Depot_Fourn'] == 'AF-KN-FAS'].copy()
        if not f_fas.empty:
            f_fas['Depot_Ventilation'] = f_fas['Libellé produit'].apply(classify_line_depot)

        f_rec = f_all[f_all['Depot_Fourn'] != 'AF-KN-FAS'].copy()
        is_conso_mask = f_rec['Réf. produit'].astype(str).str.strip().str.upper() == 'CONSO'
        f_conso = f_rec[is_conso_mask].copy()
        f_abo = f_rec[~is_conso_mask].copy()

        # --- 4. MOTEUR DE RÉCONCILIATION CONSOLIDÉ (1 TABLEAU UNIQUE + DIAGNOSTIC FINAL) ---
        sim_rec = sim_valid[sim_valid['Depot_Cible'] != 'AF-KN-FAS'].copy()
        sim_rec['Tech_ID'] = sim_rec.apply(get_erp_tech_id, axis=1)
        sim_rec['Art_Num'] = sim_rec['Code Article'].astype(str).str.extract(r'(\d+)')[0].fillna('')
        sim_rec['Prod_Lib'] = sim_rec['Libellé RFC'].apply(lambda s: str(s).split('//')[0].strip())
        sim_rec['Client_Norm'] = sim_rec['Client'].apply(norm_txt)
        sim_rec['Prod_Norm'] = sim_rec['Prod_Lib'].apply(norm_txt)

        f_abo['Tech_ID'] = f_abo['Matricule fournisseur'].fillna('').astype(str).str.replace('.0', '', regex=False).str.strip()
        f_abo['Art_Num'] = f_abo['Réf. produit'].astype(str).str.strip()
        f_abo['Client_Norm'] = f_abo['Client final'].apply(norm_txt)
        f_abo['Prod_Norm'] = f_abo['Libellé produit'].apply(norm_txt)

        sim_rec['Pass_RAF'] = ''
        sim_rec['Diagnostic'] = 'Manquant annexe fournisseur'
        sim_rec['Code_Client_Annexe'] = ''
        sim_rec['Client_Annexe'] = ''
        sim_rec['Libelle_Annexe'] = ''
        sim_rec['Ref_Produit_Annexe'] = ''
        sim_rec['Qt_Annexe'] = 0.0
        sim_rec['PU_Annexe'] = 0.0
        sim_rec['Montant_Annexe_HT'] = 0.0

        f_abo['Pass_RAF'] = ''

        for dep in ['AF-KN-Ope', 'AF-KN-MOB', 'AF-KN-Heberg']:
            s_idx = sim_rec[sim_rec['Depot_Cible'] == dep].index
            f_idx = f_abo[f_abo['Depot_Fourn'] == dep].index

            # Passe 1 : Identifiant technique + montant identique
            f_by_id = {}
            for fi in f_idx:
                tid = f_abo.at[fi, 'Tech_ID']
                if tid:
                    f_by_id.setdefault(tid, []).append(fi)

            for si in s_idx:
                tid = sim_rec.at[si, 'Tech_ID']
                if not tid or tid not in f_by_id:
                    continue
                s_ht = float(sim_rec.at[si, 'Montant HT'])
                for fi in f_by_id[tid]:
                    if f_abo.at[fi, 'Pass_RAF'] == '' and abs(s_ht - float(f_abo.at[fi, 'Total H.T.'])) <= 0.05:
                        sim_rec.at[si, 'Pass_RAF'] = 'P1'
                        sim_rec.at[si, 'Diagnostic'] = 'Conforme'
                        sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[fi, 'Ref. client'] or '')
                        sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[fi, 'Client final'] or '')
                        sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[fi, 'Libellé produit'] or '')
                        sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[fi, 'Réf. produit'] or '')
                        sim_rec.at[si, 'Qt_Annexe'] = float(f_abo.at[fi, 'Qté'])
                        sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[fi, 'P.U. HT'])
                        sim_rec.at[si, 'Montant_Annexe_HT'] = float(f_abo.at[fi, 'Total H.T.'])
                        f_abo.at[fi, 'Pass_RAF'] = 'P1'
                        break

            # Passe 1b : Identifiant technique avec écart de prix ou quantité
            for si in s_idx:
                if sim_rec.at[si, 'Pass_RAF'] != '':
                    continue
                tid = sim_rec.at[si, 'Tech_ID']
                if not tid or tid not in f_by_id:
                    continue
                for fi in f_by_id[tid]:
                    if f_abo.at[fi, 'Pass_RAF'] == '':
                        sim_rec.at[si, 'Pass_RAF'] = 'P1b'
                        sim_rec.at[si, 'Diagnostic'] = 'Écart Prix / Quantité'
                        sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[fi, 'Ref. client'] or '')
                        sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[fi, 'Client final'] or '')
                        sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[fi, 'Libellé produit'] or '')
                        sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[fi, 'Réf. produit'] or '')
                        sim_rec.at[si, 'Qt_Annexe'] = float(f_abo.at[fi, 'Qté'])
                        sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[fi, 'P.U. HT'])
                        sim_rec.at[si, 'Montant_Annexe_HT'] = float(f_abo.at[fi, 'Total H.T.'])
                        f_abo.at[fi, 'Pass_RAF'] = 'P1b'
                        break

            # Passe 2 : Client + Code Article Numérique
            s_rem = [si for si in s_idx if sim_rec.at[si, 'Pass_RAF'] == '' and sim_rec.at[si, 'Art_Num'] not in ('', '2912')]
            f_rem_ca = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    f_rem_ca.setdefault((f_abo.at[fi, 'Client_Norm'], f_abo.at[fi, 'Art_Num']), []).append(fi)
            s_by_ca = {}
            for si in s_rem:
                s_by_ca.setdefault((sim_rec.at[si, 'Client_Norm'], sim_rec.at[si, 'Art_Num']), []).append(si)

            for ca, sis in s_by_ca.items():
                if ca not in f_rem_ca:
                    continue
                fis = [fi for fi in f_rem_ca[ca] if f_abo.at[fi, 'Pass_RAF'] == '']
                if not fis:
                    continue
                sum_s = sum(float(sim_rec.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in fis)
                sum_qf = sum(float(f_abo.at[fi, 'Qté']) for fi in fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                diag_lbl = 'Conforme' if is_exact else 'Écart Prix / Quantité'
                first_fi = fis[0]
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    ratio = (s_ht / sum_s) if sum_s > 0 else (1.0 / len(sis))
                    sim_rec.at[si, 'Pass_RAF'] = 'P2'
                    sim_rec.at[si, 'Diagnostic'] = diag_lbl
                    sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[first_fi, 'Ref. client'] or '')
                    sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[first_fi, 'Client final'] or '')
                    sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[first_fi, 'Libellé produit'] or '')
                    sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[first_fi, 'Réf. produit'] or '')
                    sim_rec.at[si, 'Qt_Annexe'] = round(sum_qf * ratio, 2)
                    sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[first_fi, 'P.U. HT'])
                    sim_rec.at[si, 'Montant_Annexe_HT'] = round(sum_f * ratio, 2)
                for fi in fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P2'

            # Passe 3 : Client + Libellé Produit
            s_rem3 = [si for si in s_idx if sim_rec.at[si, 'Pass_RAF'] == '']
            f_rem_cp = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    f_rem_cp.setdefault((f_abo.at[fi, 'Client_Norm'], f_abo.at[fi, 'Prod_Norm']), []).append(fi)
            s_by_cp = {}
            for si in s_rem3:
                s_by_cp.setdefault((sim_rec.at[si, 'Client_Norm'], sim_rec.at[si, 'Prod_Norm']), []).append(si)

            for cp, sis in s_by_cp.items():
                if cp not in f_rem_cp:
                    continue
                fis = [fi for fi in f_rem_cp[cp] if f_abo.at[fi, 'Pass_RAF'] == '']
                if not fis:
                    continue
                sum_s = sum(float(sim_rec.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in fis)
                sum_qf = sum(float(f_abo.at[fi, 'Qté']) for fi in fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                diag_lbl = 'Conforme' if is_exact else 'Écart Prix / Quantité'
                first_fi = fis[0]
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    ratio = (s_ht / sum_s) if sum_s > 0 else (1.0 / len(sis))
                    sim_rec.at[si, 'Pass_RAF'] = 'P3'
                    sim_rec.at[si, 'Diagnostic'] = diag_lbl
                    sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[first_fi, 'Ref. client'] or '')
                    sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[first_fi, 'Client final'] or '')
                    sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[first_fi, 'Libellé produit'] or '')
                    sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[first_fi, 'Réf. produit'] or '')
                    sim_rec.at[si, 'Qt_Annexe'] = round(sum_qf * ratio, 2)
                    sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[first_fi, 'P.U. HT'])
                    sim_rec.at[si, 'Montant_Annexe_HT'] = round(sum_f * ratio, 2)
                for fi in fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P3'

        # Construction du tableau unique consolidé (ERP + Manquants ERP)
        df_part_erp = pd.DataFrame({
            'Code Dépôt (Origine Facture)': sim_rec['Depot_Cible'],
            'Facture Associée': sim_rec['Depot_Cible'].map({
                'AF-KN-Ope': 'Facture Opérateur',
                'AF-KN-MOB': 'Facture Mobile',
                'AF-KN-Heberg': 'Facture Hébergement'
            }),
            'Chargé Factu': sim_rec['Charge_Factu'],
            'Code Client ERP': sim_rec['Code client'].fillna(''),
            'Code Client Annexe': sim_rec['Code_Client_Annexe'],
            'Client ERP': sim_rec['Client'].fillna(''),
            'Client Annexe': sim_rec['Client_Annexe'],
            'Numéro SSC': sim_rec['Code SSC'].fillna(''),
            'Numéro RFC': sim_rec['Numero_RFC'],
            'Code Article ERP': sim_rec['Code Article'].fillna(''),
            'Réf. Produit Annexe': sim_rec['Ref_Produit_Annexe'],
            'Libellé RFC': sim_rec['Libellé RFC'].fillna(''),
            'Libellé Annexe': sim_rec['Libelle_Annexe'],
            'Identifiant Technique': sim_rec['Tech_ID'],
            'Qté ERP': sim_rec['Quantite'],
            'Qté Annexe': sim_rec['Qt_Annexe'],
            'P.U. ERP HT': sim_rec['Prix Unitaire'],
            'P.U. Annexe HT': sim_rec['PU_Annexe'],
            'Montant ERP HT': sim_rec['Montant HT'],
            'Montant Annexe HT': sim_rec['Montant_Annexe_HT'],
            'Écart HT (+ / -)': sim_rec['Montant HT'] - sim_rec['Montant_Annexe_HT'],
            'Diagnostic': sim_rec['Diagnostic']
        })

        f_un

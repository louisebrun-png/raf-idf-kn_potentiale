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

def norm_client_code(val):
    s = str(val or '').strip().upper()
    if s in ('', 'NAN', 'NONE'):
        return ''
    if '/' in s:
        s = s.split('/')[0].strip()
    s = re.sub(r'^TO_', '', s)
    return s

def extract_art_num(val):
    s = str(val or '').strip().upper()
    if s in ('', 'NAN', 'NONE'):
        return ''
    m = re.search(r'(\d+)', s)
    return m.group(1) if m else s

def to_mda_article(val):
    s = str(val or '').strip().upper()
    if s in ('', 'NAN', 'NONE'):
        return ''
    if re.fullmatch(r'KN\d+', s):
        return s
    if s.isdigit():
        return f"KN{s}"
    m = re.search(r'(\d+)', s)
    if m and not s.startswith('KN'):
        return f"KN{m.group(1)}"
    return s

def fmt_date_fr(val):
    if pd.isna(val) or val is None or str(val).strip() == '':
        return ''
    dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
    if pd.isna(dt):
        return str(val).strip()
    return dt.strftime('%d/%m/%Y')

def get_period_ym(val):
    if pd.isna(val) or val is None or str(val).strip() == '':
        return ''
    dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
    if pd.isna(dt):
        return ''
    return f"{dt.year}-{dt.month:02d}"

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

def is_internal_depot(dep):
    """Exclut UNIQUEMENT les lignes dont le code dépôt contient 026159. Tout autre dépôt est conservé."""
    return '026159' in str(dep or '').strip()

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

def classify_mismatch_diag(qt_annexe, deb_erp_str, deb_annexe_str):
    if 0 < qt_annexe < 0.999:
        return 'Problème de prorata et/ou de prix'
    if deb_erp_str and not deb_erp_str.startswith('01/'):
        return 'Problème de prorata et/ou de prix'
    if deb_annexe_str and not deb_annexe_str.startswith('01/'):
        return 'Problème de prorata et/ou de prix'
    return 'Écart Prix / Quantité'

def diag_priority_rank(diag_str):
    """Place toutes les anomalies/actions en haut (rang 0) et relègue 'Conforme' (rang 1) et 'Conforme 0€' (rang 2) tout en bas."""
    d = str(diag_str or '').strip()
    if d == 'Conforme 0€':
        return 2
    if d == 'Conforme':
        return 1
    return 0

def compute_article_correction(row):
    art_erp = str(row.get('Code Article ERP', '')).strip().upper()
    ref_annexe = str(row.get('Réf. Produit Annexe', '')).strip().upper()
    diag = str(row.get('Diagnostic', '')).strip()
    code_cli_erp = str(row.get('Code Client ERP', '')).strip()

    if diag == 'Conforme 0€':
        return "Aucune action (0,00 €)"
    if diag == 'Dossier lié via Numéro SSC':
        mda = to_mda_article(ref_annexe)
        if mda and not art_erp:
            return f"Corriger Ref. client partenaire -> {code_cli_erp} + Créer {mda}"
        return f"Pointer ce client (Ref. client partenaire à corriger -> {code_cli_erp})"
    if 'Interne / Non attribué (Code 10)' in diag:
        return "Basculer sur dépôt 026159_1 OU rattacher au client final"
    if 'Doublon — Supprimer AFIDF' in diag:
        return "Fermer ligne AFIDF (Conserver contrat chiffré)"
    if 'Doublon — Conserver contrat chiffré' in diag:
        return "Conserver (Porte la vente) — Fermer AFIDF"
    if 'Erreur date dernier achat' in diag:
        return "Corriger date dernier achat RFC"
    if diag == 'Manquant erp' or 'Régul. fournisseur' in diag:
        mda = to_mda_article(ref_annexe)
        return f"Créer / Saisir avec {mda}" if mda else "À créer / saisir"
    if not ref_annexe:
        return "Pointer ce client (Absent annexe)"

    mda_cible = to_mda_article(ref_annexe)
    num_erp = extract_art_num(art_erp)
    num_annexe = extract_art_num(ref_annexe)

    if art_erp == 'KN2912':
        return f"Remplacer KN2912 -> {mda_cible}"
    if num_erp != num_annexe and mda_cible:
        return f"Remplacer {art_erp} -> {mda_cible}"
    return "OK"

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
    fill_prorata = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    fill_regul = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
    fill_err = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_neutral = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin = Side(style='thin', color='CBD5E1')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, col_name in enumerate(df_export.columns, start=1):
        c = ws.cell(row=1, column=ci, value=col_name)
        c.font = font_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
    ws.row_dimensions[1].height = 26

    diag_col_idx = list(df_export.columns).index('Diagnostic')
    art_corr_col_idx = list(df_export.columns).index('Action / Code Article à Corriger (MDA)')

    for ri, row in enumerate(df_export.itertuples(index=False), start=2):
        ws.row_dimensions[ri].height = 20
        diag_val = str(row[diag_col_idx])
        art_corr_val = str(row[art_corr_col_idx])

        for ci, val in enumerate(row, start=1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = font_body
            c.border = bdr
            if ri % 2 == 1:
                c.fill = fill_zebra

            col_name = df_export.columns[ci - 1]
            if col_name == 'Diagnostic':
                c.font = font_bold
                if diag_val in ('Conforme', 'Conforme 0€'):
                    c.fill = fill_ok
                elif 'Interne / Non attribué' in diag_val:
                    c.fill = fill_neutral
                elif 'Manquant' in diag_val or 'Erreur date' in diag_val or 'Doublon' in diag_val:
                    c.fill = fill_err
                elif 'Régul.' in diag_val:
                    c.fill = fill_regul
                elif 'prorata' in diag_val or diag_val == 'Dossier lié via Numéro SSC':
                    c.fill = fill_prorata
                else:
                    c.fill = fill_warn
            elif col_name == 'Action / Code Article à Corriger (MDA)':
                c.font = font_bold
                if art_corr_val in ('OK', 'Aucune action (0,00 €)'):
                    c.fill = fill_ok
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
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 4, 14), 54)

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
    "**Fournisseur : Koesio Networks** | Urbanisation SI : Audit Fiscal TVA [STD Standard 20%], "
    "Exclusion Stricte Dépôt 026159, Priorité Contrats Chiffrés & Tri Anomalies en Tête"
)

with st.expander("📌 Correspondance Dépôts ERP = Factures Fournisseur (Règle d'Urbanisation Potentiale)", expanded=True):
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.info("**`AF-KN-Ope`**\n\n= **Facture Abonnements Opérateur**\n\n*(Tout dépôt autre que 026159 conservé)*")
    col_m2.info("**`AF-KN-MOB`**\n\n= **Facture Forfaits Mobiles**\n\n*(Tout dépôt autre que 026159 conservé)*")
    col_m3.info("**`AF-KN-Heberg`**\n\n= **Facture Abonnements Hébergement**\n\n*(Tout dépôt autre que 026159 conservé)*")
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
            ssc_tva['Gestion_TVA'] = ssc_tva[tva_col].fillna('Non renseigné').astype(str).str.strip()
            ssc_tva['Taux_TVA'] = ssc_tva[taux_col].fillna('Non renseigné').astype(str).str.strip() if taux_col else 'Normal - 20.00 %'
            ssc_tva['Assiette_HT'] = (
                pd.to_numeric(ssc_tva['Coût ABONNEMENT facturé'], errors='coerce').fillna(0.0) *
                pd.to_numeric(ssc_tva['Abonnement Nombre'], errors='coerce').fillna(1.0)
            )
            is_std_ok = ssc_tva['Gestion_TVA'].str.upper().str.contains('STD')
            is_20_ok = ssc_tva['Taux_TVA'].str.contains('20')
            anomalies_tva = ssc_tva[~(is_std_ok & is_20_ok)].copy()
        else:
            ssc_tva = pd.DataFrame()
            anomalies_tva = pd.DataFrame()

        # --- 2. LECTURE DYNAMIQUE DES DÉPÔTS, NUMÉROS RFC & TABLE SSC -> CLIENT ---
        ssc_dedup = ssc_active.drop_duplicates(subset=['Code SSC', 'Libellé RFC'])
        depot_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Abonnement Code dépôt'].to_dict()
        rfc_code_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Code RFC'].to_dict() if 'Code RFC' in ssc_dedup.columns else {}
        charge_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Nom chargé factu SSC'].to_dict() if 'Nom chargé factu SSC' in ssc_dedup.columns else {}

        ssc_to_client_map = {}
        for _, r_ssc in ssc_active.drop_duplicates(subset=['Code SSC']).iterrows():
            c_ssc = str(r_ssc.get('Code SSC', '')).strip()
            if c_ssc:
                ssc_to_client_map[c_ssc] = {
                    'code_client': str(r_ssc.get('Code client', '') or '').strip(),
                    'client_erp': str(r_ssc.get('Raison sociale client', '') or '').strip(),
                    'charge_factu': str(r_ssc.get('Nom chargé factu SSC', '') or '').strip()
                }

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

        # Exclusion UNIQUEMENT des lignes dont le code dépôt contient 026159
        mask_interne_erp = sim_valid['Depot_Brut'].apply(is_internal_depot)
        sim_valid = sim_valid[~mask_interne_erp].copy()
        sim_valid['Depot_Cible'] = sim_valid['Depot_Brut'].apply(normalize_depot_simple)

        # Détection automatique du mois d'étude M et du mois précédent M-1
        sim_valid['Deb_Per'] = pd.to_datetime(sim_valid['Début période facturation'], errors='coerce')
        sim_valid['Fin_Per'] = pd.to_datetime(sim_valid['Fin période facturation'], errors='coerce') if 'Fin période facturation' in sim_valid.columns else pd.NaT
        sim_valid['Deb_ERP_Fmt'] = sim_valid['Deb_Per'].apply(fmt_date_fr)
        sim_valid['Fin_ERP_Fmt'] = sim_valid['Fin_Per'].apply(fmt_date_fr)
        sim_valid['YM_ERP'] = sim_valid['Deb_Per'].apply(get_period_ym)

        valid_dates = sim_valid['Deb_Per'].dropna()
        if not valid_dates.empty:
            mode_date = valid_dates.dt.to_period('M').mode()[0]
            target_start = mode_date.to_timestamp()
            prev_month_start = (mode_date - 1).to_timestamp()
            last_day_prev = (target_start - pd.Timedelta(days=1)).strftime('%d/%m/%Y')
            mois_detecte = f"{target_start.month:02d}/{target_start.year}"
            mois_m1_label = f"{prev_month_start.month:02d}/{prev_month_start.year}"
        else:
            target_start = pd.Timestamp.today().replace(day=1)
            prev_month_start = (target_start - pd.DateOffset(months=1))
            last_day_prev = (target_start - pd.Timedelta(days=1)).strftime('%d/%m/%Y')
            mois_detecte = f"{target_start.month:02d}/{target_start.year}"
            mois_m1_label = f"{prev_month_start.month:02d}/{prev_month_start.year}"

        def check_temporalite(dt):
            if pd.isna(dt):
                return 'Date manquante'
            if dt.year == target_start.year and dt.month == target_start.month:
                return 'Mois Courant (M)'
            if dt.year == prev_month_start.year and dt.month == prev_month_start.month:
                return 'Mois Précédent (M-1)'
            if dt < prev_month_start:
                return 'Erreur Date Dernier Achat (< M-1)'
            return 'Facturation Anticipée'

        sim_valid['Statut_Temporel'] = sim_valid['Deb_Per'].apply(check_temporalite)
        nb_m = (sim_valid['Statut_Temporel'] == 'Mois Courant (M)').sum()
        nb_m1 = (sim_valid['Statut_Temporel'] == 'Mois Précédent (M-1)').sum()
        nb_err_date_erp = (sim_valid['Statut_Temporel'] == 'Erreur Date Dernier Achat (< M-1)').sum()
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
        f_all = f_all[~f_all['Ref. client'].apply(is_internal_depot)].copy()

        f_all['Total H.T.'] = pd.to_numeric(f_all['Total H.T.'], errors='coerce').fillna(0.0)
        f_all['P.U. HT'] = pd.to_numeric(f_all['P.U. HT'], errors='coerce').fillna(0.0)
        f_all['Qté'] = pd.to_numeric(f_all['Qté'], errors='coerce').fillna(1.0)
        f_all['Deb_F_Dt'] = pd.to_datetime(f_all['Date début'], dayfirst=True, errors='coerce')
        f_all['Deb_Annexe_Fmt'] = f_all['Date début'].apply(fmt_date_fr)
        f_all['Fin_Annexe_Fmt'] = f_all['Date fin'].apply(fmt_date_fr)
        f_all['YM_Annexe'] = f_all['Date début'].apply(get_period_ym)

        f_fas = f_all[f_all['Depot_Fourn'] == 'AF-KN-FAS'].copy()
        if not f_fas.empty:
            f_fas['Depot_Ventilation'] = f_fas['Libellé produit'].apply(classify_line_depot)

        f_rec = f_all[f_all['Depot_Fourn'] != 'AF-KN-FAS'].copy()
        is_conso_mask = f_rec['Réf. produit'].astype(str).str.strip().str.upper() == 'CONSO'
        f_conso = f_rec[is_conso_mask].copy()
        f_abo_all = f_rec[~is_conso_mask].copy()

        is_f_regul_ant = f_abo_all['Deb_F_Dt'].notna() & (f_abo_all['Deb_F_Dt'] < prev_month_start)
        f_regul_ant = f_abo_all[is_f_regul_ant].copy()
        f_abo = f_abo_all[~is_f_regul_ant].copy()

        # --- 4. ISOLEMENT CÔTÉ ERP DES ERREURS DE DATE DERNIER ACHAT (< M-1) ---
        sim_all_rec = sim_valid[sim_valid['Depot_Cible'] != 'AF-KN-FAS'].copy()
        sim_err_date = sim_all_rec[sim_all_rec['Statut_Temporel'] == 'Erreur Date Dernier Achat (< M-1)'].copy()
        sim_rec = sim_all_rec[sim_all_rec['Statut_Temporel'] != 'Erreur Date Dernier Achat (< M-1)'].copy()

        # --- 5. MOTEUR DE RÉCONCILIATION EN 6 PASSES ---
        sim_rec['Tech_ID'] = sim_rec.apply(get_erp_tech_id, axis=1)
        sim_rec['Sub_Client_Code'] = sim_rec['Tech_ID'].apply(norm_client_code)
        sim_rec['Code_Client_ERP_Norm'] = sim_rec['Code client'].apply(norm_client_code)
        sim_rec['Art_Num'] = sim_rec['Code Article'].apply(extract_art_num)
        sim_rec['Prod_Lib'] = sim_rec['Libellé RFC'].apply(lambda s: str(s).split('//')[0].strip())
        sim_rec['Client_Norm'] = sim_rec['Client'].apply(norm_txt)
        sim_rec['Prod_Norm'] = sim_rec['Prod_Lib'].apply(norm_txt)
        sim_rec['Is_Digit_SSC'] = sim_rec['Code SSC'].astype(str).str.strip().str[0].str.isdigit()

        f_abo['Tech_ID'] = f_abo['Matricule fournisseur'].fillna('').astype(str).str.replace('.0', '', regex=False).str.strip()
        f_abo['Code_Client_Fourn_Norm'] = f_abo['Ref. client'].apply(norm_client_code)
        f_abo['Art_Num'] = f_abo['Réf. produit'].apply(extract_art_num)
        f_abo['Client_Norm'] = f_abo['Client final'].apply(norm_txt)
        f_abo['Prod_Norm'] = f_abo['Libellé produit'].apply(norm_txt)

        erp_id_c = sim_rec[sim_rec['Tech_ID'] != ''].drop_duplicates('Tech_ID').set_index('Tech_ID')['Client_Norm'].to_dict()
        fourn_id_c = f_abo[f_abo['Tech_ID'] != ''].drop_duplicates('Tech_ID').set_index('Tech_ID')['Client_Norm'].to_dict()
        client_alias_map = {}
        for tid, ce in erp_id_c.items():
            if tid in fourn_id_c:
                client_alias_map.setdefault(ce, set()).add(fourn_id_c[tid])
        for ce in sim_rec['Client_Norm'].unique():
            client_alias_map.setdefault(ce, set()).add(ce)

        sim_rec['Pass_RAF'] = ''
        sim_rec['Diagnostic'] = 'Manquant annexe fournisseur'
        sim_rec['Code_Client_Annexe'] = ''
        sim_rec['Client_Annexe'] = ''
        sim_rec['Libelle_Annexe'] = ''
        sim_rec['Ref_Produit_Annexe'] = ''
        sim_rec['Deb_Annexe'] = ''
        sim_rec['Fin_Annexe'] = ''
        sim_rec['Qt_Annexe'] = 0.0
        sim_rec['PU_Annexe'] = 0.0
        sim_rec['Montant_Annexe_HT'] = 0.0

        f_abo['Pass_RAF'] = ''

        def bind_single_match(si, fi, pass_code, diag_str):
            sim_rec.at[si, 'Pass_RAF'] = pass_code
            sim_rec.at[si, 'Diagnostic'] = diag_str
            sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[fi, 'Ref. client'] or '')
            sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[fi, 'Client final'] or '')
            sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[fi, 'Libellé produit'] or '')
            sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[fi, 'Réf. produit'] or '')
            sim_rec.at[si, 'Deb_Annexe'] = f_abo.at[fi, 'Deb_Annexe_Fmt']
            sim_rec.at[si, 'Fin_Annexe'] = f_abo.at[fi, 'Fin_Annexe_Fmt']
            sim_rec.at[si, 'Qt_Annexe'] = float(f_abo.at[fi, 'Qté'])
            sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[fi, 'P.U. HT'])
            sim_rec.at[si, 'Montant_Annexe_HT'] = float(f_abo.at[fi, 'Total H.T.'])
            f_abo.at[fi, 'Pass_RAF'] = pass_code

        for dep in ['AF-KN-Ope', 'AF-KN-MOB', 'AF-KN-Heberg']:
            s_sub = sim_rec[sim_rec['Depot_Cible'] == dep].sort_values(by=['Is_Digit_SSC'], ascending=False)
            s_idx = s_sub.index
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
                        bind_single_match(si, fi, 'P1', 'Conforme')
                        break

            # Passe 1b : Identifiant technique avec écart (Prorata ou Prix/Qté)
            for si in s_idx:
                if sim_rec.at[si, 'Pass_RAF'] != '':
                    continue
                tid = sim_rec.at[si, 'Tech_ID']
                if not tid or tid not in f_by_id:
                    continue
                for fi in f_by_id[tid]:
                    if f_abo.at[fi, 'Pass_RAF'] == '':
                        qt_a = float(f_abo.at[fi, 'Qté'])
                        deb_a = f_abo.at[fi, 'Deb_Annexe_Fmt']
                        deb_e = sim_rec.at[si, 'Deb_ERP_Fmt']
                        bind_single_match(si, fi, 'P1b', classify_mismatch_diag(qt_a, deb_e, deb_a))
                        break

            # Passe 2 : Entités Régionales logées dans l'Identifiant Technique (FNTP)
            f_by_ccode = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    cc = f_abo.at[fi, 'Code_Client_Fourn_Norm']
                    if cc:
                        f_by_ccode.setdefault(cc, []).append(fi)

            for si in s_idx:
                if sim_rec.at[si, 'Pass_RAF'] != '':
                    continue
                sub_cc = sim_rec.at[si, 'Sub_Client_Code']
                if not sub_cc or sub_cc not in f_by_ccode:
                    continue
                s_ht = float(sim_rec.at[si, 'Montant HT'])
                art_s = sim_rec.at[si, 'Art_Num']
                cand = [fi for fi in f_by_ccode[sub_cc] if f_abo.at[fi, 'Pass_RAF'] == '' and float(f_abo.at[fi, 'Total H.T.']) >= 10.0]
                if not cand:
                    continue
                exact_c = [fi for fi in cand if abs(s_ht - float(f_abo.at[fi, 'Total H.T.'])) <= 0.05]
                same_art_c = [fi for fi in cand if f_abo.at[fi, 'Art_Num'] == art_s]
                if exact_c:
                    bind_single_match(si, exact_c[0], 'P2_Reg', 'Conforme')
                elif same_art_c:
                    fi = same_art_c[0]
                    qt_a = float(f_abo.at[fi, 'Qté'])
                    bind_single_match(si, fi, 'P2_Reg', classify_mismatch_diag(qt_a, sim_rec.at[si, 'Deb_ERP_Fmt'], f_abo.at[fi, 'Deb_Annexe_Fmt']))
                else:
                    fi = cand[0]
                    qt_a = float(f_abo.at[fi, 'Qté'])
                    bind_single_match(si, fi, 'P2_Reg', classify_mismatch_diag(qt_a, sim_rec.at[si, 'Deb_ERP_Fmt'], f_abo.at[fi, 'Deb_Annexe_Fmt']))

            # Passe 3a : Code Client Strict Identique + Code Article Numérique (KNxxxx == xxxx) + Période
            s_rem3a = [si for si in s_idx if sim_rec.at[si, 'Pass_RAF'] == '' and sim_rec.at[si, 'Code_Client_ERP_Norm'] != '' and sim_rec.at[si, 'Art_Num'] not in ('', '2912')]
            f_rem_code_am = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '' and f_abo.at[fi, 'Code_Client_Fourn_Norm'] != '':
                    key = (f_abo.at[fi, 'Code_Client_Fourn_Norm'], f_abo.at[fi, 'Art_Num'], f_abo.at[fi, 'YM_Annexe'])
                    f_rem_code_am.setdefault(key, []).append(fi)
            s_by_code_am = {}
            for si in s_rem3a:
                key = (sim_rec.at[si, 'Code_Client_ERP_Norm'], sim_rec.at[si, 'Art_Num'], sim_rec.at[si, 'YM_ERP'])
                s_by_code_am.setdefault(key, []).append(si)

            for key, sis in s_by_code_am.items():
                if key not in f_rem_code_am:
                    continue
                matched_fis = [fi for fi in f_rem_code_am[key] if f_abo.at[fi, 'Pass_RAF'] == '']
                if not matched_fis:
                    continue
                sum_s = sum(float(sim_rec.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in matched_fis)
                sum_qf = sum(float(f_abo.at[fi, 'Qté']) for fi in matched_fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                first_fi = matched_fis[0]
                deb_a = f_abo.at[first_fi, 'Deb_Annexe_Fmt']
                fin_a = f_abo.at[first_fi, 'Fin_Annexe_Fmt']
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    ratio = (s_ht / sum_s) if sum_s > 0 else (1.0 / len(sis))
                    qt_alloc = round(sum_qf * ratio, 2)
                    deb_e = sim_rec.at[si, 'Deb_ERP_Fmt']
                    diag_lbl = 'Conforme' if is_exact else classify_mismatch_diag(qt_alloc, deb_e, deb_a)
                    sim_rec.at[si, 'Pass_RAF'] = 'P3a'
                    sim_rec.at[si, 'Diagnostic'] = diag_lbl
                    sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[first_fi, 'Ref. client'] or '')
                    sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[first_fi, 'Client final'] or '')
                    sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[first_fi, 'Libellé produit'] or '')
                    sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[first_fi, 'Réf. produit'] or '')
                    sim_rec.at[si, 'Deb_Annexe'] = deb_a
                    sim_rec.at[si, 'Fin_Annexe'] = fin_a
                    sim_rec.at[si, 'Qt_Annexe'] = qt_alloc
                    sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[first_fi, 'P.U. HT'])
                    sim_rec.at[si, 'Montant_Annexe_HT'] = round(sum_f * ratio, 2)
                for fi in matched_fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P3a'

            # Passe 3b : Alias Client + Code Article Numérique (KNxxxx == xxxx) + Période
            s_rem = [si for si in s_idx if sim_rec.at[si, 'Pass_RAF'] == '' and sim_rec.at[si, 'Art_Num'] not in ('', '2912')]
            f_rem_cam = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    key = (f_abo.at[fi, 'Client_Norm'], f_abo.at[fi, 'Art_Num'], f_abo.at[fi, 'YM_Annexe'])
                    f_rem_cam.setdefault(key, []).append(fi)
            s_by_cam = {}
            for si in s_rem:
                key = (sim_rec.at[si, 'Client_Norm'], sim_rec.at[si, 'Art_Num'], sim_rec.at[si, 'YM_ERP'])
                s_by_cam.setdefault(key, []).append(si)

            for (ce, art_n, ym_e), sis in s_by_cam.items():
                matched_fis = []
                for cf in client_alias_map.get(ce, {ce}):
                    if (cf, art_n, ym_e) in f_rem_cam:
                        matched_fis.extend([fi for fi in f_rem_cam[(cf, art_n, ym_e)] if f_abo.at[fi, 'Pass_RAF'] == ''])
                if not matched_fis:
                    continue
                sum_s = sum(float(sim_rec.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in matched_fis)
                sum_qf = sum(float(f_abo.at[fi, 'Qté']) for fi in matched_fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                first_fi = matched_fis[0]
                deb_a = f_abo.at[first_fi, 'Deb_Annexe_Fmt']
                fin_a = f_abo.at[first_fi, 'Fin_Annexe_Fmt']
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    ratio = (s_ht / sum_s) if sum_s > 0 else (1.0 / len(sis))
                    qt_alloc = round(sum_qf * ratio, 2)
                    deb_e = sim_rec.at[si, 'Deb_ERP_Fmt']
                    diag_lbl = 'Conforme' if is_exact else classify_mismatch_diag(qt_alloc, deb_e, deb_a)
                    sim_rec.at[si, 'Pass_RAF'] = 'P3b'
                    sim_rec.at[si, 'Diagnostic'] = diag_lbl
                    sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[first_fi, 'Ref. client'] or '')
                    sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[first_fi, 'Client final'] or '')
                    sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[first_fi, 'Libellé produit'] or '')
                    sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[first_fi, 'Réf. produit'] or '')
                    sim_rec.at[si, 'Deb_Annexe'] = deb_a
                    sim_rec.at[si, 'Fin_Annexe'] = fin_a
                    sim_rec.at[si, 'Qt_Annexe'] = qt_alloc
                    sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[first_fi, 'P.U. HT'])
                    sim_rec.at[si, 'Montant_Annexe_HT'] = round(sum_f * ratio, 2)
                for fi in matched_fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P3b'

            # Passe 4 : Alias Client + Libellé Produit + Période (absorbe la reprise KN2912)
            s_rem4 = [si for si in s_idx if sim_rec.at[si, 'Pass_RAF'] == '']
            f_rem_cpm = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '':
                    key = (f_abo.at[fi, 'Client_Norm'], f_abo.at[fi, 'Prod_Norm'], f_abo.at[fi, 'YM_Annexe'])
                    f_rem_cpm.setdefault(key, []).append(fi)
            s_by_cpm = {}
            for si in s_rem4:
                key = (sim_rec.at[si, 'Client_Norm'], sim_rec.at[si, 'Prod_Norm'], sim_rec.at[si, 'YM_ERP'])
                s_by_cpm.setdefault(key, []).append(si)

            for (ce, prod_n, ym_e), sis in s_by_cpm.items():
                matched_fis = []
                for cf in client_alias_map.get(ce, {ce}):
                    if (cf, prod_n, ym_e) in f_rem_cpm:
                        matched_fis.extend([fi for fi in f_rem_cpm[(cf, prod_n, ym_e)] if f_abo.at[fi, 'Pass_RAF'] == ''])
                if not matched_fis:
                    continue
                sum_s = sum(float(sim_rec.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in matched_fis)
                sum_qf = sum(float(f_abo.at[fi, 'Qté']) for fi in matched_fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                first_fi = matched_fis[0]
                deb_a = f_abo.at[first_fi, 'Deb_Annexe_Fmt']
                fin_a = f_abo.at[first_fi, 'Fin_Annexe_Fmt']
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    ratio = (s_ht / sum_s) if sum_s > 0 else (1.0 / len(sis))
                    qt_alloc = round(sum_qf * ratio, 2)
                    deb_e = sim_rec.at[si, 'Deb_ERP_Fmt']
                    diag_lbl = 'Conforme' if is_exact else classify_mismatch_diag(qt_alloc, deb_e, deb_a)
                    sim_rec.at[si, 'Pass_RAF'] = 'P4'
                    sim_rec.at[si, 'Diagnostic'] = diag_lbl
                    sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[first_fi, 'Ref. client'] or '')
                    sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[first_fi, 'Client final'] or '')
                    sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[first_fi, 'Libellé produit'] or '')
                    sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[first_fi, 'Réf. produit'] or '')
                    sim_rec.at[si, 'Deb_Annexe'] = deb_a
                    sim_rec.at[si, 'Fin_Annexe'] = fin_a
                    sim_rec.at[si, 'Qt_Annexe'] = qt_alloc
                    sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[first_fi, 'P.U. HT'])
                    sim_rec.at[si, 'Montant_Annexe_HT'] = round(sum_f * ratio, 2)
                for fi in matched_fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P4'

            # Passe 5 : Code Client Exact Identique + Libellé Produit + Période
            s_rem5 = [si for si in s_idx if sim_rec.at[si, 'Pass_RAF'] == '' and sim_rec.at[si, 'Code_Client_ERP_Norm'] != '']
            f_rem_code_pm = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '' and f_abo.at[fi, 'Code_Client_Fourn_Norm'] != '':
                    key = (f_abo.at[fi, 'Code_Client_Fourn_Norm'], f_abo.at[fi, 'Prod_Norm'], f_abo.at[fi, 'YM_Annexe'])
                    f_rem_code_pm.setdefault(key, []).append(fi)
            s_by_code_pm = {}
            for si in s_rem5:
                key = (sim_rec.at[si, 'Code_Client_ERP_Norm'], sim_rec.at[si, 'Prod_Norm'], sim_rec.at[si, 'YM_ERP'])
                s_by_code_pm.setdefault(key, []).append(si)

            for key, sis in s_by_code_pm.items():
                if key not in f_rem_code_pm:
                    continue
                matched_fis = [fi for fi in f_rem_code_pm[key] if f_abo.at[fi, 'Pass_RAF'] == '']
                if not matched_fis:
                    continue
                sum_s = sum(float(sim_rec.at[si, 'Montant HT']) for si in sis)
                sum_f = sum(float(f_abo.at[fi, 'Total H.T.']) for fi in matched_fis)
                sum_qf = sum(float(f_abo.at[fi, 'Qté']) for fi in matched_fis)
                is_exact = abs(sum_s - sum_f) <= 0.10
                first_fi = matched_fis[0]
                deb_a = f_abo.at[first_fi, 'Deb_Annexe_Fmt']
                fin_a = f_abo.at[first_fi, 'Fin_Annexe_Fmt']
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    ratio = (s_ht / sum_s) if sum_s > 0 else (1.0 / len(sis))
                    qt_alloc = round(sum_qf * ratio, 2)
                    deb_e = sim_rec.at[si, 'Deb_ERP_Fmt']
                    diag_lbl = 'Conforme' if is_exact else classify_mismatch_diag(qt_alloc, deb_e, deb_a)
                    sim_rec.at[si, 'Pass_RAF'] = 'P5'
                    sim_rec.at[si, 'Diagnostic'] = diag_lbl
                    sim_rec.at[si, 'Code_Client_Annexe'] = str(f_abo.at[first_fi, 'Ref. client'] or '')
                    sim_rec.at[si, 'Client_Annexe'] = str(f_abo.at[first_fi, 'Client final'] or '')
                    sim_rec.at[si, 'Libelle_Annexe'] = str(f_abo.at[first_fi, 'Libellé produit'] or '')
                    sim_rec.at[si, 'Ref_Produit_Annexe'] = str(f_abo.at[first_fi, 'Réf. produit'] or '')
                    sim_rec.at[si, 'Deb_Annexe'] = deb_a
                    sim_rec.at[si, 'Fin_Annexe'] = fin_a
                    sim_rec.at[si, 'Qt_Annexe'] = qt_alloc
                    sim_rec.at[si, 'PU_Annexe'] = float(f_abo.at[first_fi, 'P.U. HT'])
                    sim_rec.at[si, 'Montant_Annexe_HT'] = round(sum_f * ratio, 2)
                for fi in matched_fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P5'

            # Passe 6 : Même Client (ou Code Client) + Même Montant HT Exact sur le même Dépôt
            f_by_client_ht = {}
            for fi in f_idx:
                if f_abo.at[fi, 'Pass_RAF'] == '' and float(f_abo.at[fi, 'Total H.T.']) > 0:
                    ht_r = round(float(f_abo.at[fi, 'Total H.T.']), 2)
                    f_by_client_ht.setdefault((f_abo.at[fi, 'Client_Norm'], ht_r), []).append(fi)
                    cc_f = f_abo.at[fi, 'Code_Client_Fourn_Norm']
                    if cc_f:
                        f_by_client_ht.setdefault((f"CODE_{cc_f}", ht_r), []).append(fi)

            for si in s_idx:
                if sim_rec.at[si, 'Pass_RAF'] != '':
                    continue
                ht_s = round(float(sim_rec.at[si, 'Montant HT']), 2)
                if ht_s <= 0:
                    continue
                ce = sim_rec.at[si, 'Client_Norm']
                cc_s = sim_rec.at[si, 'Code_Client_ERP_Norm']
                cand_keys = [(cf, ht_s) for cf in client_alias_map.get(ce, {ce})]
                if cc_s:
                    cand_keys.append((f"CODE_{cc_s}", ht_s))
                matched_fi = None
                for ck in cand_keys:
                    if ck in f_by_client_ht:
                        avail = [fi for fi in f_by_client_ht[ck] if f_abo.at[fi, 'Pass_RAF'] == '']
                        if avail:
                            matched_fi = avail[0]
                            break
                if matched_fi is not None:
                    bind_single_match(si, matched_fi, 'P6_Tarif', 'Conforme')

        # --- 6. RÈGLE DE PRIORITÉ DOUBLONS, REGROUPEMENT PAR SSC (CATÉGORIE UNIQUE) & CODE 10 ---
        conforme_rows = sim_rec[sim_rec['Diagnostic'] == 'Conforme']
        conforme_signatures = set(zip(
            conforme_rows['Depot_Cible'],
            conforme_rows['Client_Norm'],
            conforme_rows['Montant HT'].round(2)
        ))
        conforme_grouped_sums = conforme_rows.groupby(['Depot_Cible', 'Client_Norm', 'Art_Num'])['Montant HT'].sum().round(2).to_dict()

        f_unmatched_ssc_refs = set(
            f_abo[f_abo['Pass_RAF'] == '']['Ref. client'].dropna().astype(str).str.strip()
        ).intersection(set(ssc_to_client_map.keys()))

        for si in sim_rec[sim_rec['Pass_RAF'] == ''].index:
            code_cli_str = str(sim_rec.at[si, 'Code client']).strip()
            num_ssc_str = str(sim_rec.at[si, 'Code SSC']).strip()
            m_ht = round(float(sim_rec.at[si, 'Montant HT']), 2)
            sig = (sim_rec.at[si, 'Depot_Cible'], sim_rec.at[si, 'Client_Norm'], m_ht)
            grp_key = (sim_rec.at[si, 'Depot_Cible'], sim_rec.at[si, 'Client_Norm'], sim_rec.at[si, 'Art_Num'])

            is_dup = (sig in conforme_signatures) or (grp_key in conforme_grouped_sums and abs(conforme_grouped_sums[grp_key] - m_ht) <= 0.10)

            if is_dup:
                if not sim_rec.at[si, 'Is_Digit_SSC']:
                    sim_rec.at[si, 'Diagnostic'] = 'Doublon — Supprimer AFIDF (Contrat chiffré actif)'
                else:
                    sim_rec.at[si, 'Diagnostic'] = 'Doublon — Conserver contrat chiffré (Fermer AFIDF)'
            elif num_ssc_str in f_unmatched_ssc_refs:
                # Catégorie strictement unique pour l'écran de synthèse et le fichier
                sim_rec.at[si, 'Diagnostic'] = 'Dossier lié via Numéro SSC'
            elif code_cli_str == '10':
                sim_rec.at[si, 'Diagnostic'] = 'Interne / Non attribué (Code 10)'

        # Construction du tableau unique consolidé
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
            'Début Période ERP': sim_rec['Deb_ERP_Fmt'],
            'Fin Période ERP': sim_rec['Fin_ERP_Fmt'],
            'Début Période Annexe': sim_rec['Deb_Annexe'],
            'Fin Période Annexe': sim_rec['Fin_Annexe'],
            'Qté ERP': sim_rec['Quantite'],
            'Qté Annexe': sim_rec['Qt_Annexe'],
            'P.U. ERP HT': sim_rec['Prix Unitaire'],
            'P.U. Annexe HT': sim_rec['PU_Annexe'],
            'Montant ERP HT': sim_rec['Montant HT'],
            'Montant Annexe HT': sim_rec['Montant_Annexe_HT'],
            'Écart HT (+ / -)': sim_rec['Montant HT'] - sim_rec['Montant_Annexe_HT'],
            'Diagnostic': sim_rec['Diagnostic']
        })

        # Lignes ERP avec erreur de date dernier achat (< M-1)
        sim_err_date['Tech_ID'] = sim_err_date.apply(get_erp_tech_id, axis=1)
        df_part_err_date = pd.DataFrame({
            'Code Dépôt (Origine Facture)': sim_err_date['Depot_Cible'],
            'Facture Associée': sim_err_date['Depot_Cible'].map({
                'AF-KN-Ope': 'Facture Opérateur',
                'AF-KN-MOB': 'Facture Mobile',
                'AF-KN-Heberg': 'Facture Hébergement'
            }),
            'Chargé Factu': sim_err_date['Charge_Factu'],
            'Code Client ERP': sim_err_date['Code client'].fillna(''),
            'Code Client Annexe': '',
            'Client ERP': sim_err_date['Client'].fillna(''),
            'Client Annexe': '',
            'Numéro SSC': sim_err_date['Code SSC'].fillna(''),
            'Numéro RFC': sim_err_date['Numero_RFC'],
            'Code Article ERP': sim_err_date['Code Article'].fillna(''),
            'Réf. Produit Annexe': '',
            'Libellé RFC': sim_err_date['Libellé RFC'].fillna(''),
            'Libellé Annexe': '',
            'Identifiant Technique': sim_err_date['Tech_ID'],
            'Début Période ERP': sim_err_date['Deb_ERP_Fmt'],
            'Fin Période ERP': sim_err_date['Fin_ERP_Fmt'],
            'Début Période Annexe': '',
            'Fin Période Annexe': '',
            'Qté ERP': sim_err_date['Quantite'],
            'Qté Annexe': 0.0,
            'P.U. ERP HT': sim_err_date['Prix Unitaire'],
            'P.U. Annexe HT': 0.0,
            'Montant ERP HT': sim_err_date['Montant HT'],
            'Montant Annexe HT': 0.0,
            'Écart HT (+ / -)': sim_err_date['Montant HT'],
            'Diagnostic': f'Erreur date dernier achat ERP (Mettre au {last_day_prev})'
        })

        # Régularisations fournisseur antérieures à M-1 (< M-1)
        f_regul_ant['Tech_ID'] = f_regul_ant['Matricule fournisseur'].fillna('').astype(str).str.replace('.0', '', regex=False).str.strip()
        df_part_f_regul = pd.DataFrame({
            'Code Dépôt (Origine Facture)': f_regul_ant['Depot_Fourn'],
            'Facture Associée': f_regul_ant['Facture_Label'],
            'Chargé Factu': '',
            'Code Client ERP': '',
            'Code Client Annexe': f_regul_ant['Ref. client'].fillna(''),
            'Client ERP': '',
            'Client Annexe': f_regul_ant['Client final'].fillna(''),
            'Numéro SSC': '',
            'Numéro RFC': '',
            'Code Article ERP': '',
            'Réf. Produit Annexe': f_regul_ant['Réf. produit'].fillna(''),
            'Libellé RFC': '',
            'Libellé Annexe': f_regul_ant['Libellé produit'].fillna(''),
            'Identifiant Technique': f_regul_ant['Tech_ID'],
            'Début Période ERP': '',
            'Fin Période ERP': '',
            'Début Période Annexe': f_regul_ant['Deb_Annexe_Fmt'],
            'Fin Période Annexe': f_regul_ant['Fin_Annexe_Fmt'],
            'Qté ERP': 0.0,
            'Qté Annexe': f_regul_ant['Qté'],
            'P.U. ERP HT': 0.0,
            'P.U. Annexe HT': f_regul_ant['P.U. HT'],
            'Montant ERP HT': 0.0,
            'Montant Annexe HT': f_regul_ant['Total H.T.'],
            'Écart HT (+ / -)': -f_regul_ant['Total H.T.'],
            'Diagnostic': f'Régul. fournisseur < {mois_m1_label} (Variable sur facture)'
        })

        # Lignes fournisseur non rapprochées
        f_unmatched = f_abo[f_abo['Pass_RAF'] == ''].copy()

        def enrich_fourn_unmatched(r):
            ref_cli = str(r.get('Ref. client') or '').strip()
            tot_ht = float(r.get('Total H.T.', 0.0))
            if abs(tot_ht) < 0.005:
                return pd.Series(['', '', '', '', 'Conforme 0€'])
            if ref_cli in ssc_to_client_map:
                info = ssc_to_client_map[ref_cli]
                c_erp = info['code_client']
                return pd.Series([
                    info['charge_factu'],
                    c_erp,
                    info['client_erp'],
                    ref_cli,
                    'Dossier lié via Numéro SSC'
                ])
            return pd.Series(['', '', '', '', 'Manquant erp'])

        f_unmatched[['Charge_Link', 'Code_Cli_Link', 'Client_ERP_Link', 'Num_SSC_Link', 'Diag_Fourn']] = f_unmatched.apply(enrich_fourn_unmatched, axis=1)

        df_part_fourn = pd.DataFrame({
            'Code Dépôt (Origine Facture)': f_unmatched['Depot_Fourn'],
            'Facture Associée': f_unmatched['Facture_Label'],
            'Chargé Factu': f_unmatched['Charge_Link'],
            'Code Client ERP': f_unmatched['Code_Cli_Link'],
            'Code Client Annexe': f_unmatched['Ref. client'].fillna(''),
            'Client ERP': f_unmatched['Client_ERP_Link'],
            'Client Annexe': f_unmatched['Client final'].fillna(''),
            'Numéro SSC': f_unmatched['Num_SSC_Link'],
            'Numéro RFC': '',
            'Code Article ERP': '',
            'Réf. Produit Annexe': f_unmatched['Réf. produit'].fillna(''),
            'Libellé RFC': '',
            'Libellé Annexe': f_unmatched['Libellé produit'].fillna(''),
            'Identifiant Technique': f_unmatched['Tech_ID'],
            'Début Période ERP': '',
            'Fin Période ERP': '',
            'Début Période Annexe': f_unmatched['Deb_Annexe_Fmt'],
            'Fin Période Annexe': f_unmatched['Fin_Annexe_Fmt'],
            'Qté ERP': 0.0,
            'Qté Annexe': f_unmatched['Qté'],
            'P.U. ERP HT': 0.0,
            'P.U. Annexe HT': f_unmatched['P.U. HT'],
            'Montant ERP HT': 0.0,
            'Montant Annexe HT': f_unmatched['Total H.T.'],
            'Écart HT (+ / -)': -f_unmatched['Total H.T.'],
            'Diagnostic': f_unmatched['Diag_Fourn']
        })

        df_conso_unique = pd.concat([df_part_erp, df_part_err_date, df_part_f_regul, df_part_fourn], ignore_index=True)
        df_conso_unique['Action / Code Article à Corriger (MDA)'] = df_conso_unique.apply(compute_article_correction, axis=1)

        # Tri hiérarchique : Anomalies / Actions EN HAUT (rang 0), Conforme (rang 1) et Conforme 0€ (rang 2) TOUT EN BAS
        df_conso_unique['_rank'] = df_conso_unique['Diagnostic'].apply(diag_priority_rank)
        df_conso_unique = df_conso_unique.sort_values(
            by=['_rank', 'Code Dépôt (Origine Facture)', 'Client ERP', 'Numéro SSC', 'Diagnostic', 'Client Annexe', 'Numéro RFC', 'Début Période ERP'],
            ascending=[True, True, True, True, True, True, True, True]
        ).drop(columns=['_rank'])

        # ======================================================================
        # NAVIGATION EN 4 ONGLETS MÉTIERS (STANDARD POTENTIALE)
        # ======================================================================
        tab1, tab2, tab3, tab4 = st.tabs([
            "🚨 1. Audit TVA (STD Standard 20%)",
            "📊 2. Abonnements Récurrents HT",
            "📦 3. Frais Ponctuels par Dépôt (FAS)",
            "📞 4. Consommations & Régul. par Dépôt"
        ])

        # ----------------------------------------------------------------------
        # ONGLET 1 : AUDIT TVA (STD STANDARD 20%)
        # ----------------------------------------------------------------------
        with tab1:
            st.markdown("## 🚨 Audit Fiscal TVA — Contrôle Régime STD Standard (20 %)")
            st.info("**Règle d'Audit Potentiale :** Toutes les RFC actives d'abonnement sur Artis doivent être paramétrées en **Gestion TVA = `STD Standard`** avec un **Taux TVA = `Normal - 20.00 %`**. Toute autre valeur détectée sur le listing SSC remonte en anomalie.")

            if not ssc_tva.empty:
                nb_anom_tva = len(anomalies_tva)
                assiette_totale = ssc_tva['Assiette_HT'].sum()
                assiette_anom = anomalies_tva['Assiette_HT'].sum()

                k1, k2, k3 = st.columns(3)
                k1.metric("Règles Actives Contrôlées", f"{len(ssc_tva):,}".replace(',', ' '))
                k2.metric("Assiette Mensuelle HT Auditée", f"{assiette_totale:,.2f} €")
                k3.metric(
                    "Anomalies TVA Détectées",
                    f"{nb_anom_tva} ligne(s)",
                    delta=f"{assiette_anom:,.2f} € HT" if nb_anom_tva > 0 else "0,00 € (100 % Conforme)",
                    delta_color="inverse" if nb_anom_tva > 0 else "normal"
                )

                if nb_anom_tva == 0:
                    st.success("✅ **Audit Fiscal 100 % Conforme :** Toutes les RFC actives du listing SSC sont bien paramétrées en `STD Standard` à `20.00 %`.")
                else:
                    st.error(f"⚠️ **Attention :** {nb_anom_tva} RFC présente(nt) un code ou taux TVA différent de `STD Standard / 20.00 %` :")
                    st.dataframe(anomalies_tva[['Code SSC', 'Raison sociale client', 'Code RFC', 'Libellé RFC', 'Gestion_TVA', 'Taux_TVA', 'Assiette_HT']], use_container_width=True, hide_index=True)

        # ----------------------------------------------------------------------
        # ONGLET 2 : ABONNEMENTS RÉCURRENTS HT
        # ----------------------------------------------------------------------
        with tab2:
            st.markdown(f"## ⏳ Audit de Temporalité ERP (Clôture {mois_detecte} — Prorata {mois_m1_label} Admis)")
            t1, t2, t3, t4 = st.columns(4)
            t1.metric(f"🟢 Mois Courant ({mois_detecte})", f"{nb_m} lignes")
            t2.metric(f"🔵 Prorata M-1 ({mois_m1_label})", f"{nb_m1} lignes")
            t3.metric(f"🚨 Erreur Date Dernier Achat (< {mois_m1_label})", f"{nb_err_date_erp} lignes")
            t4.metric("⏳ Anomalies Modules / Dates", f"{nb_retards} lignes")

            st.markdown("---")
            st.markdown(f"## 🧮 Synthèse Macro-Financière HT (Périmètre {mois_detecte} & {mois_m1_label} Hors Dépôt 026159)")

            tot_sim_perimetre_ht = sim_rec['Montant HT'].sum()
            tot_err_date_ht = sim_err_date['Montant HT'].sum()
            tot_fourn_facture_ht = f_abo['Total H.T.'].sum()
            ecart_net_global = tot_sim_perimetre_ht - tot_fourn_facture_ht

            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"Simulation Artis ({mois_detecte} + {mois_m1_label})", f"{tot_sim_perimetre_ht:,.2f} €")
            m2.metric("Facturé Fournisseur (Hors Conso)", f"{tot_fourn_facture_ht:,.2f} €")
            m3.metric(
                "Écart Net Récurrent (+ / -)",
                format_ecart_signe(ecart_net_global),
                delta="🚨 Attention (> 500 €)" if abs(ecart_net_global) > SEUIL_ALERTE_HT else "✅ Dans le seuil",
                delta_color="inverse" if abs(ecart_net_global) > SEUIL_ALERTE_HT else "normal"
            )
            m4.metric(
                f"Erreur Date ERP (< {mois_m1_label})",
                f"{tot_err_date_ht:,.2f} €",
                delta=f"{nb_err_date_erp} lignes à caler au {last_day_prev}",
                delta_color="inverse"
            )

            if abs(ecart_net_global) > SEUIL_ALERTE_HT:
                st.warning(
                    f"🚨 **Attention (Écart > 500 €) :** L'écart net sur le périmètre courant ({mois_detecte} & {mois_m1_label}) atteint **{format_ecart_signe(ecart_net_global)} HT** "
                    "(un signe **`+`** indique un surplus de lignes actives dans l'ERP — résiliations d'achats à passer ; "
                    "un signe **`-`** indique des prestations facturées absentes de l'ERP)."
                )

            st.markdown("### 📋 Comparatif Détaillé par Dépôt (= Facture Fournisseur Hors Consommations)")
            rows_dep = []
            for dep, fact_nom in [
                ('AF-KN-Ope', 'Facture Abonnements Opérateur'),
                ('AF-KN-MOB', 'Facture Forfaits Mobiles'),
                ('AF-KN-Heberg', 'Facture Abonnements Hébergement')
            ]:
                s_d = sim_rec[sim_rec['Depot_Cible'] == dep]
                s_err = sim_err_date[sim_err_date['Depot_Cible'] == dep]
                f_d = f_abo[f_abo['Depot_Fourn'] == dep]
                sim_m = s_d['Montant HT'].sum()
                err_m = s_err['Montant HT'].sum()
                fourn_m = f_d['Total H.T.'].sum()
                ecart_m = sim_m - fourn_m
                alerte_txt = "🚨 ATTENTION (> 500 €)" if abs(ecart_m) > SEUIL_ALERTE_HT else "✅ OK (< 500 €)"
                rows_dep.append({
                    'Code Dépôt': dep,
                    'Facture Associée': fact_nom,
                    f'Simulation ({mois_detecte}+{mois_m1_label}) HT': round(sim_m, 2),
                    'Facture Fournisseur HT': round(fourn_m, 2),
                    'Écart Net (+ / -)': format_ecart_signe(ecart_m),
                    f'Erreur Date ERP (< {mois_m1_label}) HT': round(err_m, 2),
                    'Contrôle Seuil 500 €': alerte_txt
                })
            df_comp_dep = pd.DataFrame(rows_dep)
            st.dataframe(df_comp_dep, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("### 📊 Bilan Volumétrique Assaini du Rapprochement (Anomalies en Haut, Conformes & 0€ en Bas)")
            diag_counts = df_conso_unique.groupby('Diagnostic', as_index=False).agg(
                Nombre_Lignes=('Diagnostic', 'count'),
                Montant_ERP_HT=('Montant ERP HT', 'sum'),
                Montant_Annexe_HT=('Montant Annexe HT', 'sum'),
                Ecart_HT=('Écart HT (+ / -)', 'sum')
            )
            diag_counts['_rank'] = diag_counts['Diagnostic'].apply(diag_priority_rank)
            diag_counts = diag_counts.sort_values(by=['_rank', 'Diagnostic'], ascending=[True, True]).drop(columns=['_rank'])
            st.dataframe(diag_counts, use_container_width=True, hide_index=True)

        # ----------------------------------------------------------------------
        # ONGLET 3 : FRAIS PONCTUELS PAR DÉPÔT (FAS)
        # ----------------------------------------------------------------------
        with tab3:
            st.markdown("## 📦 Ventilation Analytique des Frais Ponctuels par Dépôt (Facture FAS)")
            st.info("**Règle de Répartition Potentiale :** Les frais de mise en service (`FAS`) sont isolés de la facturation récurrente et ventilés par dépôt cible pour contrôle sur bons de livraison.")

            if not f_fas.empty:
                fas_ope = f_fas[f_fas['Depot_Ventilation'] == 'AF-KN-Ope']['Total H.T.'].sum()
                fas_mob = f_fas[f_fas['Depot_Ventilation'] == 'AF-KN-MOB']['Total H.T.'].sum()
                fas_heb = f_fas[f_fas['Depot_Ventilation'] == 'AF-KN-Heberg']['Total H.T.'].sum()

                fc1, fc2, fc3, fc4 = st.columns(4)
                fc1.metric("Total Facture FAS HT", f"{f_fas['Total H.T.'].sum():,.2f} €")
                fc2.metric("Ventilation AF-KN-Ope", f"{fas_ope:,.2f} € HT")
                fc3.metric("Ventilation AF-KN-MOB", f"{fas_mob:,.2f} € HT")
                fc4.metric("Ventilation AF-KN-Heberg", f"{fas_heb:,.2f} € HT")

                st.dataframe(
                    f_fas[['Depot_Ventilation', 'Ref. client', 'Client final', 'Site', 'Réf. produit', 'Libellé produit', 'Matricule fournisseur', 'Deb_Annexe_Fmt', 'Fin_Annexe_Fmt', 'Qté', 'P.U. HT', 'Total H.T.']].sort_values(by=['Depot_Ventilation', 'Total H.T.'], ascending=[True, False]),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("Aucune ligne de frais ponctuels (FAS) détectée dans les annexes chargées.")

        # ----------------------------------------------------------------------
        # ONGLET 4 : CONSOMMATIONS & RÉGULARISATIONS FOURNISSEUR PAR DÉPÔT
        # ----------------------------------------------------------------------
        with tab4:
            st.markdown("## 📞 Ventilation des Consommations & Régularisations Fournisseur (< M-1) par Dépôt")
            st.info("**Règle de Répartition Potentiale :** Les consommations variables (`CONSO`) ainsi que les éventuelles régularisations rétroactives fournisseur antérieures à M-1 (impossibles à anticiper) sont ventilées par dépôt et famille d'article pour saisie en variable sur la facture d'achat.")

            conso_ope = f_conso[f_conso['Depot_Fourn'] == 'AF-KN-Ope']['Total H.T.'].sum()
            conso_mob = f_conso[f_conso['Depot_Fourn'] == 'AF-KN-MOB']['Total H.T.'].sum()
            conso_heb = f_conso[f_conso['Depot_Fourn'] == 'AF-KN-Heberg']['Total H.T.'].sum()

            cc1, cc2, cc3 = st.columns(3)
            cc1.metric("AF-KN-Ope (Consommations Opérateur)", f"{conso_ope:,.2f} € HT", delta="À ajouter sur facture d'achat Ope")
            cc2.metric("AF-KN-MOB (Consommations Mobiles)", f"{conso_mob:,.2f} € HT", delta="À ajouter sur facture d'achat MOB")
            cc3.metric("AF-KN-Heberg (Consommations Cloud)", f"{conso_heb:,.2f} € HT", delta="À ajouter sur facture d'achat Heberg")

            st.markdown("### 📋 Synthèse des Variables de Consommation à Saisir dans Artis")
            df_conso_synth = f_conso.groupby(['Depot_Fourn', 'Facture_Label', 'Libellé produit'], as_index=False).agg(
                Nombre_Lignes=('Total H.T.', 'count'),
                Montant_HT_A_Ajouter=('Total H.T.', 'sum')
            )
            st.dataframe(df_conso_synth, use_container_width=True, hide_index=True)

            st.markdown(f"### 🟣 Régularisations Fournisseur Antérieures (< {mois_m1_label}) par Famille Article (Variables à Saisir)")
            if not f_regul_ant.empty:
                df_regul_synth = f_regul_ant.groupby(['Depot_Fourn', 'Réf. produit', 'Libellé produit'], as_index=False).agg(
                    Nombre_Lignes=('Total H.T.', 'count'),
                    Montant_HT_A_Saisir=('Total H.T.', 'sum')
                )
                st.dataframe(df_regul_synth, use_container_width=True, hide_index=True)
            else:
                st.success(f"✅ Aucune régularisation fournisseur chiffrée antérieure à {mois_m1_label} sur ce cycle.")

            st.markdown("### 🔎 Détail des Lignes de Consommation par Client")
            st.dataframe(
                f_conso[['Depot_Fourn', 'Ref. client', 'Client final', 'Site', 'Matricule fournisseur', 'Deb_Annexe_Fmt', 'Fin_Annexe_Fmt', 'Total H.T.']].sort_values(by=['Depot_Fourn', 'Total H.T.'], ascending=[True, False]),
                use_container_width=True,
                hide_index=True
            )

        # ======================================================================
        # EXPORT EXCEL UNIQUE : Rapprochement_Abonnements_HT_IDF_KN.xlsx
        # ======================================================================
        st.markdown("---")
        excel_bytes = build_single_tab_excel(df_conso_unique)
        st.download_button(
            label="📥 Télécharger Rapprochement_Abonnements_HT_IDF_KN.xlsx (Onglet Unique Consolidé)",
            data=excel_bytes,
            file_name="Rapprochement_Abonnements_HT_IDF_KN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👈 Importez la **Simulation Achat Artis**, le **Listing des SSC Artis** et vos **4 Annexes Fournisseur** dans le volet de gauche pour lancer l'analyse.")

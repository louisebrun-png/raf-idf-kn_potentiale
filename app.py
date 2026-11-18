import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import xml.etree.ElementTree as ET
import unicodedata
import re
import io
from datetime import datetime

# ==============================================================================
# CONFIGURATION DE LA PAGE — CABINET POTENTIALE
# ==============================================================================
st.set_page_config(
    page_title="Potentiale — Outil Métier RAF (Koesio Networks)",
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
        return 'FAS', 'Facture FAS (Frais de mise en service)'
    if 'VMWARE' in libs or 'PROXMOX' in libs or 'VHDD' in libs or 'VRAM' in libs or 'VCPU' in libs or 'VDOM' in libs:
        return 'AF-KN-Heberg', 'Facture Hébergement'
    if 'CALL CONNECT' in libs or 'CALL PRO' in libs or 'MVNO' in libs:
        return 'AF-KN-MOB', 'Facture Mobile'
    return 'AF-KN-Ope', 'Facture Opérateur'

# ==============================================================================
# BARRE LATÉRALE — IMPORTATION DES FICHIERS (HARMONISÉE POTENTIALE)
# ==============================================================================
with st.sidebar:
    st.markdown("### 📁 Importation des Fichiers")
    file_sim = st.file_uploader("1. Simulation Achat Artis (.xlsx)", type=['xlsx'])
    file_ssc = st.file_uploader("2. Listing des SSC Artis (.xlsx) — Requis TVA & Dépôts", type=['xlsx'])
    files_pdf = st.file_uploader("3. Factures PDF Koesio Networks (Optionnel)", type=['pdf'], accept_multiple_files=True)
    files_csv = st.file_uploader("4. Annexes Fournisseur (.csv)", type=['csv'], accept_multiple_files=True)

    st.markdown("---")
    mois_etude = st.text_input("📅 Mois d'étude (MM/AAAA)", value="10/2026")
    seuil_alerte = st.number_input("🚨 Seuil d'alerte écart HT (€)", value=500.0, step=50.0)

# ==============================================================================
# EN-TÊTE PRINCIPAL — IDENTITÉ POTENTIALE
# ==============================================================================
st.markdown("# 💼 Potentiale — Outil Métier RAF (Rapprochement Abonnements Fournisseurs)")
st.markdown(
    "**Fournisseur : Koesio Networks** | Urbanisation SI : Audit Fiscal TVA [S 20% vs AE 0%], "
    "Lecture Dynamique des Dépôts SSC & Réconciliation Multi-Passes"
)

with st.expander("📌 Correspondance Dépôts ERP = Factures Koesio Networks (Règle d'Urbanisation Potentiale)", expanded=True):
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.info("**`AF-KN-Ope`**\n\n= **Facture Abonnements Opérateur**\n\n*(Liens, SIP, SDA, Centrex)*")
    col_m2.info("**`AF-KN-MOB`**\n\n= **Facture Forfaits Mobiles**\n\n*(Lignes mobiles & options 5G)*")
    col_m3.info("**`AF-KN-Heberg`**\n\n= **Facture Abonnements Hébergement**\n\n*(VM, Cloud, VDOM, Sauvegarde)*")
    col_m4.warning("**`Flux Ponctuels (FAS)`**\n\n= **Facture Frais de Mise en Service**\n\n*(Pointage BL hors récurrent)*")

st.markdown("---")

# ==============================================================================
# TRAITEMENT MÉTIER SI LES FICHIERS SONT CHARGÉS
# ==============================================================================
if file_sim and file_ssc and files_csv:
    with st.spinner("Exécution de l'audit fiscal, temporel et financier en cours..."):
        df_sim = read_xlsx_robust(file_sim)
        df_ssc = read_xlsx_robust(file_ssc)

        # Parsing mois d'étude
        try:
            m_str, y_str = mois_etude.strip().split('/')
            target_start = pd.Timestamp(int(y_str), int(m_str), 1)
            target_end = target_start + pd.offsets.MonthEnd(1)
        except Exception:
            target_start = pd.Timestamp(2026, 10, 1)
            target_end = pd.Timestamp(2026, 10, 31)

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

        # --- 2. LECTURE DYNAMIQUE DES DÉPÔTS DEPUIS SSC ---
        ssc_dedup = ssc_active.drop_duplicates(subset=['Code SSC', 'Libellé RFC'])
        depot_map = ssc_dedup.set_index(['Code SSC', 'Libellé RFC'])['Abonnement Code dépôt'].to_dict()
        ssc_active['k_norm'] = (
            ssc_active['Code SSC'].astype(str).str.strip() + "||" +
            ssc_active['Libellé RFC'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        )
        map_norm = ssc_active.drop_duplicates('k_norm').set_index('k_norm')['Abonnement Code dépôt'].to_dict()

        sim_valid = df_sim[pd.to_numeric(df_sim['Montant HT'], errors='coerce').notna()].copy()
        sim_valid['Montant HT'] = pd.to_numeric(sim_valid['Montant HT'], errors='coerce')
        sim_valid['Quantite'] = pd.to_numeric(sim_valid['Quantite'], errors='coerce').fillna(1.0)
        sim_valid['Prix Unitaire'] = pd.to_numeric(sim_valid['Prix Unitaire'], errors='coerce').fillna(0.0)

        sim_valid['Depot_Brut'] = sim_valid.apply(
            lambda r: depot_map.get((str(r['Code SSC']).strip(), str(r['Libellé RFC']).strip()), None),
            axis=1
        )
        miss = sim_valid['Depot_Brut'].isna()
        sim_valid.loc[miss, 'Depot_Brut'] = sim_valid.loc[miss].apply(
            lambda r: map_norm.get(str(r['Code SSC']).strip() + "||" + " ".join(str(r['Libellé RFC']).split()), 'AF-KN-Ope'),
            axis=1
        )
        sim_valid['Depot_Cible'] = sim_valid['Depot_Brut'].apply(normalize_depot_simple)

        # Audit de temporalité sur le mois d'étude
        sim_valid['Deb_Per'] = pd.to_datetime(sim_valid['Début période facturation'], errors='coerce')
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
        f_conso = f_rec[f_rec['Réf. produit'].astype(str).str.contains('CONSO', case=False, na=False)].copy()
        f_abo = f_rec[~f_rec['Réf. produit'].astype(str).str.contains('CONSO', case=False, na=False)].copy()

        # --- 4. MOTEUR DE RÉCONCILIATION MULTI-PASSES (RÉCURRENT) ---
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
        sim_rec['Fourn_Reconcilie_HT'] = 0.0
        f_abo['Pass_RAF'] = ''

        for dep in ['AF-KN-Ope', 'AF-KN-MOB', 'AF-KN-Heberg']:
            s_idx = sim_rec[sim_rec['Depot_Cible'] == dep].index
            f_idx = f_abo[f_abo['Depot_Fourn'] == dep].index

            # Passe 1 : Identifiant technique + montant
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
                        sim_rec.at[si, 'Fourn_Reconcilie_HT'] = float(f_abo.at[fi, 'Total H.T.'])
                        f_abo.at[fi, 'Pass_RAF'] = 'P1'
                        break

            # Passe 2 : Client + Code Article
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
                for si in sis:
                    s_ht = float(sim_rec.at[si, 'Montant HT'])
                    sim_rec.at[si, 'Pass_RAF'] = 'P2'
                    sim_rec.at[si, 'Fourn_Reconcilie_HT'] = (sum_f * (s_ht / sum_s)) if sum_s > 0 else 0.0
                for fi in fis:
                    f_abo.at[fi, 'Pass_RAF'] = 'P2'

            # Passe 3 : Client + Libellé
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
                if abs(sum_s - sum_f) <= 0.10:
                    for si in sis:
                        sim_rec.at[si, 'Pass_RAF'] = 'P3'
                        sim_rec.at[si, 'Fourn_Reconcilie_HT'] = float(sim_rec.at[si, 'Montant HT'])
                    for fi in fis:
                        f_abo.at[fi, 'Pass_RAF'] = 'P3'

        # ======================================================================
        # NAVIGATION EN 4 ONGLETS MÉTIERS (STANDARD POTENTIALE)
        # ======================================================================
        tab1, tab2, tab3, tab4 = st.tabs([
            "🚨 1. Audit TVA (S 20% vs AE 0%)",
            "📊 2. Abonnements Récurrents HT",
            "📦 3. Frais Ponctuels par Dépôt (FAS)",
            "📞 4. Consommations par Dépôt"
        ])

        # ----------------------------------------------------------------------
        # ONGLET 1 : AUDIT TVA (S 20% VS AE 0%)
        # ----------------------------------------------------------------------
        with tab1:
            st.markdown("## 🚨 Audit Fiscal TVA — Contrôle Régime Standard (20%) vs Autoliquidation (AE)")
            st.info("**Règle d'Audit Potentiale :** L'ensemble des prestations facturées par Koesio Networks relève exclusivement du régime français **Standard à 20 %**. Toute ligne paramétrée en **AE (0 %)** dans l'ERP constitue une anomalie fiscale bloquante.")

            if not ssc_tva.empty:
                nb_ae = len(anomalies_tva)
                assiette_totale = ssc_tva['Assiette_HT'].sum()
                assiette_ae = anomalies_tva['Assiette_HT'].sum()

                k1, k2, k3 = st.columns(3)
                k1.metric("Règles Actives Contrôlées", f"{len(ssc_tva):,}".replace(',', ' '))
                k2.metric("Assiette Mensuelle HT Auditée", f"{assiette_totale:,.2f} €")
                k3.metric("Anomalies AE (0%) Détectées", f"{nb_ae} ligne(s)", delta=f"{assiette_ae:,.2f} € HT" if nb_ae > 0 else "0,00 € (Conforme)", delta_color="inverse")

                if nb_ae == 0:
                    st.success("✅ **Audit Fiscal 100 % Conforme :** Aucune règle active n'est paramétrée en autoliquidation (AE 0%). Toutes les lignes sont bien configurées en régime Standard 20 %.")
                else:
                    st.error(f"⚠️ **Attention :** {nb_ae} règle(s) présentent un paramétrage TVA non conforme à basculer en Standard 20 % :")
                    st.dataframe(anomalies_tva[['Code SSC', 'Raison sociale client', 'Code RFC', 'Libellé RFC', 'Gestion_TVA', 'Taux_TVA', 'Assiette_HT']], use_container_width=True, hide_index=True)

        # ----------------------------------------------------------------------
        # ONGLET 2 : ABONNEMENTS RÉCURRENTS HT
        # ----------------------------------------------------------------------
        with tab2:
            st.markdown(f"## ⏳ Audit de Temporalité ERP (Mois : {mois_etude})")
            t1, t2, t3, t4 = st.columns(4)
            t1.metric("🟢 Cycles Conformes", f"{nb_conformes} lignes")
            t2.metric("🔵 Régularisations Échues", f"{nb_echues} lignes")
            t3.metric("ℹ️ Facturations Anticipées", f"{nb_anticipees} lignes")
            t4.metric("🚨 Retards / Anomalies Dates", f"{nb_retards} lignes")

            st.markdown("---")
            st.markdown("## 🧮 Synthèse Macro-Financière HT (Abonnements Récurrents)")

            tot_sim_ht = sim_rec['Montant HT'].sum()
            tot_fourn_facture_ht = f_abo['Total H.T.'].sum()
            tot_fourn_reconcilie_ht = sim_rec['Fourn_Reconcilie_HT'].sum()
            ecart_net_global = tot_sim_ht - tot_fourn_facture_ht

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Simulation Artis (Achats)", f"{tot_sim_ht:,.2f} €")
            m2.metric("Total Facturé Fournisseur (Abo)", f"{tot_fourn_facture_ht:,.2f} €")
            m3.metric("Total Réconcilié Fournisseur", f"{tot_fourn_reconcilie_ht:,.2f} €")
            m4.metric("Écart Net Global HT (+ / -)", f"{ecart_net_global:+,.2f} €", delta="Attention > 500 €" if abs(ecart_net_global) > seuil_alerte else "Dans le seuil", delta_color="inverse" if abs(ecart_net_global) > seuil_alerte else "normal")

            if abs(ecart_net_global) > seuil_alerte:
                st.warning(f"🚨 **Alerte Seuil (> {seuil_alerte:,.0f} €) :** L'écart net global entre la simulation ERP et la facturation fournisseur atteint **{ecart_net_global:+,.2f} € HT**.")

            st.markdown("### 📋 Comparatif Détaillé par Dépôt (= Facture Fournisseur)")
            rows_dep = []
            for dep, fact_nom in [
                ('AF-KN-Ope', 'Facture Abonnements Opérateur'),
                ('AF-KN-MOB', 'Facture Forfaits Mobiles'),
                ('AF-KN-Heberg', 'Facture Abonnements Hébergement')
            ]:
                s_d = sim_rec[sim_rec['Depot_Cible'] == dep]
                f_d = f_abo[f_abo['Depot_Fourn'] == dep]
                sim_m = s_d['Montant HT'].sum()
                fourn_m = f_d['Total H.T.'].sum()
                rec_m = s_d['Fourn_Reconcilie_HT'].sum()
                ecart_m = sim_m - fourn_m
                alerte_txt = f"🚨 ATTENTION (> {seuil_alerte:.0f} €)" if abs(ecart_m) > seuil_alerte else "✅ OK (< 500 €)"
                rows_dep.append({
                    'Dépôt ERP': dep,
                    'Facture Associée': fact_nom,
                    'Total Simulation HT': round(sim_m, 2),
                    'Total Facturé Fournisseur HT': round(fourn_m, 2),
                    'Fournisseur Réconcilié HT': round(rec_m, 2),
                    'Écart Net (+ / -)': round(ecart_m, 2),
                    'Contrôle Seuil 500 €': alerte_txt
                })
            df_comp_dep = pd.DataFrame(rows_dep)
            st.dataframe(df_comp_dep, use_container_width=True, hide_index=True)

            # Détail des écarts résiduels client par client
            st.markdown("### 🔍 Analyse des Écarts Résiduels (Client par Client & Article)")
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                st.markdown("**➕ Présent Fournisseur / Absent ERP (RFC manquante ou date)**")
                f_reste = f_abo[f_abo['Pass_RAF'] == ''][['Depot_Fourn', 'Client final', 'Réf. produit', 'Libellé produit', 'Tech_ID', 'Qté', 'Total H.T.']].sort_values(by=['Depot_Fourn', 'Total H.T.'], ascending=[True, False])
                st.dataframe(f_reste, use_container_width=True, hide_index=True)
            with col_e2:
                st.markdown("**➖ Présent Simulation ERP / Absent Fournisseur (Résiliation / Mouvement)**")
                s_reste = sim_rec[sim_rec['Pass_RAF'] == ''][['Depot_Cible', 'Client', 'Code SSC', 'Code Article', 'Libellé RFC', 'Tech_ID', 'Quantite', 'Montant HT']].sort_values(by=['Depot_Cible', 'Montant HT'], ascending=[True, False])
                st.dataframe(s_reste, use_container_width=True, hide_index=True)

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
                    f_fas[['Depot_Ventilation', 'Client final', 'Site', 'Libellé produit', 'Date début', 'Qté', 'P.U. HT', 'Total H.T.']].sort_values(by=['Depot_Ventilation', 'Total H.T.'], ascending=[True, False]),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("Aucune ligne de frais ponctuels (FAS) détectée dans les annexes chargées.")

        # ----------------------------------------------------------------------
        # ONGLET 4 : CONSOMMATIONS PAR DÉPÔT
        # ----------------------------------------------------------------------
        with tab4:
            st.markdown("## 📞 Ventilation Analytique des Consommations par Dépôt (Pied de Facture)")
            st.info("**Règle de Répartition Potentiale :** Les consommations hors forfait (`CONSO`) sont automatiquement isolées par dépôt cible pour saisie directe en pied de facture dans Artis.")

            conso_ope = f_conso[f_conso['Depot_Fourn'] == 'AF-KN-Ope']['Total H.T.'].sum()
            conso_mob = f_conso[f_conso['Depot_Fourn'] == 'AF-KN-MOB']['Total H.T.'].sum()
            conso_heb = f_conso[f_conso['Depot_Fourn'] == 'AF-KN-Heberg']['Total H.T.'].sum()

            cc1, cc2, cc3 = st.columns(3)
            cc1.metric("AF-KN-Ope (Consommations Voix / Data)", f"{conso_ope:,.2f} € HT")
            cc2.metric("AF-KN-MOB (Consommations Mobiles)", f"{conso_mob:,.2f} € HT")
            cc3.metric("AF-KN-Heberg (Consommations Cloud)", f"{conso_heb:,.2f} € HT")

            st.markdown("### 📋 Synthèse des Montants de Consommation à Saisir par Dépôt")
            df_conso_synth = f_conso.groupby(['Depot_Fourn', 'Facture_Label', 'Libellé produit'], as_index=False).agg(
                Nombre_Lignes=('Total H.T.', 'count'),
                Montant_HT_Total=('Total H.T.', 'sum')
            )
            st.dataframe(df_conso_synth, use_container_width=True, hide_index=True)

        # ======================================================================
        # EXPORT EXCEL COMPLET DE CLÔTURE
        # ======================================================================
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_comp_dep.to_excel(writer, sheet_name='1_Synthese_Recurrent_HT', index=False)
            f_reste.to_excel(writer, sheet_name='2_Fourn_Absent_ERP', index=False)
            s_reste.to_excel(writer, sheet_name='3_ERP_Absent_Fourn', index=False)
            if not f_fas.empty:
                f_fas.to_excel(writer, sheet_name='4_FAS_Par_Depot', index=False)
            if not f_conso.empty:
                f_conso.to_excel(writer, sheet_name='5_Conso_Par_Depot', index=False)
            if not anomalies_tva.empty:
                anomalies_tva.to_excel(writer, sheet_name='6_Anomalies_TVA', index=False)

        st.download_button(
            label="📥 Télécharger le Dossier Complet de Rapprochement & Audit (.xlsx)",
            data=output.getvalue(),
            file_name=f"Potentiale_RAF_KoesioNetworks_{mois_etude.replace('/', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👈 Importez la **Simulation Achat Artis**, le **Listing des SSC Artis** et vos **4 Annexes Fournisseur** dans le volet de gauche pour lancer l'analyse.")

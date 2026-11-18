import streamlit as st
import pandas as pd

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="Potentiale — Outil Métier RAF",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- EN-TÊTE INSTITUTIONNEL ---
st.title("💼 Potentiale — Outil Métier RAF (Rapprochement Abonnements Fournisseurs)")
st.caption("Urbanisation SI : Pivot Identifiant Technique, Traçabilité d'Activation & Ventilation Analytique")
st.divider()

# --- BARRE LATÉRALE : IMPORTATION DES FICHIERS ---
with st.sidebar:
    st.header("📁 Importation des Fichiers")
    
    st.subheader("1. Simulation Achat ERP (.xlsx)")
    file_erp = st.file_uploader(
        "Déposer l'export simulation ERP", 
        type=["xlsx"], 
        key="erp_sim"
    )
    
    st.subheader("2. Annexe Fournisseur (.xlsx / .csv)")
    file_annexe = st.file_uploader(
        "Déposer l'annexe détaillée partenaire", 
        type=["xlsx", "csv"], 
        key="annexe_fourn"
    )
    
    st.subheader("3. Facture PDF (.pdf)")
    file_pdf = st.file_uploader(
        "Déposer la facture de contrôle", 
        type=["pdf"], 
        key="facture_pdf"
    )

# --- NAVIGATION PAR ONGLETS MÉTIER ---
tab1, tab2, tab3 = st.tabs([
    "📊 1. Abonnements Récurrents HT",
    "📦 2. Frais Ponctuels & BL",
    "📞 3. Consommations & Ventilation"
])

with tab1:
    st.subheader("Rapprochement des Abonnements Mensuels Récurrents (Simulation vs Fournisseur)")
    
    if file_erp and file_annexe:
        st.success("✅ Fichiers détectés ! Prêt pour l'exécution du moteur de rapprochement.")
    else:
        st.info("⏳ En attente du chargement de la simulation ERP et de l'annexe fournisseur dans le menu latéral.")
    
    st.divider()
    st.markdown("### 🧮 Synthèse Macro-Financière HT")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Simulation ERP (HT)", "0.00 €")
    col2.metric("Total Annexe Fournisseur (HT)", "0.00 €")
    col3.metric("Écart Net Global (HT)", "0.00 €")

with tab2:
    st.subheader("Ventilation des Frais Ponctuels par Entité & Famille")
    st.info("Module prêt pour l'isolement des prestations ponctuelles et le pointage des bons de livraison.")

with tab3:
    st.subheader("Ventilation Analytique des Consommations")
    st.info("Module prêt pour la répartition analytique des lignes variables.")
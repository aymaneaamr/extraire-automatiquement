import streamlit as st
import pandas as pd
from io import BytesIO
from utils.extraction import extraire_texte_pdf, chercher_champ

# Configuration de la page
st.set_page_config(page_title="Extraction Factures & BL", layout="wide")
st.title("📄 Remplissage automatique d'Excel depuis factures et bons de livraison")

# --- Initialisation des variables de session ---
if "df_final" not in st.session_state:
    st.session_state.df_final = pd.DataFrame(columns=[
        "fournisseur", "date", "commande", "bon_de_livraison", "numero_facture", "montant_facture"
    ])

# --- Barre latérale : paramètres d'extraction ---
with st.sidebar:
    st.header("⚙️ Paramètres d'extraction")
    st.markdown("Saisissez les expressions régulières pour chaque champ. Utilisez des parenthèses capturantes pour isoler la valeur recherchée.")
    
    patterns = {
        "fournisseur": st.text_input(
            "Fournisseur",
            r"Fournisseur\s*[:\-]?\s*(.+)",
            help="Exemple: Fournisseur : SARL Dupont → motif : Fournisseur\s*[:\-]?\s*(.+)"
        ),
        "date": st.text_input(
            "Date",
            r"Date\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
            help="Format attendu JJ/MM/AAAA. Exemple: Date : 15/03/2025"
        ),
        "commande": st.text_input(
            "N° commande",
            r"(?:Commande|Order|N° Commande)\s*[:\-]?\s*(\w+)",
            help="Exemple: Commande n° 12345"
        ),
        "bon_de_livraison": st.text_input(
            "N° bon de livraison",
            r"(?:BL?|Livraison|Bon de livraison)\s*[:\-]?\s*(\w+)",
            help="Exemple: BL : BL-2025-001"
        ),
        "numero_facture": st.text_input(
            "N° facture",
            r"(?:Facture|Invoice|N° Facture)\s*[:\-]?\s*(\w+)",
            help="Exemple: Facture n° F2025-01"
        ),
        "montant_facture": st.text_input(
            "Montant facture",
            r"(?:Total|Montant|Amount)\s*[:\-]?\s*([\d\s,\.]+(?:€|EUR)?)",
            help="Exemple: Total TTC : 125,50 €"
        )
    }
    
    st.markdown("---")
    st.markdown("💡 **Astuce** : vous pouvez affiner les regex après avoir vu le texte extrait.")

# --- Zone principale ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📎 Fichiers PDF à traiter")
    fichiers_pdf = st.file_uploader(
        "Sélectionnez un ou plusieurs PDF (factures, bons de livraison, ou documents combinés)",
        type=["pdf"],
        accept_multiple_files=True
    )
    
    if fichiers_pdf:
        st.info(f"{len(fichiers_pdf)} fichier(s) sélectionné(s). Chaque fichier générera une ligne dans le tableau.")

with col2:
    st.subheader("📊 Fichier Excel existant (optionnel)")
    excel_file = st.file_uploader("Téléchargez un fichier Excel avec les colonnes attendues", type=["xlsx", "xls"])
    if excel_file:
        df_existant = pd.read_excel(excel_file)
        # Vérifier que les colonnes sont présentes
        if set(df_existant.columns) == set(st.session_state.df_final.columns):
            st.session_state.df_final = df_existant
            st.success("Fichier chargé avec succès.")
        else:
            st.error("Le fichier ne contient pas les bonnes colonnes. Utilisation d'un DataFrame vide.")

# --- Bouton de traitement ---
if st.button("🚀 Extraire et ajouter les données"):
    if not fichiers_pdf:
        st.warning("Veuillez uploader au moins un fichier PDF.")
    else:
        nouvelles_lignes = []
        
        # Barre de progression
        progress_bar = st.progress(0)
        for i, fichier in enumerate(fichiers_pdf):
            with st.spinner(f"Traitement de {fichier.name}..."):
                # Extraire le texte
                texte = extraire_texte_pdf(fichier)
                
                # Afficher un extrait pour debug (optionnel, peut être désactivé)
                with st.expander(f"Aperçu du texte extrait de {fichier.name}"):
                    st.text(texte[:1000] + ("..." if len(texte) > 1000 else ""))
                
                # Créer une ligne de données
                ligne = {}
                for champ, pattern in patterns.items():
                    valeur = chercher_champ(texte, pattern)
                    ligne[champ] = valeur
                
                nouvelles_lignes.append(ligne)
            
            # Mettre à jour la progression
            progress_bar.progress((i + 1) / len(fichiers_pdf))
        
        # Ajouter les nouvelles lignes au DataFrame final
        df_nouv = pd.DataFrame(nouvelles_lignes)
        st.session_state.df_final = pd.concat([st.session_state.df_final, df_nouv], ignore_index=True)
        st.success(f"{len(nouvelles_lignes)} ligne(s) ajoutée(s) avec succès.")

# --- Affichage et édition du DataFrame ---
st.subheader("📋 Données consolidées")
st.dataframe(st.session_state.df_final, use_container_width=True)

# Option de modification manuelle
if st.checkbox("✏️ Modifier les données manuellement"):
    df_edit = st.data_editor(st.session_state.df_final, num_rows="dynamic")
    if st.button("Mettre à jour le tableau"):
        st.session_state.df_final = df_edit
        st.success("Mise à jour effectuée.")

# --- Téléchargement du fichier Excel ---
if not st.session_state.df_final.empty:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        st.session_state.df_final.to_excel(writer, index=False, sheet_name="Factures")
    output.seek(0)
    
    st.download_button(
        label="📥 Télécharger le fichier Excel",
        data=output,
        file_name="factures_extraites.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

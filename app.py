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
    st.markdown("Saisissez les expressions régulières pour chaque champ. Utilisez des parenthèses capturantes.")
    
    patterns = {
        "fournisseur": st.text_input("Fournisseur", r"Fournisseur\s*[:\-]?\s*(.+)", help="Exemple: Fournisseur : SARL Dupont"),
        "date": st.text_input("Date", r"Date\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", help="Format attendu JJ/MM/AAAA"),
        "commande": st.text_input("N° commande", r"(?:Commande|Order)\s*[:\-]?\s*(\w+)"),
        "bon_de_livraison": st.text_input("N° bon de livraison", r"(?:BL?|Livraison)\s*[:\-]?\s*(\w+)"),
        "numero_facture": st.text_input("N° facture", r"(?:Facture|Invoice)\s*[:\-]?\s*(\w+)"),
        "montant_facture": st.text_input("Montant facture", r"(?:Total|Montant|Amount)\s*[:\-]?\s*([\d\s,\.]+(?:€|EUR)?)", help="Inclure symbole monétaire si présent")
    }
    
    st.markdown("---")
    st.markdown("💡 **Astuce** : vous pouvez affiner les regex après avoir vu le texte extrait.")

# --- Zone principale ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📎 Fichiers sources")
    fichiers_invoice = st.file_uploader("Factures (PDF)", type=["pdf"], accept_multiple_files=True, key="inv")
    fichiers_bl = st.file_uploader("Bons de livraison (PDF)", type=["pdf"], accept_multiple_files=True, key="bl")
    
    # Option pour associer manuellement les fichiers si leur nombre diffère
    if fichiers_invoice and fichiers_bl and len(fichiers_invoice) != len(fichiers_bl):
        st.warning("Le nombre de factures et de BL ne correspond pas. Ils seront traités indépendamment.")

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
    if not fichiers_invoice and not fichiers_bl:
        st.warning("Veuillez uploader au moins un fichier (facture ou BL).")
    else:
        nouvelles_lignes = []
        
        # On va traiter chaque facture (ou chaque BL) comme une ligne.
        # Si les deux types sont fournis, on suppose que l'ordre correspond.
        nb_inv = len(fichiers_invoice) if fichiers_invoice else 0
        nb_bl = len(fichiers_bl) if fichiers_bl else 0
        nb_lignes = max(nb_inv, nb_bl)
        
        for i in range(nb_lignes):
            with st.spinner(f"Traitement de la ligne {i+1}..."):
                ligne = {"fournisseur": "", "date": "", "commande": "", 
                         "bon_de_livraison": "", "numero_facture": "", "montant_facture": ""}
                
                # Facture
                if i < nb_inv:
                    texte_inv = extraire_texte_pdf(fichiers_invoice[i])
                    st.text(f"Texte extrait de la facture {i+1} : {texte_inv[:200]}...")  # aperçu
                    for champ in ["fournisseur", "date", "commande", "numero_facture", "montant_facture"]:
                        valeur = chercher_champ(texte_inv, patterns[champ])
                        if valeur:
                            ligne[champ] = valeur
                    # Si le BL est dans la même facture, on peut aussi l'extraire ici
                    if "bon_de_livraison" in patterns:
                        ligne["bon_de_livraison"] = chercher_champ(texte_inv, patterns["bon_de_livraison"])
                
                # Bon de livraison séparé
                if i < nb_bl:
                    texte_bl = extraire_texte_pdf(fichiers_bl[i])
                    # On extrait seulement les champs pertinents pour le BL (souvent date, BL, commande)
                    # On peut surcharger ceux déjà trouvés si on veut prioriser le BL
                    # Ici on complète juste si vide
                    if not ligne["bon_de_livraison"]:
                        ligne["bon_de_livraison"] = chercher_champ(texte_bl, patterns["bon_de_livraison"])
                    if not ligne["date"]:
                        ligne["date"] = chercher_champ(texte_bl, patterns["date"])
                    if not ligne["commande"]:
                        ligne["commande"] = chercher_champ(texte_bl, patterns["commande"])
                
                nouvelles_lignes.append(ligne)
        
        # Ajouter les nouvelles lignes au DataFrame final
        df_nouv = pd.DataFrame(nouvelles_lignes)
        st.session_state.df_final = pd.concat([st.session_state.df_final, df_nouv], ignore_index=True)
        st.success(f"{len(nouvelles_lignes)} ligne(s) ajoutée(s) avec succès.")

# --- Affichage et édition du DataFrame ---
st.subheader("📋 Données consolidées")
st.dataframe(st.session_state.df_final, use_container_width=True)

# Option de modification manuelle (simple via éditeur de données)
if st.checkbox("✏️ Modifier les données manuellement"):
    df_edit = st.data_editor(st.session_state.df_final, num_rows="dynamic")
    if st.button("Mettre à jour"):
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

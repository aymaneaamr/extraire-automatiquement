import streamlit as st
import pandas as pd
import re
import os
import tempfile
from io import BytesIO
from PIL import Image
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes

# -------------------------------------------------------------------
# Fonctions d'extraction (intégrées pour simplifier)
# -------------------------------------------------------------------
def extraire_texte_avec_ocr(fichier, extension, lang='fra+eng'):
    """
    Extrait le texte d'un fichier uploadé (PDF ou image) en utilisant
    pdfplumber (PDF textuel) ou OCR (PDF scanné / image).
    """
    texte = ""
    try:
        if extension.lower() == 'pdf':
            with pdfplumber.open(fichier) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        texte += page_text
                    else:
                        # Convertir la page en image pour OCR
                        pil_image = page.to_image(resolution=300).original
                        texte += pytesseract.image_to_string(pil_image, lang=lang)
        else:
            # Image directe (jpg, png, etc.)
            image = Image.open(fichier)
            texte = pytesseract.image_to_string(image, lang=lang)
    except Exception as e:
        st.error(f"Erreur lors de l'extraction : {e}")
        texte = ""
    return texte

def chercher_champ(texte, pattern, groupe=1, flags=re.IGNORECASE):
    """
    Cherche un motif regex dans le texte et retourne le groupe capturé.
    Si le groupe demandé n'existe pas, retourne la correspondance entière.
    """
    if not texte:
        return ""
    match = re.search(pattern, texte, flags)
    if match:
        try:
            return match.group(groupe).strip()
        except IndexError:
            return match.group(0).strip()
    return ""

# -------------------------------------------------------------------
# Interface Streamlit
# -------------------------------------------------------------------
st.set_page_config(page_title="Extraction Factures & BL (OCR)", layout="wide")
st.title("📄 Remplissage automatique d'Excel depuis factures et bons de livraison (PDF ou image)")

# Vérification rapide de Tesseract
try:
    pytesseract.get_tesseract_version()
except Exception:
    st.sidebar.error(
        "⚠️ Tesseract n'est pas installé ou introuvable. "
        "L'OCR ne fonctionnera pas. Installez-le depuis https://github.com/tesseract-ocr/tesseract"
    )

# Initialisation du DataFrame en session
if "df_final" not in st.session_state:
    st.session_state.df_final = pd.DataFrame(columns=[
        "fournisseur", "date", "commande", "bon_de_livraison", "numero_facture", "montant_facture"
    ])

# -------------------------------------------------------------------
# Barre latérale : paramètres d'extraction
# -------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Paramètres d'extraction")
    st.markdown("Saisissez les expressions régulières pour chaque champ. Utilisez des parenthèses capturantes.")
    
    patterns = {
        "fournisseur": st.text_input(
            "Fournisseur",
            r"(?:Fournisseur|Supplier|Vendor|Client)\s*[:\-]?\s*(.+)",
            help="Exemple: Fournisseur : SARL Dupont"
        ),
        "date": st.text_input(
            "Date",
            r"(?:Date|Facture date|Invoice date)\s*[:\-]?\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})",
            help="Format JJ/MM/AAAA ou JJ-MM-AA"
        ),
        "commande": st.text_input(
            "N° commande",
            r"(?:Commande|Order|N° Commande|PO Number)\s*[:\-]?\s*([A-Z0-9\-]+)",
            help="Exemple: Commande n° 12345"
        ),
        "bon_de_livraison": st.text_input(
            "N° bon de livraison",
            r"(?:BL|Livraison|Delivery|Bon de livraison)\s*[:\-]?\s*([A-Z0-9\-]+)",
            help="Exemple: BL-2025-001"
        ),
        "numero_facture": st.text_input(
            "N° facture",
            r"(?:Facture|Invoice|N° Facture|Invoice Number)\s*[:\-]?\s*([A-Z0-9\-]+)",
            help="Exemple: F2025-01"
        ),
        "montant_facture": st.text_input(
            "Montant facture",
            r"(?:Total|Montant|Amount|TOTAL TTC|Net à payer)\s*[:\-]?\s*([\d\s,\.]+\s*(?:€|EUR|USD)?)",
            help="Exemple: Total TTC : 125,50 €"
        )
    }
    
    st.markdown("---")
    lang_ocr = st.selectbox("Langue OCR", ["fra+eng", "eng", "fra"], index=0,
                            help="Langues pour Tesseract (doivent être installées)")
    st.markdown("💡 **Astuce** : après extraction, ouvrez l'aperçu du texte pour ajuster les regex.")

# -------------------------------------------------------------------
# Zone principale
# -------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("📎 Fichiers à traiter")
    fichiers = st.file_uploader(
        "Sélectionnez un ou plusieurs fichiers (PDF, JPG, PNG, TIFF...)",
        type=["pdf", "jpg", "jpeg", "png", "tiff"],
        accept_multiple_files=True
    )
    if fichiers:
        st.info(f"{len(fichiers)} fichier(s) sélectionné(s). Chaque fichier générera une ligne.")

with col2:
    st.subheader("📊 Fichier Excel existant (optionnel)")
    excel_file = st.file_uploader("Téléchargez un fichier Excel avec les colonnes attendues", type=["xlsx", "xls"])
    if excel_file:
        try:
            df_existant = pd.read_excel(excel_file)
            if set(df_existant.columns) == set(st.session_state.df_final.columns):
                st.session_state.df_final = df_existant
                st.success("Fichier chargé avec succès.")
            else:
                st.error("Le fichier ne contient pas les bonnes colonnes. Utilisation d'un DataFrame vide.")
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")

# -------------------------------------------------------------------
# Bouton de traitement
# -------------------------------------------------------------------
if st.button("🚀 Extraire et ajouter les données"):
    if not fichiers:
        st.warning("Veuillez uploader au moins un fichier.")
    else:
        nouvelles_lignes = []
        progress_bar = st.progress(0)
        
        for i, fichier in enumerate(fichiers):
            with st.spinner(f"Traitement de {fichier.name}..."):
                # Déterminer l'extension
                ext = os.path.splitext(fichier.name)[1][1:].lower()
                
                # Lire le contenu du fichier (car après lecture, le pointeur est consommé)
                contenu = fichier.read()
                fichier_bytes = BytesIO(contenu)
                
                # Extraire le texte
                texte = extraire_texte_avec_ocr(fichier_bytes, ext, lang=lang_ocr)
                
                # Afficher un aperçu (dans un expandeur)
                with st.expander(f"Aperçu du texte extrait de {fichier.name}"):
                    st.text(texte[:2000] + ("..." if len(texte) > 2000 else ""))
                
                # Chercher chaque champ
                ligne = {}
                for champ, pattern in patterns.items():
                    valeur = chercher_champ(texte, pattern)
                    ligne[champ] = valeur
                
                nouvelles_lignes.append(ligne)
            
            # Mise à jour de la barre de progression
            progress_bar.progress((i + 1) / len(fichiers))
        
        # Ajout au DataFrame final
        if nouvelles_lignes:
            df_nouv = pd.DataFrame(nouvelles_lignes)
            st.session_state.df_final = pd.concat([st.session_state.df_final, df_nouv], ignore_index=True)
            st.success(f"{len(nouvelles_lignes)} ligne(s) ajoutée(s) avec succès.")
        else:
            st.warning("Aucune donnée extraite. Vérifiez les regex ou le contenu des fichiers.")

# -------------------------------------------------------------------
# Affichage et édition des données
# -------------------------------------------------------------------
st.subheader("📋 Données consolidées")
st.dataframe(st.session_state.df_final, use_container_width=True)

if st.checkbox("✏️ Modifier les données manuellement"):
    df_edit = st.data_editor(st.session_state.df_final, num_rows="dynamic")
    if st.button("Mettre à jour le tableau"):
        st.session_state.df_final = df_edit
        st.success("Mise à jour effectuée.")

# -------------------------------------------------------------------
# Téléchargement du fichier Excel
# -------------------------------------------------------------------
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

import streamlit as st
import pandas as pd
import re
import os
import shutil
from io import BytesIO
from PIL import Image
import pdfplumber
import pytesseract

# -------------------------------------------------------------------
# Vérification de Tesseract
# -------------------------------------------------------------------
try:
    pytesseract.get_tesseract_version()
    st.sidebar.success("✅ Tesseract trouvé")
except Exception:
    st.sidebar.error("⚠️ Tesseract non installé. L'OCR ne fonctionnera pas.")

# -------------------------------------------------------------------
# Fonctions d'extraction
# -------------------------------------------------------------------
def extraire_texte_avec_ocr(fichier, extension, lang='fra+eng'):
    texte = ""
    try:
        if extension.lower() == 'pdf':
            with pdfplumber.open(fichier) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        texte += page_text
                    else:
                        pil_image = page.to_image(resolution=300).original
                        texte += pytesseract.image_to_string(pil_image, lang=lang)
        else:
            image = Image.open(fichier)
            texte = pytesseract.image_to_string(image, lang=lang)
    except Exception as e:
        st.error(f"Erreur OCR : {e}")
    return texte

def chercher_champ(texte, pattern):
    if not texte:
        return ""
    match = re.search(pattern, texte, re.IGNORECASE)
    if match:
        try:
            return match.group(1).strip()
        except IndexError:
            return match.group(0).strip()
    return ""

# -------------------------------------------------------------------
# Interface Streamlit
# -------------------------------------------------------------------
st.set_page_config(page_title="Extraction Factures & BL", layout="wide")
st.title("📄 Remplissage automatique d'Excel depuis factures et bons de livraison")

if "df_final" not in st.session_state:
    st.session_state.df_final = pd.DataFrame(columns=[
        "fournisseur", "date", "commande", "bon_de_livraison", "numero_facture", "montant_facture"
    ])

with st.sidebar:
    st.header("⚙️ Paramètres d'extraction")
    st.markdown("Ajustez les expressions régulières en fonction du texte extrait.")

    patterns = {
        "fournisseur": st.text_input(
            "Fournisseur",
            r"(?:Fournisseur|Supplier|Vendor|Client)\s*[:\-]?\s*([A-Z][A-Z\s\-\.]+(?:S\.?A\.?|SARL|SAS)?)",
            help="Ex: Fournisseur : SARL DUPONT"
        ),
        "date": st.text_input(
            "Date",
            r"(?:Date|Facture\s*du|Invoice\s*date)\s*[:\-]?\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})",
            help="Format JJ/MM/AAAA"
        ),
        "commande": st.text_input(
            "N° commande",
            r"(?:Commande|Order|N°\s*Commande|PO\s*Number|Référence\s*commande)\s*[:\-]?\s*([A-Z0-9\-/]{5,})",
            help="Au moins 5 caractères alphanumériques"
        ),
        "bon_de_livraison": st.text_input(
            "N° bon de livraison",
            r"(?:BL|Bon\s*de\s*livraison|Delivery\s*note|N°\s*BL)\s*[:\-]?\s*([A-Z0-9\-/]{3,})",
            help="Ex: BL-2025-001"
        ),
        "numero_facture": st.text_input(
            "N° facture",
            r"(?:Facture|Invoice|N°\s*Facture|Invoice\s*Number)\s*[:\-]?\s*([A-Z0-9\-/]{3,})",
            help="Ex: FA-2025-01"
        ),
        "montant_facture": st.text_input(
            "Montant",
            r"(?:Total|Montant|Amount|TOTAL\s*TTC|Net\s*à\s*payer)\s*[:\-]?\s*([\d\s,\.]+\s*(?:€|EUR)?)",
            help="Ex: Total TTC : 125,50 €"
        )
    }

    lang_ocr = st.selectbox("Langue OCR", ["fra+eng", "eng", "fra"], index=0)

col1, col2 = st.columns(2)
with col1:
    fichiers = st.file_uploader(
        "Fichiers (PDF, JPG, PNG, TIFF...)",
        type=["pdf", "jpg", "jpeg", "png", "tiff"],
        accept_multiple_files=True
    )
with col2:
    excel_file = st.file_uploader("Fichier Excel existant (optionnel)", type=["xlsx", "xls"])
    if excel_file:
        try:
            df_existant = pd.read_excel(excel_file)
            if set(df_existant.columns) == set(st.session_state.df_final.columns):
                st.session_state.df_final = df_existant
                st.success("Fichier Excel chargé.")
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")

if st.button("🚀 Extraire et ajouter les données"):
    if not fichiers:
        st.warning("Veuillez sélectionner au moins un fichier.")
    else:
        nouvelles_lignes = []
        progress = st.progress(0)
        for i, f in enumerate(fichiers):
            with st.spinner(f"Traitement de {f.name}..."):
                ext = os.path.splitext(f.name)[1][1:].lower()
                contenu = f.read()
                texte = extraire_texte_avec_ocr(BytesIO(contenu), ext, lang=lang_ocr)

                with st.expander(f"Aperçu de {f.name}"):
                    st.text(texte[:2000] + ("..." if len(texte) > 2000 else ""))

                ligne = {champ: chercher_champ(texte, pattern) for champ, pattern in patterns.items()}
                nouvelles_lignes.append(ligne)
            progress.progress((i+1)/len(fichiers))

        if nouvelles_lignes:
            df_nouv = pd.DataFrame(nouvelles_lignes)
            st.session_state.df_final = pd.concat([st.session_state.df_final, df_nouv], ignore_index=True)
            st.success(f"{len(nouvelles_lignes)} ligne(s) ajoutée(s).")

st.subheader("📋 Données consolidées")
st.dataframe(st.session_state.df_final, width='stretch')

if st.checkbox("✏️ Modifier les données manuellement"):
    df_edit = st.data_editor(st.session_state.df_final, num_rows="dynamic")
    if st.button("Mettre à jour"):
        st.session_state.df_final = df_edit

if not st.session_state.df_final.empty:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        st.session_state.df_final.to_excel(writer, index=False)
    st.download_button("📥 Télécharger Excel", data=output.getvalue(), file_name="factures_extraites.xlsx")

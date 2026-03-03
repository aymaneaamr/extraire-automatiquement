import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
import pdfplumber

st.set_page_config(page_title="Extraction Factures & BL", layout="wide")
st.title("📄 Remplissage automatique d'Excel (version sans OCR)")

if "df_final" not in st.session_state:
    st.session_state.df_final = pd.DataFrame(columns=[
        "fournisseur", "date", "commande", "bon_de_livraison", "numero_facture", "montant_facture"
    ])

with st.sidebar:
    st.header("⚙️ Paramètres d'extraction")
    patterns = {
        "fournisseur": st.text_input("Fournisseur", r"(?:Fournisseur|Supplier)\s*[:\-]?\s*(.+)"),
        "date": st.text_input("Date", r"(?:Date)\s*[:\-]?\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})"),
        "commande": st.text_input("N° commande", r"(?:Commande|Order)\s*[:\-]?\s*([A-Z0-9\-/]+)"),
        "bon_de_livraison": st.text_input("N° BL", r"(?:BL|Bon de livraison)\s*[:\-]?\s*([A-Z0-9\-/]+)"),
        "numero_facture": st.text_input("N° facture", r"(?:Facture|Invoice)\s*[:\-]?\s*([A-Z0-9\-/]+)"),
        "montant_facture": st.text_input("Montant", r"(?:Total|Montant)\s*[:\-]?\s*([\d\s,\.]+(?:€|EUR)?)")
    }

def extraire_texte_pdf(fichier):
    texte = ""
    with pdfplumber.open(fichier) as pdf:
        for page in pdf.pages:
            texte += page.extract_text() or ""
    return texte

def chercher_champ(texte, pattern):
    match = re.search(pattern, texte, re.IGNORECASE)
    return match.group(1).strip() if match and match.groups() else (match.group(0).strip() if match else "")

col1, col2 = st.columns(2)
with col1:
    fichiers = st.file_uploader("Fichiers PDF", type="pdf", accept_multiple_files=True)
with col2:
    excel_file = st.file_uploader("Excel existant", type=["xlsx", "xls"])
    if excel_file:
        df_existant = pd.read_excel(excel_file)
        if set(df_existant.columns) == set(st.session_state.df_final.columns):
            st.session_state.df_final = df_existant

if st.button("Extraire"):
    if fichiers:
        nouvelles = []
        for f in fichiers:
            texte = extraire_texte_pdf(f)
            with st.expander(f"Aperçu {f.name}"):
                st.text(texte[:1000])
            ligne = {champ: chercher_champ(texte, pattern) for champ, pattern in patterns.items()}
            nouvelles.append(ligne)
        st.session_state.df_final = pd.concat([st.session_state.df_final, pd.DataFrame(nouvelles)], ignore_index=True)

st.dataframe(st.session_state.df_final, width='stretch')
if not st.session_state.df_final.empty:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        st.session_state.df_final.to_excel(writer, index=False)
    st.download_button("Télécharger Excel", data=output.getvalue(), file_name="factures.xlsx")

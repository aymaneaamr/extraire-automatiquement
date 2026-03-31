"""
Application Streamlit pour l'extraction automatique d'informations de factures via l'API Claude d'Anthropic.
"""

import streamlit as st
import anthropic
import base64
import json
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime
from typing import List, Dict, Optional
import re
import sys

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Extraction Factures",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        color: white;
        padding: 20px 30px;
        border-radius: 12px;
        margin-bottom: 25px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 { margin: 0; font-size: 2rem; font-weight: 600; }
    .main-header p  { margin: 5px 0 0; opacity: 0.85; font-size: 1rem; }
    .info-box {
        background: #eff6ff;
        border-left: 4px solid #3b82f6;
        padding: 12px 16px;
        border-radius: 6px;
        margin-bottom: 15px;
        font-size: 0.9rem;
        color: #1e3a8a;
    }
    .success-box {
        background: #f0fdf4;
        border-left: 4px solid #22c55e;
        padding: 12px 16px;
        border-radius: 6px;
        margin: 10px 0;
        color: #166534;
    }
    .error-box {
        background: #fef2f2;
        border-left: 4px solid #ef4444;
        padding: 12px 16px;
        border-radius: 6px;
        margin: 10px 0;
        color: #991b1b;
    }
    .stButton > button {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 24px;
        font-size: 1rem;
        font-weight: 600;
        width: 100%;
        transition: all .2s;
    }
    .stButton > button:hover {
        opacity: .85;
        transform: translateY(-1px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .footer {
        text-align: center;
        margin-top: 40px;
        padding: 20px;
        color: #6b7280;
        font-size: 0.875rem;
        border-top: 1px solid #e5e7eb;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🧾 Extraction Automatique de Factures</h1>
    <p>Importez vos factures PDF/image et exportez les données structurées vers Excel</p>
</div>
""", unsafe_allow_html=True)

# ── Session state initialization ──────────────────────────────────────────────
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = []
if "api_key_configured" not in st.session_state:
    st.session_state.api_key_configured = False

# ── Résolution de la clé API (secrets Streamlit Cloud OU saisie manuelle) ────
def get_api_key() -> Optional[str]:
    """Récupère la clé API depuis st.secrets ou depuis la session."""
    # Priorité 1 : st.secrets (pour Streamlit Cloud)
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        pass
    # Priorité 2 : saisie manuelle via la sidebar
    return st.session_state.get("manual_api_key")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("---")

    # Vérifier si la clé est déjà dans les secrets
    key_from_secrets = False
    try:
        _ = st.secrets["ANTHROPIC_API_KEY"]
        key_from_secrets = True
    except (KeyError, FileNotFoundError):
        pass

    if key_from_secrets:
        st.success("✅ Clé API chargée depuis les secrets")
        st.session_state.api_key_configured = True
    else:
        api_key_input = st.text_input(
            "Clé API Anthropic",
            type="password",
            help="Entrez votre clé API Anthropic (sk-ant-...)",
            placeholder="sk-ant-..."
        )
        if api_key_input:
            st.session_state.manual_api_key = api_key_input
            st.session_state.api_key_configured = True
            st.success("✅ Clé API configurée")
        else:
            st.session_state.api_key_configured = False
            st.warning("⚠️ Clé API requise pour l'extraction")

    st.markdown("---")
    st.markdown("### 📋 Instructions")
    st.markdown("""
    1. Configurez votre clé API Anthropic
    2. Importez vos factures (PDF, PNG, JPG)
    3. Cliquez sur **Extraire**
    4. Téléchargez le fichier Excel

    **Formats supportés :** PDF · PNG · JPG/JPEG
    """)

    st.markdown("---")
    st.markdown(f"**Version :** 1.1.0  \n**Date :** {datetime.now().strftime('%d/%m/%Y')}")

# ── Helpers ───────────────────────────────────────────────────────────────────
def encode_file(file_bytes: bytes) -> str:
    return base64.standard_b64encode(file_bytes).decode("utf-8")


def extract_invoice_info(file_bytes: bytes, mime: str, filename: str) -> Dict:
    """Extrait les informations d'une facture via l'API Claude."""
    try:
        api_key = get_api_key()
        if not api_key:
            raise ValueError("Clé API non configurée.")

        client = anthropic.Anthropic(api_key=api_key)

        prompt = """Analyse cette facture et extrais les informations suivantes en JSON UNIQUEMENT (pas de texte avant/après):
{
  "fournisseur": "nom du fournisseur/vendeur",
  "date": "date de la facture au format JJ/MM/AAAA",
  "commande": "numéro de commande ou bon de commande",
  "bon_de_livraison": "numéro du bon de livraison",
  "numero_facture": "numéro de la facture",
  "montant_facture": "montant total TTC en chiffres uniquement (ex: 1500.00)"
}

Règles :
- Si une information est absente, mets null
- Pour le montant, extrais uniquement le nombre, sans symbole monétaire
- La date doit être au format JJ/MM/AAAA
Réponds UNIQUEMENT avec le JSON, rien d'autre."""

        if mime == "application/pdf":
            content = [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": encode_file(file_bytes)}},
                {"type": "text", "text": prompt}
            ]
        else:
            content = [
                {"type": "image", "source": {"type": "base64", "media_type": mime, "data": encode_file(file_bytes)}},
                {"type": "text", "text": prompt}
            ]

        response = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=1000,
            temperature=0,
            messages=[{"role": "user", "content": content}]
        )

        raw = response.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        raw = re.sub(r'\s+', ' ', raw)

        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            json_match = re.search(r'\{.*\}', raw, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group())
            else:
                raise ValueError(f"Impossible de parser le JSON : {raw[:200]}")

        for field in ["fournisseur", "date", "commande", "bon_de_livraison", "numero_facture", "montant_facture"]:
            if field not in data or data[field] in ["null", "NULL", "None", ""]:
                data[field] = None

        data["fichier"] = filename
        return data

    except Exception as e:
        return {
            "fournisseur": None, "date": None, "commande": None,
            "bon_de_livraison": None, "numero_facture": None,
            "montant_facture": None, "fichier": filename, "erreur": str(e)
        }


def parse_montant(value) -> float:
    """Convertit une valeur montant (str ou number) en float."""
    if value is None or value in ("null", ""):
        return 0.0
    try:
        if isinstance(value, str):
            value = re.sub(r'[^\d.,-]', '', value)
            value = value.replace(',', '.')
            if value.count('.') > 1:
                value = value.replace('.', '', value.count('.') - 1)
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def build_excel(records: List[Dict]) -> bytes:
    """Construit un fichier Excel à partir des données extraites."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Factures"

    header_fill = PatternFill("solid", fgColor="1e3a8a")
    alt_fill    = PatternFill("solid", fgColor="dbeafe")
    white_fill  = PatternFill("solid", fgColor="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="d1d5db"),
        right=Side(style="thin", color="d1d5db"),
        top=Side(style="thin", color="d1d5db"),
        bottom=Side(style="thin", color="d1d5db"),
    )

    columns = [
        ("Fournisseur",      "fournisseur",     25),
        ("Date",             "date",            14),
        ("N° Commande",      "commande",        18),
        ("Bon de Livraison", "bon_de_livraison", 20),
        ("N° Facture",       "numero_facture",  20),
        ("Montant TTC",      "montant_facture", 15),
        ("Fichier Source",   "fichier",         30),
    ]

    for col_idx, (label, _, width) in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 30

    montant_total = 0.0
    for row_idx, record in enumerate(records, 2):
        fill = alt_fill if row_idx % 2 == 0 else white_fill
        for col_idx, (_, key, _) in enumerate(columns, 1):
            value = record.get(key)
            if key == "montant_facture":
                value = parse_montant(value)
                montant_total += value
            cell = ws.cell(row=row_idx, column=col_idx,
                           value=value if value not in (None, "null", "") else "")
            cell.font = Font(name="Arial", size=10)
            cell.fill = fill
            cell.alignment = Alignment(
                horizontal="right" if key == "montant_facture" else "left",
                vertical="center"
            )
            cell.border = thin_border
            if key == "montant_facture" and isinstance(value, float):
                cell.number_format = '#,##0.00 "MAD"'
        ws.row_dimensions[row_idx].height = 20

    last_row = len(records) + 2
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
    lbl = ws.cell(row=last_row, column=1, value="TOTAL GÉNÉRAL")
    lbl.font = Font(bold=True, name="Arial", size=11)
    lbl.fill = PatternFill("solid", fgColor="fbbf24")
    lbl.border = thin_border
    tot = ws.cell(row=last_row, column=6, value=montant_total)
    tot.font = Font(bold=True, name="Arial", size=11)
    tot.number_format = '#,##0.00 "MAD"'
    tot.alignment = Alignment(horizontal="right")
    tot.fill = PatternFill("solid", fgColor="fbbf24")
    tot.border = thin_border

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Interface principale ──────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown("### 📤 Importer des Factures")
    st.markdown("""
    <div class="info-box">
        <strong>Formats acceptés :</strong> PDF, PNG, JPG, JPEG<br>
        <strong>Taille max :</strong> 200 MB par fichier
    </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Choisissez vos fichiers",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        help="Glissez-déposez ou cliquez pour sélectionner des fichiers"
    )

    if uploaded_files:
        st.markdown(f"""
        <div class="success-box">✅ {len(uploaded_files)} fichier(s) sélectionné(s)</div>
        """, unsafe_allow_html=True)
        with st.expander("📋 Voir les fichiers sélectionnés"):
            for f in uploaded_files:
                st.text(f"📄 {f.name} ({f.size / 1024:.1f} KB)")

    extract_btn = st.button(
        "🔍 Extraire les Informations",
        disabled=not (uploaded_files and st.session_state.api_key_configured),
        use_container_width=True
    )

    if st.session_state.extracted_data:
        st.markdown("---")
        st.markdown("### 💾 Exporter")
        excel_bytes = build_excel(st.session_state.extracted_data)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        col_dl, col_clear = st.columns(2)
        with col_dl:
            st.download_button(
                label="📥 Télécharger Excel",
                data=excel_bytes,
                file_name=f"factures_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_clear:
            if st.button("🗑️ Effacer", use_container_width=True):
                st.session_state.extracted_data = []
                st.rerun()

with col_right:
    st.markdown("### 📊 Données Extraites")

    if extract_btn and uploaded_files:
        mime_map = {
            "pdf": "application/pdf",
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
        }
        progress_bar = st.progress(0)
        status_text  = st.empty()
        newly_extracted = []

        for i, uploaded_file in enumerate(uploaded_files):
            ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
            mime_type = mime_map.get(ext, "application/octet-stream")
            status_text.info(f"⏳ Traitement de **{uploaded_file.name}**… ({i+1}/{len(uploaded_files)})")

            try:
                info = extract_invoice_info(uploaded_file.read(), mime_type, uploaded_file.name)
                newly_extracted.append(info)
                if "erreur" not in info:
                    st.markdown(f'<div class="success-box">✅ {uploaded_file.name} traité avec succès</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="error-box">❌ Erreur — {uploaded_file.name} : {info["erreur"]}</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erreur inattendue pour {uploaded_file.name} : {e}")

            progress_bar.progress((i + 1) / len(uploaded_files))

        st.session_state.extracted_data.extend(newly_extracted)
        progress_bar.empty()
        status_text.empty()
        st.success(f"✅ Traitement terminé ! {len(newly_extracted)} fichier(s) traité(s)")
        st.rerun()

    if st.session_state.extracted_data:
        df = pd.DataFrame(st.session_state.extracted_data)
        column_mapping = {
            'fournisseur': 'Fournisseur', 'date': 'Date',
            'commande': 'Commande', 'bon_de_livraison': 'Bon Livraison',
            'numero_facture': 'N° Facture', 'montant_facture': 'Montant (MAD)',
            'fichier': 'Fichier'
        }
        df_display = df.rename(columns=column_mapping)
        display_cols = [c for c in ['Fournisseur', 'Date', 'N° Facture', 'Montant (MAD)', 'Fichier'] if c in df_display.columns]
        st.dataframe(df_display[display_cols], use_container_width=True, height=400, hide_index=True)

        st.markdown("---")
        c1, c2, c3, c4 = st.columns(4)
        total_amount  = sum(parse_montant(r.get("montant_facture")) for r in st.session_state.extracted_data)
        valid_invoices = sum(1 for r in st.session_state.extracted_data if parse_montant(r.get("montant_facture")) > 0)
        c1.metric("📄 Factures traitées", len(st.session_state.extracted_data))
        c2.metric("💰 Montant total", f"{total_amount:,.2f} MAD")
        c3.metric("✅ Extractions réussies", f"{valid_invoices}/{len(st.session_state.extracted_data)}")
        c4.metric("📅 Dernière extraction", datetime.now().strftime("%d/%m/%Y %H:%M"))

        with st.expander("📋 Voir toutes les données détaillées"):
            st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        st.markdown("""
        <div style="text-align:center;padding:60px 20px;background:#f8fafc;border-radius:8px;">
            <div style="font-size:5rem;margin-bottom:20px;">🧾</div>
            <p style="font-size:1.2rem;color:#1e293b;margin-bottom:10px;">Aucune donnée extraite</p>
            <p style="color:#64748b;">Importez des factures dans le panneau de gauche<br>et cliquez sur <strong>Extraire</strong> pour commencer</p>
        </div>
        """, unsafe_allow_html=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    <p>🧾 Extraction Automatique de Factures | Propulsé par Claude API et Streamlit</p>
    <p style="font-size:0.8rem;margin-top:5px;">© 2024 - Tous droits réservés</p>
</div>
""", unsafe_allow_html=True)

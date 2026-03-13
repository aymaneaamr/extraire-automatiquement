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
import re

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Extraction Factures",
    page_icon="🧾",
    layout="wide"
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
    }
    .main-header h1 { margin: 0; font-size: 2rem; }
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
    .stDataFrame { border-radius: 8px; overflow: hidden; }
    .stButton > button {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 24px;
        font-size: 1rem;
        font-weight: 600;
        width: 100%;
        transition: opacity .2s;
    }
    .stButton > button:hover { opacity: .85; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🧾 Extraction Automatique de Factures</h1>
    <p>Importez vos factures PDF/image et exportez les données vers Excel</p>
</div>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = []

# ── Helper: encode file to base64 ─────────────────────────────────────────────
def encode_file(file_bytes: bytes, mime: str) -> str:
    return base64.standard_b64encode(file_bytes).decode("utf-8")

# ── Helper: extract invoice info via Claude API ───────────────────────────────
def extract_invoice_info(file_bytes: bytes, mime: str, filename: str) -> dict:
    client = anthropic.Anthropic()

    prompt = """Analyse cette facture et extrais les informations suivantes en JSON UNIQUEMENT (pas de texte avant/après):
{
  "fournisseur": "nom du fournisseur/vendeur",
  "date": "date de la facture au format JJ/MM/AAAA",
  "commande": "numéro de commande ou bon de commande",
  "bon_de_livraison": "numéro du bon de livraison",
  "numero_facture": "numéro de la facture",
  "montant_facture": "montant total TTC en chiffres uniquement (ex: 1500.00)"
}

Si une information est absente, mets null.
Réponds UNIQUEMENT avec le JSON, rien d'autre."""

    if mime == "application/pdf":
        content = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": encode_file(file_bytes, mime)
                }
            },
            {"type": "text", "text": prompt}
        ]
    else:
        content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": mime,
                    "data": encode_file(file_bytes, mime)
                }
            },
            {"type": "text", "text": prompt}
        ]

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{"role": "user", "content": content}]
    )

    raw = response.content[0].text.strip()
    # strip markdown code fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)
    data["fichier"] = filename
    return data

# ── Helper: build Excel workbook ──────────────────────────────────────────────
def build_excel(records: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Factures"

    # colours
    header_fill  = PatternFill("solid", fgColor="1e3a8a")
    alt_fill     = PatternFill("solid", fgColor="dbeafe")
    white_fill   = PatternFill("solid", fgColor="FFFFFF")
    thin_border  = Border(
        left=Side(style="thin", color="d1d5db"),
        right=Side(style="thin", color="d1d5db"),
        top=Side(style="thin", color="d1d5db"),
        bottom=Side(style="thin", color="d1d5db"),
    )

    columns = [
        ("Fournisseur",       "fournisseur",      25),
        ("Date",              "date",             14),
        ("Commande",          "commande",         18),
        ("Bon de Livraison",  "bon_de_livraison", 20),
        ("Numéro de Facture", "numero_facture",   20),
        ("Montant de Facture","montant_facture",  20),
        ("Fichier Source",    "fichier",          30),
    ]

    # header row
    for col_idx, (label, _, width) in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 30

    # data rows
    for row_idx, record in enumerate(records, 2):
        fill = alt_fill if row_idx % 2 == 0 else white_fill
        for col_idx, (_, key, _) in enumerate(columns, 1):
            value = record.get(key)
            if key == "montant_facture" and value not in (None, "null", ""):
                try:
                    value = float(str(value).replace(",", ".").replace(" ", ""))
                except ValueError:
                    pass
            cell = ws.cell(row=row_idx, column=col_idx, value=value if value not in (None, "null") else "")
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = fill
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border    = thin_border
            if key == "montant_facture" and isinstance(value, float):
                cell.number_format = '#,##0.00 "MAD"'
                cell.alignment = Alignment(horizontal="right")
        ws.row_dimensions[row_idx].height = 20

    # totals row
    last = len(records) + 2
    ws.cell(row=last, column=1, value="TOTAL").font = Font(bold=True, name="Arial")
    total_cell = ws.cell(
        row=last, column=6,
        value=f"=SUM(F2:F{last-1})"
    )
    total_cell.font         = Font(bold=True, name="Arial")
    total_cell.number_format = '#,##0.00 "MAD"'
    total_cell.alignment    = Alignment(horizontal="right")
    total_cell.fill         = PatternFill("solid", fgColor="dbeafe")
    total_cell.border       = thin_border

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── UI ────────────────────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown("### 📤 Importer des Factures")
    st.markdown('<div class="info-box">Formats acceptés : PDF, PNG, JPG, JPEG</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Glissez vos fichiers ici",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if uploaded_files:
        st.markdown(f"**{len(uploaded_files)} fichier(s) sélectionné(s)**")

    extract_btn = st.button("🔍 Extraire les Informations", disabled=not uploaded_files)

    if st.session_state.extracted_data:
        st.markdown("---")
        st.markdown("### 💾 Exporter")
        excel_bytes = build_excel(st.session_state.extracted_data)
        timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="⬇️ Télécharger Excel",
            data=excel_bytes,
            file_name=f"factures_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if st.button("🗑️ Effacer les données"):
            st.session_state.extracted_data = []
            st.rerun()

with col_right:
    st.markdown("### 📊 Données Extraites")

    if extract_btn and uploaded_files:
        progress = st.progress(0)
        status   = st.empty()
        newly    = []

        mime_map = {
            "pdf":  "application/pdf",
            "png":  "image/png",
            "jpg":  "image/jpeg",
            "jpeg": "image/jpeg",
        }

        for i, uf in enumerate(uploaded_files):
            ext  = uf.name.rsplit(".", 1)[-1].lower()
            mime = mime_map.get(ext, "image/jpeg")
            status.info(f"⏳ Traitement : **{uf.name}** …")
            try:
                info = extract_invoice_info(uf.read(), mime, uf.name)
                newly.append(info)
                st.markdown(f'<div class="success-box">✅ {uf.name} traité avec succès</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erreur pour {uf.name} : {e}")
            progress.progress((i + 1) / len(uploaded_files))

        st.session_state.extracted_data.extend(newly)
        status.empty()
        progress.empty()

    if st.session_state.extracted_data:
        df = pd.DataFrame(st.session_state.extracted_data)
        df.columns = [c.replace("_", " ").title() for c in df.columns]

        df_display = df.rename(columns={
            "Fournisseur":      "Fournisseur",
            "Date":             "Date",
            "Commande":         "Commande",
            "Bon De Livraison": "Bon de Livraison",
            "Numero Facture":   "N° Facture",
            "Montant Facture":  "Montant",
            "Fichier":          "Fichier",
        })

        st.dataframe(df_display, use_container_width=True, height=400)

        # KPIs
        k1, k2, k3 = st.columns(3)
        total = 0
        for r in st.session_state.extracted_data:
            try:
                total += float(str(r.get("montant_facture", 0) or 0).replace(",", ".").replace(" ", ""))
            except ValueError:
                pass

        k1.metric("📄 Factures traitées", len(st.session_state.extracted_data))
        k2.metric("💰 Montant total", f"{total:,.2f} MAD")
        k3.metric("📅 Date d'extraction", datetime.now().strftime("%d/%m/%Y"))
    else:
        st.markdown("""
        <div style="text-align:center; padding:60px 20px; color:#94a3b8;">
            <div style="font-size:4rem;">🧾</div>
            <p style="font-size:1.1rem; margin-top:10px;">
                Importez des factures et cliquez sur <strong>Extraire</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)

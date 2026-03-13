"""
Application Streamlit pour l'extraction automatique d'informations de factures
via l'API Claude d'Anthropic.
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
import re
import sys
import subprocess
import pkg_resources

# ── Vérification et installation des dépendances nécessaires ─────────────────
def check_and_install_dependencies():
    """Vérifie et installe les versions correctes des dépendances"""
    required_packages = {
        'altair': '4.2.0',
        'streamlit': '1.28.0',
        'pandas': '2.0.3',
        'openpyxl': '3.1.2',
        'anthropic': '0.18.0'
    }
    
    for package, version in required_packages.items():
        try:
            installed = pkg_resources.get_distribution(package).version
            if installed != version:
                st.warning(f"Mise à jour de {package} de {installed} vers {version}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", f"{package}=={version}"])
        except pkg_resources.DistributionNotFound:
            st.warning(f"Installation de {package}=={version}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", f"{package}=={version}"])

# Vérifier les dépendances au démarrage (optionnel, peut être commenté en production)
# check_and_install_dependencies()

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
    .main-header h1 { 
        margin: 0; 
        font-size: 2rem;
        font-weight: 600;
    }
    .main-header p  { 
        margin: 5px 0 0; 
        opacity: 0.85; 
        font-size: 1rem;
    }
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
    .stDataFrame { 
        border-radius: 8px; 
        overflow: hidden;
        border: 1px solid #e5e7eb;
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
        border: 1px solid transparent;
    }
    .stButton > button:hover { 
        opacity: .85;
        transform: translateY(-1px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .stButton > button:disabled {
        opacity: 0.5;
        cursor: not-allowed;
    }
    .upload-text {
        text-align: center;
        padding: 20px;
        border: 2px dashed #3b82f6;
        border-radius: 8px;
        background: #f8fafc;
    }
    div[data-testid="stFileUploader"] {
        width: 100%;
    }
    div[data-testid="stFileUploader"] section {
        padding: 0;
    }
    div[data-testid="stFileUploader"] button {
        background: #3b82f6;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 6px;
        font-weight: 500;
    }
    .metric-card {
        background: white;
        padding: 16px;
        border-radius: 8px;
        border: 1px solid #e5e7eb;
        text-align: center;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
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

# ── Session state initialization ─────────────────────────────────────────────
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = []
if "processing" not in st.session_state:
    st.session_state.processing = False
if "api_key_configured" not in st.session_state:
    st.session_state.api_key_configured = False

# ── Sidebar for configuration ────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("---")
    
    # API Key configuration
    api_key = st.text_input(
        "Clé API Anthropic",
        type="password",
        help="Entrez votre clé API Anthropic pour activer l'extraction",
        placeholder="sk-ant-..."
    )
    
    if api_key:
        st.session_state.api_key_configured = True
        # Set environment variable for the API key
        import os
        os.environ["ANTHROPIC_API_KEY"] = api_key
        st.success("✅ Clé API configurée")
    else:
        st.session_state.api_key_configured = False
        st.warning("⚠️ Clé API requise pour l'extraction")
    
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    st.markdown("""
    1. Configurez votre clé API Anthropic
    2. Importez vos factures (PDF, PNG, JPG)
    3. Cliquez sur 'Extraire'
    4. Téléchargez le fichier Excel
    
    **Format supportés :**
    - PDF
    - PNG
    - JPG/JPEG
    """)
    
    st.markdown("---")
    st.markdown("### ℹ️ Informations")
    st.markdown(f"**Version:** 1.0.0")
    st.markdown(f"**Date:** {datetime.now().strftime('%d/%m/%Y')}")

# ── Helper: encode file to base64 ─────────────────────────────────────────────
def encode_file(file_bytes: bytes) -> str:
    """Encode des bytes en base64"""
    return base64.standard_b64encode(file_bytes).decode("utf-8")

# ── Helper: extract invoice info via Claude API ───────────────────────────────
def extract_invoice_info(file_bytes: bytes, mime: str, filename: str) -> dict:
    """
    Extrait les informations d'une facture via l'API Claude
    
    Args:
        file_bytes: Contenu du fichier en bytes
        mime: Type MIME du fichier
        filename: Nom du fichier
    
    Returns:
        Dictionnaire contenant les informations extraites
    """
    try:
        client = anthropic.Anthropic()
        
        prompt = """Analyse cette facture et extrais les informations suivantes en JSON UNIQUEMENT (pas de texte avant/après):
{
  "fournisseur": "nom du fournisseur/vendeur (entreprise ou personne)",
  "date": "date de la facture au format JJ/MM/AAAA",
  "commande": "numéro de commande ou bon de commande (PO number)",
  "bon_de_livraison": "numéro du bon de livraison (delivery note)",
  "numero_facture": "numéro de la facture (invoice number)",
  "montant_facture": "montant total TTC en chiffres uniquement (ex: 1500.00)"
}

Règles importantes:
- Si une information est absente ou illisible, mets null
- Pour le montant, extrais uniquement le nombre, sans symbole monétaire
- La date doit être au format JJ/MM/AAAA
- Le fournisseur doit être le nom de l'entreprise émettrice

Réponds UNIQUEMENT avec le JSON, rien d'autre."""

        # Préparation du contenu selon le type de fichier
        if mime == "application/pdf":
            content = [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": encode_file(file_bytes)
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
                        "data": encode_file(file_bytes)
                    }
                },
                {"type": "text", "text": prompt}
            ]

        # Appel à l'API Claude
        response = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=1000,
            temperature=0,
            messages=[{"role": "user", "content": content}]
        )

        # Extraction et nettoyage de la réponse
        raw = response.content[0].text.strip()
        
        # Nettoyage des marqueurs markdown
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        
        # Nettoyage des espaces et sauts de ligne excessifs
        raw = re.sub(r'\s+', ' ', raw)
        
        # Parsing JSON
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            # Tentative de correction si le JSON est mal formé
            # Recherche d'un pattern JSON dans le texte
            json_match = re.search(r'\{.*\}', raw, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group())
            else:
                raise

        # Validation et nettoyage des données
        expected_fields = ["fournisseur", "date", "commande", "bon_de_livraison", "numero_facture", "montant_facture"]
        for field in expected_fields:
            if field not in data:
                data[field] = None
            elif data[field] in ["null", "NULL", "None", ""]:
                data[field] = None
        
        # Ajout du nom du fichier
        data["fichier"] = filename
        
        return data
        
    except Exception as e:
        st.error(f"Erreur lors de l'extraction pour {filename}: {str(e)}")
        # Retourne un dictionnaire avec des valeurs null en cas d'erreur
        return {
            "fournisseur": None,
            "date": None,
            "commande": None,
            "bon_de_livraison": None,
            "numero_facture": None,
            "montant_facture": None,
            "fichier": filename,
            "erreur": str(e)
        }

# ── Helper: build Excel workbook ──────────────────────────────────────────────
def build_excel(records: list[dict]) -> bytes:
    """
    Construit un fichier Excel à partir des données extraites
    
    Args:
        records: Liste des dictionnaires contenant les données
    
    Returns:
        Bytes du fichier Excel
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Factures"

    # Couleurs et styles
    header_fill = PatternFill("solid", fgColor="1e3a8a")
    alt_fill = PatternFill("solid", fgColor="dbeafe")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="d1d5db"),
        right=Side(style="thin", color="d1d5db"),
        top=Side(style="thin", color="d1d5db"),
        bottom=Side(style="thin", color="d1d5db"),
    )

    # Définition des colonnes
    columns = [
        ("Fournisseur", "fournisseur", 25),
        ("Date", "date", 14),
        ("N° Commande", "commande", 18),
        ("Bon de Livraison", "bon_de_livraison", 20),
        ("N° Facture", "numero_facture", 20),
        ("Montant TTC", "montant_facture", 15),
        ("Fichier Source", "fichier", 30),
    ]

    # En-têtes
    for col_idx, (label, _, width) in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 30

    # Lignes de données
    montant_total = 0
    for row_idx, record in enumerate(records, 2):
        fill = alt_fill if row_idx % 2 == 0 else white_fill
        
        for col_idx, (_, key, _) in enumerate(columns, 1):
            value = record.get(key)
            
            # Traitement spécial pour le montant
            if key == "montant_facture" and value not in (None, "null", ""):
                try:
                    # Nettoyage et conversion du montant
                    if isinstance(value, str):
                        value = re.sub(r'[^\d.,-]', '', value)
                        value = value.replace(',', '.')
                        # Gestion des séparateurs de milliers
                        if '.' in value and value.count('.') > 1:
                            # Format européen (ex: 1.234,56 -> 1234.56)
                            value = value.replace('.', '')
                            value = value.replace(',', '.')
                    value = float(value)
                    montant_total += value
                except (ValueError, TypeError):
                    value = 0.0
            
            cell = ws.cell(row=row_idx, column=col_idx, 
                          value=value if value not in (None, "null", "") else "")
            cell.font = Font(name="Arial", size=10)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="left" if key != "montant_facture" else "right", 
                                      vertical="center")
            cell.border = thin_border
            
            if key == "montant_facture" and isinstance(value, (int, float)):
                cell.number_format = '#,##0.00 "MAD"'
        
        ws.row_dimensions[row_idx].height = 20

    # Ligne de total
    last_row = len(records) + 2
    total_label = ws.cell(row=last_row, column=1, value="TOTAL GÉNÉRAL")
    total_label.font = Font(bold=True, name="Arial", size=11)
    total_label.fill = PatternFill("solid", fgColor="fbbf24")
    total_label.border = thin_border
    
    total_cell = ws.cell(row=last_row, column=6, value=montant_total)
    total_cell.font = Font(bold=True, name="Arial", size=11)
    total_cell.number_format = '#,##0.00 "MAD"'
    total_cell.alignment = Alignment(horizontal="right")
    total_cell.fill = PatternFill("solid", fgColor="fbbf24")
    total_cell.border = thin_border

    # Fusion des cellules pour le label total
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)

    # Ajustement automatique des hauteurs
    for row in ws.rows:
        ws.row_dimensions[row[0].row].height = max(ws.row_dimensions[row[0].row].height or 0, 20)

    # Sauvegarde dans un buffer
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Interface principale ─────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown("### 📤 Importer des Factures")
    st.markdown("""
    <div class="info-box">
        <strong>Formats acceptés :</strong> PDF, PNG, JPG, JPEG<br>
        <strong>Taille max :</strong> 200MB par fichier
    </div>
    """, unsafe_allow_html=True)

    # Zone d'upload
    uploaded_files = st.file_uploader(
        "Choisissez vos fichiers",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        help="Glissez-déposez ou cliquez pour sélectionner des fichiers"
    )

    if uploaded_files:
        st.markdown(f"""
        <div class="success-box">
            ✅ {len(uploaded_files)} fichier(s) sélectionné(s)
        </div>
        """, unsafe_allow_html=True)
        
        # Aperçu des fichiers sélectionnés
        with st.expander("📋 Voir les fichiers sélectionnés"):
            for file in uploaded_files:
                st.text(f"📄 {file.name} ({(file.size / 1024):.1f} KB)")

    # Bouton d'extraction
    extract_btn = st.button(
        "🔍 Extraire les Informations", 
        disabled=not (uploaded_files and st.session_state.api_key_configured),
        use_container_width=True
    )

    # Zone d'export
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
        if not st.session_state.api_key_configured:
            st.error("⚠️ Veuillez configurer votre clé API Anthropic dans le menu latéral")
        else:
            # Traitement des fichiers
            progress_bar = st.progress(0)
            status_text = st.empty()
            newly_extracted = []

            mime_map = {
                "pdf": "application/pdf",
                "png": "image/png",
                "jpg": "image/jpeg",
                "jpeg": "image/jpeg",
            }

            for i, uploaded_file in enumerate(uploaded_files):
                # Détermination du type MIME
                file_extension = uploaded_file.name.rsplit(".", 1)[-1].lower()
                mime_type = mime_map.get(file_extension, "application/octet-stream")
                
                # Mise à jour du statut
                status_text.info(f"⏳ Traitement de **{uploaded_file.name}**... ({i+1}/{len(uploaded_files)})")
                
                try:
                    # Lecture et extraction
                    file_bytes = uploaded_file.read()
                    extracted_info = extract_invoice_info(file_bytes, mime_type, uploaded_file.name)
                    newly_extracted.append(extracted_info)
                    
                    # Message de succès
                    if "erreur" not in extracted_info:
                        st.markdown(f"""
                        <div class="success-box">
                            ✅ {uploaded_file.name} traité avec succès
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div class="error-box">
                            ❌ Erreur pour {uploaded_file.name}: {extracted_info['erreur']}
                        </div>
                        """, unsafe_allow_html=True)
                        
                except Exception as e:
                    st.error(f"❌ Erreur inattendue pour {uploaded_file.name}: {str(e)}")
                
                # Mise à jour de la progression
                progress_bar.progress((i + 1) / len(uploaded_files))

            # Ajout des nouvelles données à la session
            st.session_state.extracted_data.extend(newly_extracted)
            
            # Nettoyage des indicateurs de progression
            progress_bar.empty()
            status_text.empty()
            
            st.success(f"✅ Traitement terminé ! {len(newly_extracted)} fichier(s) traité(s)")
            st.rerun()

    # Affichage des données extraites
    if st.session_state.extracted_data:
        # Création du DataFrame
        df = pd.DataFrame(st.session_state.extracted_data)
        
        # Renommage des colonnes pour l'affichage
        column_mapping = {
            'fournisseur': 'Fournisseur',
            'date': 'Date',
            'commande': 'Commande',
            'bon_de_livraison': 'Bon Livraison',
            'numero_facture': 'N° Facture',
            'montant_facture': 'Montant (MAD)',
            'fichier': 'Fichier'
        }
        
        df_display = df.rename(columns=column_mapping)
        
        # Sélection des colonnes à afficher
        display_columns = ['Fournisseur', 'Date', 'N° Facture', 'Montant (MAD)', 'Fichier']
        available_columns = [col for col in display_columns if col in df_display.columns]
        
        # Affichage du DataFrame
        st.dataframe(
            df_display[available_columns],
            use_container_width=True,
            height=400,
            hide_index=True
        )

        # Métriques et KPIs
        st.markdown("---")
        col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

        # Calcul du total
        total_amount = 0
        valid_invoices = 0
        
        for record in st.session_state.extracted_data:
            montant = record.get("montant_facture")
            if montant and montant not in ("null", None, ""):
                try:
                    if isinstance(montant, str):
                        # Nettoyage de la chaîne
                        montant_clean = re.sub(r'[^\d.,-]', '', montant)
                        montant_clean = montant_clean.replace(',', '.')
                        # Gestion des séparateurs de milliers
                        if '.' in montant_clean and montant_clean.count('.') > 1:
                            montant_clean = montant_clean.replace('.', '')
                            montant_clean = montant_clean.replace(',', '.')
                        montant_float = float(montant_clean)
                    else:
                        montant_float = float(montant)
                    
                    total_amount += montant_float
                    valid_invoices += 1
                except (ValueError, TypeError):
                    pass

        # Affichage des métriques
        col_kpi1.metric(
            "📄 Factures traitées",
            len(st.session_state.extracted_data),
            delta=None
        )
        
        col_kpi2.metric(
            "💰 Montant total",
            f"{total_amount:,.2f} MAD",
            delta=None
        )
        
        col_kpi3.metric(
            "✅ Extractions réussies",
            f"{valid_invoices}/{len(st.session_state.extracted_data)}",
            delta=None
        )
        
        col_kpi4.metric(
            "📅 Dernière extraction",
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            delta=None
        )

        # Option pour voir les détails
        with st.expander("📋 Voir toutes les données détaillées"):
            st.dataframe(df_display, use_container_width=True, hide_index=True)

    else:
        # Message quand aucune donnée n'est disponible
        st.markdown("""
        <div style="text-align:center; padding:60px 20px; background:#f8fafc; border-radius:8px;">
            <div style="font-size:5rem; margin-bottom:20px;">🧾</div>
            <p style="font-size:1.2rem; color:#1e293b; margin-bottom:10px;">
                Aucune donnée extraite
            </p>
            <p style="color:#64748b;">
                Importez des factures dans le panneau de gauche<br>
                et cliquez sur <strong>Extraire</strong> pour commencer
            </p>
        </div>
        """, unsafe_allow_html=True)

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    <p>🧾 Extraction Automatique de Factures | Propulsé par Claude API et Streamlit</p>
    <p style="font-size:0.8rem; margin-top:5px;">© 2024 - Tous droits réservés</p>
</div>
""", unsafe_allow_html=True)

# ── Gestion des erreurs non capturées ───────────────────────────────────────
def handle_uncaught_exceptions(exc_type, exc_value, exc_traceback):
    """Gestionnaire d'exceptions global"""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    st.error(f"Une erreur inattendue s'est produite: {exc_value}")

sys.excepthook = handle_uncaught_exceptions

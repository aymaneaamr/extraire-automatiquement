import base64
import io
import os
import re
import shutil
import tempfile
from datetime import datetime

import dash
from dash import dcc, html, dash_table, Input, Output, State, callback, no_update
import dash_bootstrap_components as dbc
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image
import plotly.express as px  # optionnel, pour d'éventuels graphiques

# -------------------------------------------------------------------
# Vérification de Tesseract (affichage d'un avertissement si absent)
# -------------------------------------------------------------------
if shutil.which("tesseract") is None:
    tesseract_ok = False
    print("⚠️ Tesseract non trouvé. L'OCR ne fonctionnera pas.")
else:
    tesseract_ok = True
    pytesseract.pytesseract.tesseract_cmd = shutil.which("tesseract")

# -------------------------------------------------------------------
# Fonctions d'extraction
# -------------------------------------------------------------------
def extraire_texte_fichier(contenu_bytes, nom_fichier, lang='fra+eng'):
    """
    Extrait le texte d'un fichier (PDF ou image) en utilisant pdfplumber ou OCR.
    """
    ext = os.path.splitext(nom_fichier)[1].lower()
    texte = ""

    try:
        if ext == '.pdf':
            with pdfplumber.open(io.BytesIO(contenu_bytes)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        texte += page_text
                    else:
                        # OCR sur l'image de la page
                        pil_image = page.to_image(resolution=300).original
                        texte += pytesseract.image_to_string(pil_image, lang=lang)
        else:
            # Image (jpg, png, etc.)
            image = Image.open(io.BytesIO(contenu_bytes))
            texte = pytesseract.image_to_string(image, lang=lang)
    except Exception as e:
        texte = f"[ERREUR] {e}"
    return texte

def chercher_champ(texte, pattern):
    """
    Applique une expression régulière et retourne le premier groupe capturé
    ou la correspondance entière si aucun groupe.
    """
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
# Initialisation de l'application Dash
# -------------------------------------------------------------------
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server  # pour le déploiement

# Layout de l'application
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("📄 Extraction de factures et bons de livraison"), className="mt-4 mb-2")
    ]),

    dbc.Row([
        dbc.Col([
            html.H5("1. Paramètres d'extraction"),
            html.P("Expressions régulières (utilisez des parenthèses capturantes) :"),
            dbc.Label("Fournisseur"),
            dbc.Input(id="regex-fournisseur", value=r"(?:Fournisseur|Supplier|Vendor|Client)\s*[:\-]?\s*([A-Z][A-Z\s\-\.]+(?:S\.?A\.?|SARL|SAS)?)"),
            dbc.Label("Date"),
            dbc.Input(id="regex-date", value=r"(?:Date|Facture\s*du|Invoice\s*date)\s*[:\-]?\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})"),
            dbc.Label("N° commande"),
            dbc.Input(id="regex-commande", value=r"(?:Commande|Order|N°\s*Commande|PO\s*Number|Référence\s*commande)\s*[:\-]?\s*([A-Z0-9\-/]{5,})"),
            dbc.Label("N° bon de livraison"),
            dbc.Input(id="regex-bl", value=r"(?:BL|Bon\s*de\s*livraison|Delivery\s*note|N°\s*BL)\s*[:\-]?\s*([A-Z0-9\-/]{3,})"),
            dbc.Label("N° facture"),
            dbc.Input(id="regex-facture", value=r"(?:Facture|Invoice|N°\s*Facture|Invoice\s*Number)\s*[:\-]?\s*([A-Z0-9\-/]{3,})"),
            dbc.Label("Montant"),
            dbc.Input(id="regex-montant", value=r"(?:Total|Montant|Amount|TOTAL\s*TTC|Net\s*à\s*payer)\s*[:\-]?\s*([\d\s,\.]+\s*(?:€|EUR)?)"),
            html.Hr(),
            dbc.Label("Langue OCR"),
            dcc.Dropdown(
                id='lang-ocr',
                options=[
                    {'label': 'Français + Anglais', 'value': 'fra+eng'},
                    {'label': 'Anglais', 'value': 'eng'},
                    {'label': 'Français', 'value': 'fra'}
                ],
                value='fra+eng'
            ),
            html.Br(),
            dbc.Alert("⚠️ Tesseract n'est pas installé – l'OCR ne fonctionnera pas.", color="warning") if not tesseract_ok else None
        ], width=4),

        dbc.Col([
            html.H5("2. Fichiers à traiter"),
            dcc.Upload(
                id='upload-fichiers',
                children=html.Div([
                    'Glissez-déposez ou ',
                    html.A('sélectionnez des fichiers (PDF, images)')
                ]),
                style={
                    'width': '100%', 'height': '60px', 'lineHeight': '60px',
                    'borderWidth': '1px', 'borderStyle': 'dashed',
                    'borderRadius': '5px', 'textAlign': 'center', 'margin': '10px'
                },
                multiple=True
            ),
            html.Div(id='liste-fichiers'),
            html.Hr(),
            dbc.Button("🚀 Extraire et ajouter les données", id="btn-extraire", color="primary", className="mt-2"),
            dcc.Loading(id="loading", type="circle", children=html.Div(id="loading-output")),
            html.Hr(),
            html.H5("3. Données extraites"),
            dash_table.DataTable(
                id='table-donnees',
                columns=[
                    {"name": "Fournisseur", "id": "fournisseur", "editable": True},
                    {"name": "Date", "id": "date", "editable": True},
                    {"name": "Commande", "id": "commande", "editable": True},
                    {"name": "Bon de livraison", "id": "bon_de_livraison", "editable": True},
                    {"name": "N° Facture", "id": "numero_facture", "editable": True},
                    {"name": "Montant", "id": "montant_facture", "editable": True}
                ],
                data=[],
                editable=True,
                row_deletable=True,
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left'},
            ),
            html.Hr(),
            dbc.Button("📥 Télécharger Excel", id="btn-download", color="success", className="mr-2"),
            dcc.Download(id="download-excel"),
            dcc.Store(id='store-donnees', data=[])  # stockage des données en JSON
        ], width=8)
    ])
], fluid=True)

# -------------------------------------------------------------------
# Callbacks
# -------------------------------------------------------------------
@callback(
    Output('liste-fichiers', 'children'),
    Input('upload-fichiers', 'contents'),
    State('upload-fichiers', 'filename')
)
def afficher_fichiers(liste_contents, liste_noms):
    """Affiche la liste des fichiers sélectionnés"""
    if liste_noms is None:
        return ""
    return html.Ul([html.Li(f"{nom} ({len(contenu)} caractères)") for nom, contenu in zip(liste_noms, liste_contents)])

@callback(
    Output('store-donnees', 'data'),
    Output('loading-output', 'children'),
    Input('btn-extraire', 'n_clicks'),
    State('upload-fichiers', 'contents'),
    State('upload-fichiers', 'filename'),
    State('regex-fournisseur', 'value'),
    State('regex-date', 'value'),
    State('regex-commande', 'value'),
    State('regex-bl', 'value'),
    State('regex-facture', 'value'),
    State('regex-montant', 'value'),
    State('lang-ocr', 'value'),
    prevent_initial_call=True
)
def extraire_donnees(n_clicks, liste_contents, liste_noms,
                     pat_fourn, pat_date, pat_cmd, pat_bl, pat_fact, pat_mont, lang):
    if not liste_contents:
        return no_update, "Veuillez sélectionner des fichiers."

    nouvelles_lignes = []
    patterns = {
        "fournisseur": pat_fourn,
        "date": pat_date,
        "commande": pat_cmd,
        "bon_de_livraison": pat_bl,
        "numero_facture": pat_fact,
        "montant_facture": pat_mont
    }

    for content, filename in zip(liste_contents, liste_noms):
        # décoder le contenu base64
        content_type, content_string = content.split(',')
        decoded = base64.b64decode(content_string)

        # extraire le texte
        texte = extraire_texte_fichier(decoded, filename, lang)

        # appliquer les regex
        ligne = {}
        for champ, pattern in patterns.items():
            ligne[champ] = chercher_champ(texte, pattern)

        nouvelles_lignes.append(ligne)

    # Récupérer les données existantes depuis le store (s'il y en a)
    # Note : pour simplifier, on écrase les précédentes. Si on veut cumuler,
    # il faudrait utiliser un store existant et concaténer.
    return nouvelles_lignes, f"{len(nouvelles_lignes)} ligne(s) extraite(s)."

@callback(
    Output('table-donnees', 'data'),
    Input('store-donnees', 'data')
)
def mettre_a_jour_table(data):
    return data if data else []

@callback(
    Output('download-excel', 'data'),
    Input('btn-download', 'n_clicks'),
    State('table-donnees', 'data'),
    prevent_initial_call=True
)
def telecharger_excel(n_clicks, data):
    if not data:
        return dash.no_update
    df = pd.DataFrame(data)
    # Créer un fichier Excel en mémoire
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Factures')
    output.seek(0)
    return dcc.send_bytes(output.getvalue(), filename="factures_extraites.xlsx")

# -------------------------------------------------------------------
# Lancement en local
# -------------------------------------------------------------------
if __name__ == '__main__':
    app.run_server(debug=True)

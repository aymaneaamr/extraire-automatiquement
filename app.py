import streamlit as st
import cv2
import numpy as np
from collections import defaultdict
import pandas as pd
from datetime import datetime
import json
import base64
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import re
import sqlite3
import os
import pickle

# ==================== Configuration de la page ====================
st.set_page_config(
    page_title="Gestionnaire d'Inventaire Multi-Pièces",
    page_icon="📦",
    layout="wide"
)

# ==================== Fonctions de persistance SQLite ====================

def init_database():
    """Initialise la base de données SQLite"""
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    
    # Table des articles
    c.execute('''CREATE TABLE IF NOT EXISTS articles
                 (code TEXT PRIMARY KEY,
                  libelle TEXT,
                  emplacement TEXT,
                  date_creation TEXT)''')
    
    # Table des photos
    c.execute('''CREATE TABLE IF NOT EXISTS photos
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  code_article TEXT,
                  timestamp TEXT,
                  nb_pieces INTEGER,
                  image_originale TEXT,
                  image_analyse TEXT,
                  FOREIGN KEY (code_article) REFERENCES articles(code))''')
    
    conn.commit()
    conn.close()

def charger_donnees():
    """Charge les données depuis SQLite"""
    gestionnaire = GestionnairePieces()
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    
    # Charger les articles
    c.execute("SELECT code, libelle, emplacement, date_creation FROM articles")
    articles = c.fetchall()
    
    for code, libelle, emplacement, date_creation in articles:
        gestionnaire.articles[code] = {
            'libelle': libelle,
            'photos': [],
            'emplacement': emplacement,
            'date_creation': date_creation
        }
    
    # Charger les photos
    c.execute("SELECT code_article, timestamp, nb_pieces, image_originale, image_analyse, id FROM photos ORDER BY timestamp")
    photos = c.fetchall()
    
    for code_article, timestamp, nb_pieces, img_originale, img_analyse, photo_id in photos:
        if code_article in gestionnaire.articles:
            photo_data = {
                'timestamp': timestamp,
                'nb_pieces': nb_pieces,
                'image_originale': img_originale,
                'image_analyse': img_analyse,
                'id': len(gestionnaire.articles[code_article]['photos'])
            }
            gestionnaire.articles[code_article]['photos'].append(photo_data)
    
    conn.close()
    return gestionnaire

def sauvegarder_article(code, libelle, emplacement, date_creation):
    """Sauvegarde un article dans SQLite"""
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO articles (code, libelle, emplacement, date_creation) VALUES (?, ?, ?, ?)",
              (code, libelle, emplacement, date_creation))
    conn.commit()
    conn.close()

def sauvegarder_photo(code_article, timestamp, nb_pieces, image_originale, image_analyse):
    """Sauvegarde une photo dans SQLite"""
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    c.execute("INSERT INTO photos (code_article, timestamp, nb_pieces, image_originale, image_analyse) VALUES (?, ?, ?, ?, ?)",
              (code_article, timestamp, nb_pieces, image_originale, image_analyse))
    conn.commit()
    conn.close()

def supprimer_article_db(code):
    """Supprime un article et ses photos de la base"""
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    c.execute("DELETE FROM photos WHERE code_article = ?", (code,))
    c.execute("DELETE FROM articles WHERE code = ?", (code,))
    conn.commit()
    conn.close()

def supprimer_photo_db(photo_id):
    """Supprime une photo de la base"""
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    c.execute("DELETE FROM photos WHERE id = ?", (photo_id,))
    conn.commit()
    conn.close()

def get_photo_db_id(code_article, timestamp):
    """Récupère l'ID SQLite d'une photo à partir de son timestamp"""
    conn = sqlite3.connect('inventaire.db')
    c = conn.cursor()
    c.execute("SELECT id FROM photos WHERE code_article = ? AND timestamp = ?", (code_article, timestamp))
    result = c.fetchone()
    conn.close()
    return result[0] if result else None

# ==================== JavaScript pour confirmation avant actualisation ====================
def add_refresh_confirmation():
    has_data = 'true' if 'gestionnaire' in st.session_state and len(st.session_state.gestionnaire.articles) > 0 else 'false'
    refresh_html = f"""
    <div id="refresh-confirmation" style="display:none;"></div>
    <script>
    function hasData() {{
        return {has_data};
    }}
    window.addEventListener('beforeunload', function (e) {{
        if (hasData()) {{
            var confirmationMessage = '⚠️ Attention ! Si vous actualisez la page, toutes les données non exportées seront perdues.\\n\\nVoulez-vous vraiment continuer ?';
            e.returnValue = confirmationMessage;
            return confirmationMessage;
        }}
    }});
    document.addEventListener('keydown', function(e) {{
        if (hasData()) {{
            if (e.key === 'F5' || (e.ctrlKey && e.key === 'r') || (e.ctrlKey && e.key === 'R')) {{
                e.preventDefault();
                var confirmRefresh = confirm('⚠️ Attention ! Si vous actualisez la page, toutes les données non exportées seront perdues.\\n\\nVoulez-vous vraiment actualiser ?');
                if (confirmRefresh) {{
                    window.location.reload();
                }}
            }}
        }}
    }});
    setInterval(function() {{
        if (typeof hasData === 'function') {{
        }}
    }}, 1000);
    </script>
    """
    st.components.v1.html(refresh_html, height=0)

# CSS personnalisé
st.markdown("""
<style>
    .success-box {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #28a745;
        margin: 1rem 0;
    }
    .location-badge {
        background: #17a2b8;
        color: white;
        padding: 0.2rem 0.5rem;
        border-radius: 5px;
        font-size: 0.8rem;
        margin-left: 0.5rem;
    }
    .label-badge {
        background: #28a745;
        color: white;
        padding: 0.2rem 0.5rem;
        border-radius: 5px;
        font-size: 0.8rem;
        margin-left: 0.5rem;
    }
    .import-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px dashed #6c757d;
        margin: 1rem 0;
    }
    .selection-box {
        background: #fff3cd;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #ffc107;
        margin: 1rem 0;
    }
    .warning-box {
        background: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #ffc107;
        margin: 1rem 0;
        font-weight: bold;
    }
    .database-info {
        background: #d1ecf1;
        color: #0c5460;
        padding: 0.5rem;
        border-radius: 5px;
        border-left: 5px solid #17a2b8;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

class GestionnairePieces:
    def __init__(self):
        """Initialise le gestionnaire de pièces"""
        self.articles = {}  # Dictionnaire {code_article: {"libelle": "", "photos": [], "emplacement": ""}}
        self.reset_article_courant()
    
    def reset_article_courant(self):
        """Réinitialise l'article en cours de saisie"""
        self.article_courant = {
            'code': '',
            'libelle': '',
            'emplacement': '',
            'photos': [],
            'total_pieces': 0
        }
    
    def creer_nouvel_article(self, code_article, libelle="", emplacement=""):
        """Crée un nouvel article dans l'inventaire avec son libellé et emplacement"""
        if code_article and code_article not in self.articles:
            date_creation = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.articles[code_article] = {
                'libelle': libelle,
                'photos': [],
                'emplacement': emplacement,
                'date_creation': date_creation
            }
            # Sauvegarder dans la base de données
            sauvegarder_article(code_article, libelle, emplacement, date_creation)
            return True
        return False
    
    def nettoyer_articles_mal_importes(self):
        """Supprime les articles qui ont des libellés d'en-tête"""
        a_supprimer = []
        for code, data in self.articles.items():
            libelle = data.get('libelle', '').upper()
            if 'COLONNE' in libelle or 'CODE ARTICLE' in libelle or 'LIBELLÉ' in libelle or 'EMPLACEMENT' in libelle:
                a_supprimer.append(code)
        
        for code in a_supprimer:
            # Supprimer de la base de données
            supprimer_article_db(code)
            del self.articles[code]
        
        return len(a_supprimer)
    
    def importer_articles_excel(self, df, col_code, col_libelle, col_emplacement, skip_first_row=True):
        """Importe des articles avec sélection manuelle des colonnes - version complète"""
        articles_importes = 0
        articles_existants = 0
        erreurs = 0
        
        # Barre de progression simple
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Déterminer l'index de début (0 ou 1)
        start_idx = 1 if skip_first_row else 0
        total_lignes = len(df) - start_idx
        
        # Dictionnaire pour suivre les codes déjà vus (pour éviter les doublons dans le fichier)
        codes_vus = set()
        
        for index in range(start_idx, len(df)):
            # Mettre à jour la progression
            progression = (index - start_idx + 1) / total_lignes
            progress_bar.progress(progression)
            status_text.text(f"Import en cours... {index - start_idx + 1}/{total_lignes}")
            
            row = df.iloc[index]
            try:
                # Récupérer le code
                code_value = row[col_code]
                if pd.isna(code_value) or str(code_value).strip() == '':
                    continue
                
                code = str(code_value).strip()
                
                # Vérifier que le code n'est pas un en-tête de colonne
                if code.lower() in ['code article', 'code', 'article', 'réf', 'ref', '0', '1', '2', '3', '4']:
                    continue
                
                # Éviter les doublons dans le même fichier
                if code in codes_vus:
                    continue
                codes_vus.add(code)
                
                # Récupérer le libellé
                libelle = ""
                if col_libelle and col_libelle != "(Aucune)" and col_libelle in row.index:
                    libelle_value = row[col_libelle]
                    if pd.notna(libelle_value):
                        libelle = str(libelle_value).strip()
                        # Nettoyer les valeurs "None"
                        if libelle.lower() == 'none':
                            libelle = ""
                
                # Récupérer l'emplacement
                emplacement = ""
                if col_emplacement and col_emplacement != "(Aucune)" and col_emplacement in row.index:
                    emp_value = row[col_emplacement]
                    if pd.notna(emp_value):
                        emp_str = str(emp_value).strip()
                        if emp_str.lower() not in ['none', 'nan', '']:
                            emplacement = emp_str
                
                # Créer l'article
                if code and code not in self.articles:
                    date_creation = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.articles[code] = {
                        'libelle': libelle,
                        'photos': [],
                        'emplacement': emplacement,
                        'date_creation': date_creation
                    }
                    # Sauvegarder dans la base de données
                    sauvegarder_article(code, libelle, emplacement, date_creation)
                    articles_importes += 1
                elif code in self.articles:
                    articles_existants += 1
                        
            except Exception as e:
                erreurs += 1
                continue
        
        # Nettoyer les éléments de progression
        progress_bar.empty()
        status_text.empty()
        
        return articles_importes, articles_existants, erreurs
    
    def ajouter_photo_article(self, code_article, frame_original, frame_analyse, nb_pieces):
        """Ajoute une photo analysée à un article existant"""
        if code_article in self.articles:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Convertir les images en base64
            _, buffer_original = cv2.imencode('.jpg', frame_original)
            _, buffer_analyse = cv2.imencode('.jpg', frame_analyse)
            
            img_originale_b64 = base64.b64encode(buffer_original).decode('utf-8')
            img_analyse_b64 = base64.b64encode(buffer_analyse).decode('utf-8')
            
            photo_data = {
                'timestamp': timestamp,
                'nb_pieces': nb_pieces,
                'image_originale': img_originale_b64,
                'image_analyse': img_analyse_b64,
                'id': len(self.articles[code_article]['photos'])
            }
            
            self.articles[code_article]['photos'].append(photo_data)
            
            # Sauvegarder dans la base de données
            sauvegarder_photo(code_article, timestamp, nb_pieces, img_originale_b64, img_analyse_b64)
            
            return True
        return False
    
    def get_total_article(self, code_article):
        """Retourne le total de pièces pour un article donné"""
        if code_article in self.articles:
            return sum(photo['nb_pieces'] for photo in self.articles[code_article]['photos'])
        return 0
    
    def get_photos_article(self, code_article):
        """Retourne toutes les photos d'un article"""
        if code_article in self.articles:
            return self.articles[code_article]['photos']
        return []
    
    def get_emplacement_article(self, code_article):
        """Retourne l'emplacement d'un article"""
        if code_article in self.articles:
            return self.articles[code_article].get('emplacement', '')
        return ''
    
    def get_libelle_article(self, code_article):
        """Retourne le libellé d'un article"""
        if code_article in self.articles:
            return self.articles[code_article].get('libelle', '')
        return ''
    
    def supprimer_photo(self, code_article, photo_id):
        """Supprime une photo d'un article"""
        if code_article in self.articles and 0 <= photo_id < len(self.articles[code_article]['photos']):
            # Récupérer le timestamp pour trouver l'ID SQLite
            timestamp = self.articles[code_article]['photos'][photo_id]['timestamp']
            db_id = get_photo_db_id(code_article, timestamp)
            
            # Supprimer de la base de données
            if db_id:
                supprimer_photo_db(db_id)
            
            # Supprimer de la mémoire
            del self.articles[code_article]['photos'][photo_id]
            
            # Réindexer les IDs
            for i, photo in enumerate(self.articles[code_article]['photos']):
                photo['id'] = i
            return True
        return False
    
    def supprimer_article(self, code_article):
        """Supprime complètement un article"""
        if code_article in self.articles:
            # Supprimer de la base de données
            supprimer_article_db(code_article)
            
            # Supprimer de la mémoire
            del self.articles[code_article]
            return True
        return False
    
    def get_tous_les_totaux(self):
        """Retourne un dictionnaire avec tous les totaux par article"""
        return {code: self.get_total_article(code) for code in self.articles}
    
    def get_tous_emplacements(self):
        """Retourne un dictionnaire avec tous les emplacements par article"""
        return {code: self.get_emplacement_article(code) for code in self.articles}
    
    def get_tous_libelles(self):
        """Retourne un dictionnaire avec tous les libellés par article"""
        return {code: self.get_libelle_article(code) for code in self.articles}
    
    def generer_excel(self):
        """Génère un fichier Excel avec l'inventaire complet"""
        # Créer un nouveau classeur Excel
        output = BytesIO()
        workbook = openpyxl.Workbook()
        
        # Feuille principale - Résumé
        sheet_resume = workbook.active
        sheet_resume.title = "Inventaire"
        
        # En-têtes
        headers = ["Code Article", "Libellé", "Emplacement", "Quantité totale", "Nombre de photos", "Dernière mise à jour"]
        for col, header in enumerate(headers, 1):
            cell = sheet_resume.cell(row=1, column=col)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center")
        
        # Données du résumé
        row = 2
        for code_article, data in self.articles.items():
            total = sum(p['nb_pieces'] for p in data['photos'])
            nb_photos = len(data['photos'])
            derniere_date = data['photos'][-1]['timestamp'] if data['photos'] else data.get('date_creation', 'N/A')
            emplacement = data.get('emplacement', '')
            libelle = data.get('libelle', '')
            
            sheet_resume.cell(row=row, column=1).value = code_article
            sheet_resume.cell(row=row, column=2).value = libelle
            sheet_resume.cell(row=row, column=3).value = emplacement
            sheet_resume.cell(row=row, column=4).value = total
            sheet_resume.cell(row=row, column=5).value = nb_photos
            sheet_resume.cell(row=row, column=6).value = derniere_date
            row += 1
        
        # Ajuster la largeur des colonnes
        sheet_resume.column_dimensions['A'].width = 20
        sheet_resume.column_dimensions['B'].width = 40
        sheet_resume.column_dimensions['C'].width = 20
        sheet_resume.column_dimensions['D'].width = 15
        sheet_resume.column_dimensions['E'].width = 15
        sheet_resume.column_dimensions['F'].width = 22
        
        # Feuille de détail
        sheet_detail = workbook.create_sheet("Détail des photos")
        
        # En-têtes détail
        detail_headers = ["Code Article", "Libellé", "Emplacement", "Photo #", "Date", "Nombre de pièces"]
        for col, header in enumerate(detail_headers, 1):
            cell = sheet_detail.cell(row=1, column=col)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        # Données détaillées
        row = 2
        for code_article, data in self.articles.items():
            libelle = data.get('libelle', '')
            emplacement = data.get('emplacement', '')
            for i, photo in enumerate(data['photos'], 1):
                sheet_detail.cell(row=row, column=1).value = code_article
                sheet_detail.cell(row=row, column=2).value = libelle
                sheet_detail.cell(row=row, column=3).value = emplacement
                sheet_detail.cell(row=row, column=4).value = f"Photo {i}"
                sheet_detail.cell(row=row, column=5).value = photo['timestamp']
                sheet_detail.cell(row=row, column=6).value = photo['nb_pieces']
                row += 1
        
        # Ajuster les colonnes du détail
        sheet_detail.column_dimensions['A'].width = 20
        sheet_detail.column_dimensions['B'].width = 40
        sheet_detail.column_dimensions['C'].width = 20
        sheet_detail.column_dimensions['D'].width = 12
        sheet_detail.column_dimensions['E'].width = 22
        sheet_detail.column_dimensions['F'].width = 18
        
        workbook.save(output)
        output.seek(0)
        return output
    
    def reinitialiser_tout(self):
        """Réinitialise complètement l'inventaire"""
        # Supprimer le fichier de base de données
        if os.path.exists('inventaire.db'):
            os.remove('inventaire.db')
        self.articles = {}

# Fonction pour détecter les pièces dans une image
def detecter_pieces(image):
    """Détecte et compte les pièces dans une image"""
    resultat = image.copy()
    
    # Conversion en niveaux de gris
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Flou pour réduire le bruit
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    
    # Détection des contours
    edges = cv2.Canny(blur, 50, 150)
    
    # Dilatation et érosion
    kernel = np.ones((3, 3), np.uint8)
    edges = cv2.dilate(edges, kernel, iterations=2)
    edges = cv2.erode(edges, kernel, iterations=1)
    
    # Trouver les contours
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Filtrer les petits contours (bruit)
    pieces_valides = []
    for contour in contours:
        aire = cv2.contourArea(contour)
        if aire > 200:  # Seuil minimum
            pieces_valides.append(contour)
    
    nb_pieces = len(pieces_valides)
    
    # Dessiner les contours
    for contour in pieces_valides:
        # Dessiner le contour en vert
        cv2.drawContours(resultat, [contour], -1, (0, 255, 0), 2)
        
        # Ajouter un point au centre
        M = cv2.moments(contour)
        if M["m00"] != 0:
            cx = int(M["m10"] / M["m00"])
            cy = int(M["m01"] / M["m00"])
            cv2.circle(resultat, (cx, cy), 3, (0, 0, 255), -1)
    
    # Ajouter le compteur
    cv2.putText(resultat, f"Pieces: {nb_pieces}", (10, 30),
                cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
    
    return resultat, nb_pieces

# Nouvelle fonction pour recadrer l'image au format 4/3
def recadrer_4_3(image):
    """Recadre l'image pour obtenir un ratio 4:3 (largeur/hauteur) en conservant le centre."""
    h, w = image.shape[:2]
    ratio_cible = 4.0 / 3.0
    ratio_actuel = w / h
    if abs(ratio_actuel - ratio_cible) < 0.01:
        return image  # déjà proche de 4:3
    if ratio_actuel > ratio_cible:
        # Image trop large : on recadre horizontalement
        nouvelle_largeur = int(h * ratio_cible)
        debut_x = (w - nouvelle_largeur) // 2
        return image[:, debut_x:debut_x+nouvelle_largeur]
    else:
        # Image trop haute : on recadre verticalement
        nouvelle_hauteur = int(w / ratio_cible)
        debut_y = (h - nouvelle_hauteur) // 2
        return image[debut_y:debut_y+nouvelle_hauteur, :]

# Fonction pour décoder l'image base64
def base64_to_image(base64_string):
    img_data = base64.b64decode(base64_string)
    nparr = np.frombuffer(img_data, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    return img

# Initialisation de la base de données
init_database()

# Initialisation des états
if 'gestionnaire' not in st.session_state:
    st.session_state.gestionnaire = charger_donnees()
if 'page' not in st.session_state:
    st.session_state.page = "saisie"
if 'article_selectionne' not in st.session_state:
    st.session_state.article_selectionne = None
if 'photo_selectionnee' not in st.session_state:
    st.session_state.photo_selectionnee = None
if 'show_import' not in st.session_state:
    st.session_state.show_import = False
if 'photo_temp' not in st.session_state:
    st.session_state.photo_temp = None
if 'ajout_photo' not in st.session_state:
    st.session_state.ajout_photo = False
if 'search_query' not in st.session_state:
    st.session_state.search_query = ""

gestionnaire = st.session_state.gestionnaire

# Ajouter la confirmation d'actualisation
add_refresh_confirmation()

# Afficher un avertissement si des données sont présentes
if len(gestionnaire.articles) > 0:
    st.markdown("""
    <div class="warning-box">
        ⚠️ <strong>Attention :</strong> Les données sont stockées temporairement. 
        Pensez à exporter votre inventaire en Excel avant de quitter ou d'actualiser la page !
    </div>
    """, unsafe_allow_html=True)

# Afficher l'information de persistance (modifiée)
st.markdown("""
<div class="database-info">
    💾 <strong>Persistance active :</strong> Les données sont automatiquement sauvegardées
</div>
""", unsafe_allow_html=True)

# Interface principale
st.title("📦 Gestionnaire d'Inventaire Multi-Pièces")
st.markdown("""
Cette application permet de gérer l'inventaire de plusieurs types de pièces :
1. **Importer** un fichier Excel avec vos articles (code, libellé, emplacement)
2. **Ajouter** plusieurs photos pour chaque article (avec possibilité d'ajuster le comptage)
3. **Exporter** un fichier Excel avec tous les totaux
""")

# Barre latérale avec la liste des articles
with st.sidebar:
    st.header("📋 Articles en inventaire")
    
    # UN SEUL bouton pour importer Excel (toujours visible)
    if st.button("📥 Importer des articles Excel", use_container_width=True):
        st.session_state.show_import = True
        st.rerun()
    
    if gestionnaire.articles:
        st.write(f"**{len(gestionnaire.articles)} articles**")
        
        # Bouton pour nettoyer les articles mal importés
        if st.button("🧹 Nettoyer les articles mal importés", use_container_width=True):
            nb_supprimes = gestionnaire.nettoyer_articles_mal_importes()
            if nb_supprimes > 0:
                st.success(f"✅ {nb_supprimes} articles supprimés")
                st.rerun()
            else:
                st.info("Aucun article à nettoyer")
        
        # ---- Champ de recherche ----
        search_query = st.text_input(
            "🔍 Rechercher un article",
            value=st.session_state.search_query,
            placeholder="Code, libellé ou emplacement...",
            key="search_input"
        ).lower().strip()
        st.session_state.search_query = search_query
        
        # Filtrer les articles
        codes_filtres = []
        if search_query:
            for code, data in gestionnaire.articles.items():
                libelle = data.get('libelle', '').lower()
                emplacement = data.get('emplacement', '').lower()
                if (search_query in code.lower() or 
                    search_query in libelle or 
                    search_query in emplacement):
                    codes_filtres.append(code)
        else:
            codes_filtres = list(gestionnaire.articles.keys())
        
        codes_filtres.sort()
        
        if not codes_filtres:
            st.info("Aucun article ne correspond à votre recherche")
        
        # Afficher les articles filtrés
        for code_article in codes_filtres:
            total = gestionnaire.get_total_article(code_article)
            libelle = gestionnaire.get_libelle_article(code_article)
            emplacement = gestionnaire.get_emplacement_article(code_article)
            
            with st.container():
                col1, col2 = st.columns([3, 1])
                with col1:
                    if st.button(f"📦 {code_article}", key=f"select_{code_article}", use_container_width=True):
                        st.session_state.article_selectionne = code_article
                        st.session_state.page = "details"
                        st.rerun()
                with col2:
                    st.write(f"**{total}**")
            
            # Afficher les badges séparément
            if libelle or emplacement:
                badge_text = ""
                if libelle:
                    badge_text += f"📝 {libelle[:30]}{'...' if len(libelle) > 30 else ''}"
                if libelle and emplacement:
                    badge_text += " | "
                if emplacement:
                    badge_text += f"📍 {emplacement}"
                
                if badge_text:
                    st.caption(badge_text)
        
        st.divider()
        
        # Bouton pour retourner à la saisie
        if st.button("➕ Nouvel article manuel", use_container_width=True):
            st.session_state.page = "saisie"
            st.session_state.article_selectionne = None
            st.rerun()
        
        st.divider()
        
        # Export Excel
        if gestionnaire.articles:
            st.header("📊 Export")
            excel_file = gestionnaire.generer_excel()
            st.download_button(
                label="📥 Télécharger Excel",
                data=excel_file,
                file_name=f"inventaire_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Réinitialisation
            if st.button("🔄 Tout réinitialiser", type="primary", use_container_width=True):
                gestionnaire.reinitialiser_tout()
                st.session_state.page = "saisie"
                st.session_state.article_selectionne = None
                st.rerun()
    else:
        st.info("Aucun article pour le moment")

# Section d'import Excel
if st.session_state.show_import:
    st.markdown("---")
    st.markdown('<div class="import-section">', unsafe_allow_html=True)
    st.header("📥 Importer des articles depuis Excel")
    st.markdown("""
    ### Sélectionnez les colonnes correspondantes :
    Choisissez quelle colonne de votre fichier correspond à chaque information.
    """)
    
    uploaded_excel = st.file_uploader("Choisir un fichier Excel", type=['xlsx', 'xls'], key="import_excel")
    
    if uploaded_excel:
        try:
            # Lire le fichier Excel
            df = pd.read_excel(uploaded_excel)
            
            # Afficher un aperçu du fichier original
            st.subheader("Aperçu du fichier original")
            st.dataframe(df.head(10))  # Afficher les 10 premières lignes
            
            # Sélection des colonnes
            cols = df.columns.tolist()
            
            # Trouver automatiquement les bonnes colonnes
            default_code_index = 0
            default_libelle_index = 1
            default_emplacement_index = 2
            
            for i, col in enumerate(cols):
                col_lower = col.lower()
                if 'emplacement' in col_lower:
                    default_emplacement_index = i + 1  # +1 car on a "(Aucune)" en première position
            
            col1, col2, col3 = st.columns(3)
            with col1:
                col_code = st.selectbox("📌 Colonne pour CODE article *", cols, index=default_code_index)
            with col2:
                col_libelle = st.selectbox("📝 Colonne pour LIBELLÉ *", ["(Aucune)"] + cols, index=default_libelle_index + 1)
            with col3:
                col_emplacement = st.selectbox("📍 Colonne pour EMPLACEMENT (optionnel)", ["(Aucune)"] + cols, index=default_emplacement_index)
            
            # Option pour ignorer la première ligne
            skip_first = st.checkbox("Ignorer la première ligne (en-têtes)", value=True, 
                                   help="Cochez cette case si votre fichier contient des en-têtes de colonnes")
            
            # Afficher un aperçu des données à importer (TOUTES les lignes)
            st.subheader("Aperçu des données à importer :")
            
            # Aperçu en ignorant ou non la première ligne
            start_preview = 1 if skip_first else 0
            preview_data = {}
            
            # Code - prendre toutes les lignes
            preview_data['Code'] = df[col_code].iloc[start_preview:].values
            
            # Libellé
            if col_libelle != "(Aucune)":
                preview_data['Libellé'] = df[col_libelle].iloc[start_preview:].values
            
            # Emplacement
            if col_emplacement != "(Aucune)":
                preview_data['Emplacement'] = df[col_emplacement].iloc[start_preview:].values
            
            # Créer le DataFrame d'aperçu
            apercu = pd.DataFrame(preview_data)
            st.dataframe(apercu)
            
            # Statistiques détaillées
            total_lignes = len(df) - (1 if skip_first else 0)
            codes_non_vides = df[col_code].iloc[start_preview:].notna().sum()
            codes_uniques = df[col_code].iloc[start_preview:].nunique()
            
            col_s1, col_s2, col_s3, col_s4 = st.columns(4)
            with col_s1:
                st.metric("📊 Lignes totales", total_lignes)
            with col_s2:
                st.metric("✅ Codes valides", codes_non_vides)
            with col_s3:
                st.metric("🆔 Codes uniques", codes_uniques)
            with col_s4:
                st.metric("📝 Articles actuels", len(gestionnaire.articles))
            
            # Bouton d'import
            if st.button("✅ Confirmer l'import", use_container_width=True, type="primary"):
                with st.spinner("Import en cours..."):
                    # Convertir "(Aucune)" en None
                    col_lib = col_libelle if col_libelle != "(Aucune)" else None
                    col_emp = col_emplacement if col_emplacement != "(Aucune)" else None
                    
                    importes, existants, erreurs = gestionnaire.importer_articles_excel(df, col_code, col_lib, col_emp, skip_first)
                    
                    # Afficher le résultat
                    st.markdown("---")
                    st.subheader("📊 Résultat de l'import")
                    
                    col_r1, col_r2, col_r3, col_r4 = st.columns(4)
                    with col_r1:
                        st.metric("✅ Importés", importes)
                    with col_r2:
                        st.metric("⚠️ Déjà existants", existants)
                    with col_r3:
                        st.metric("❌ Erreurs", erreurs)
                    with col_r4:
                        st.metric("📊 Total après import", len(gestionnaire.articles))
                    
                    if importes > 0:
                        st.success(f"✅ {importes} articles importés avec succès !")
                        st.balloons()
                        
                        # FERMER AUTOMATIQUEMENT l'import
                        st.session_state.show_import = False
                        st.rerun()
                    else:
                        st.warning("⚠️ Aucun article n'a été importé. Vérifiez que :")
                        st.warning("   - La colonne CODE contient bien des valeurs")
                        st.warning("   - Les codes ne sont pas déjà dans l'inventaire")
                        st.warning("   - Vous avez bien sélectionné les bonnes colonnes")
        
        except Exception as e:
            st.error(f"❌ Erreur lors de la lecture du fichier : {str(e)}")
    
    # Bouton pour fermer manuellement
    if st.button("❌ Fermer l'import", use_container_width=True):
        st.session_state.show_import = False
        st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("---")

# Contenu principal
if st.session_state.page == "saisie" and not st.session_state.show_import:
    # Page de saisie d'un nouvel article (sans scan)
    st.header("➕ Ajouter un nouvel article")
    
    # Formulaire de saisie manuelle
    st.markdown("### 📝 Informations de l'article")
    
    # Trois colonnes pour le code, le libellé et l'emplacement
    col_code, col_lib, col_emp = st.columns([2, 2, 1])
    
    with col_code:
        code_article = st.text_input(
            "Code article *",
            placeholder="Code article (obligatoire)",
            key="code_article_input"
        )
    
    with col_lib:
        libelle = st.text_input(
            "Libellé (optionnel)",
            value="",
            placeholder="Description de l'article",
            key="libelle_input"
        )
    
    with col_emp:
        emplacement = st.text_input(
            "Emplacement (optionnel)",
            value="",
            placeholder="Ex: A-12, Rayon 3...",
            key="emplacement_input"
        )
    
    st.caption("* Champ obligatoire")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Créer l'article", use_container_width=True):
            if code_article:
                if gestionnaire.creer_nouvel_article(code_article, libelle, emplacement):
                    st.success(f"✅ Article '{code_article}' créé avec succès!")
                    if libelle:
                        st.info(f"📝 Libellé: {libelle}")
                    if emplacement:
                        st.info(f"📍 Emplacement: {emplacement}")
                    st.session_state.article_selectionne = code_article
                    st.session_state.page = "details"
                    st.rerun()
                else:
                    if code_article in gestionnaire.articles:
                        st.error("❌ Ce code article existe déjà")
                    else:
                        st.error("❌ Erreur lors de la création de l'article")
            else:
                st.error("❌ Veuillez entrer un code article")
    with col2:
        if st.button("❌ Annuler", use_container_width=True):
            st.rerun()

elif st.session_state.page == "details" and st.session_state.article_selectionne:
    # Page de détails d'un article
    code_article = st.session_state.article_selectionne
    photos = gestionnaire.get_photos_article(code_article)
    total = gestionnaire.get_total_article(code_article)
    libelle = gestionnaire.get_libelle_article(code_article)
    emplacement = gestionnaire.get_emplacement_article(code_article)
    
    # En-tête avec libellé et emplacement
    col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([2, 1, 1, 1, 1])
    with col_h1:
        st.header(f"📦 {code_article}")
        if libelle:
            st.markdown(f"<span class='label-badge'>📝 {libelle}</span>", unsafe_allow_html=True)
        if emplacement:
            st.markdown(f"<span class='location-badge'>📍 {emplacement}</span>", unsafe_allow_html=True)
    with col_h2:
        st.metric("Total pièces", total)
    with col_h3:
        st.metric("Photos", len(photos))
    with col_h4:
        if libelle:
            st.metric("Libellé", libelle[:20] + "..." if len(libelle) > 20 else libelle)
    with col_h5:
        if emplacement:
            st.metric("Emplacement", emplacement)
    
    # Afficher un badge si le code est un code-barres (juste pour info)
    if re.match(r'^[A-Z0-9-]+$', code_article):
        st.info(f"🔖 Code produit: {code_article}")
    
    # Options
    col_o1, col_o2, col_o3 = st.columns(3)
    with col_o1:
        if st.button("⬅️ Retour à la saisie", use_container_width=True):
            st.session_state.page = "saisie"
            st.rerun()
    with col_o2:
        if st.button("📸 Ajouter une photo", use_container_width=True):
            st.session_state.ajout_photo = True
            st.session_state.photo_temp = None
            st.rerun()
    with col_o3:
        if st.button("🗑️ Supprimer cet article", use_container_width=True, type="primary"):
            if gestionnaire.supprimer_article(code_article):
                st.success(f"✅ Article '{code_article}' supprimé")
                st.session_state.page = "saisie"
                st.rerun()
    
    st.divider()
    
    # Ajout de photo avec options de calcul
    if st.session_state.get('ajout_photo', False):
        st.subheader("📸 Ajouter une photo")
        
        col_p1, col_p2 = st.columns([2, 1])
        with col_p2:
            if st.button("❌ Annuler"):
                st.session_state.ajout_photo = False
                st.session_state.photo_temp = None
                st.rerun()
        
        with col_p1:
            source = st.radio("Source", ["📸 Prendre une photo", "🖼️ Choisir une image"], horizontal=True, key="photo_source")
        
        # Gestion de la capture/upload
        img_file = None
        if source == "📸 Prendre une photo":
            img_file = st.camera_input("Prendre une photo", key="camera_photo")
        else:
            img_file = st.file_uploader("Choisir une image", type=['jpg', 'jpeg', 'png'], key="upload_photo")
        
        # Si une nouvelle image est fournie, on l'analyse et on stocke temporairement
        if img_file is not None:
            with st.spinner("Analyse de l'image..."):
                bytes_data = img_file.getvalue()
                frame = cv2.imdecode(np.frombuffer(bytes_data, np.uint8), cv2.IMREAD_COLOR)
                # Appliquer le recadrage 4/3
                frame = recadrer_4_3(frame)
                resultat, nb_pieces = detecter_pieces(frame)
                st.session_state.photo_temp = {
                    'original': frame,
                    'analyse': resultat,
                    'detected': nb_pieces
                }
        
        # Si des données temporaires existent, on affiche l'aperçu et les options
        if st.session_state.photo_temp is not None:
            temp = st.session_state.photo_temp
            st.image(cv2.cvtColor(temp['analyse'], cv2.COLOR_BGR2RGB), 
                     caption=f"Analyse - {temp['detected']} pièces détectées", use_container_width=True)
            
            st.markdown("### Options de comptage")
            col_opt1, col_opt2, col_opt3 = st.columns(3)
            
            with col_opt1:
                operation = st.selectbox("Opération", 
                                         ["Utiliser détection", "Remplacer", "Additionner", "Multiplier"],
                                         index=0)
            with col_opt2:
                manuel = st.number_input("Valeur manuelle", min_value=0, value=0, step=1)
            with col_opt3:
                st.write("")  # espace
                st.write("")  # espace
                if st.button("✅ Ajouter cette photo", use_container_width=True):
                    # Calcul du nombre final selon l'opération
                    detected = temp['detected']
                    if operation == "Utiliser détection":
                        nb_final = detected
                    elif operation == "Remplacer":
                        nb_final = manuel if manuel > 0 else detected
                    elif operation == "Additionner":
                        nb_final = detected + manuel
                    elif operation == "Multiplier":
                        nb_final = detected * manuel if manuel > 0 else detected
                    else:
                        nb_final = detected
                    
                    # Ajouter la photo avec le nombre calculé
                    if gestionnaire.ajouter_photo_article(code_article, temp['original'], temp['analyse'], nb_final):
                        st.success(f"✅ Photo ajoutée avec {nb_final} pièces!")
                        st.session_state.ajout_photo = False
                        st.session_state.photo_temp = None
                        st.rerun()
    
    # Affichage des photos existantes
    if photos:
        st.subheader("📸 Photos enregistrées")
        
        # Options d'affichage
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            tri = st.selectbox("Trier par", ["Plus récente", "Plus ancienne", "Plus de pièces", "Moins de pièces"])
        
        # Trier les photos
        photos_affichees = photos.copy()
        if tri == "Plus récente":
            photos_affichees = list(reversed(photos_affichees))
        elif tri == "Plus ancienne":
            photos_affichees = photos_affichees
        elif tri == "Plus de pièces":
            photos_affichees = sorted(photos_affichees, key=lambda x: x['nb_pieces'], reverse=True)
        elif tri == "Moins de pièces":
            photos_affichees = sorted(photos_affichees, key=lambda x: x['nb_pieces'])
        
        # Afficher les photos en grille
        cols = st.columns(3)
        for i, photo in enumerate(photos_affichees):
            with cols[i % 3]:
                # Afficher la miniature
                img = base64_to_image(photo['image_analyse'])
                img_mini = cv2.resize(img, (200, 150))
                st.image(cv2.cvtColor(img_mini, cv2.COLOR_BGR2RGB), use_column_width=True)
                
                # Informations
                st.caption(f"📅 {photo['timestamp'][:10]}")
                st.caption(f"🔢 {photo['nb_pieces']} pièces")
                
                # Boutons
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    if st.button("🔍 Voir", key=f"view_{code_article}_{i}"):
                        st.session_state.photo_selectionnee = photo['id']
                        st.session_state.page = "photo_detail"
                        st.rerun()
                with col_b2:
                    if st.button("🗑️", key=f"del_{code_article}_{i}"):
                        if gestionnaire.supprimer_photo(code_article, photo['id']):
                            st.rerun()
    
    else:
        st.info("📸 Aucune photo pour cet article. Cliquez sur 'Ajouter une photo' pour commencer.")

elif st.session_state.page == "photo_detail" and st.session_state.article_selectionne and st.session_state.photo_selectionnee is not None:
    # Détail d'une photo spécifique
    code_article = st.session_state.article_selectionne
    photos = gestionnaire.get_photos_article(code_article)
    photo_id = st.session_state.photo_selectionnee
    
    if 0 <= photo_id < len(photos):
        photo = photos[photo_id]
        libelle = gestionnaire.get_libelle_article(code_article)
        
        st.header(f"🔍 Détail de la photo - {code_article}")
        if libelle:
            st.subheader(libelle)
        
        # Afficher les deux images
        col_img1, col_img2 = st.columns(2)
        
        with col_img1:
            st.subheader("📸 Image originale")
            img_originale = base64_to_image(photo['image_originale'])
            st.image(cv2.cvtColor(img_originale, cv2.COLOR_BGR2RGB), use_column_width=True)
        
        with col_img2:
            st.subheader(f"🔍 Analyse - {photo['nb_pieces']} pièces")
            img_analyse = base64_to_image(photo['image_analyse'])
            st.image(cv2.cvtColor(img_analyse, cv2.COLOR_BGR2RGB), use_column_width=True)
        
        # Informations
        st.metric("Nombre de pièces", photo['nb_pieces'])
        st.caption(f"Date: {photo['timestamp']}")
        
        # Boutons
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("⬅️ Retour à l'article", use_container_width=True):
                st.session_state.page = "details"
                st.session_state.photo_selectionnee = None
                st.rerun()
        with col_b2:
            if st.button("🗑️ Supprimer cette photo", use_container_width=True, type="primary"):
                if gestionnaire.supprimer_photo(code_article, photo_id):
                    st.session_state.page = "details"
                    st.session_state.photo_selectionnee = None
                    st.rerun()
    else:
        st.error("Photo non trouvée")
        if st.button("Retour"):
            st.session_state.page = "details"
            st.session_state.photo_selectionnee = None
            st.rerun()

# Pied de page (modifié)
st.markdown("---")
col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
with col_f1:
    st.caption("📦 Gestionnaire d'Inventaire")
with col_f2:
    total_global = sum(gestionnaire.get_tous_les_totaux().values())
    st.caption(f"🧩 Total global: {total_global} pièces")
with col_f3:
    st.caption(f"📊 Articles: {len(gestionnaire.articles)}")
with col_f4:
    emplacements_renseignes = sum(1 for e in gestionnaire.get_tous_emplacements().values() if e)
    st.caption(f"📍 Emplacements: {emplacements_renseignes}/{len(gestionnaire.articles)}")
with col_f5:
    libelles_renseignes = sum(1 for l in gestionnaire.get_tous_libelles().values() if l)
    st.caption(f"📝 Libellés: {libelles_renseignes}/{len(gestionnaire.articles)}")

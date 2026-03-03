import pdfplumber
import re

def extraire_texte_pdf(fichier):
    """Extrait tout le texte d'un PDF uploadé (fichier binaire)"""
    texte = ""
    with pdfplumber.open(fichier) as pdf:
        for page in pdf.pages:
            texte += page.extract_text() or ""
    return texte

def chercher_champ(texte, pattern, groupe=1, flags=re.IGNORECASE):
    """
    Cherche un motif dans le texte et retourne le groupe capturé.
    Si le motif n'est pas trouvé, retourne une chaîne vide.
    """
    match = re.search(pattern, texte, flags)
    if match:
        return match.group(groupe).strip()
    return ""

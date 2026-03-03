import pdfplumber
import re

def extraire_texte_pdf(fichier):
    """Extrait tout le texte d'un fichier PDF uploadé (objet binaire)"""
    texte = ""
    with pdfplumber.open(fichier) as pdf:
        for page in pdf.pages:
            texte += page.extract_text() or ""
    return texte

def chercher_champ(texte, pattern, groupe=1, flags=re.IGNORECASE):
    """
    Cherche un motif dans le texte et retourne le groupe capturé.
    Si le motif est trouvé mais que le groupe demandé n'existe pas,
    retourne la correspondance entière (groupe 0).
    Si le motif n'est pas trouvé, retourne une chaîne vide.
    """
    match = re.search(pattern, texte, flags)
    if match:
        # Vérifier si le groupe demandé existe parmi les groupes capturés
        if groupe <= len(match.groups()):
            return match.group(groupe).strip()
        else:
            # Fallback : retourner la chaîne entière correspondante
            return match.group(0).strip()
    return ""

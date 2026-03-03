import pdfplumber
import pytesseract
from PIL import Image
import io
import re
from pdf2image import convert_from_bytes

def extraire_texte_pdf_ou_image(fichier, extension):
    """
    Extrait le texte d'un fichier uploadé.
    - Si c'est un PDF textuel : utilise pdfplumber.
    - Si c'est un PDF scanné ou une image : utilise OCR (Tesseract).
    """
    texte = ""
    if extension.lower() == 'pdf':
        # Essayer d'abord l'extraction textuelle
        with pdfplumber.open(fichier) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    texte += page_text
                else:
                    # Si pas de texte, faire OCR sur la page
                    pil_image = page.to_image(resolution=300).original
                    texte += pytesseract.image_to_string(pil_image, lang='fra+eng')
    else:
        # Image (jpg, png, etc.)
        image = Image.open(fichier)
        texte = pytesseract.image_to_string(image, lang='fra+eng')
    return texte

def chercher_champ(texte, pattern, groupe=1, flags=re.IGNORECASE):
    """Même fonction que précédemment"""
    match = re.search(pattern, texte, flags)
    if match:
        if groupe <= len(match.groups()):
            return match.group(groupe).strip()
        else:
            return match.group(0).strip()
    return ""

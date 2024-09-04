import pytesseract
import cv2
import pdf2image
import numpy as np
from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams
import re

# Configura el OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\P108\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def ocr_image(image):
    text = pytesseract.image_to_string(image, lang='spa')
    return text

def extract_from_pdf(pdf_path):
    # Extraer texto ya presente en el PDF
    text = extract_text(pdf_path, laparams=LAParams())
    
    # Convertir PDF a imágenes
    poppler_path = r'C:\poppler\bin'  # Asegúrate de que esta ruta sea correcta
    try:
        pages = pdf2image.convert_from_path(pdf_path, poppler_path=poppler_path)
    except Exception as e:
        print(f'Error al convertir el PDF a imágenes: {e}')
        return ""
    
    ocr_text = ""
    
    for page in pages:
        # Aplicar OCR en cada página
        open_cv_image = cv2.cvtColor(np.array(page), cv2.COLOR_RGB2BGR)
        ocr_text += ocr_image(open_cv_image)
    
    full_text = text + ocr_text
    return full_text

def extract_info(text):
    # Extraer NIC
    nic = re.search(r'NIC:\s*(\d+)', text)
    nic = nic.group(1) if nic else "No encontrado"
    
    # Extraer Total a pagar mes
    total_pagar = re.search(r'Total a pagar mes:\s*\$([\d,.]+)', text)
    total_pagar = total_pagar.group(1) if total_pagar else "No encontrado"
    
    # Extraer No. Facturas vencidas
    facturas_vencidas = re.search(r'No. Facturas vencidas:\s*(\d+)', text)
    facturas_vencidas = facturas_vencidas.group(1) if facturas_vencidas else "No encontrado"
    
    # Extraer Periodo facturado
    periodo_facturado = re.search(r'Periodo facturado\s*([\d/]+)\s*a\s*([\d/]+)', text)
    periodo_facturado = periodo_facturado.group(0) if periodo_facturado else "No encontrado"
    
    # Extraer Energía
    energia = re.search(r'Consumo\s*([\d,.]+)', text)
    energia = energia.group(1) if energia else "No encontrado"

    # Repetir para otros valores como Aseo, Alumbrado Público, etc.
    
    return {
        "NIC": nic,
        "Total a pagar mes": total_pagar,
        "No. Facturas vencidas": facturas_vencidas,
        "Periodo facturado": periodo_facturado,
        "Consumo Energía": energia,
        # Agregar otros valores aquí
    }

pdf_path = r'C:\Users\P108\Documents\ARA_PY\afinia_model.pdf'
text = extract_from_pdf(pdf_path)
info = extract_info(text)

for key, value in info.items():
    print(f"{key}: {value}")

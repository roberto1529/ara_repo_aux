# import os
# import time
# import logging
# import multiprocessing as mp
# import fitz  # PyMuPDF
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from commons.commons import start_logging, read_excel_electrohuila, process_error
# from funciones import generar_arbol_carpetas

# # Configuración del logging
# logger = start_logging("LOGS_EHUILA", mode="dev")
# comercializadora = 'Electrohuila'
# def configurar_navegador():
#     """
#     Configura el navegador Chrome para la descarga del PDF.
#     """
#     options = Options()
#     options.add_argument("--no-sandbox")
#     options.add_argument("--disable-dev-shm-usage")
#     options.add_argument("--disable-gpu")
#     options.add_argument("--disable-software-rasterizer")
#     options.add_argument("--disable-web-security")
#     options.add_argument("--allow-running-insecure-content")
#     options.add_argument("--disable-blink-features=AutomationControlled")
#     options.add_argument("--disable-extensions")
#     options.add_argument("--disable-popup-blocking")
#     options.add_argument("--disable-infobars")
#     options.add_argument("--ignore-certificate-errors")
#     options.add_argument("--allow-insecure-localhost")
#     options.add_argument("--disable-features=InsecureDownloadWarnings")
#     options.add_argument("--window-size=800x600")

#     # Configuración de preferencias de descarga
#     download_path = f'C:\\Users\\P108\\Documents\\PyDocto\\{comercializadora}'
#     prefs = {
#         "download.prompt_for_download": False,
#         "download.directory_upgrade": True,
#         "safebrowsing.enabled": True,
#         "plugins.always_open_pdf_externally": True,
#         "profile.managed_default_content_settings.images": 2,  # Bloquear imágenes
#         "download.default_directory": download_path
#     }
#     options.add_experimental_option("prefs", prefs)
    
#     return options, download_path

# def descargar_pdf(contrato):
#     """
#     Usa Selenium para descargar el PDF del contrato.
#     """
#     options, download_path = configurar_navegador()
#     service = Service('chromedriver.exe')
#     driver = webdriver.Chrome(service=service, options=options)

#     logging.info("Navegador Chrome iniciado")

#     try:
#         # Carga la página local
#         driver.get('https://enlinea.electrohuila.com.co/generate-invoice/')
#         logging.info("Página cargada")

#         # Espera a que el campo de matrícula esté presente y visible
#         input_element = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.ID, 'txtAccount'))
#         )
#         conStr = str(contrato)
#         input_element.send_keys(conStr)
#         logging.info(f"Texto escrito en el input: {conStr}")

#         # Espera a que el botón de consultar esté presente y visible
#         consult_button = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.ID, 'btnSearch'))
#         )
#         consult_button.click()
#         logging.info("Botón 'Consultar' clicado")

#         # Espera a que el PDF se cargue completamente dentro del iframe
#         time.sleep(15)  # Ajusta este tiempo según lo que tarde en cargar

#         # Obtén el src del iframe y redirige a esa URL para descargar el PDF
#         iframe = driver.find_element(By.ID, 'iframeInvoice')
#         pdf_url = iframe.get_attribute('src')
#         driver.get(pdf_url)
#         logging.info(f"Redirigido al PDF en: {pdf_url}")

#         # Espera un tiempo para asegurar que el PDF se descargue
#         time.sleep(10)

#     except Exception as e:
#         logging.error(f"Ocurrió un error durante la descarga del PDF: {str(e)}")
#     finally:
#         # Cierra el navegador
#         driver.quit()

#     return os.path.join(download_path, f'{contrato}.pdf')

# def leer_pdf(pdf_path):
#     """
#     Lee el contenido del PDF y busca el texto después de "Período".
#     Detiene la búsqueda después de encontrar el primer resultado.
#     """
#     logging.info(f"Buscando el archivo PDF en: {pdf_path}")

#     if not os.path.exists(pdf_path):
#         logging.error(f"El archivo {pdf_path} no existe.")
#         return None

#     try:
#         # Abre el archivo PDF
#         pdf_document = fitz.open(pdf_path)
#         logging.info("Archivo PDF abierto")

#         # Itera sobre las páginas del PDF
#         for page_num in range(len(pdf_document)):
#             page = pdf_document.load_page(page_num)
#             text = page.get_text()
#             logging.info(f"Texto extraído de la página {page_num + 1}:\n{text}")

#             # Busca el texto "Período" y extrae el texto que sigue
#             periodo_index = text.find("Período")
#             if periodo_index != -1:
#                 start_index = periodo_index + len("Período")
#                 # Extrae el texto que sigue a "Período" (ajusta el rango según el formato del PDF)
#                 extracted_text = text[start_index:start_index + 100].strip()  # Ajusta el rango según sea necesario
#                 logging.info(f"Texto encontrado después de 'Período': {extracted_text}")
#                 return extracted_text  # Retorna el texto encontrado y sale del bucle
                
#             else:
#                 logging.info("Texto 'Período' no encontrado en la página.")

#         return None  # Si no se encuentra el texto "Período"

#     except Exception as e:
#         logging.error(f"Ocurrió un error al procesar el PDF: {str(e)}")
#         return None
#     finally:
#         # Cierra el documento PDF
#         pdf_document.close()

# def mover_pdf(pdf_path, destino):
#     """
#     Mueve el archivo PDF a la nueva ubicación especificada por destino.
#     """
#     try:
#         os.makedirs(os.path.dirname(destino), exist_ok=True)
#         os.rename(pdf_path, destino)
#         logging.info(f"PDF movido a: {destino}")
#     except Exception as e:
#         logging.error(f"No se pudo mover el archivo: {str(e)}")

# def procesar_facturas(contrato):
#     """
#     Procesa un contrato descargando el PDF, extrayendo información relevante,
#     generando una estructura de carpetas y moviendo el archivo PDF.
#     """
#     pdf_path = descargar_pdf(contrato)
#     texto_encontrado = leer_pdf(pdf_path)

#     if texto_encontrado:
#         # Genera la estructura de carpetas basada en el texto encontrado
#         nueva_ruta = generar_arbol_carpetas(texto_encontrado,comercializadora)
#         destino = os.path.join(nueva_ruta, f'{contrato}.pdf')
        
#         # Mueve el PDF a la nueva ubicación
#         mover_pdf(pdf_path, destino)
#     else:
#         logging.warning(f"No se encontró texto relevante en el PDF para el contrato {contrato}")

# def download_contratos():
#     """
#     Recolecta todos los contratos y ejecuta la función 'procesar_facturas' usando multiprocessing.
#     """
#     df = read_excel_electrohuila()
#     contratos = df["CONTRATO"].tolist()

#     with mp.Pool(processes=1) as pool:
#         pool.map(procesar_facturas, contratos)

# if __name__ == "__main__":
#     try:
#         download_contratos()
#     except Exception as e:
#         process_error("warning")
#         logger.error(f"Ocurrió un error al cargar la aplicación: {e}")


import os
import time
import logging
import multiprocessing as mp
import fitz  # PyMuPDF
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from commons.commons import start_logging, read_excel_electrohuila, process_error
from funciones import generar_arbol_carpetas

# Configuración del logging
logger = start_logging("LOGS_EHUILA", mode="dev")
comercializadora = 'Electrohuila'

def configurar_navegador():
    options = Options()
    options.add_argument("--headless")  # Ejecuta en modo headless para ahorrar recursos
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=800x600")

    download_path = f'C:\\Users\\P108\\Documents\\PyDocto\\{comercializadora}'
    prefs = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.managed_default_content_settings.images": 2,
        "download.default_directory": download_path
    }
    options.add_experimental_option("prefs", prefs)
    
    return options, download_path

def descargar_pdf(contrato):
    options, download_path = configurar_navegador()
    service = Service('chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    logging.info("Navegador Chrome iniciado")

    try:
        driver.get('https://enlinea.electrohuila.com.co/generate-invoice/')
        input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'txtAccount'))
        )
        input_element.send_keys(str(contrato))
        logging.info(f"Texto escrito en el input: {contrato}")

        consult_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btnSearch'))
        )
        consult_button.click()
        logging.info("Botón 'Consultar' clicado")

        time.sleep(15)

        iframe = driver.find_element(By.ID, 'iframeInvoice')
        pdf_url = iframe.get_attribute('src')
        driver.get(pdf_url)
        logging.info(f"Redirigido al PDF en: {pdf_url}")

        time.sleep(10)

    except Exception as e:
        logging.error(f"Ocurrió un error durante la descarga del PDF: {str(e)}")
    finally:
        driver.quit()

    return os.path.join(download_path, f'{contrato}.pdf')

def leer_pdf(pdf_path):
    logging.info(f"Buscando el archivo PDF en: {pdf_path}")
    if not os.path.exists(pdf_path):
        logging.error(f"El archivo {pdf_path} no existe.")
        return None

    try:
        pdf_document = fitz.open(pdf_path)
        logging.info("Archivo PDF abierto")

        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()
            logging.info(f"Texto extraído de la página {page_num + 1}:\n{text}")

            periodo_index = text.find("Período")
            if periodo_index != -1:
                start_index = periodo_index + len("Período")
                extracted_text = text[start_index:start_index + 100].strip()
                logging.info(f"Texto encontrado después de 'Período': {extracted_text}")
                return extracted_text
                
            else:
                logging.info("Texto 'Período' no encontrado en la página.")

        return None

    except Exception as e:
        logging.error(f"Ocurrió un error al procesar el PDF: {str(e)}")
        return None
    finally:
        pdf_document.close()

def mover_pdf(pdf_path, destino):
    try:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        os.rename(pdf_path, destino)
        logging.info(f"PDF movido a: {destino}")
    except Exception as e:
        logging.error(f"No se pudo mover el archivo: {str(e)}")

def procesar_facturas(contrato):
    pdf_path = descargar_pdf(contrato)
    texto_encontrado = leer_pdf(pdf_path)

    if texto_encontrado:
        nueva_ruta = generar_arbol_carpetas(texto_encontrado, comercializadora)
        destino = os.path.join(nueva_ruta, f'{contrato}.pdf')
        mover_pdf(pdf_path, destino)
    else:
        logging.warning(f"No se encontró texto relevante en el PDF para el contrato {contrato}")

def download_contratos():
    df = read_excel_electrohuila()
    contratos = df["CONTRATO"].tolist()
    
    with mp.Pool(processes=mp.cpu_count()) as pool:
        pool.map(procesar_facturas, contratos)

if __name__ == "__main__":
    try:
        download_contratos()
    except Exception as e:
        process_error("warning")
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")

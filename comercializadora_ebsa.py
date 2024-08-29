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
from commons.commons import start_logging, read_excel_ebsa, read_excel_homologacion
from funciones import generar_arbol_carpetas, renombrar_pdf, conexion_correo, mover_pdf, eliminar_archivo_con_motivo, verificar_descarga, registrar_descarga

# Configuración del logging
logger = start_logging("LOGS_EBSA", mode="dev")
comercializadora = 'EBSA'


def obtener_mes_numero(mes_texto):
    """
    Convierte el nombre del mes en formato de texto a su número correspondiente.
    Soporta nombres completos y abreviaturas en español, tanto en mayúsculas como en minúsculas.
    """
    meses = {
        'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04', 'mayo': '05',
        'junio': '06', 'julio': '07', 'agosto': '08', 'septiembre': '09', 'octubre': '10',
        'noviembre': '11', 'diciembre': '12',
        'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05',
        'jun': '06', 'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10',
        'nov': '11', 'dic': '12',
        '01': '01', '02': '02', '03': '03', '04': '04', '05': '05',
        '06': '06', '07': '07', '08': '08', '09': '09', '10': '10',
        '11': '11', '12': '12'
    }
    return meses.get(mes_texto.lower(), None)


def configurar_navegador():
    """
    Configura el navegador Chrome para la descarga del PDF.
    """
    options = Options()
    download_path = f'C:\\Users\\P108\\Documents\\PyDocto\\{comercializadora}'
    prefs = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.managed_default_content_settings.images": 2,  # Bloquear imágenes
        "download.default_directory": download_path
    }
    options.add_experimental_option("prefs", prefs)
    
    return options, download_path

def descargar_pdf(contrato):
    """
    Usa Selenium para descargar el PDF del contrato si aún no ha sido descargado.
    """
    options, download_path = configurar_navegador()
    service = Service('chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    log_file = os.path.join(download_path, 'descargas.log')
    pdf_path = os.path.join(download_path, 'mpdf.pdf')

    if verificar_descarga(contrato['CONTRATO'], download_path):
        logging.info(f"El archivo para el contrato {contrato['CONTRATO']} ya ha sido descargado anteriormente.")
        driver.quit()  # Asegúrate de cerrar el navegador si no se descarga el PDF
        return pdf_path

    logging.info("Navegador Chrome iniciado")

    try:
        # Carga la página local
        driver.get('https://factura.ebsa.com.co/')
        logging.info("Página cargada")

        # Interacción con la página...
        input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ndc'))
        )
        input_element.send_keys(str(contrato['CONTRATO']))
        logging.info(f"Texto escrito en el input: {contrato['CONTRATO']}")

        # Clic en el botón de consulta...
        consult_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[text()='Consultar factura']"))
        )
        consult_button.click()
        logging.info("Botón 'Consultar factura' clicado")

        # Espera a que cargue el contenido del PDF...
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'invoice-area'))
        )

        # Encuentra y haz clic en el enlace para descargar el PDF
        pdf_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/factura/')]"))
        )
        pdf_url = pdf_link.get_attribute('href')
        driver.get(pdf_url)
        logging.info(f"Redirigido al PDF en: {pdf_url}")
        time.sleep(10)

    except Exception as e:
        logging.error(f"Ocurrió un error durante la descarga del PDF: {str(e)}")
    finally:
        driver.quit()

    return pdf_path

def leer_pdf(pdf_path):
    """
    Lee el contenido del PDF y busca el texto después de "PERIODO FACTURADO".
    """
    logging.info(f"Buscando el archivo PDF en: {pdf_path}")

    if not os.path.exists(pdf_path):
        eliminar_archivo_con_motivo(pdf_path, "Factura no encontrada")
        return None

    try:
        pdf_document = fitz.open(pdf_path)
        logging.info("Archivo PDF abierto")

        if len(pdf_document) == 0:
            raise Exception("El archivo PDF está vacío o dañado.")

        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()
            logging.info(f"Texto extraído de la página {page_num + 1}:\n{text}")

            periodo_index = text.find("PERIODO FACTURADO")
            if periodo_index != -1:
                start_index = periodo_index + len("PERIODO FACTURADO")
                extracted_text = text[start_index:start_index + 100].strip()
                logging.info(f"Texto encontrado después de 'PERIODO FACTURADO': {extracted_text}")
                return extracted_text  
            else:
                logging.info("Texto 'PERIODO FACTURADO' no encontrado en la página.")

        return None

    except Exception as e:
        eliminar_archivo_con_motivo(pdf_path, "Factura no encontrada")
        err = "Factura no encontrada: " + str(e)
        conexion_correo('Error de ejecución en el bot', err)
        return None
    finally:
        pdf_document.close()

def obtener_periodo_y_anio(texto_extraido):
    """
    Extrae el período y el año del texto encontrado en el PDF.
    """
    partes = texto_extraido.split(' A ')
    if len(partes) == 2:
        mes_anio_inicio = partes[0].strip().split('-')
        mes_anio_fin = partes[1].strip().split('-')

        if len(mes_anio_inicio) == 2 and len(mes_anio_fin) == 2:
            mes_inicio_texto = mes_anio_inicio[0]  # Ejemplo: 'AGO'
            anio_inicio = mes_anio_inicio[1]         # Ejemplo: '2024'
            
            mes_numero = obtener_mes_numero(mes_inicio_texto)
            if mes_numero:
                return mes_numero, anio_inicio
            else:
                logging.error(f"No se pudo convertir el mes '{mes_inicio_texto}' a número.")
        else:
            logging.error(f"No se encontró el formato esperado en el texto: {texto_extraido}")
    else:
        logging.error(f"El texto no tiene el formato esperado: {texto_extraido}")

    return None, None

def procesar_facturas(contrato):
    """
    Procesa un contrato descargando el PDF, extrayendo información relevante,
    renombrando el archivo y moviéndolo a la carpeta correspondiente.
    """
    pdf_path = descargar_pdf(contrato)
    if os.path.exists(pdf_path):
        texto_encontrado = leer_pdf(pdf_path)

        if texto_encontrado:
            logging.info(f"Texto encontrado en el PDF: {texto_encontrado}")
            mes_numero, anio = obtener_periodo_y_anio(texto_encontrado)
            if mes_numero and anio:
                nuevo_nombre = renombrar_pdf(contrato['SAP'], comercializadora, f"{mes_numero}-{anio}")
                logging.info(f"Nuevo nombre del archivo: {nuevo_nombre}")
                if nuevo_nombre:
                    nueva_ruta = generar_arbol_carpetas(f"{mes_numero}-{anio}", comercializadora)
                    destino = os.path.join(nueva_ruta, nuevo_nombre)
                    ruta_final = mover_pdf(pdf_path, destino)

                    if ruta_final:  # Verifica si el archivo fue movido exitosamente
                        log_file = os.path.join(nueva_ruta, 'descargas.log')
                        registrar_descarga(contrato['CONTRATO'], ruta_final, log_file)
            else:
                logging.error(f"No se pudo extraer mes o año del texto: {texto_encontrado}")
        else:
            logging.error(f"No se encontró texto relevante en el PDF para el contrato {contrato}")
    else:
        logging.error(f"El archivo PDF {pdf_path} no se descargó correctamente.")

def download_contratos():
    """
    Recolecta todos los contratos y ejecuta la función 'procesar_facturas' usando multiprocessing.
    """
    df = read_excel_ebsa()
    hm = read_excel_homologacion()
    dfm = df.merge(hm, on='Supplier', how='left')

    contratos = dfm.to_dict('records')  # Convertir DataFrame a lista de diccionarios

    with mp.Pool(processes=1) as pool:
        pool.starmap(procesar_facturas, [(contrato,) for contrato in contratos])

if __name__ == "__main__":
    try:
        download_contratos()
    except Exception as e:
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")
        err = "error de funcionamiento: " + str(e)
        conexion_correo('Error de ejecución en el bot', err)

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
from commons.commons import start_logging, read_excel_electrohuila, read_excel_homologacion
from funciones import generar_arbol_carpetas, renombrar_pdf, conexion_correo, mover_pdf, eliminar_archivo_con_motivo, verificar_descarga, registrar_descarga

# Configuración del logging
logger = start_logging("LOGS_EHUILA", mode="dev")
comercializadora = 'Electrohuila'

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
    pdf_path = os.path.join(download_path, f'{contrato["CONTRATO"]}.pdf')

    if verificar_descarga(contrato['CONTRATO'], download_path):
        logging.info(f"El archivo para el contrato {contrato['CONTRATO']} ya ha sido descargado anteriormente.")
        driver.quit()  # Asegúrate de cerrar el navegador si no se descarga el PDF
        return pdf_path

    logging.info("Navegador Chrome iniciado")

    try:
        # Carga la página local
        driver.get('https://enlinea.electrohuila.com.co/generate-invoice/')
        logging.info("Página cargada")

        # Interacción con la página...
        input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'txtAccount'))
        )
        conStr = str(contrato['CONTRATO'])
        input_element.send_keys(conStr)
        logging.info(f"Texto escrito en el input: {conStr}")

        # Clic en el botón de consulta...
        consult_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btnSearch'))
        )
        consult_button.click()
        logging.info("Botón 'Consultar' clicado")

        # Espera la carga del PDF y descarga...
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

    return pdf_path

def leer_pdf(pdf_path):
    """
    Lee el contenido del PDF y busca el texto después de "Período".
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
        eliminar_archivo_con_motivo(pdf_path, "Factura no encontrada")
        err = "Factura no encontrada: " + str(e)
        conexion_correo('Error de ejecución en el bot', err);
        return None
    finally:
        pdf_document.close()

def procesar_facturas(contrato):
    """
    Procesa un contrato descargando el PDF, extrayendo información relevante,
    renombrando el archivo y moviéndolo a la carpeta correspondiente.
    """
    pdf_path = descargar_pdf(contrato)
    if os.path.exists(pdf_path):
        texto_encontrado = leer_pdf(pdf_path)

        if texto_encontrado:
            nuevo_nombre = renombrar_pdf(contrato['SAP'], contrato['Comercializadora'], texto_encontrado)
            if nuevo_nombre:
                nueva_ruta = generar_arbol_carpetas(texto_encontrado, comercializadora)
                destino = os.path.join(nueva_ruta, nuevo_nombre)
                ruta_final = mover_pdf(pdf_path, destino)

                if ruta_final:  # Verifica si el archivo fue movido exitosamente
                    log_file = os.path.join(nueva_ruta, 'descargas.log')
                    registrar_descarga(contrato['CONTRATO'], ruta_final, log_file)
        else:
            # No se encontró texto relevante, pero no se registra el error
            print(f"No se encontró texto relevante en el PDF para el contrato {contrato}")
    else:
        # El archivo PDF no se descargó correctamente, pero no se registra el error
        print(f"El archivo PDF {pdf_path} no se descargó correctamente.")

def download_contratos():
    """
    Recolecta todos los contratos y ejecuta la función 'procesar_facturas' usando multiprocessing.
    """
    df = read_excel_electrohuila()
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
        conexion_correo('Error de ejecución en el bot', err);

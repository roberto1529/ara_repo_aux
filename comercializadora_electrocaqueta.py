import multiprocessing as mp
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from undetected_chromedriver import Chrome, ChromeOptions
from selenium.webdriver.common.keys import Keys

import win32com.client as win32
import datetime
import os
import traceback as tr
import time
import toml
import logging
import calendar
import shutil  
import requests 
import fitz
import re
import psycopg2

from commons.commons import start_logging, read_excel_electrocaqueta, process_error, send_email

# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('LOGS_ELECTROCAQUETA', mode='dev')

def send_email2(subject, body):
    """
        Se encarga de enviar una notificación vía correo electronico
        
        Entradas: 
            - subject: Hace referencia al asunto que llevará el correo.
            - body: Hace referencia al cuerpo que llevará el correo.
    """
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.To = config["EMAIL_SEND_AIRE"]["email_recept"]
        mail.Send()
        logger.info("Correo enviado exitosamente.")
    except Exception as e:
        logger.error(f"Error al enviar el correo: {e}")
        
        
def send_error_email(subject, body):
    """
        Envía un correo de error al cliente.
        
        Entradas:
            - subject: Asunto del correo donde muestra donde fue producido el error.
            - body: Cuerpo del correo donde generalmente va el mensaje del error.
    """
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.To = config["EMAIL_SEND_AIRE"]["email_recept"]
        mail.Send()
        logger.info("Correo enviado exitosamente.")
        
    except Exception as e:
        logger.error(f"Error al enviar el correo: {e}")
        
        


def extraer_periodo_facturado(pdf_path):
    """
    Extrae el período facturado de un archivo PDF.
    
    Entradas:
        - pdf_path: Ruta al archivo PDF.
        
    Salidas:
        - Un string con el período facturado en formato 'dd-mm-yy a dd-mm-yy'.
    """
    # Abrir el archivo PDF
    try:
        documento = fitz.open(pdf_path)
        
        # Iterar sobre cada página del documento PDF
        for pagina_num in range(documento.page_count):
            pagina = documento[pagina_num]
            texto = pagina.get_text()
            
            # Buscar el patrón de período facturado en el texto
            match = re.search(r"(\d{2}-\d{2}-\d{2})\s*a\s*(\d{2}-\d{2}-\d{2})", texto)
            if match:
                fecha_inicio = match.group(1)
                fecha_fin = match.group(2)
                
                return fecha_inicio, fecha_fin

        return None  # Si no se encuentra el período facturado

    except Exception as e:
        logger.error(f"Error al procesar el PDF {pdf_path}: {e}")
        return None
        
    
def es_factura_de_julio(fecha_inicio, fecha_fin):
    # print(f"Tipo de fecha_inicio: {type(fecha_inicio)}")  # Depuración
    # print(f"Tipo de fecha_fin: {type(fecha_fin)}") 
    
    fecha_inicio = datetime.datetime.strptime(fecha_inicio, "%d-%m-%y")
    fecha_fin = datetime.datetime.strptime(fecha_fin, "%d-%m-%y")
    
    # Definir los límites del mes de julio
    inicio_julio = datetime.datetime(fecha_inicio.year, 7, 1)
    fin_julio = datetime.datetime(fecha_fin.year, 7, 31)
    
    # Verificar si el periodo cae dentro de julio
    if fecha_inicio <= fin_julio and fecha_fin >= inicio_julio:
        return True
    return False




def process_contract(contrato, contador):
    """
    Se encarga de procesar cada contrato por el número de contrato (NIC).

    Entradas:
        - contrato: Hace referencia al número de contrato a procesar.
        - contador: Número de iteración o identificador único para el proceso.
    """
    carpeta_facturas_electrocaqueta = config["CARPETA_FACTURAS"]["carpeta_facturas_electrocaqueta"]
    
    # Crear un directorio de descarga único para este proceso
    directorio_descarga_temporal = os.path.join(os.path.expanduser("~"), "Downloads", f"temp_download_electrocaqueta_{contador}")

    # Crear el directorio si no existe
    os.makedirs(directorio_descarga_temporal, exist_ok=True)

    options = ChromeOptions()

    preferences = {
        "download.default_directory": directorio_descarga_temporal,
        "directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_setting_values.notifications": 2,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.mixed_content": 1,
        "safebrowsing.enabled": True
    }

    options.add_experimental_option("prefs", preferences)
    options.add_argument("--log-level=3") 
    options.add_argument("--kiosk-printing")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-web-security")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-infobars")
    options.add_argument("--start-maximized")
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-insecure-localhost')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--disable-features=InsecureDownloadWarnings')
    options.add_argument("--unsafely-treat-insecure-origin-as-secure=http://exampleDomain1.com, http://200.21.186.74:85/")

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()

    try:
        logger.info(f"Iteración número: {contador}")
        
        logger.info("=" * 100)
        
        logger.info(f"Procesando contrato: {contrato}")
        
        driver.get("https://www.electrocaqueta.com.co/")
        
        logger.info(f"Ingresando a la página")

        original_window = driver.current_window_handle

        opcion_factura = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'FACTURA')]/ancestor::a")))
        
        time.sleep(3)
        
        driver.execute_script("arguments[0].scrollIntoView(true);", opcion_factura)
        
        time.sleep(3)
        
        opcion_factura.click()
        
        logger.info("Clic!")

        WebDriverWait(driver, 10).until(EC.new_window_is_opened)
        
        windows = driver.window_handles
        
        for window in windows:
            if window != original_window:
                driver.switch_to.window(window)
                break
        
        time.sleep(2)
        
        input_nic = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "form-control")))
        input_nic.send_keys(contrato)
        time.sleep(2)
        
        descargar_factura = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-primary"))).click()
        logging.info("Clic en descargar factura")
        
        driver.execute_script("window.print();")
        
        logging.info(f"Descargando factura...")
        
        pdfs = [f for f in os.listdir(directorio_descarga_temporal) if f.endswith('.pdf')]
        
        if pdfs:
            pdf_path = os.path.join(directorio_descarga_temporal, pdfs[-1])
            
            fecha_inicio, fecha_fin = extraer_periodo_facturado(pdf_path)
            
            #fecha_fin = datetime.datetime.strptime(fecha_fin, "%d-%m-%y")
            
            mes_facturado_pdf = f"{fecha_inicio} a {fecha_fin}"
            
            # mes = fecha_fin.month
            # año = fecha_fin.year
            
            mes = fecha_fin[3:5]  # Los caracteres en la posición 3 y 4 corresponden al mes
            año = fecha_fin[6:8]  # Los caracteres en la posición 6 y 7 corresponden al año

            # print(f"Mes: {mes}")
            # print(f"Año: {año}")
            
            nombre_comercializadora = "Electrocaquetá"
                        
            ruta_descarga = os.path.join(
                config["CARPETA_FACTURAS"]["ruta"], 
                nombre_comercializadora,
                f"20{año}",
                mes
            )
                        
            os.makedirs(ruta_descarga, exist_ok=True)
            
            if mes_facturado_pdf:
                logger.info(f"Período Facturado encontrado: {mes_facturado_pdf}")
            else:
                logger.info("No se encontró el Período Facturado en el PDF.")
                
            if es_factura_de_julio(fecha_inicio, fecha_fin):
                logger.info(f"Procesando factura con periodo {fecha_inicio} a {fecha_fin}")
                
                #Mover los archivos PDF a la carpeta de destino
                destino_final = os.path.join(ruta_descarga, f"{contrato}.pdf")
                shutil.move(pdf_path, destino_final)
                
                time.sleep(3)
                
                shutil.rmtree(directorio_descarga_temporal)
            else:
                logger.info(f"Factura con periodo {fecha_inicio} a {fecha_fin} no pertenece a Julio")

        else:
            logger.info("No se encontró ningún archivo PDF.")
        
        time.sleep(10)
    
    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
    
    finally:
        fin = datetime.datetime.now()
        contador = contador + 1
        #logger.info(f"Información guardada en la BD.")
        driver.quit()

def download_electrocaqueta():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la librería multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_electrocaqueta()
    contratos = df['CONTRATO'].tolist()
    
    # Comienza el proceso en paralelo con un identificador único para cada proceso
    with mp.Pool(processes=3) as pool:
        pool.starmap(process_contract, [(contrato, i) for i, contrato in enumerate(contratos, start=1)])

if __name__ == "__main__":
    try:
        download_electrocaqueta()
    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")

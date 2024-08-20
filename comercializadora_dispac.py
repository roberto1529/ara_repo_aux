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

from commons.commons import start_logging, read_excel_dispac, process_error, send_email

# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('LOGS_DISPAC', mode='dev')

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
        print("Correo enviado exitosamente.")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")
        
        
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

        
        
def numero_a_nombre_mes(numero_mes):
    """
        Convierte un número de mes al nombre correspondiente de dicho mes.
        
        Entrada: 
            - Número del mes en formato string (ej: "01" sería enero).
        
        Salida:
            - Retorna el nombre del mes.
    """
    
    meses = {
        '01': 'Enero',
        '02': 'Febrero',
        '03': 'Marzo',
        '04': 'Abril',
        '05': 'Mayo',
        '06': 'Junio',
        '07': 'Julio',
        '08': 'Agosto',
        '09': 'Septiembre',
        '10': 'Octubre',
        '11': 'Noviembre',
        '12': 'Diciembre'
    }
    
    return meses.get(numero_mes.zfill(2), 'Mes no válido')


def leer_pdf_y_verificar(pdf_path, periodo_a_consultar):
    """
        Lee un archivo PDF y verifica si contiene una mención del mes actual.
        
        Entradas: 
            - pdf_path: Ruta del archivo PDF a leer.
            - nombre_mes_actual: Corresponde al mes actual para verificar en el PDF.
            
        Salida:
            - Retorna el mes en dado caso que lo haya en contrado en el PDF, de lo contrario retorna un None.
    """
    
    try:
        doc = fitz.open(pdf_path)
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()
        doc.close()
        
        # Compila el patrón de expresión regular
        patron_fecha = re.compile(r"\b(?:Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre) de \d{4}\b")
        
        for linea in texto.splitlines():
            
            match = patron_fecha.search(linea)
            
            if match:
                
                periodo = match.group()
                
                logger.info(f"Periodo encontrado en la factura: {periodo}")
                
                # Verificar si el mes coincide con el mes actual
                logger.info(f"Periodo actual: {periodo_a_consultar}")
                
                if periodo == periodo_a_consultar:
                    print(f"Periodo encontrado en el PDF: {periodo}")
                    return periodo
                    
                print(f"Periodo no coincide: {periodo}")
                return None
            
            else:
                logger.warning(f"No se encontró una fecha en la línea: {linea}")
                
        return None
    
    except Exception as e:
        logger.warning(f"Error al leer el PDF {pdf_path}: {e}")
        return False
    


def connect_db():
    try:
        conection = psycopg2.connect(
            host = 'localhost',
            user = 'postgres',
            password='admin',
            database='adutoria_comercializadoras',
            port = 5432
        )
        
        logger.info(f"Conexión exitosa!")
        return conection
        
    except Exception as e:
        logger.error(f"Ocurrio un error: {e}")
    
    
def registrar_ejecucion(id, comercializadora, exito, inicio, fin, contrato):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO registros (id, comercializadora, fecha_inicio, fecha_fin, exito, contrato)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (id, comercializadora, inicio, fin, exito, contrato))
    conn.commit()
    cursor.close()
    conn.close()
    

def process_contract(contrato):
    """
        Se encarga de procesar cada contrato por el número de contrato (NIC).
        
        Entradas:
            - contrato: Hace referencia al número de contrato a procesar.
    """
    
    carpeta_facturas_dispac = config["CARPETA_FACTURAS"]["carpeta_facturas_dispac"]
    contador = 1
    
    options = ChromeOptions()
    
    preferences = {
        "download.default_directory": carpeta_facturas_dispac,
        "directory_upgrade": True,
        "safebrowsing.enabled": False,  # Desactiva el análisis de seguridad
        "safebrowsing.disable_download_protection": True,  # Desactiva la protección de descargas
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_setting_values.notifications": 2,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.mixed_content": 1, # Permite contenido mixto (HTTP en HTTPS)
        "safebrowsing.enabled": True
    }
    
    # Opciones de configuración para el navegador.
    options.add_experimental_option("prefs", preferences)
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
    options.add_argument("--unsafely-treat-insecure-origin-as-secure=http://exampleDomain1.com, http://apps.dispac.com.co/FacturaExpress/servlet/com.dispac.facturaexpress.invoiceexpress")
    
    #options.add_argument("--headless")
    
    inicio = datetime.datetime.now()
    exito = False

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()
    
    try:
        
        logger.info(f"Iteración número: {contador}")
        
        logger.info("=" * 100)
        
        logger.info(f"Procesando contrato: {contrato}")
        
        driver.get("http://apps.dispac.com.co/FacturaExpress/servlet/com.dispac.facturaexpress.invoiceexpress")
        
        logger.info(f"Ingresando a la página")
        
        time.sleep(2)

        input_nic = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "vCLICUE_SEARCH")))
        
        input_nic.send_keys(contrato)
        
        logger.info(f"Diligenciando número de NIC")
        
        descargar_factura = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "DOINVOICE"))
        ).click()
        
        logger.info("Consultando factura...")
        
        time.sleep(5)
        
        descargar_aceptar = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "DOCUMENTOEQUIVALENTE")) 
        )
        
        time.sleep(1)
        
        driver.execute_script("arguments[0].scrollIntoView(true);", descargar_aceptar)
        
        time.sleep(1)
        
        descargar_aceptar.click()
        
        logger.info("Clic en descargar factura")
        
        time.sleep(10)
        
        # Se obtienen todos los pdfs que se encuentran en una carpeta
        pdfs = [f for f in os.listdir(carpeta_facturas_dispac) if f.endswith('.pdf')]
        
        pdf_path = os.path.join(carpeta_facturas_dispac, pdfs[-1])
        
        periodo_actual = datetime.datetime.now().strftime('%Y/%m')
        
        mes_actual = datetime.datetime.now().strftime('%m')
        
        año_actual = datetime.datetime.now().strftime('%Y')
        
        nombre_mes_actual = numero_a_nombre_mes(mes_actual)
        
        nombre_mes_actual = "Junio"

        periodo_a_consultar = f"{nombre_mes_actual} de {año_actual}"
        
        logger.info(f"Periodo a Consultar: {periodo_a_consultar}")
        
        logger.info(f"Nombre del mes actual: {nombre_mes_actual}")
        
        if leer_pdf_y_verificar(pdf_path, periodo_a_consultar):
            
            logger.info(f"La factura: {pdf_path} es del mes actual")
        
        else:
            send_email2("PRUEBA ASUNTO DISPAC", f"El presente correo es para informar que la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {periodo_actual}")
            
            logger.info(f"la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {periodo_actual}")
            
            os.remove(pdf_path)
            
            logger.info(f"Eliminando la factura...")  
            
        time.sleep(130)
        
    except Exception as e:
        
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
        
    
    finally:
        
        fin = datetime.datetime.now()
        
        # Registrando en la base de datos el resultado de la ejecución.
        #registrar_ejecucion(contador, "Dispac", exito, inicio, fin, contrato)
        contador = contador + 1
        
        logger.info(f"Información guardada en la BD.")
        
def download_dispac():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la libreria multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_dispac()
    contratos = df['CONTRATO'].tolist()

    # Comienza el proceso en paralelo
    with mp.Pool(processes = 1) as pool:
        pool.map(process_contract, contratos)

if __name__ == "__main__":
    try:
        
        download_dispac()
    except Exception as e:
        
        process_error('warning')
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")
        
import multiprocessing as mp
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from undetected_chromedriver import Chrome, ChromeOptions

import win32com.client as win32
import datetime
import os
import traceback as tr
import time
import toml
import fitz  
import logging
import calendar
import shutil  # Import shutil for moving files
import requests
import re
import psycopg2

from commons.commons import start_logging, read_excel_energuaviare, process_error, send_email

# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('LOGS_EMSA', mode='dev')

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


def numero_a_nombre_mes(numero_mes):
    """
        Convierte un número de mes al nombre correspondiente de dicho mes.
        
        Entrada: 
            - Número del mes en formato string (ej: "01" sería enero).
        
        Salida:
            - Retorna el nombre del mes.
    """
    
    meses = {
        '01': 'ENE',
        '02': 'FEB',
        '03': 'MAR',
        '04': 'ABR',
        '05': 'MAY',
        '06': 'JUN',
        '07': 'JUL',
        '08': 'AGO',
        '09': 'SEP',
        '10': 'OCT',
        '11': 'NOV',
        '12': 'DIC'
    }
    
    return meses.get(numero_mes.zfill(2), 'Mes no válido')


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


def leer_pdf_y_verificar(pdf_path, nombre_mes_actual):
    """
        Lee un archivo PDF y verifica si contiene una mención del mes actual.
        
        Entradas: 
            - pdf_path: Ruta del archivo PDF a leer.
            - periodo_actual: Corresponde al periodo actual para verificar en el PDF.
            
        Salida:
            - Retorna el mes en dado caso que lo haya en contrado en el PDF, de lo contrario retorna un None.
    """
    
    try:
        doc = fitz.open(pdf_path)
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()
        doc.close()
        
        patron_fecha = re.compile(r'\b(\d{2})/([A-Z]{3})/(\d{4})\b')
        
        for linea in texto.splitlines():
            
            match = patron_fecha.search(linea)
            
            if match:
                
                dia, mes, ano = match.groups()
                
                logger.info(f"Fecha encontrada en el PDF: {dia}/{mes}/{ano}")
                
                # Verificar si el mes coincide con el mes actual
                logger.info(f"NOMBRE DEL MES ACTUAL: {nombre_mes_actual}")
                if mes == nombre_mes_actual:
                    
                    print(f"Mes encontrado en el PDF: {mes}")
                    
                    return mes
                    
                print(f"Mes no coincide: {mes}")
                
                return None
            
            else:
                logger.warning(f"No se encontró una fecha en la línea: {linea}")
                
        return None
    
    except Exception as e:
        logger.warning(f"Error al leer el PDF {pdf_path}: {e}")
        return False

def process_contract(contrato):
    """
        Se encarga de procesar cada contrato por el número de contrato (NIC).
        
        Entradas:
            - contrato: Hace referencia al número de contrato a procesar.
    """
    
    carpeta_facturas_energuaviare = config["CARPETA_FACTURAS"]["carpeta_facturas_energuaviare"]
    contador = 1
    
    options = ChromeOptions()
    
    preferences = {
        "download.default_directory": carpeta_facturas_energuaviare,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
        "useAutomationExtension": False,
        "profile.default_content_setting_values.notifications": 2,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
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
        
        # Inicia el proceso de procesamiento de facturas
        logger.info(f"Procesando contrato: {contrato}")
        
        driver.get("http://factura.energuaviare.com:82/factura/web/")
        
        # URL Comercializadora energuaviare.
        url = f'http://factura.energuaviare.com:82/factura/web/{contrato}'
        
        logger.info(f"Ingresando a la página")
        
        time.sleep(2)
        
        input_nic = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "numerodecuenta")))
        
        input_nic.send_keys(contrato)
        
        logger.info(f"Diligenciando número de NIC")
        
        time.sleep(3)
        
        # descargar_factura = WebDriverWait(driver, 10).until(
        #     EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn btn-secondary') and text()='Consultar factura']"))
        # ).click()
        
        # Haciendo una petición a la API para la descarga de facturas
        response = requests.get(url, stream=True, verify=False)
        
        # Verifica que la descarga se haya realizado correctamente y que además sea un PDF
        if response.status_code == 200 and 'application/pdf' in response.headers.get('Content-Type', ''):
            
            archivo_nombre = f'{contrato}.pdf'
            
            shutil.move(archivo_nombre, carpeta_facturas_energuaviare)
            
            logger.info(f"Factura movida exitosamente a: {carpeta_facturas_energuaviare}")
            
            with open(archivo_nombre, 'wb') as file:
                file.write(response.content)
            
            logger.info(f"Factura descargada exitosamente como: {archivo_nombre}")
            
        
        if response.status_code == 404 or response.status_code == 504 or response.status_code == 502 or response.status_code == 500 or response.status_code == 503:
            logger.info(f"La página no se  encuentra disponible en este momento.") 
            send_email2(f"HTTP CODE: {response.satus_code}", "La página no se encuentra disponible en este momento.")
            
        logger.info(f"Descargando factura...")
        
        time.sleep(10)
        
        # Se obtienen todos los pdfs que se encuentran en una carpeta
        pdfs = [f for f in os.listdir(carpeta_facturas_energuaviare) if f.endswith('.pdf')]
        
        # Path de cada pdf
        pdf_path = os.path.join(carpeta_facturas_energuaviare, pdfs[-1])
        
        periodo_actual = datetime.datetime.now().strftime('%Y/%m')
        
        mes_actual = datetime.datetime.now().strftime('%m')
        
        nombre_mes_actual = numero_a_nombre_mes(mes_actual)
        
        logger.info(f"AAAAAAAAAAAAAAAAAAAAAAA: {nombre_mes_actual}")
        
        if leer_pdf_y_verificar(pdf_path, nombre_mes_actual):
            
            # Aca se valida si es necesario el mes nuevamente.
            logger.info(f"La factura: {pdf_path} es del mes actual")
            
            #shutil.move(pdf_path, os.path.join(config["CARPETA_FACTURAS"]["carpeta_facturas_energuaviare"], os.path.basename(pdf_path)))
            
        else:
            
            send_email2("PRUEBA ASUNTO ENERGUAVIARE", f"El presente correo es para informar que la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {periodo_actual}")
            
            logger.info(f"la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {periodo_actual}")
            
            os.remove(pdf_path)
        
        logger.info("Proceso finalizado!")
        
    except Exception as e:
        
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
        
    finally:
        
        fin = datetime.datetime.now()
        
        # Registrando en la base de datos el resultado de la ejecución.
        registrar_ejecucion(contador, "Energuaviare", exito, inicio, fin, contrato)
        contador = contador + 1
        
        logger.info(f"Información guardada en la BD.")


def download_energuaviare():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la libreria multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_energuaviare()
    contratos = df['CONTRATO'].tolist()

    # Comienza el proceso en paralelo
    with mp.Pool(processes = 1) as pool:
        pool.map(process_contract, contratos)

if __name__ == "__main__":
    try:
        
        download_energuaviare()
    except Exception as e:
        
        process_error('warning')
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")
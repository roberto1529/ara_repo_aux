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

import win32com.client as win32
import datetime
import os
import traceback as tr
import time
import toml
import logging
import calendar
import shutil  # Import shutil for moving files
import requests  # Import requests to check page accessibility
import locale

from commons.commons import start_logging, read_excel_emsa, process_error, send_email

# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('LOGS_EMSA', mode='dev')

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

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
    
    return meses.get(numero_mes, 'Mes no válido')


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

def verificar_status_pagina(url, timeout=30):
    """
        Verifica si la página se encuentra accesible para la descarga de facturas.
        
        Entrada: 
            - url: Hace referencia al enlace del sitio web donde se aplicara la verificación.
            - timeout: Hace referencia al limite de tiempo que tarda en hacer la petición.
            
        Salida:
            La función retorna un True si la página se encuentra accesible o un False en caso que no esté accesible.
    """
    try:
        
        response = requests.get(url, timeout=timeout, verify=False)
        
        if response.status_code == 200:
            return True
        if response.satus_code == 404 or response.status_code == 504 or response.status_code == 502 or response.status_code == 500 or response.status_code == 503:
            return False
        
    except Exception as e:
        logger.error(f"Error al verificar la accesibilidad de la página: {e}")
        return False

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
        


def process_contract(contrato):
    """
        Se encarga de procesar cada contrato por el número de contrato (NIC).
        
        Entradas:
            - contrato: Hace referencia al número de contrato a procesar.
    """    
    process_id = mp.current_process().pid
    download_dir = os.path.join(config["CARPETA_FACTURAS"]["carpeta_facturas_emsa"], f"process_{process_id}")
    os.makedirs(download_dir, exist_ok=True)

    options = ChromeOptions()
    preferences = {
        "download.default_directory": download_dir,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "useAutomationExtension": False,
        "profile.default_content_setting_values.notifications": 2,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    
    options.add_experimental_option("prefs", preferences)

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()
    
    url = "https://www.emsa-esp.com.co:441/factura/"
    
    
    if not verificar_status_pagina(url):
        
        logger.error(f"La página de EMSA ({url}) no se encuentra disponible en estos momentos.")
        
        send_email2(f"No se pudo acceder la página de EMSA", "La página no se encuentra disponible en este momento.")
        
        driver.quit()
        
        return

    try:
        logger.info("=" * 100)
        
        logger.info(f"Procesando contrato: {contrato}")

        # Ingresar a página de EMSA #
        driver.get(url)
        
        time.sleep(5)
        
        original_window = driver.current_window_handle

        # Obteniendo el Input del NIC
        input_nic = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "factura")))

        # Ingresando el número de NIC
        input_nic.send_keys(contrato)
        
        time.sleep(5)

        # Obteniendo el mes actual
        mes_actual = datetime.datetime.now().month
        
        mes_actual = str(mes_actual).zfill(2)
        
        logger.info(mes_actual)

        # Obteniendo el nombre del mes a partir del número del mes
        nombre_mes = numero_a_nombre_mes(mes_actual)

        # Obteniendo el periodo mes-año
        periodo_actual = f"{nombre_mes}-{datetime.datetime.now().strftime('%Y')}"
        
        logger.info(f"Periodo actual: {periodo_actual}")

        logger.info("Ingresando número de NIC")

        # Obteniendo y clic en el botón consultar factura 
        boton_consultar = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.NAME, "boton"))
        ).click()
        
        time.sleep(5)
        
        logger.info("Clic en consultar factura")

        # Obteniendo la tabla donde se encuentra el historico de facturas
        tabla_facturas = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//table[@class='table']"))
        )
        
        time.sleep(5)
        
        # Obteniendo todas las filas de la tabla facturas
        tabla_facturas = tabla_facturas.find_elements(By.XPATH, ".//tbody/tr")
        
        time.sleep(5)

        # Iterando cada fila de la tabla facturas
        for fila in tabla_facturas:
        
            time.sleep(5)
            # Obteniendo el texto que se encuentra en la columna número 4
            fecha_factura = fila.find_element(By.XPATH, "./td[4]").text
            #fecha_factura = datetime.datetime.strptime(fecha_factura, "%d-%b-%Y").strftime("%m-%Y")
            
            time.sleep(5)
            
            partes = fecha_factura.split('-')
            
            mes = partes[1]
            año = partes[2]
            
            fecha_factura = f"{mes}-{año}"

            logger.info(f"Fecha de la factura en la página: {fecha_factura}")
            
            time.sleep(3)
            
            if periodo_actual == fecha_factura:
                ver_factura = WebDriverWait(fila, 50).until(
                    EC.element_to_be_clickable((By.XPATH, ".//td[5]/a"))
                )

                time.sleep(3)

                # Scrolleando al botón ver facturas
                driver.execute_script("arguments[0].scrollIntoView(true);", ver_factura)

                time.sleep(3)

                # Clic en el botón ver factura
                ver_factura.click()  # Cuando se le da clic a ver facturas abre una nueva ventana mostrando la factura seleccionada

                time.sleep(5)
                
                # Obteniendo el número de ventanas
                ventanas = driver.window_handles
                
                time.sleep(5)
                
                
                windows = driver.window_handles

                for window in windows:
                    if window != original_window:
                        driver.switch_to.window(window)
                        break
            else:
                send_email2("PRUEBA ASUNTO EMSA", f"El presente correo es para informar que la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {periodo_actual}")
                logger.info(f"La factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {periodo_actual}")

        logger.info(f"Proceso de descarga terminado")
        
        main_folder = config["CARPETA_FACTURAS"]["carpeta_facturas_emsa"]
        
        nombre_comercializadora = "Emsa"
        
        logger.info("No se encontro el nombre de la comercializadora.")
        
        ruta_descarga = os.path.join(
            config["CARPETA_FACTURAS"]["ruta"],
            nombre_comercializadora,
            año,
            mes
        )
                        
        os.makedirs(ruta_descarga, exist_ok=True)

        for filename in os.listdir(download_dir):
            if filename.endswith(".pdf"):
                source_path = os.path.join(download_dir, filename)
                new_filename = f"doc_{contrato}.pdf"  # Nuevo nombre del archivo
                destination_path = os.path.join(ruta_descarga, new_filename)
                
                os.rename(source_path, destination_path)  # Renombrando archivo con el número de NIC
                logger.info(f"Renombrando pdf")
                
                # shutil.move(new_filename, ruta_descarga)
                # logger.info(f"Moviendo a la carpeta: {ruta_descarga}")

        # Eliminando la carpeta temporal 
        shutil.rmtree(download_dir)

        
        time.sleep(5)

    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
        
        
    time.sleep(10)
        
def download_emsa():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la libreria multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_emsa()
    contratos = df['CONTRATO'].tolist()

    with mp.Pool(processes=3) as pool:
        pool.map(process_contract, contratos)

if __name__ == "__main__":
    try:
        download_emsa()
    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")
import multiprocessing
import os
import time
import datetime
import traceback as tr
import toml
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from dateutil.relativedelta import relativedelta
from commons.commons import start_logging, read_excel, process_error, send_email
import win32com.client as win32
import win32com.client
import shutil

# Importar archivo config.toml
with open("config.toml", "r") as f:
    config = toml.load(f)

# Configuración del logger
logger = start_logging('prueba', mode='dev')

# Lock global para manejar escritura en el archivo de log
log_lock = multiprocessing.Lock()

def download_aire_for_nic(contrato):
    """
        Descarga la factura para un NIC específico.
        
        Entrada: 
            - Número de contrato (NIC) para el cual se descargará la factura.
    """
    options = Options()
    
    # Configura la carpeta temporal para descargas
    directorio_descarga_temporal = os.path.join(os.path.expanduser("~"), "Downloads", "temp_download_aire")
    
    preferences = {
        "download.default_directory": directorio_descarga_temporal,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "useAutomationExtension": False,
        "profile.default_content_setting_values.notifications": 2
    }

    options.add_experimental_option("prefs", preferences)

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()

    max_intentos = 300

    try:
        intentos = 1
        success = False

        while intentos <= max_intentos and not success:
            try:
                with log_lock:
                    logger.info("=" * 100)
                    logger.info(f"Iteración número: {contrato}")

                # Ingresar a página de Air-e
                driver.get("https://procesamiento.datosydisenos.com/micrositioaire/faces/index.xhtml")

                with log_lock:
                    logger.info(f"Ingresando a oficina virtual de Air-e con NIC: {contrato}")

                # Ingresar NIC
                input_nic = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "form:nic")))
                input_nic.send_keys(contrato)
                time.sleep(3)

                # Clic en consultar
                consultar_factura = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "form:j_idt25"))).click()
                with log_lock:
                    logger.info("Clic en consultar factura")

                # Obtener periodo actual - AJUSTAR AQUI
                periodo_actual = datetime.date.today()
                periodo_actual_format = periodo_actual.strftime('%Y-%m')
                periodo_actual_format = "2024-07"

                with log_lock:
                    logger.info(f"Periodo actual: {periodo_actual_format}")

                # Calcular mes anterior
                # mes_anterior = periodo_actual - relativedelta(months=1)
                # mes_anterior = mes_anterior.strftime("%Y-%m")

                # Verificar si hay mensaje de error por el captcha
                try:
                    error_message = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-messages-error-detail")))
                    if error_message.is_displayed():
                        with log_lock:
                            logger.warning(f"Error validando el captcha, reintentando... {intentos}")
                        intentos += 1

                        if intentos % 5 == 0:
                            with log_lock:
                                logger.warning("Demasiados intentos fallidos, pausando por 5 minutos...")
                            driver.delete_all_cookies()
                            driver.execute_script("window.localStorage.clear();")
                            driver.refresh()
                            time.sleep(300)

                        driver.refresh()
                        time.sleep(3)
                        continue
                except Exception as e:
                    pass

                # Obtiene todas las filas de la tabla
                facturas = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, "//table/tbody/tr")))
                factura_encontrada = False

                # Itera sobre cada fila de la tabla
                for factura in facturas:

                    # Obtiene los valores de la columna periodo
                    period_cell = factura.find_element(By.XPATH, "./td[1]")
                    periodo_pagina = period_cell.text.strip()

                    if periodo_pagina == periodo_actual_format:
                        with log_lock:
                            logger.info(f"Periodo encontrado: {periodo_pagina}")
                            logger.info("Descargando factura")
                            
                        
                        año, mes = periodo_pagina.split("-")
                        
                        nombre_comercializadora = "Air-e"
                        
                        ruta_descarga = os.path.join(
                            config["CARPETA_FACTURAS"]["ruta"], 
                            nombre_comercializadora, 
                            año, 
                            mes
                        )
                        
                        os.makedirs(ruta_descarga, exist_ok=True)
                        
                        descargar_factura = factura.find_element(By.XPATH, ".//button[contains(@id, 'j_idt')]").click()
                        
                        factura_encontrada = True
                        
                        time.sleep(10)
                        logger.info("LO ENCONTRO")
                        success = True
                        
                        # Mover el archivo descargado a la ruta deseada
                        for filename in os.listdir(directorio_descarga_temporal):
                            file_path = os.path.join(directorio_descarga_temporal, filename)
                            if os.path.isfile(file_path):
                                nuevo_ruta = os.path.join(ruta_descarga, filename)
                                shutil.move(file_path, nuevo_ruta)
                                
                                # Renombrar al solicitado por caro
                                #nuevo_nombre = f"t{num}_AIRE_{periodo_pagina}.pdf"
                                #nuevo_nombre_path = os.path.join(ruta_descarga, nuevo_nombre)
                                #os.rename(nuevo_ruta, nuevo_nombre_path)
                                
                                logger.info(f"Archivo movido a: {nuevo_ruta}")
                        break

                if factura_encontrada == False:
                    
                    send_email2("PRUEBA ASUNTO AIR-E", f"No se ha encontrado la factura al mes correspondiente con el número de NIC: {contrato}")
                    with log_lock:
                        logger.warning(f"No se encuentra la factura del mes vencido para NIC: {contrato}")
                    
                    success = True  # No es un error, simplemente no encontró la factura

                time.sleep(5)
            except Exception as e:
                with log_lock:
                    logger.warning(f"Error al cargar la página: {e}")
                intentos += 1
                
                if intentos >= max_intentos:
                    with log_lock:
                        process_error('prueba')
    finally:
        driver.quit()
        
        
def send_email2(subject, body):
    """
        Envía un correo electrónico para notificar a la persona deseada.

        Entradas:
            - subject: Asunto del correo.
            - body: Cuerpo del correo.
    """
    
    try:
        # Crea una instancia de la aplicación Outlook
        outlook = win32.Dispatch("Outlook.Application")
        
        # Crea un nuevo correo
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.To = config["EMAIL_SEND_AIRE"]["email_recept"]

        # Envía el correo
        mail.Send()
        print("Correo enviado exitosamente.")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")


def download_aire_multiprocessing():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'download_aire_for_nic' haciendo uso de la libreria multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel()

    # numero de contrato (NICS)
    contratos = df['CONTRATO'].tolist()
    
    numero_de_registros = df.shape[0]
    
    logger.info(f"El número total de registros es: {numero_de_registros}")
    
    # Crear un pool de procesos
    with multiprocessing.Pool(processes=2) as pool: 
        
        # Ejecutar la función en paralelo
        pool.map(download_aire_for_nic, contratos)

if __name__ == "__main__":
    try:
        download_aire_multiprocessing()
    except Exception as e:
        process_error('prueba')
        with log_lock:
            logger.warning(f"Error al cargar la aplicación: {e}")

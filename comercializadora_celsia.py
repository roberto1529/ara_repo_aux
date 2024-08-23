from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from undetected_chromedriver import Chrome, ChromeOptions
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from time import sleep
import win32com.client as win32
import datetime
import os
import traceback as tr
import time
import toml
import fitz  
import shutil
import requests
from commons.commons import start_logging
from commons.commons import read_excel_celsia
from commons.commons import process_error
from commons.commons import send_email
import multiprocessing as mp


# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('prueba', mode='dev')

def numero_a_nombre_mes(numero_mes):
    """
        Convierte un número de mes al nombre correspondiente de dicho mes.
        
        Entrada: Número del mes en formato string (ej: "01" para enero).
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

def send_email2(subject, body):
    """
        Se encarga de enviar una notificación vía correo electronico
        
        Entradas: 
            - subject: Hace referencia al asunto que llevará el correo.
            - body: Hace referencia al cuerpo que llevará el correo.
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

def leer_pdf_y_verificar(pdf_path, nombre_mes_actual):
    """
        Lee un archivo PDF y verifica si contiene una mención del mes actual.
        
        Entradas: 
            - pdf_path: Ruta del archivo PDF a leer.
            - nombre_mes_actual: Nombre del mes actual para verificar en el PDF.
            
        Salida:
            - Retorna el mes siempre y cuando lo encuentre en el PDF, de lo contrario retora un None.
    """
    
    try:
        doc = fitz.open(pdf_path)
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()
        doc.close()
        
        for linea in texto.splitlines():
            if linea.startswith("MES:"):
                mes = linea.split(":")[1].strip()
                logger.info(f"Mes encontrado en el pdf: {mes}")
                return mes == nombre_mes_actual
        return None
    except Exception as e:
        logger.warning(f"Error al leer el PDF {pdf_path}: {e}")
        return False

def send_email2(subject, body):
    """
        Envía un correo electrónico para notificar a la persona deseada.

        Entradas:
            - subject: Asunto del correo.
            - body: Cuerpo del correo.
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



def setup_driver(download_dir):
    """
        Configura y devuelve una instancia del controlador Chrome con configuraciones específicas para la descarga de archivos.

        Entradas:
            - download_dir: Directorio de descarga para los archivos.

        Salidas:
            - webdriver.Chrome: Instancia de web driver Chrome configurado.
    """
    
    options = ChromeOptions()
    
    # Configura la carpeta temporal para descargas
    directorio_descarga_temporal = os.path.join(os.path.expanduser("~"), "Downloads", "temp_download_celsia")
    
    preferences = {
        "download.default_directory": directorio_descarga_temporal,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "useAutomationExtension": False,
        "profile.default_content_setting_values.notifications": 2
    }
    
    options.add_experimental_option("prefs", preferences)
    options.add_argument('--disable-blink-features=AutomationControlled')

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    return driver


def download_factura(contrato):
    """
        Descarga la factura para un contrato específico (NIC) y verifica si la factura es del mes actual.
        
        Entradas:
            - contrato: Número de contrato (NIC) para el cual se desea descargar la factura.
    """
    
    download_dir = os.path.join(config["CARPETA_FACTURAS"]["carpeta_facturas_celsia"], f"temp_{mp.current_process().name}")
    os.makedirs(download_dir, exist_ok=True)
    
    driver = setup_driver(download_dir)
    try:
        
        logger.info(f"Procesando contrato: {contrato}")
        
        # Ingresar a página de Celsia
        driver.get("https://nube.celsia.com:4443/clientes/paga-tus-facturas")
        
        # Ingresar NIC
        input_nic = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "nicABuscar")))
        
        # Limpiar campo del NIC
        input_nic.clear()
        
        # Ingresar número de NIC
        input_nic.send_keys(contrato)
        
        time.sleep(5)
        
        # Seleccionando el chechbox
        checkbox = driver.find_element(By.ID, "mat-checkbox-2-input")
        
        # Creando una acción
        action = ActionChains(driver)
        
        # Moviendo el cursor y haciendo click al checkbox
        action.move_to_element(checkbox).click().perform()
        
        time.sleep(5)
        
        logger.info("Clic en acepto términos y condiciones")
        
        # Seleccionando boton para consultar el NIC
        boton_consultar = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "buscarCodigoCuenta")))
        
        # Valida si el botón está activo
        if boton_consultar.is_enabled():
            boton_consultar.click()
            logger.info("Clic en consultar factura")
        else:
            logger.info("El botón se encuentra deshabilitado")
            
        time.sleep(8)
        
        # Obteniendo la fecha/periodo actual
        fecha_actual = datetime.datetime.now().strftime('%Y/%m')
        
        # Obteniendo el mes actual
        mes_actual = datetime.datetime.now().strftime('%m')
        
        # Transformando el número de mes a nombre de mes
        nombre_mes_actual = numero_a_nombre_mes(mes_actual)
        
        logger.info(f"Mes convertido por la función: {nombre_mes_actual}")
        logger.info(f"Periodo actual: {fecha_actual}")
        
        time.sleep(8)
        
        # Obteniendo la opción duplicado factura
        opcion_duplicado_factura = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//h4[contains(@class, 'billing-main-info') and contains(text(), 'Duplicado de factura')]")))
        
        # Scroll hacia la opción duplicado factura
        driver.execute_script("arguments[0].scrollIntoView(true);", opcion_duplicado_factura)
        
        time.sleep(7)
        
        # Clic en la opción duplicado factura
        opcion_duplicado_factura.click()
        
        logger.info("Clic en Duplicado de factura")
        
        time.sleep(9)
        
        # Clic en el botón descargar factura
        boton_descargar = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'mat-flat-button') and contains(@class, 'mat-button-base') and contains(@class, 'w-50') and contains(@class, 'bg-transparent') and contains(@class, 'border-dark') and contains(@class, 'color-dark')]"))).click()
        
        logger.info("Descargando factura...")
        
        time.sleep(75)
        
        # Clic en aceptar
        boton_aceptar = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'mat-flat-button') and contains(@class, 'mat-button-base') and contains(@class, 'mat-primary') and contains(@class, 'custom-button') and contains(text(), 'Aceptar')]"))).click()
        
        logger.info("Validando mes de la factura...")
        
        # Lista de todos los pdfs de la carpeta facturascelsia
        pdfs = [f for f in os.listdir(download_dir) if f.endswith('.pdf')]
        
        # Path de cada pdf
        pdf_path = os.path.join(download_dir, pdfs[-1])
        
        
        # Verifica que la fecha de cada factura sea del mes actual
        # if leer_pdf_y_verificar(pdf_path, nombre_mes_actual):
        #     logger.info(f"La factura: {pdf_path} es del mes actual")
        #     shutil.move(pdf_path, os.path.join(config["CARPETA_FACTURAS"]["carpeta_facturas_celsia"], os.path.basename(pdf_path)))
        # else:
            
        #     send_email2("PRUEBA ASUNTO CELSIA", f"El presente correo es para informar que la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {fecha_actual}")
        #     logger.info(f"la factura perteneciente al NIC {contrato} no se encuentra disponible para el periodo {fecha_actual}")
            
        #     os.remove(pdf_path)
        
        shutil.move(pdf_path, os.path.join(config["CARPETA_FACTURAS"]["carpeta_facturas_celsia"], os.path.basename(pdf_path)))

            
        logger.info("Proceso finalizado!")
        
        time.sleep(10)
        
    except Exception as e:
        logger.warning(f"Error en contrato {contrato}: {e}")
    finally:
        driver.quit()
        shutil.rmtree(download_dir)

def main():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la libreria multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_celsia()
    contratos = df['CONTRATO'].tolist()
    
    with mp.Pool(processes = 3) as pool:
        pool.map(download_factura, contratos)

if __name__ == "__main__":
    try:
        main()
    except:
        process_error('prueba')
        print("error en cargue")

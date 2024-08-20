import multiprocessing as mp
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import win32com.client as win32
import datetime
import os
import traceback as tr
import time
import toml
import fitz  
import logging
import calendar
import shutil
import requests
import re

from commons.commons import start_logging, read_excel_enelar, process_error, send_email

# Importar archivo config.toml
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('LOGS_EMSA', mode='dev')

def send_email2(subject, body):
    """
    Se encarga de enviar una notificación vía correo electrónico.

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

def solve_captcha(api_key, site_key, url):
    """
    Resuelve el CAPTCHA utilizando el servicio Anti-Captcha.

    Entradas:
        - api_key: Clave API de Anti-Captcha.
        - site_key: La clave del sitio de Google reCAPTCHA.
        - url: La URL del sitio web que contiene el CAPTCHA.

    Salidas:
        - token: El token generado por Anti-Captcha que resuelve el CAPTCHA.
    """
    # Paso 1: Crear la tarea
    task_data = {
        "clientKey": api_key,
        "task": {
            "type": "NoCaptchaTaskProxyless",
            "websiteURL": url,
            "websiteKey": site_key
        }
    }
    
    create_task_url = "https://api.anti-captcha.com/createTask"
    response = requests.post(create_task_url, json=task_data)
    result = response.json()
    
    if result.get("errorId") != 0:
        print(f"Error al crear la tarea: {result.get('errorDescription')}")
        return None
    
    task_id = result.get("taskId")
    
    # Paso 2: Obtener el resultado
    check_result_url = "https://api.anti-captcha.com/getTaskResult"
    while True:
        result_data = {
            "clientKey": api_key,
            "taskId": task_id
        }
        result_response = requests.post(check_result_url, json=result_data)
        result = result_response.json()
        
        if result.get("errorId") != 0:
            print(f"Error al obtener el resultado: {result.get('errorDescription')}")
            return None
        
        if result.get("status") == "ready":
            return result.get("solution", {}).get("gRecaptchaResponse")
        
        time.sleep(5)  # Esperar antes de intentar nuevamente

def process_contract(contrato):
    """
    Se encarga de procesar cada contrato por el número de contrato (NIC).

    Entradas:
        - contrato: Hace referencia al número de contrato a procesar.
    """
    carpeta_facturas_energuaviare = config["CARPETA_FACTURAS"]["carpeta_facturas_energuaviare"]
    api_key = "400ac675b3cf154155d9fc68e6a40929"
    url = "https://enelar.net.co:9876/consultar-factura/"
    site_key = "6LdVX18aAAAAAIKpX6IbUmGJqDEl0M0K-ecKoL71"  # Debes reemplazar con el site_key correcto del CAPTCHA

    options = Options()
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

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()
    
    try:
        
        logger.info("=" * 100)
        
        # Inicia el proceso de procesamiento de facturas
        logger.info(f"Procesando contrato: {contrato}")
        
        driver.get(url)
        
        time.sleep(2)
        
        input_nic = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'mat-input-0'))
        )
        
        input_nic.send_keys(contrato)
        
        time.sleep(3)
        
        # Obtener el token del CAPTCHA usando Anti-Captcha
        #captcha_token = solve_captcha(api_key, site_key, url)
        
        #logger.info(f"Token del CAPTCHA obtenido: {captcha_token}")
        # Insertar el token en el campo correspondiente en la página web
        #driver.execute_script(f"document.getElementById('g-recaptcha-response').innerHTML = '{captcha_token}';")
        #driver.execute_script("document.getElementById('g-recaptcha-response').dispatchEvent(new Event('input', { bubbles: true }));")
        #driver.execute_script("document.getElementById('g-recaptcha-response').dispatchEvent(new Event('change', { bubbles: true }));")
        
        
        # Aquí puedes agregar la lógica para enviar el formulario y continuar con el proceso.
        # Aqui se agrega la lógica para todo el envío de formulario.
        time.sleep(3)
            
        boton = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.mat-mini-fab"))
        )
            
        time.sleep(3)
            
        driver.execute_script("""
                arguments[0].removeAttribute('disabled');
                arguments[0].classList.remove('mat-button-disabled');
            """, boton)
            
        time.sleep(3)
            
        # Opcional: Verifica si el botón ya está habilitado y si la clase ha sido eliminada
        activar_boton = driver.execute_script("return arguments[0].hasAttribute('disabled');", boton)
        
        quitar_clase_css = driver.execute_script("return arguments[0].classList.contains('mat-button-disabled');", boton)
            
        logger.info(f"El botón está {'habilitado' if not activar_boton else 'deshabilitado'}")
        
        logger.info(f"La clase 'mat-button-disabled' está {'presente' if quitar_clase_css else 'eliminada'}")
            
        logger.info(f"Clic en descargar: {boton}")
            
        boton.click()
                
        time.sleep(70)
            
        logger.info("Descargando!")

        time.sleep(5)
    
    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
    finally:
        driver.quit()


def download_enelar():
    """
    Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la librería multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_enelar()
    contratos = df['CONTRATO'].tolist()

    with mp.Pool(processes=1) as pool:
        pool.map(process_contract, contratos)

if __name__ == "__main__":
    try:
        download_enelar()
    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")

import multiprocessing as mp
from threading import Thread
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

from commons.commons import start_logging, read_excel_enelar, process_error, send_email

# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)


logger = start_logging('LOGS_EMSA', mode='dev')


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


def process_contract(contrato):
    """
        Se encarga de procesar cada contrato por el número de contrato (NIC).
        
        Entradas:
            - contrato: Hace referencia al número de contrato a procesar.
    """    
    
    carpeta_facturas_energuaviare = config["CARPETA_FACTURAS"]["carpeta_facturas_energuaviare"]
    
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

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()
    
    try:
        logger.info("=" * 100)
        
        # Inicia el proceso de procesamiento de facturas
        logger.info(f"Procesando contrato: {contrato}")
        
        driver.get("https://enelar.net.co:9876/consultar-factura/")
        
        time.sleep(2)
        
        input_nic = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'mat-input-0'))
        )
        
        input_nic.send_keys(contrato)
        
        time.sleep(3)
        
        try:
            
            boton = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.mat-mini-fab"))
            )
            
            logger.info(f"Obteniendo boton: {boton}")
            
            driver.execute_script("""
                arguments[0].removeAttribute('disabled');
                arguments[0].classList.remove('mat-button-disabled');
            """, boton)
            
            logger.info(f"Captcha resuelto.")
            
            is_disabled = driver.execute_script("return arguments[0].hasAttribute('disabled');", boton)
            has_class = driver.execute_script("return arguments[0].classList.contains('mat-button-disabled');", boton)
            
            logger.info(f"El botón está {'habilitado' if not is_disabled else 'deshabilitado'}")
            logger.info(f"La clase 'mat-button-disabled' está {'presente' if has_class else 'eliminada'}")
            
            logger.info(f"Clic en descargar: {boton}")
            
            boton.click()
            
            time.sleep(50)
            
        except Exception as e:
            logger.error(f"Error durante el manejo del CAPTCHA: {e}")
    
        time.sleep(5)
    
    except Exception as e:
        
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
    finally:
        driver.quit()

# Funcion multiproccesing.

def download_enelar():
    """
        Se encarga de recolectar todos los contratos y ejecutar la función 'process_contract' haciendo uso de la libreria multiprocessing para crear varios procesos en paralelo.
    """
    df = read_excel_enelar()
    contratos = df['CONTRATO'].tolist()


    with mp.Pool(processes = 1) as pool:
        pool.map(process_contract, contratos)

if __name__ == "__main__":
    try:
        download_enelar()
    except Exception as e:
        
        process_error('warning')
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")
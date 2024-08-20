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
import shutil
import requests
import re
import random  # Importar random para tiempos de espera aleatorios

from commons.commons import start_logging, read_excel_enelar, process_error, send_email

# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

logger = start_logging('LOGS_EMSA', mode='dev')

def send_email2(subject, body):
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

def human_sleep(min_seconds, max_seconds):
    time.sleep(random.uniform(min_seconds, max_seconds))

def process_contract(contrato):
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

    if os.path.isfile("chromedriver.exe"):
        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()
    
    try:
        logger.info("=" * 100)
        
        logger.info(f"Procesando contrato: {contrato}")
        
        driver.get("https://checweb.cadenaportalgestion.com/SegUsuario/Login?ReturnUrl=%2f")
        
        human_sleep(2, 5)
        
        driver.switch_to.frame(0)
        
        human_sleep(2, 5)
        
        recaptcha_checkbox = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "recaptcha-checkbox-border"))
        )
        
        action = webdriver.common.action_chains.ActionChains(driver)
        
        action.move_to_element_with_offset(recaptcha_checkbox, random.uniform(-5, 5), random.uniform(-5, 5))
        
        action.click()
        
        action.perform()
        
        human_sleep(30, 60)
        
    except Exception as e:
        process_error('warning')
        logger.error(f"Ocurrió un error al procesar el contrato {contrato}: {e}")
    finally:
        driver.quit()

def download_enelar():
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

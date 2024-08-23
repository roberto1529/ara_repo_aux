from dateutil.relativedelta import relativedelta
from multiprocessing import Pool
from selenium import webdriver # CREAR EL DRIVER: el objeto que nos va permitir manejar el navegador #
from selenium.webdriver import ActionChains # Nos sirve para crear acciones o eventos en el navegador
from selenium.webdriver.chrome.service import Service  # Nos sirve para iniciar o detener el chromedriver.
from selenium.webdriver.chrome.options import Options# Nos sirve para configurar las opciones del navegador
from selenium.webdriver.common.by import By# Nos sirve para identificar los elementos en una pagina web, tales como: botones, slides, checkbox, radiobuttoms, text inputs, entre otros
from selenium.webdriver.support.ui import WebDriverWait# Nos sirve para configurar los tiempos de espera en el navegador dependiendo de un acción #
from selenium.webdriver.support import expected_conditions as EC# Nos sirve para manejar las excepciones cuando esperamos algún compartamiento y este no sucede #
from webdriver_manager.chrome import ChromeDriverManager # Nos sirve para tener siempre la versión chromedriver requerida para nuestro navegador #
from time import sleep


import datetime
import os
import pandas as pd
import traceback as tr
import time
import toml
import logging


# Importar archivo config. toml #
with open("config.toml", "r") as f:
    config = toml.load(f)

def strip_accents(s: str):
    import unicodedata

    return "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def listado_nics():

    df = pd.read_excel("./datos/listado_contratos.xlsx")
    #strip_accents(listado)
    df = df.loc[df["comercializador"] == "AFINIA"]
    df.drop(columns="comercializador",inplace=True,index=None)
    nics=df["contrato"].tolist()
    
    return nics

listado_nics()

def download_duplicado_afinia(nic):
    """
    Función que permite ingresar a la plataforma oficina virtual de afinia
    """

    options = Options()

    # options.add_argument("--no-sandbox")

    # options.add_argument("--headless")

    # options.add_argument("--disable-dev-shm-usage")

    # options.add_argument("--disable-gpu")

    # options.add_argument("--disable-extensions")

    # options.add_argument("--allow-insecure-localhost")

    # options.add_argument("--ignore-certificate-errors")

    # options.add_argument("--incognito")

    # options.add_argument("--log-level=3")

    preferences = {"directory_upgrade": True,
        "safebrowsing.enabled": True,
        "useAutomationExtension": False,
        "profile.default_content_setting_values.notifications": 2,}

    options.add_experimental_option("prefs", preferences)

    if os.path.isfile("chromedriver.exe"):

        driver = webdriver.Chrome(service=Service(executable_path="chromedriver.exe"), options=options)

    else:

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.maximize_window()


    try:

        # Ingresar a página de afinia#
        driver.get("https://caribemar.facture.co/")

        print("Cargue de oficina virtual de afinia")

        # Clic en cnsultar factura
        factura = (WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//strong[contains(.,'Consulte su Factura')]"))).click())

        print("Clic en consulte su factura")

        # Ingresar NIC
        input_nic = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@name,'ref')]")))

        input_nic.send_keys(nic)

        print("Diligenciando nic")

        # Clic en consultar
        consultar = (WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                    (By.XPATH,"//span[@class='ng-binding ng-scope'][contains(.,'Consultar')]",))).click())

        print("Clic en consultar")

        time.sleep(5)

        # Ingresar NIC nuevamente
        input_nic_again = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@required,'required')]")))

        input_nic_again.send_keys({nic})

        print("Diligenciar nic nuevamente")

        # Consultar nuevamente
        consultar_again = (WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                            (By.XPATH,"//span[@class='ng-binding ng-scope'][contains(.,'Siguiente')]",))).click())

        print("Clic en consultar nuevamente")

        time.sleep(5)

        # Obtener periodo actual
        periodo_actual = datetime.date.today()

        # obtener periodo vencido
        mes_anterior = periodo_actual - relativedelta(months=1)

        mes_anterior = mes_anterior.strftime("%Y/%m")

        print(f"Periodo a consultar: {mes_anterior}")

        # obtener periodo de ultima factura cargada
        period_cell = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[5]"))
        )

        periodo_pagina = period_cell.text

        print(f"Periodo de ultima factura cargada: {periodo_pagina}")

        # validar que se encuentre la factura del periodo vencido
        if periodo_pagina != mes_anterior:

            print("No se encuentra la factura del mes vencido")
            # driver.quit()

        else:

            print("Se encuentra disponible la factura del mes vencido")

            id_documento = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[3]")))

            id_documento = id_documento.text

            print(f"id factura:{id_documento}") 

            descarga_factura = driver.get(f'https://caribemar.facture.co/DesktopModules/GatewayOficinaVirtual.Commons/API/Pdf/Get?id={id_documento}')

            time.sleep(20)

            #Cerrar navegador
            driver.quit()

            # time.sleep(70)

    except:

        print("Error al cargar la página")

        print(tr.format_exc())

        logging.warning(f"Error al cargar la página")

        # driver.refresh()

    return None



if __name__ == "__main__":

    #mode = "prod" if "prod" in sys.argv else "dev"

    try:

        nic_list = listado_nics()

        with Pool(5) as p:  
        # Ajusta el número de procesos según tu necesidad
            p.map(download_duplicado_afinia, nic_list)

    except:
        print("error en cargue")

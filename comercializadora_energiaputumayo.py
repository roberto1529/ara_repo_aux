import multiprocessing as mp
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import datetime
import logging
import glob
from commons.commons import read_excel_energiaputumayo, process_error, send_email

def procesar_facturas(contrato):
    # Configuración del logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Configuración del navegador
    options = Options()
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
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--disable-features=InsecureDownloadWarnings")
    options.add_argument("--window-size=800x600")  # Reducir tamaño de ventana

    # Configuración de preferencias de descarga
    download_path = 'C:\\Users\\P108\\Documents\\PyDocto\\energiaputumayo'
    prefs = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.managed_default_content_settings.images": 2,  # Bloquear imágenes
        "download.default_directory": download_path
    }
    options.add_experimental_option("prefs", prefs)

    # Configurar el servicio de ChromeDriver
    service = Service('chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)

    logging.info("Navegador Chrome iniciado")

    # Inicia el proceso de procesamiento de facturas
    logging.info(f"Procesando contrato: {contrato}")

    # Carga la página local
    driver.get('https://www.energiaputumayo.com/Backup/factura/ConsultaFactura1.php')
    logging.info("Página cargada")

    try:
        # Espera a que el campo de matrícula esté presente y visible
        input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'matricula'))
        )
        conStr = str(contrato)
        input_element.send_keys(conStr)

        logging.info(f"Texto escrito en el input: {conStr}")

        # Espera a que el botón de consultar esté presente y visible
        consult_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'Enviar'))
        )
        consult_button.click()
        logging.info("Botón 'Consultar' clicado")

        # Espera a que la tabla de resultados esté disponible
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'tablaFactura'))
        )

        # Encuentra la celda que contiene "Mes Facturado" y extrae el texto del siguiente <td> en la misma fila
        mes_facturado_element = driver.find_element(By.XPATH, '//table[@id="tablaFactura"]//tr[2]/td[2]')
        mes_facturado_text = mes_facturado_element.text
        download_path = generar_arbol_carpetas(mes_facturado_text)

        # Actualiza las preferencias de descarga
        prefs["download.default_directory"] = download_path
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_path
        })

        logging.info("Mes capturado: %s", mes_facturado_text)
        logging.info("Árbol generado: %s", download_path)

        # Encuentra el formulario dentro de la tabla y el botón "Imprimir"
        print_button = driver.find_element(By.XPATH, '//table[@id="tablaFactura"]//input[@type="submit" and @value="Imprimir"]')
        print_button.click()
        logging.info("Botón 'Imprimir' clicado")

        # Espera a que se abra la nueva ventana con el PDF
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

        # Cambia el foco a la nueva ventana
        driver.switch_to.window(driver.window_handles[1])
        logging.info("Cambiado a la nueva ventana")

        # Espera adicional para asegurar que el PDF se descargue
        logging.info("Esperando la descarga del PDF")
        time.sleep(10)

        # Esperar a que el archivo se descargue completamente
        archivos_pdf = [f for f in os.listdir(download_path) if f.endswith('.pdf')]
        archivos_doc = [f for f in archivos_pdf if 'doc.pdf' in f]
        
        for archivo in archivos_doc:
            nuevo_nombre = f"{contrato}.pdf"
            nuevo_nombre_path = os.path.join(download_path, nuevo_nombre)

            # Evitar sobrescritura, agregar sufijo incremental si el archivo ya existe
            if os.path.exists(nuevo_nombre_path):
                base, ext = os.path.splitext(nuevo_nombre)
                contador = 1
                while os.path.exists(nuevo_nombre_path):
                    nuevo_nombre = f"{base}_{contador}{ext}"
                    nuevo_nombre_path = os.path.join(download_path, nuevo_nombre)
                    contador += 1
            os.rename(os.path.join(download_path, archivo), nuevo_nombre_path)
            

            logging.info(f"Archivo renombrado a: {nuevo_nombre}")
       
        if not archivos_doc:
            logging.warning("No se encontraron archivos PDF con el nombre 'doc.pdf'.")

        # Cierra la ventana del PDF y vuelve a la ventana principal
        # driver.close()
        # driver.switch_to.window(driver.window_handles[0])
     
            pattern = os.path.join(download_path, '*.crdownload')
            files_to_remove = glob.glob(pattern)
            # Eliminar cada archivo encontrado
            for file_path in files_to_remove:
                    try:
                        os.remove(file_path)
                        print(f"Eliminado: {file_path}")
                    except Exception as e:
                        print(f"Error al eliminar {file_path}: {e}")
                   

    except Exception as e:
        logging.error(f"Ocurrió un error: {str(e)}")
    finally:
        # Cierra el navegador
        driver.quit()
        

def obtener_mes_numero(mes_texto):
    meses = {
        "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
        "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12
    }

    try:
        mes_numero = int(mes_texto)
        if 1 <= mes_numero <= 12:
            return mes_numero
        else:
            print(f"Número de mes {mes_numero} no es válido")
            return None
    except ValueError:
        mes_texto = mes_texto.lower()[:3]
        return meses.get(mes_texto, None)

def generar_arbol_carpetas(texto_fecha):
    if '-' in texto_fecha:
        mes_texto = texto_fecha.split('-')[0]
    else:
        mes_texto = texto_fecha[:3]

    mes_numero = obtener_mes_numero(mes_texto)

    if not mes_numero:
        return None

    anio_actual = datetime.datetime.now().year

    ruta = os.path.join(
        'C:\\Users\\P108\\Documents\\PyDocto\\energiaputumayo',
        str(anio_actual),
        f'{mes_numero:02d}'
    )

    if not os.path.exists(ruta):
        os.makedirs(ruta)
        print(f"Directorio {ruta} creado!")
    else:
        print(f"Directorio {ruta} ya existe")

    return ruta

def download_contratos():
    df = read_excel_energiaputumayo()
    contratos = df['CONTRATO'].tolist()

    for contrato in contratos:
        procesar_facturas(contrato)
        time.sleep(2)

if __name__ == "__main__":
    try:
        download_contratos()
    except Exception as e:
        logging.error(f"Ocurrió un error al cargar la aplicación: {e}")

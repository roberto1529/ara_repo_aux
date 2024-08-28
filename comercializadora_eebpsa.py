import os
import time
import logging
import multiprocessing as mp
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from commons.commons import start_logging, read_excel_eebpsa, read_excel_homologacion
from funciones import generar_arbol_carpetas, renombrar_pdf, mover_pdf, eliminar_archivo_con_motivo, verificar_descarga, registrar_descarga, conexion_correo
import glob

# Configuración del logging
logger = start_logging("LOGS_EEBPSA", mode="dev")
comercializadora = 'EEBPSA'

def configurar_navegador():
    """
    Configura el navegador Chrome para la descarga del PDF.
    """
    options = Options()
    download_path = f'C:\\Users\\P108\\Documents\\PyDocto\\{comercializadora}'
    prefs = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.managed_default_content_settings.images": 2,  # Bloquear imágenes
        "download.default_directory": download_path
    }
    options.add_experimental_option("prefs", prefs)
    
    return options, download_path

def extraer_mes_anio(valor_opcion):
    """
    Extrae el mes y el año del valor del option en formato "dd/mm/yy".
    """
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    try:
        mes_num = valor_opcion.split('/')[1]
        anio = '20' + valor_opcion.split('/')[2]  # Ajusta el año si es necesario
        mes = meses[int(mes_num) - 1]
        mes_num_formateado = f'{int(mes_num):02d}'
        return mes, mes_num_formateado, anio
    except IndexError:
        logger.error(f"Error al extraer mes y año del valor: {valor_opcion}")
        return None, None, None

def descargar_pdf(contrato, mes, anio):
    """
    Usa Selenium para descargar el PDF del contrato si aún no ha sido descargado.
    """
    options, download_path = configurar_navegador()
    service = Service('chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    log_file = os.path.join(download_path, 'descargas.log')

    pdf_nombre_base = f'{contrato["CONTRATO"]}_{mes}_{anio}.pdf'
    pdf_path = os.path.join(download_path, pdf_nombre_base)

    if verificar_descarga(contrato['CONTRATO'], download_path):
        logger.info(f"El archivo para el contrato {contrato['CONTRATO']} ya ha sido descargado anteriormente.")
        return pdf_path

    logger.info("Navegador Chrome iniciado")

    try:
        driver.get('https://sie.eebpsa.com.co:8043/atencliente/')
        logger.info("Página cargada")
        
        # Verificar si la página se cargó correctamente
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'formConsultaFacturaPeriodo-nroSuscripcion'))
        )
        
        input_element = driver.find_element(By.ID, 'formConsultaFacturaPeriodo-nroSuscripcion')
        input_element.send_keys(contrato['CONTRATO'])
        logger.info(f"Texto escrito en el input: {contrato['CONTRATO']}")

        # Esperar a que las opciones del select estén disponibles
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'formConsultaFacturaPeriodo-periodoFacturado'))
        )

        select_element = driver.find_element(By.ID, 'formConsultaFacturaPeriodo-periodoFacturado')

        # Obtener el valor de la opción para seleccionar
        opciones = select_element.find_elements(By.TAG_NAME, 'option')
        for opcion in opciones:
            valor = opcion.get_attribute('value')
            texto = opcion.text
            if valor and texto:
                mes_texto, mes_num, anio = extraer_mes_anio(valor)
                if mes_texto and mes_num and anio:
                    logger.info(f"Intentando seleccionar: {texto}")

                    # Utilizar JavaScript para seleccionar la opción
                    driver.execute_script("arguments[0].value = arguments[1];", select_element, valor)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", select_element)
                    logger.info(f"Mes y año seleccionados: {texto}")

                    consult_button = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, 'formConsultaFacturaPeriodo-btnConsultar'))
                    )
                    consult_button.click()
                    logger.info("Botón 'Consultar' clicado")

                    # Esperar a que el enlace de descarga esté disponible y hacer clic
                    WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//a[@href="javascript:$(\'#formConsultaFacturaPeriodo-descargarFactura\').click();"]'))
                    )
                    descargar_link = driver.find_element(By.XPATH, '//a[@href="javascript:$(\'#formConsultaFacturaPeriodo-descargarFactura\').click();"]')
                    driver.execute_script("arguments[0].click();", descargar_link)
                    logger.info("Enlace de descarga clicado")

                    # Esperar a que la descarga se complete
                    tiempo_espera_descarga = 30  # Ajusta el tiempo si es necesario
                    time.sleep(tiempo_espera_descarga)

                    # Verificar si el archivo PDF fue descargado
                    pdf_descargado = max(glob.glob(os.path.join(download_path, 'factura-eebp-*.pdf')), key=os.path.getctime)
                    if os.path.exists(pdf_descargado):
                        logger.info(f"Archivo descargado en: {pdf_descargado}")

                        # Renombrar y mover el archivo
                        nuevo_nombre = renombrar_pdf(contrato['SAP'], contrato['Comercializadora'], f'{mes_texto} {anio}')
                        if nuevo_nombre:
                            nueva_ruta = generar_arbol_carpetas(f'{mes_texto} {anio}', comercializadora)
                            if nueva_ruta:
                                destino = os.path.join(nueva_ruta, nuevo_nombre)
                                ruta_final = mover_pdf(pdf_descargado, destino)

                                if ruta_final:
                                    logger.info(f"Archivo movido a: {ruta_final}")
                                    registrar_descarga(contrato['CONTRATO'], ruta_final, log_file)
                                else:
                                    logger.error(f"No se pudo mover el archivo {pdf_descargado} a {destino}")
                            else:
                                logger.error(f"No se pudo generar la carpeta para {mes_texto} {anio}")
                        else:
                            logger.error(f"No se pudo renombrar el archivo para el contrato {contrato['CONTRATO']}")
                    else:
                        logger.error(f"El archivo PDF no se descargó correctamente para el contrato {contrato['CONTRATO']}")
                    break  # Salir del loop después de procesar la opción
    except Exception as e:
        logger.error(f"Ocurrió un error durante la descarga del PDF: {str(e)}")
        err = "Ocurrió un error durante la descarga del PDF. " + str(e)
        conexion_correo('Error de ejecución en el bot', err);
    finally:
        driver.quit()

def procesar_facturas(contrato):
    """
    Procesa un contrato descargando el PDF, extrayendo información relevante,
    renombrando el archivo y moviéndolo a la carpeta correspondiente.
    """
    try:
        # Primero extraemos las opciones del select y luego iteramos
        options, _ = configurar_navegador()
        service = Service('chromedriver.exe')
        driver = webdriver.Chrome(service=service, options=options)

        driver.get('https://sie.eebpsa.com.co:8043/atencliente/')
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'formConsultaFacturaPeriodo-periodoFacturado'))
        )

        select_element = driver.find_element(By.ID, 'formConsultaFacturaPeriodo-periodoFacturado')
        opciones = select_element.find_elements(By.TAG_NAME, 'option')

        for opcion in opciones:
            valor = opcion.get_attribute('value')
            texto = opcion.text
            if valor and texto:
                mes_texto, mes_num, anio = extraer_mes_anio(valor)
                if mes_texto and mes_num and anio:
                    logger.info(f"Procesando opción: {texto}")
                    descargar_pdf(contrato, mes_num, anio)
                    time.sleep(5)  # Tiempo para permitir que el navegador procese la descarga
        driver.quit()
    except Exception as e:
        logger.error(f"Ocurrió un error durante el procesamiento del contrato: {str(e)}")
        err = "Ocurrió un error durante el procesamiento del contrato. " + str(e)
        conexion_correo('Error de ejecución en el bot', err);

def download_contratos():
    """
    Recolecta todos los contratos y ejecuta la función 'procesar_facturas' usando multiprocessing.
    """
    df = read_excel_eebpsa()
    hm = read_excel_homologacion()
    dfm = df.merge(hm, on='Supplier', how='left')

    contratos = dfm.to_dict('records')  # Convertir DataFrame a lista de diccionarios

    with mp.Pool(processes=4) as pool:  # Ajusta el número de procesos según tu capacidad
        pool.map(procesar_facturas, contratos)

if __name__ == "__main__":
    try:
        download_contratos()
    except Exception as e:
        logger.error(f"Ocurrió un error al cargar la aplicación: {e}")
        err = "Ocurrió un error al cargar la aplicación: " + str(e)
        conexion_correo('Error de ejecución en el bot', err);

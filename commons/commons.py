from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from logging.handlers import RotatingFileHandler

import logging
import multiprocessing as mp
import os
import smtplib
import ssl
import time
import toml 
import traceback
import pandas as pd

#config_path = 'C:\\Users\\P340\ISES S.A.S\\Administrador App - Analitica de datos\\Proyectos_Automatizaciones\\bot_descarga_pdf_repo_auditoria_enel\\config.toml'
config_path = 'C:\\Users\\P108\\Documents\\ARA_PY\\ara_repo_aux\\config.toml'

with open(config_path, 'r') as f:
    config = toml.load(f)

app_logger = logging.getLogger(config['app_parameters']['LOGGER_NAME'])


def send_email(send_to, subject, message):
    # launch process_email
    process = mp.Process(
        target=process_email,
        args=(send_to, subject, message,),
    )
    process.start()


def process_email(send_to, subject, message):
    # build message
    msg = MIMEMultipart()
    msg['From'] = config['DICT_APP_MAIL']['app_mail']
    msg['To'] = ", ".join(send_to)
    msg['Subject'] = subject
    body = message
    msg.attach(MIMEText(body, 'plain'))
    # context
    context = ssl.create_default_context()
    # send email
    with smtplib.SMTP(host=config['DICT_APP_MAIL']['smtp_server'], 
                    port=config['DICT_APP_MAIL']['smtp_port']) as server:
        for _ in range(10):
            try:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(config['DICT_APP_MAIL']['app_mail'], config['DICT_APP_MAIL']['email_pass'])
                server.sendmail(config['DICT_APP_MAIL']['app_mail'], send_to, msg.as_string())
                break
            except Exception as error:
                print(f"ERROR connecting to smtp server: {error}")
                time.sleep(0.1)


def process_error(logger_name):
    logger = logging.getLogger(logger_name)
    # get error trace
    error_trace = traceback.format_exc().splitlines()[::-1]
    # log error
    error_lines = []
    for error_line in error_trace:
        error_lines.append(error_line.strip())
        if 'line' in error_line:
            break
    logger.error(" --> ".join(error_lines[::-1]))
    # send email to admins
    err_message = "\n".join(error_lines[::-1])
    send_email(config['app_parameters']['ON_ERROR_EMAIL'], 'BOT Ara: Error Detected', err_message)
    logger.info("email sent to dev-admins with error")


def start_logging(logger_name, mode='dev'):
    if mode == 'prod':
        log_formatter = logging.Formatter(
            '%(asctime)s | %(lineno)s | %(levelname)s | %(message)s')
        log_handler = logging.StreamHandler()
        log_handler.setFormatter(log_formatter)
        logger = logging.getLogger(logger_name)
        logger.addHandler(log_handler)
        logger.setLevel("DEBUG")
    else:

        logger = logging.getLogger(logger_name)
        log_formatter = logging.Formatter('%(asctime)s | %(lineno)s | %(levelname)s | %(message)s')
        # Configurar el nivel de registro para la consola
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)  
        console_handler.setFormatter(log_formatter)
        logger.addHandler(console_handler)
        # Configurar el nivel de registro para el archivo log
        logging.captureWarnings(True)
        # validate if log path exist
        log_path = 'logs/'
        if not (os.path.exists(log_path) and os.path.isdir(log_path)):
            os.mkdir(log_path)
        # log file
        today = time.localtime()
        date = time.strftime("%Y%m%d", today)
        log_file = f"{log_path}log_{logger_name}_{date}.log"        
        #log_file = f"{log_path}{os.sep}{logger_name}{date}.log"
        # handler
        # log_formatter = logging.Formatter(
        #     '%(asctime)s | %(lineno)s | %(levelname)s | %(message)s')
        log_handler = RotatingFileHandler(
            log_file, mode='w', maxBytes=1024 * 1024,
            backupCount=1, encoding=None, delay=0)
        log_handler.setFormatter(log_formatter)
        # logger
        # logger = logging.getLogger(logger_name)
        logger.addHandler(log_handler)
        logger.setLevel(config['app_parameters']['LOG_LEVEL'])
    return logger


def exception_handler_and_timing(func):
    def wrapper(*args, **kwargs):
        #app_logger.info(f"Function {func.__name__} started")
        start_time = time.time()
        try:
            result = func(*args, **kwargs)
            return result
        except Exception:
            process_error(app_logger.name)
        finally:
            end_time = time.time()
            execution_time = end_time - start_time
            app_logger.info(f"{func.__name__} process finished after {execution_time:.2f} seconds")

    return wrapper


def read_excel():
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Limpiar columna CONTRATO
        #df['CONTRATO'] = df['CONTRATO'].str.replace('.', '', regex=False)
        
        # Filtrar columna contrato != vacio and != no aplica
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA')]
        
        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('AIRE|CARIBESOL DE LA COSTA SAS ESP', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")



def read_excel_celsia():
    
    
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Limpiar columna CONTRATO
        #df['CONTRATO'] = df['CONTRATO'].str.replace('.', '', regex=False)
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN') & (df['CONTRATO'] != '105663') & (df['CONTRATO'] != '470444') & (df['CONTRATO'] != '68303')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('CELSIA|CELSIA COLOMBIA SA ESP|CELSIA TOLIMA SA ESP', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")
        
        
def read_excel_emsa():
    
    try:
        
        
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('EMSA', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")



def read_excel_energuaviare():
    
    
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('EMPRESA DE ENERGÍA ELÉCTRICA DEL GUAVIAR', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")


def read_excel_enelar():
    
    
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('ENELAR', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")
        
        
def read_excel_ceosp():
    
    
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('CEO', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")
        

def read_excel_dispac():
    
    
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('DISPAC', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")
        
        
def read_excel_electrocaqueta():
    
    
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']
    
        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')
        
        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']
        
        # Filtrar columna contrato != vacio and != no aplica and no cruza 
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA') & (df['CONTRATO'] != 'NO CRUZAN')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('ELECTROCAQUETA', case=False, na=False)]
        
        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']
        
        df = df[columnas_a_mantener]
        
        return df
        
    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")

def read_excel_energiaputumayo():
    try:
        data = config['CARPETA_DATA']['RUTA_EXCEL']

        df = pd.read_excel(data, sheet_name='HOJA DE TRABAJO')

        columnas_a_mantener = ['Supplier', 'AVI', 'CONTRATO']

        # Limpiar columna CONTRATO
        # df['CONTRATO'] = df['CONTRATO'].str.replace('.', '', regex=False)

        # Filtrar columna contrato != vacio and != no aplica
        df = df[df['CONTRATO'].notna() & (df['CONTRATO'] != 'NO APLICA')]

        # Filtrar columna supplier para encontrar la comercializadora
        df = df[df['Supplier'].str.contains('EMPRESA DE ENERGIA DEL PUTUMAYO SA ESP', case=False, na=False)]

        # Filtrar columna AVI == si
        df = df[df['AVI'] == 'SI']

        df = df[columnas_a_mantener]

        return df

    except Exception as e:
        print(f"Ha ocurrido un error al leer el informe: {e}")

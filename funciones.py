import psycopg2
import toml
import win32com
import logging
import pandas as pd
import querys
import traceback as tr
import os
from sqlalchemy import create_engine
import datetime
config_path = "./config.toml"

with open(config_path, "r") as f:

    config = toml.load(f)


def conexion():
    """Función para crear conexión a base de datos de motor postgres"""

    config_bd = config["database"]

    try:
        # Conectarse a la base de datos PostgreSQL
        db_connect = psycopg2.connect(
                            host = config_bd[ "host"],
                            database = config_bd["database"],
                            user = config_bd["user"],
                            password = config_bd["password"]
        )
        
        return db_connect
    
    except:

        print("error")



def consulta_tienda_id(query,valor):
    """Función para realizar consulta en la base de datos de motor postgres"""

    try:    
        cnx = conexion()

        cnx = cnx.cursor()

        query_tienda_id = (query)

        cnx.execute(query_tienda_id,(valor,))

        existe = cnx.fetchall()

        return existe[0][0]
        
    except:
        None
        print("mal") #Retornar error



def consulta(query):
    """Función para realizar consulta en la base de datos de motor postgres"""

    try:    
        cnx = conexion()

        cursor = cnx.cursor()

        cursor.execute(query)

        resultado = cursor.fetchall()

        return resultado
        
    except:
        None
        print("mal")
        print(tr.format_exc())
    
    finally:
        # Asegurarse de cerrar el cursor y la conexión
        if cursor is not None:
            cursor.close()
        if cnx is not None:
            cnx.close()


def insertar_datos(insert_query,datos,tabla):

    if datos is not None and len(datos) > 0 :

        try:

            cnx = conexion()

            cursor = cnx.cursor()
            
            # Ejecutar la consulta SQL con los valores individuales
            cursor.execute(insert_query,datos)

            cnx.commit()

            print(f"Datos insertados exitosamente en la tabla {tabla}")

        except Exception as e:

            print(f"Error al insertar datos en la tabla {tabla}: {e}")
            print(tr.format_exc())
            cnx.rollback()

        finally:

            # Cerrar la conexión
            cursor.close()

            cnx.close()
    else:
        print(f'tupla {datos} vacia')
    



def conexion_correo(asunto_notificacion, cuerpo_notificacion):
    try:
       
        cuenta_correo = 'amaldonado@ises.com.co'
       
        # Inicializar cliente de correo #
        outlookApp = win32com.client.Dispatch("Outlook.Application")
        outlookNamespace = outlookApp.GetNamespace("MAPI")
       
        # Configurar bandeja/carpeta a revisar #
        inbox = outlookNamespace.Folders(cuenta_correo).Folders("Inbox")
       
        # Envío de correo de notificación #
        mailItem = outlookApp.CreateItem(0)  # Crear el ítem de correo
        mailItem.Subject = asunto_notificacion
        #mailItem.BodyFormat = formato_del_correo  # Asegúrate de que esto esté configurado adecuadamente
        mailItem.Body = cuerpo_notificacion
       
        # Agregar destinatarios #
        for correo in config['correo']['remitente']:
            mailItem.Recipients.Add(correo)
       
        mailItem.Save()
        mailItem.Send()
        logging.info("Correo enviado con éxito")
   
    except Exception as e:
        logging.error("Ocurrió un error al enviar el correo: %s", str(e))
        logging.exception("Detalles del error:")
 
 

def generar_arbol_carpetas(texto_fecha, comercializadora):
    """
    Genera la ruta de carpetas basándose en el texto_fecha y la comercializadora.
    """
    if '-' in texto_fecha:
        periodo = texto_fecha.split('-')[0].strip()
    else:
        periodo = texto_fecha[:3].strip()  # Asume formato 'MMM' para el mes
    
    mes_texto = periodo  # Directamente usar el texto del mes
    
    mes_numero = obtener_mes_numero(mes_texto)

    if not mes_numero:
        logging.error(f"No se pudo determinar el mes a partir de {texto_fecha}")
        return None

    anio_actual = datetime.datetime.now().year  # Obtener el año actual

    ruta = os.path.join('C:\\Users\\P108\\Documents\\PyDocto\\', str(comercializadora), str(anio_actual), f'{mes_numero:02d}')

    if not os.path.exists(ruta):
        os.makedirs(ruta)
        print(f"Directorio {ruta} creado!")
    else:
        print(f"Directorio {ruta} ya existe")

    return ruta
# funcion para generar arbol de carpetas / complemento para generar numero de mes de carperta
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

def renombrar_pdf(sap, supplier, texto_fecha):
    """
    Renombra el archivo PDF basado en el año y mes extraído de texto_fecha.
    """
   
    
    # Extraer el período y el año
    if '-' in texto_fecha:
        periodo = texto_fecha.split('-')[0].strip()
        anio = texto_fecha.split('-')[1].strip()  # Año después del guion
    else:
        periodo = texto_fecha[:7].strip()  # Asume el formato 'MM-YYYY' o 'MMM-YYYY'
        anio = texto_fecha[7:].strip()  # Año después del período
    
    # Extraer mes y año del período
    mes_texto = periodo[:3]  # Extrae el texto del mes
    anio = anio[:4]  # Asegura que el año tenga cuatro dígitos
    
    # Convertir el mes a número
    mes_numero = obtener_mes_numero(mes_texto)
    if mes_numero:
        mes_formateado = f"{mes_numero:02d}"  # Asegura que el mes tenga dos dígitos
        nuevo_nombre = f"{sap}_{supplier}_{anio}{mes_formateado}.pdf"
        return nuevo_nombre
    else:
        logging.error(f"No se pudo determinar el mes a partir de {texto_fecha}")
        return None
    

def mover_pdf(pdf_path, destino):
    """
    Mueve el archivo PDF a la nueva ubicación especificada por destino.
    """
    try:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        os.rename(pdf_path, destino)
        logging.info(f"PDF movido a: {destino}")
        return destino  # Devuelve la ruta final del archivo
    except Exception as e:
        eliminar_archivo_con_motivo(pdf_path, str(e))
        err = "Error al organizar factura: " + str(e)
        conexion_correo('Error de ejecución en el bot', err);
        return None

def eliminar_archivo_con_motivo(pdf_path, motivo):
    """
    Elimina el archivo PDF y muestra un mensaje indicando el motivo, solo con el nombre del archivo.
    """
    try:
        # Extrae solo el nombre del archivo desde la ruta
        nombre_archivo = os.path.basename(pdf_path)
        os.remove(pdf_path)
        print(f"Archivo eliminado: {nombre_archivo}. Motivo: {motivo}")
    except Exception as e:
        print(f"No se pudo eliminar el archivo {nombre_archivo}: {str(e)}")

def verificar_descarga(contrato, download_path):
    """
    Verifica si el contrato ya ha sido descargado consultando las carpetas.
    """
    pdf_path = os.path.join(download_path, f'{contrato}.pdf')
    return os.path.exists(pdf_path)

def registrar_descarga(contrato, pdf_path, log_file):
    """
    Registra el contrato descargado y la ubicación en el archivo de log.
    """
    with open(log_file, 'a') as log:
        log.write(f"{contrato}.pdf,{pdf_path}\n")

def extraer_anio(texto):
    # Asumimos que el año está compuesto por los últimos dos caracteres
    anio_abreviado = texto[-2:]
    
    # Convertimos el año abreviado en un año completo (asumiendo que está en el rango de 2000 a 2099)
    anio = int("20" + anio_abreviado)
    
    return anio
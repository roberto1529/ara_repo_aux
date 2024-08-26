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
    



def conexion_correo():
    
    asunto_notificacion = "Variables sin valor"
    formato_del_correo = f"""El presente correo es para informar que las siguientes variables no fueron encontradas 
    en el pdf (coloca nombre del pdf) perteneciente a la comercializadora: (nombre coercializadora)"""
    cuerpo_notificacion = 1
    
    try:

        cuenta_correo = 'calvarez@ises.com.co'
        # Inicializar cliente correo #
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Configurar bandeja/carpeta a revisar #
        inbox = outlook.Folders(cuenta_correo).Folders("Bandeja de entrada")

        # Envío de correo de notificación
        outlookApp = win32com.Dispatch("Outlook.Application").GetNameSpace("MAPI")

        mailItem = outlookApp.CreateItem(0)
        mailItem.Subject = asunto_notificacion
        mailItem.BodyFormat = formato_del_correo
        mailItem.Body = cuerpo_notificacion

        for correo in config["correo_remitente"]:
            mailItem.Recipients.Add(correo)

        mailItem.Save()
        mailItem.Send()
    
    except:
        logging.warning(
            f"Ocurrió un error",
            tr.format_exc(),
        )

# funcion para generar arbol de carpetas
def generar_arbol_carpetas(texto_fecha, comercializadora):
    if '-' in texto_fecha:
        mes_texto = texto_fecha.split('-')[0]
    else:
        mes_texto = texto_fecha[:3]

    mes_numero = obtener_mes_numero(mes_texto)

    if not mes_numero:
        return None

    anio_actual = datetime.datetime.now().year

    ruta = os.path.join('C:\\Users\\P108\\Documents\\PyDocto\\',  str(comercializadora),  str(anio_actual),   f'{mes_numero:02d}')

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

# def remonbrar_arc(archivo):
      # Esperar a que el archivo se descargue completamente
        # archivos_pdf = [f for f in os.listdir(archivo) if f.endswith('.pdf')]
        # archivos_doc = [f for f in archivos_pdf if 'doc.pdf' in f]
        
        # for archivo in archivos_doc:
        #     nuevo_nombre = f"{contrato}.pdf"
        #     nuevo_nombre_path = os.path.join(archivo, nuevo_nombre)

        #     # Evitar sobrescritura, agregar sufijo incremental si el archivo ya existe
        #     if os.path.exists(nuevo_nombre_path):
        #         base, ext = os.path.splitext(nuevo_nombre)
        #         contador = 1
        #         while os.path.exists(nuevo_nombre_path):
        #             nuevo_nombre = f"{base}_{contador}{ext}"
        #             nuevo_nombre_path = os.path.join(download_path, nuevo_nombre)
        #             contador += 1
        #     os.rename(os.path.join(download_path, archivo), nuevo_nombre_path)
            

        #     logging.info(f"Archivo renombrado a: {nuevo_nombre}")
       
        # if not archivos_doc:
        #     logging.warning("No se encontraron archivos PDF con el nombre 'doc.pdf'.")
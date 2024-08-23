import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
import shutil
import os
import subprocess
import toml
import traceback


# Cargue de la configuracion de TOML
with open("config.toml", "r") as f:
    config = toml.load(f)
    print(f"Carpeta donde esta watchdog => {config['CARPETA_FACTURAS']['ruta']}")
    

class ManejadorArchivos(FileSystemEventHandler):    
   
    # Funcion que detectar el movimiento en la carpeta Facturas.
    def on_created(self, event):
        
        nombre_archivo = event.src_path.split('\\')[-1]
        nombre_archivo_final = nombre_archivo[::-1]
        
        def validar_nombre_vatia(nombre):
            patron = config["PATRON"]["vatia"]
            if re.match(patron, nombre):
                return True
            else:
                return False
        
        def validar_nombre_neu(nombre):
            patron_neu = config["PATRON"]["neu"]
            if re.match(patron_neu, nombre):
                return True
            else:
                return False
            
        def validar_nombre_afinia(nombre):
            patron_afinia = config["PATRON"]["afinia"]
            if re.match(patron_afinia, nombre):
                return True
            else:
                return False
        
        if(validar_nombre_vatia(nombre_archivo_final[0]) == True and isinstance(nombre_archivo_final[0].split('.')[0], str) == True):
            try:
                if(os.path.exists(config["CARPETA_FACTURAS"]["carpeta_facturas_vatia"])):
                    os.chmod(config["CARPETA_FACTURAS"]["carpeta_facturas_vatia"], 0o777)
                    time.sleep(1)
                    shutil.move(event.src_path, f'C:\\Users\\P340\\ISES S.A.S\\Administrador App - Analitica de datos\\Proyecto ARA\\Facturasvatia\\{nombre_archivo_final[0]}')
                    print(f"Archivo fue movido a la direccion => {config['CARPETA_FACTURAS']['carpeta_facturas_vatia']}")
            except Exception as e:
                print("Error al mover archivo a vatia => ", e)
        elif(validar_nombre_neu(nombre_archivo_final[0]) == True):
            try:
                if(os.path.exists(config['CARPETA_FACTURAS']['carpeta_facturas_neu'])):
                    os.chmod(config["CARPETA_FACTURAS"]["carpeta_facturas_neu"], 0o777) #
                    time.sleep(2)
                    shutil.move(event.src_path, f'C:\\Users\\P340\\ISES S.A.S\\Administrador App - Analitica de datos\\Proyecto ARA\\Facturasneu\\{nombre_archivo_final[0]}')
                    print(f"Archivo se movio a la direccion {config['CARPETA_FACTURAS']['carpeta_facturas_neu']}")
            except Exception as e:
                print("Error al mover archivo a neu => ", e)
        elif(validar_nombre_afinia(nombre_archivo_final[0]) == True):
            try:
                if(os.path.exists(config["CARPETA_FACTURAS"]["carpeta_facturas_afinia"])):
                    os.chmod(config['CARPETA_FACTURAS']['carpeta_facturas_afinia'], 0o777) #
                    time.sleep(2)
                    shutil.move(event.src_path, f'C:\\Users\\P340\\ISES S.A.S\\Administrador App - Analitica de datos\\Proyecto ARA\\Facturasafinia\\{nombre_archivo_final[0]}')
                    print(f"Archivo fue movido a la direccion => {config['CARPETA_FACTURAS']['carpeta_facturas_afinia']}")
            except Exception as e:
                print(f"Error al mover archivo a afinia => ", e)         
        else:
            if(os.path.exists(config["CARPETA_FACTURAS"]["carpeta_facturas_otros"])):
                os.chmod(config["CARPETA_FACTURAS"]["carpeta_facturas_otros"], 0o777)
                time.sleep(1)
                shutil.move(event.src_path, f'C:\\Users\\P340\\ISES S.A.S\\Administrador App - Analitica de datos\\Proyecto ARA\\Facturasotros\\{nombre_archivo_final[0]}')

            
observador = Observer()
    
observador.schedule(ManejadorArchivos(), path=config["CARPETA_FACTURAS"]["ruta"], recursive=True)

observador.start()

try:
    while True:
        time.sleep(1)
        
except KeyboardInterrupt as e:
    observador.stop()
    
observador.join()
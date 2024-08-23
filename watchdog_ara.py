import re
import os
import shutil
import time
import toml
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# Cargue de la configuracion de TOML
with open("config.toml", "r") as f:
    config = toml.load(f)
    print(f"Carpeta donde esta watchdog => {config['CARPETA_FACTURAS']['ruta']}")

class ManejadorArchivos(FileSystemEventHandler):
    """definir métodos para validar nombres de archivos basados en patrones específicos"""    

    def on_created(self, event):
        nombre_archivo = event.src_path.split('\\')[-1]
        
        if self.validar_nombre_vatia(nombre_archivo):
            self.mover_archivo(event.src_path, config["CARPETA_FACTURAS"]["carpeta_facturas_vatia"], nombre_archivo)
        elif self.validar_nombre_neu(nombre_archivo):
            self.mover_archivo(event.src_path, config["CARPETA_FACTURAS"]["carpeta_facturas_neu"], nombre_archivo)
        elif self.validar_nombre_afinia(nombre_archivo):
            self.mover_archivo(event.src_path, config["CARPETA_FACTURAS"]["carpeta_facturas_afinia"], nombre_archivo)
        else:
            self.mover_archivo(event.src_path, config["CARPETA_FACTURAS"]["carpeta_facturas_otros"], nombre_archivo)
    
    def validar_nombre_vatia(self, nombre):
        patron = config["PATRON"]["vatia"]
        return re.match(patron, nombre) is not None
    
    def validar_nombre_neu(self, nombre):
        patron_neu = config["PATRON"]["neu"]
        return re.match(patron_neu, nombre) is not None
    
    def validar_nombre_afinia(self, nombre):
        patron_afinia = config["PATRON"]["afinia"]
        return re.match(patron_afinia, nombre) is not None

    def mover_archivo(self, src_path, dest_folder, nombre_archivo):
        try:
            if os.path.exists(dest_folder):
                os.chmod(dest_folder, 0o777)#permiso para leer, escribir y ejecutar
                time.sleep(1)
                shutil.move(src_path, os.path.join(dest_folder, nombre_archivo))
                print(f"Archivo movido a {dest_folder}")
        except Exception as e:
            print(f"Error al mover archivo a {dest_folder}: {e}")

observador = Observer()
observador.schedule(ManejadorArchivos(), path=config["CARPETA_FACTURAS"]["ruta"], recursive=True)

observador.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observador.stop()

observador.join()

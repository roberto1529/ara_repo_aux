import os
import pandas as pd
import re
import toml
import unicodedata

from datetime import datetime
from PyPDF2 import PdfReader



config_path = './config.toml'

with open(config_path, 'r') as f:

    cfg = toml.load(f)

def strip_accents(s):
   s=s.lower()
   return ''.join(c for c in unicodedata.normalize('NFD', s)
                  if unicodedata.category(c) != 'Mn')

def limpiar_texto(texto):
    texto = texto.lower()  # Convertir a minúsculas
    texto = unicodedata.normalize('NFKD', texto)  # Normalizar texto
    texto = ''.join([c for c in texto if not unicodedata.combining(c)])  # Eliminar tildes
    return texto



directory = './ModeloFacturas'

# Recorrer todos los archivos en el directorio
def factura_vatia():

    #Recorrer cada pdf que se encuentra en la carpeta
    for filename in os.listdir(directory):

        if filename.endswith(".pdf"):

            filepath = os.path.join(directory, filename)                 

            print(f"Leyendo archivo: {filename}")

            #texto = leer_pdf(filepath)

            with open(filepath, 'rb'):
                
                lector_pdf = PdfReader(filepath)
                
                #Extraer solo las paginas 2 y 3.
                texto=lector_pdf.pages[1].extract_text()  
            
            #eliminar tildes y convertir todo texto a minusculas.
            texto=strip_accents(texto)
            #texto = limpiar_texto(texto)

            print(f"Texto extraído del archivo {filename}:\n{texto}\n")

            #calcular periodo actual.
            periodo = datetime.now().strftime('%Y-%m-%d %H:%M')
            print(f'periodo actual: {periodo}')

            ## Extraer información comercial ##
            #extraer nic
            nic = texto[:5]
            print(f'nic: {nic}')

            ## Extraer infromación de factura ##
            #extraer periodo de facturación
            periodo_fact = re.search(r'\d{2}/\d{2}/\d{4} a \d{2}/\d{2}/\d{4}\s*(\d+)', texto)
            periodo_fact = periodo_fact.group(0) if periodo_fact else None
            print(periodo_fact)

            #extraer periodo de facturación inicial
            periodo_fact_inicial = periodo_fact[:11]
            print(f'periodo facturado inicial: {periodo_fact_inicial}')

            #extraer periodo de facturación final
            periodo_fact_final = periodo_fact[13:23]
            print(f'periodo facturado final: {periodo_fact_final}')

            #extraer días de facturación
            dias_fact = periodo_fact[24:]
            print(f'dias facturados: {dias_fact}')
            

            ## Extraer infromación de consumo ##
            #extraer consumo activa
            csmo_act = texto[15:25]
            #print(f'consumo ativa: {csmo_act}')

            #extraer consumo reactiva
            csmo_react = texto[15:25]
            #print(f'csmo reactiva: {csmo_react}')

            #extraer consumo activa
            csmo_reliq_act = texto[15:25]
            #print(f'csmo_reliq_act: {csmo_reliq_act}')

            csmo_reliq_react = texto[15:25]
            #print(f'csmo_reliq_react: {csmo_reliq_react}')

            ##Detalles de cobro de energía ##
            datos = {}

                #
            patron_activa = re.compile(r'activa kwh\s+([\d.]+)\s+\$\s+([\d.,]+)\s+\$\s+([\d.,]+)')
            match_activa = patron_activa.search(texto)

            if match_activa:
                datos['activa_kwh'] = {
                    'cantidad': match_activa.group(1),
                    'valor_energia': match_activa.group(2),
                    'valor_contribucion': match_activa.group(3)}

            patron_ind_facturada = re.compile(r'reactiva ind facturada kvarh\s+([\d.]+)\s+([\d.]+)\s+\$\s+([\d.,]+)\s+\$\s+([\d.,]+)')
            match_ind_facturada = patron_ind_facturada.search(texto)

            if match_ind_facturada:
                datos['reactiva_ind_facturada_kvarh'] = {
                    'cantidad': match_ind_facturada.group(1),
                    'factor': match_ind_facturada.group(2),
                    'valor_energia': match_ind_facturada.group(3),
                    'valor_contribucion': match_ind_facturada.group(4)
                }

            patron_capacitiva = re.compile(r'reactiva capacitiva kvarh\s+([\d.]+)\s+([\d.]+)\s+\$\s+([\d.,]+)\s+\$\s+([\d.,]+)')
            match_capacitiva = patron_capacitiva.search(texto)

            if match_capacitiva:
                datos['reactiva_capacitiva_kvarh'] = {
                    'cantidad': match_capacitiva.group(1),
                    'factor': match_capacitiva.group(2),
                    'valor_energia': match_capacitiva.group(3),
                    'valor_contribucion': match_capacitiva.group(4)
                }
            subtotal_energia = re.compile(r'subtotal energia \s+\$\s+([\d.,]+)\s+\$\s+([\d.,]+)')
            match_subtotal = subtotal_energia.search(texto)
                
            if match_subtotal:
                datos['subtotal_energia'] = {
                    'valor_energia': match_subtotal.group(1),
                    'valor_contribucion': match_subtotal.group(2)
                }

            energia_contribucion = re.compile(r'subtotal energia \s+\$\s+([\d.,]+)')
            match_ene_contr= energia_contribucion.search(texto)

            if match_ene_contr:
                datos['energia_contribucion'] = {
                    'valor_contribucion': match_ene_contr.group(1)
                }

            print(datos)



    return texto


factura_vatia()


def factura_neu():

    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            filepath = os.path.join(directory, filename)
            print(f"Leyendo archivo: {filename}")
            texto = leer_pdf(filepath)
            print(f"Texto extraído del archivo {filename}:\n{texto}\n")
            texto.lower()
            print(f'prueba: {texto[:5]}')
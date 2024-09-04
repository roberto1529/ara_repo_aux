import logging
import os
import re
import unicodedata
from pdfminer.high_level import extract_text
import time
from os import system

# Configurar el logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def strip_accents(s):
    """Elimina acentos de un texto y lo convierte a minúsculas."""
    s = s.lower()
    return "".join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def formatear_numero(numero):
    """Formatea el número eliminando puntos y convirtiendo comas en puntos."""
    numero = numero.replace('.', '').replace(',', '.')
    numero = numero.split('.')[0]
    return numero

def extraer_datos(texto):
    """Extrae datos clave del texto utilizando expresiones regulares mejoradas."""
    # Limpiar el texto
    texto = strip_accents(texto)
    
    # Buscar el periodo facturado
    periodo_extract = re.compile(r'periodo\s+facturado\s+(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE | re.MULTILINE)
    match_periodo_fact = periodo_extract.search(texto)
    if match_periodo_fact:
        fecha_inicio = match_periodo_fact.group(1)
        fecha_fin = match_periodo_fact.group(2)
    else:
        fecha_inicio = 'No encontrado'
        fecha_fin = 'No encontrado'
    
    # Extraer el valor total alumbrado
    valor_total_alumbrado_re = re.compile(r'valor\s+total\s+alumbrado\s*([\d.,]+)', re.IGNORECASE)
    valor_total_alumbrado = valor_total_alumbrado_re.search(texto)
    if valor_total_alumbrado:
        valor_total_alum = valor_total_alumbrado.group(1).replace('.', '').replace(',', '.')
        valor_total_alum = valor_total_alum.split('.')[0]
    else:
        valor_total_alum = 'No encontrado'

    # Extraer sección de datos clave
    seccion_datos = re.search(r'consumo\s*reactiva\s*contribucion\s*aproximacion a decenas\s*(.*)', texto, re.DOTALL)
    
    if seccion_datos:
        seccion_datos = seccion_datos.group(1)
      
        # Expresión regular para encontrar el valor de 'Consumo'
        consumo_re = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', seccion_datos)
        if consumo_re:
            consumo = formatear_numero(consumo_re.group(1))
        else:
            consumo = 'No encontrado'
        
         # Buscar todos los valores numéricos en la sección
        # Obtener todos los valores numéricos con formato adecuado
        valores_re = re.findall(r'(\d{1,3}(?:\.\d{3})*,\d{2})', seccion_datos)
       

        # Selecciona el segundo valor de la lista si existe
        if len(valores_re) >= 2:
            # El segundo valor será la contribución
            contribucion = formatear_numero(valores_re[1])  # Segundo valor en la lista
        else:
            contribucion = 'No encontrado'

    else:
        consumo = 'No encontrado'
        contribucion = 'No encontrado'
    
    # Extraer "Fecha lectura actual"
    fecha_lectura_actual_re = re.compile(r'fecha\s+lectura\s+actual:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
    fecha_lectura_actual = fecha_lectura_actual_re.search(texto)
    fecha_lectura_actual = fecha_lectura_actual.group(1) if fecha_lectura_actual else 'No encontrada'
    logging.info(f'Fecha de lectura actual: {fecha_lectura_actual}')

    # Extraer "Días facturados"
    dias_facturados_re = re.compile(r'dias\s+facturados\s+(\d+)\s+estimado', re.IGNORECASE)
    dias_facturados = dias_facturados_re.search(texto)
    dias_facturados = dias_facturados.group(1) if dias_facturados else 'No encontrado'
    logging.info(f'Días facturados: {dias_facturados}')

    # Extraer "Lectura anterior" para Activa BT y Reactiva BT
    lectura_anterior_re = re.compile(r'lectura\s+anterior\s+(\d+)\s+(\d+)', re.IGNORECASE)
    lectura_anterior = lectura_anterior_re.search(texto)
    if lectura_anterior:
        lectura_anterior_activa = lectura_anterior.group(1)
        lectura_anterior_reactiva = lectura_anterior.group(2)
    else:
        lectura_anterior_activa = 'No encontrada'
        lectura_anterior_reactiva = 'No encontrada'
    logging.info(f'Lectura anterior Activa BT: {lectura_anterior_activa}')
    logging.info(f'Lectura anterior Reactiva BT: {lectura_anterior_reactiva}')

    # Extraer "Factor múltiplo" para Activa BT y Reactiva BT
    factor_multiplo_re = re.compile(r'factor\s+multiplo\s+(\d+)\s+(\d+)', re.IGNORECASE)
    factor_multiplo = factor_multiplo_re.search(texto)
    if factor_multiplo:
        factor_multiplo_activa = factor_multiplo.group(1)
        factor_multiplo_reactiva = factor_multiplo.group(2)
    else:
        factor_multiplo_activa = 'No encontrado'
        factor_multiplo_reactiva = 'No encontrado'
    logging.info(f'Factor múltiplo Activa BT: {factor_multiplo_activa}')
    logging.info(f'Factor múltiplo Reactiva BT: {factor_multiplo_reactiva}')

    # Extraer "Consumo kWh" para Activa BT y Reactiva BT
    consumo_kwh_re = re.compile(r'consumo\s+kwh\s+(\d+)\s+(\d+)', re.IGNORECASE)
    consumo_kwh = consumo_kwh_re.search(texto)
    if consumo_kwh:
        consumo_kwh_activa = consumo_kwh.group(1)
        consumo_kwh_reactiva = consumo_kwh.group(2)
    else:
        consumo_kwh_activa = 'No encontrado'
        consumo_kwh_reactiva = 'No encontrado'
    logging.info(f'Consumo kWh Activa BT: {consumo_kwh_activa}')
    logging.info(f'Consumo kWh Reactiva BT: {consumo_kwh_reactiva}')

    # Extraer número del medidor
    medidor_re = re.compile(r'medidor\s+(\d+)', re.IGNORECASE)
    medidor = medidor_re.search(texto)
    numero_medidor = medidor.group(1) if medidor else 'No encontrado'
    logging.info(f'Número del medidor: {numero_medidor}')


   # Texto fragmentado proporcionado
    texts = """
    20/08/2024

    21/08/2024
    1
    $ 0
    01/08/2024
    41102408001508

    1000074059 - 89

    suspension a partir de: 
    """

  # Expresión regular para encontrar las fechas que preceden a "suspension a partir de:"
    patron = r'(\d{2}/\d{2}/\d{4})\s*(\d{2}/\d{2}/\d{4})\s*[\d\D]*?suspension\s+a\s+partir\s+de:'

    # Buscar las fechas utilizando la expresión regular
    coincidencias = re.search(patron, texto, re.IGNORECASE | re.DOTALL)

    if coincidencias:
        fechalimite, fechaSuspencion = coincidencias.groups()

    else:
        print("Fechas de suspensión no encontradas.")
    
     # Expresión regular para capturar el valor de "cu"
    patron_cu = re.compile(r'cu\s([\d,.]+)')
    coincidencia_cu = patron_cu.search(texto)
    valor_cu = coincidencia_cu.group(1) if coincidencia_cu else None

    # Expresión regular para capturar las etiquetas y los valores
    patron_etiquetas = re.compile(r'([a-z]+)\s*\n')
    patron_valores = re.compile(r'(\d+,\d+)\s*\n')

    # Buscar todas las coincidencias en el texto
    etiquetas = patron_etiquetas.findall(texto)
    valores = patron_valores.findall(texto)

    # Convertir listas de etiquetas y valores en un diccionario
    etiquetas_valores = dict(zip(etiquetas, valores))

    # Listar todas las etiquetas en el orden deseado
    orden_deseado = ['g', 't', 'pr', 'r', 'd', 'c']

    # Ordenar los valores en el orden de las etiquetas deseadas
    resultados_ordenados = {etiqueta: etiquetas_valores.get(etiqueta, 'No encontrado') for etiqueta in orden_deseado}

    # Imprimir los resultados en el formato deseado
    for etiqueta in orden_deseado:
        print(f'valor_{etiqueta}', resultados_ordenados[etiqueta])
        
    print("Ancla")

    print('General -> ', texto)
    
    return {
        "Fecha de inicio en periodo facturado": fecha_inicio,
        "Fecha de fin en periodo facturado": fecha_fin,
        "Valor Total Alumbrado": valor_total_alum,
        "Contribucion": contribucion,
        "Consumo": consumo,
        "fecha_lectura_actual": fecha_lectura_actual,
        "dias_facturados": dias_facturados,
        "lectura_anterior_activa": lectura_anterior_activa,
        "lectura_anterior_reactiva": lectura_anterior_reactiva,
        "factor_multiplo_activa": factor_multiplo_activa,
        "factor_multiplo_reactiva": factor_multiplo_reactiva,
        "consumo_kwh_activa": consumo_kwh_activa,
        "consumo_kwh_reactiva": consumo_kwh_reactiva,
        "numero_medidor": numero_medidor,
        "Fecha oporturna de pago": fechalimite,
        "Fecha de suspención de servicio": fechaSuspencion,
        "tarifa (CU)":valor_cu
    }

def procesar_pdf(directory):
    """Procesa cada archivo PDF en el directorio especificado."""
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            filepath = os.path.join(directory, filename)
            logging.info(f"Leyendo archivo: {filename}")
            try:
                with open(filepath, 'rb') as file:
                    texto = extract_text(file)

                # Extraer datos clave
                datos = extraer_datos(texto)
                logging.info(f'Datos extraídos: {datos} \n')

            except Exception as e:
                logging.error(f'Error procesando el archivo {filename}: {str(e)}')

            break  # Eliminar o comentar esta línea si deseas procesar todos los archivos PDF en el directorio

# Directorio donde están los archivos PDF
directory = 'C:\\Users\\P108\\Documents\\ARA_PY'

# Procesar los archivos PDF
if __name__ == "__main__":
    try:
        system("cls")
        time.sleep(2)
        procesar_pdf(directory)
    except Exception as e:
        print("Error", e)
# RC

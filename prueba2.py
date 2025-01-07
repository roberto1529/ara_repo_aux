import os
import pyodbc
import pandas as pd
import logging
from dotenv import load_dotenv
# link para descargar contralador : https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#download-for-windows

# Configurar logging
logging.basicConfig(
    filename='analisis_irregularidades.log',
    level=logging.INFO,
    format='%(asctime)s:%(levelname)s:%(message)s'
)

def cargar_variables_entorno():
    """Cargar variables de entorno desde el archivo .env."""
    load_dotenv()
    return {
        'server': os.getenv('DB_SERVER'),
        'database': os.getenv('DB_DATABASE'),
        'username': os.getenv('DB_USERNAME'),
        'password': os.getenv('DB_PASSWORD'),
        'port': os.getenv('DB_PORT'),
        'secret_key': os.getenv('SECRET_KEY')
    }

def conectar_bd(config):
    """Establecer conexión con la base de datos SQL Server."""
    try:
        conn = pyodbc.connect(
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={config["server"]},{config["port"]};'
            f'DATABASE={config["database"]};'
            f'UID={config["username"]};'
            f'PWD={config["password"]}'
        )
        logging.info("Conexión a la base de datos establecida con éxito.")
        return conn
    except pyodbc.Error as e:
        logging.error(f"Error al conectar a la base de datos: {e}")
        raise

def obtener_datos(conn, consulta):
    """Ejecutar una consulta SQL y retornar un DataFrame."""
    try:
        df = pd.read_sql(consulta, conn)
        logging.info(f"Consulta ejecutada exitosamente: {consulta}")
        return df
    except Exception as e:
        logging.error(f"Error al ejecutar la consulta: {e}")
        raise

def procesar_irregularidades(df_acta):
    """Procesar datos para calcular probabilidades y clientes recurrentes."""
    try:
        # Seleccionar columnas relevantes
        columnas = [
            'Póliza',
            '¿Se_encuentra_irregularidad_en_acometida?',
            '¿Se_encuentra_irregularidad_en_medidor?'
        ]
        df_irregularidades = df_acta[columnas].copy()

        # Convertir valores 'Si'/'No' a binarios
        df_irregularidades['irregularidad_acometida'] = df_irregularidades['¿Se_encuentra_irregularidad_en_acometida?'].map({'Si': 1, 'No': 0})
        df_irregularidades['irregularidad_medidor'] = df_irregularidades['¿Se_encuentra_irregularidad_en_medidor?'].map({'Si': 1, 'No': 0})

        # Manejar valores faltantes
        df_irregularidades['irregularidad_acometida'].fillna(0, inplace=True)
        df_irregularidades['irregularidad_medidor'].fillna(0, inplace=True)

        # Calcular total de irregularidades por registro
        df_irregularidades['total_irregularidades'] = df_irregularidades[['irregularidad_acometida', 'irregularidad_medidor']].sum(axis=1)

        # Calcular probabilidad promedio de irregularidades por cliente
        df_probabilidad = df_irregularidades.groupby('Póliza')['total_irregularidades'].mean().reset_index()
        df_probabilidad.rename(columns={'total_irregularidades': 'Probabilidad de Irregularidades'}, inplace=True)

        # Identificar clientes recurrentes con irregularidades
       # Contar clientes recurrentes con irregularidades
        df_recurrentes_count = df_irregularidades[df_irregularidades['total_irregularidades'] > 0] \
            .groupby('Póliza').size().reset_index(name='Cantidad de Irregularidades')

        # Ordenar clientes por mayor probabilidad de irregularidades
        df_top_clientes = df_probabilidad.sort_values(by='Probabilidad de Irregularidades', ascending=False)

        logging.info("Procesamiento de datos completado exitosamente.")
        return df_probabilidad, df_recurrentes_count, df_top_clientes

    except Exception as e:
        logging.error(f"Error al procesar los datos: {e}")
        raise

def generar_informes(df_prob, df_recurrentes, df_top_clientes):
    """Guardar los DataFrames procesados en archivos CSV."""
    try:
        df_prob.to_csv('probabilidad_irregularidades.csv', index=False)
        df_recurrentes.to_csv('clientes_recurrentes.csv', index=False)
        df_top_clientes.to_csv('top_clientes_irregularidades.csv', index=False)
        logging.info("Informes generados y guardados correctamente.")
    except Exception as e:
        logging.error(f"Error al generar los informes: {e}")
        raise

def main():
    """Función principal para ejecutar el flujo completo."""
    try:
        # Cargar configuraciones
        config = cargar_variables_entorno()

        # Conectar a la base de datos
        conn = conectar_bd(config)

        # Consultas SQL optimizadas (seleccionar solo columnas necesarias)
        query_inspeccion_tecnica = """
            SELECT  * from ForMapDW.TRIPLEA.Vista_InspeccionTecnica vit
        """

        query_acta_inspeccion = """
           SELECT  * from ForMapDW.TRIPLEA.Vista_ActaInspeccion vai
        """

        # Obtener datos
        df_inspeccion_tecnica = obtener_datos(conn, query_inspeccion_tecnica)
        df_acta_inspeccion = obtener_datos(conn, query_acta_inspeccion)

        # Cerrar la conexión
        conn.close()
        logging.info("Conexión a la base de datos cerrada.")

        # Procesar datos
        df_probabilidad, df_recurrentes, df_top_clientes = procesar_irregularidades(df_acta_inspeccion)

        # Generar informes
        generar_informes(df_probabilidad, df_recurrentes, df_top_clientes)

        print("Informe generado con éxito.")
        logging.info("Script ejecutado correctamente.")

    except Exception as e:
        logging.error(f"Error en la ejecución del script: {e}")
        print("Ocurrió un error durante la ejecución. Revisa el log para más detalles.")

if __name__ == "__main__":
    main()

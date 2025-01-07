import psycopg2
import time  # Opcional: para simular un pequeño retraso y que se vea mejor el progreso

def actualizar_estado_indicador(lista_indicadores):
    # Conexión a la base de datos con la configuración proporcionada
    conn = psycopg2.connect(
        host="35.83.166.238", 
        database="ciap", 
        user="ciapapp", 
        password="C1@p1S3s*+20584",
        port="33500"  # Especifica el puerto correctamente
    )
    
    try:
        # Crear un cursor
        cursor = conn.cursor()

        total_indicadores = len(lista_indicadores)
        indicador_procesado = 0
        
        # Loop para cada indicador_proyecto_id en la lista
        for indicador_proyecto_id in lista_indicadores:
            # Loop para los periodos del 1 al 9
            for periodo_id in range(1, 10):
                # Consulta de actualización
                query = """
                UPDATE data.indicador_resultado
                SET estado = 0
                WHERE indicador_proyecto_id = %s
                AND periodo_id = %s
                AND id != (
                    SELECT id
                    FROM data.indicador_resultado ir
                    WHERE ir.indicador_proyecto_id = %s
                    AND ir.periodo_id = %s
                    AND estado = 1
                    LIMIT 1
                );
                """
                # Ejecutar la consulta
                cursor.execute(query, (indicador_proyecto_id, periodo_id, indicador_proyecto_id, periodo_id))

            # Incrementar el contador de indicadores procesados
            indicador_procesado += 1

            # Calcular el porcentaje completado
            porcentaje_completado = (indicador_procesado / total_indicadores) * 100

            # Mostrar progreso en consola
            print(f"Indicador {indicador_procesado}/{total_indicadores} procesado ({porcentaje_completado:.2f}% completado). ID indicador: {indicador_proyecto_id}")
            
            # Opcional: para ralentizar un poco el proceso y que sea más visible
            # time.sleep(0.5)

        # Confirmar la transacción
        conn.commit()

        print("Actualización completada para todos los indicadores en la lista.")

    except Exception as e:
        print(f"Error al actualizar: {e}")
        conn.rollback()  # Revertir la transacción en caso de error

    finally:
        # Cerrar el cursor y la conexión
        cursor.close()
        conn.close()

# Lista de indicador_proyecto_id que quieres procesar
lista_indicadores = [
  147,148,109,96,79,80,81,83,84,85,86,87,88,89,91,92,93,94,95,98,99,100,101,102,103,105,106,107,108,110,111,113,114,116,119,120,121,123,
  126,127,128,129,130,131,132,134,135,136,139,141,144,145,146,152,153,154,155,151,124,140,125,138,142,150,143,149,82,104,97,90,112,133,
  137,115,117,504,122,464,468,472,476,480,484,488,492,496,500,508,118,560,512,516,520,524,528,532,536,540,544,548,552,564,568,572,576,580
]

# Llamada a la función con la lista
actualizar_estado_indicador(lista_indicadores)

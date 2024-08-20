from datetime import datetime

def es_factura_de_julio(fecha_inicio, fecha_fin):
    # Convertir las fechas a objetos datetime
    fecha_inicio = datetime.strptime(fecha_inicio, "%d/%m/%y")
    fecha_fin = datetime.strptime(fecha_fin, "%d/%m/%y")
    
    # Definir los límites del mes de julio
    inicio_julio = datetime(fecha_inicio.year, 7, 1)
    fin_julio = datetime(fecha_fin.year, 7, 31)
    
    
    
    # Verificar si el periodo cae dentro de julio
    if fecha_inicio <= fin_julio and fecha_fin <= fin_julio and fecha_inicio >= datetime(fecha_inicio.year, 6, 1):
        return True
    return False

# Ejemplos de uso
facturas = [
    ("29/06/24", "20/07/24"),  # Pertenece a julio
    ("25/06/24", "25/07/24"),  # Pertenece a julio
    ("23/06/24", "23/07/24"),  # Pertenece a julio
    ("30/06/24", "30/07/24"),  # Pertenece a julio
    ("29/07/24", "29/08/24"),  # No pertenece a julio
]

for factura in facturas:
    if es_factura_de_julio(factura[0], factura[1]):
        print(f"Descargando factura con periodo {factura[0]} a {factura[1]}")
    else:
        print(f"Factura con periodo {factura[0]} a {factura[1]} no pertenece a julio")

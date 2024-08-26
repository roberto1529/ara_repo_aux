

config_path = "./config.toml"

query_tienda = f'SELECT * FROM public.tienda'
query_tienda_id = f'SELECT * FROM public.tienda where numero_cuenta = %s'
query_comercializadora_id = f'SELECT * FROM public.comercializadora where comercialziadora_id = %s'
query_cobros = f'SELECT * FROM public.info_cobros'


query_comercialziadora = f'SELECT comercializadora_id, nombre FROM public.tienda'


insertar_datos_factura = """
INSERT INTO info_factura (
    numero_factura, tienda_id, periodo_factura, fecha_lectura_anterior,
    fecha_lectura_actual, dias_facturados, fecha_expedicion, fecha_vencimiento,
    facturas_vencidas, motivo_corte, fecha_suspension, consumo_facturado,
    valor_a_pagar, financiaciones_pendiente, monto_financiacion, tasa_mora, total_compensacion
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""


insertar_datos_consumo = """
INSERT INTO info_consumo (
    tienda_id, numero_factura, numero_medidor, factor, consumo_activo,
    consumo_estimado, consumo_reactivo, consumo_capacitiva, consumo_inductiva,
    consumo_activo_reliquidado, consumo_reactivo_reliquidado, causa_reliquidacion
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

insertar_datos_cobros = """
INSERT INTO info_cobros (
    tienda_id, numero_factura, energia_activa, energia_reactiva, energia_capacitiva,
    energia_inductiva, energia_penalizada, subtotal_energia, contribucion_activa,
    contribucion_reactiva, subtotal_contribucion, total_energia_contribucion,
    plan_datos, iva_plan_datos, total_alumbrado_publico, total_aseo, impuesto_seguridad,
    total_otros_cobros, subtotal_a_pagar, ajuste_decena, valor_total_pagar
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""                        


insertar_datos_tarifa = """
INSERT INTO info_tarifa (
    tienda_id, numero_factura, mercado, clase_servicio, nivel_tension,
    propiedad_activo, generacion, transmision, distribucion, comercializacion, 
    perdidas, restricciones, cot, costo_unitario, cu_cot, contribucion
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

insertar_datos_comercializadora = """
INSERT INTO comercialziadora (
    numero_factura, tienda_id, periodo_factura, fecha_lectura_anterior,
    fecha_lectura_actual, dias_facturados, fecha_expedicion, fecha_vencimiento,
    facturas_vencidas, motivo_corte, fecha_suspension, consumo_facturado,
    valor_a_pagar, financiaciones_pendiente, monto_financiacion, tasa_mora, total_compensacion
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)

"""

insertar_datos_tienda = """
INSERT INTO tienda (
    nombre, nit, numero_cuenta, comercializadora_id, estado, zona, area_fisica, fecha_creacion,
    rango_facturacion, tipo_tienda, departamento, municipio, direccion_suministro,
    direccion_envio, latitud, longitud, fecha_de_energizacion,codigo_sic, ceco, codigo_sap, operador_de_red,
    eliminado
    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

insertar_descarga_factura = """
INSERT INTO descarga_facturas (
    tienda_id, numero_factura, fecha_factura, estado_descarga, fecha_descarga
    ) VALUES (%s, %s, %s, %s, %s)
"""


## BD PRUEBAS

# config_path = "./config.toml"

# query_tienda = f'SELECT * FROM "Tienda"."Tienda" t'
# query_tienda_id = f'SELECT * FROM "Tienda"."Tienda" t where numero_cuenta = %s'
# query_comercializadora_id = f'SELECT * FROM "Comercializadora" c where comercializadora_id = %s'
# query_cobros = f'SELECT * FROM public.info_cobros'


# query_comercialziadora = f'SELECT comercializadora_id, nombre FROM public.tienda'


# insertar_datos_factura = """
# INSERT INTO "Tienda"."InfoFactura" if2 (
#     "periodo_factura","fecha_lectura_anterior","fecha_lectura_actual","dias_facturados",
    # "fecha_expedicion","fecha_vencimiento","causa_no_lectura","facturas_vencidad",
    # "motivo_corte","fecha_suspension","consumo_facturado","valora_pagar",
    # "financiamiento_pendiente","monto_financiacion","tasa_mora","total_compensacion",
    # "eliminado","fecha_sistema","fecha_fact_inicial","fecha_fact_final"
# ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
# """


# insertar_datos_consumo = """
# INSERT INTO info_consumo (
#     tienda_id, numero_factura, numero_medidor, factor, consumo_activo,
#     consumo_estimado, consumo_reactivo, consumo_capacitiva, consumo_inductiva,
#     consumo_activo_reliquidado, consumo_reactivo_reliquidado, causa_reliquidacion
# ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
# """

# insertar_datos_cobros = """
# INSERT INTO info_cobros (
#     tienda_id, numero_factura, energia_activa, energia_reactiva, energia_capacitiva,
#     energia_inductiva, energia_penalizada, subtotal_energia, contribucion_activa,
#     contribucion_reactiva, subtotal_contribucion, total_energia_contribucion,
#     plan_datos, iva_plan_datos, total_alumbrado_publico, total_aseo, impuesto_seguridad,
#     total_otros_cobros, subtotal_a_pagar, ajuste_decena, valor_total_pagar
# ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
# """                        


# insertar_datos_tarifa = """
# INSERT INTO info_tarifa (
#     tienda_id, numero_factura, mercado, clase_servicio, nivel_tension,
#     propiedad_activo, generacion, transmision, distribucion, comercializacion, 
#     perdidas, restricciones, cot, costo_unitario, cu_cot, contribucion
# ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
# """

# insertar_datos_comercializadora = """
# INSERT INTO "Tienda"."Comercializadora" c (
#     "nombre","estado","url_descarga_factura","usuario_descarga",
    # "password_descarga","fecha_ultimo_cambio","eliminado"
# ) VALUES (%s, %s, %s, %s, %s, %s, %s)

# """

# insertar_datos_tienda = """
# INSERT INTO tienda (
#     nombre, nit, numero_cuenta, comercializadora_id, estado, zona, area_fisica, fecha_creacion,
#     rango_facturacion, tipo_tienda, departamento, municipio, direccion_suministro,
#     direccion_envio, latitud, longitud, fecha_de_energizacion,codigo_sic, ceco, codigo_sap, operador_de_red,
#     eliminado
#     ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
# """

# insertar_descarga_factura = """
# INSERT INTO descarga_facturas (
#     tienda_id, numero_factura, fecha_factura, estado_descarga, fecha_descarga
#     ) VALUES (%s, %s, %s, %s, %s)
# """
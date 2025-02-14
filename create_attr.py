#!/usr/bin/python3

from xmlrpc import client
import openpyxl

# Configuración de conexión
url = 'http://localhost:8269'  # URL del servidor Odoo
dbname = 'pos_name_TEST'  # Nombre de la base de datos
user = 'admin'  # Usuario de Odoo
pwd = 'admin'  # Contraseña del usuario de Odoo

# Autenticación con Odoo
common = client.ServerProxy(f'{url}/xmlrpc/2/common')  # Conexión al endpoint de autenticación de Odoo
uid = common.authenticate(dbname, user, pwd, {})  # Autenticación con Odoo, obteniendo el ID de usuario
models = client.ServerProxy(f'{url}/xmlrpc/2/object', allow_none=True)  # Conexión al endpoint de modelos de Odoo

# Función para crear un atributo y su valor
def crear_atributo_valores(attr_name, valores):
    attr_id = None
    value_ids = []

    # Crear o buscar el atributo
    attr_id = models.execute_kw(dbname, uid, pwd, 'product.attribute', 'search', [[['name', '=', attr_name]]])
    if not attr_id:  # Si el atributo no existe, se crea uno nuevo
        attr_id = models.execute_kw(dbname, uid, pwd, 'product.attribute', 'create', [{'name': attr_name}])
        print(f"[INFO] Atributo creado: {attr_name} con ID {attr_id}")
    else:  # Si existe, se obtiene el ID del atributo existente
        attr_id = attr_id[0]
        print(f"[INFO] Atributo encontrado: {attr_name} con ID {attr_id}")

    # Procesar cada valor y agregarlo al atributo
    for attr_value in valores:
        # Crear o buscar cada valor del atributo
        value_id = models.execute_kw(dbname, uid, pwd, 'product.attribute.value', 'search', [[['attribute_id', '=', attr_id], ['name', '=', attr_value]]])
        if not value_id:  # Si el valor no existe, se crea
            value_id = models.execute_kw(dbname, uid, pwd, 'product.attribute.value', 'create', [{'name': attr_value, 'attribute_id': attr_id}])
            print(f"[INFO] Valor de atributo creado: {attr_value} para el atributo {attr_name} con ID {value_id}")
        else:  # Si existe, se obtiene el ID del valor existente
            value_id = value_id[0]
            print(f"[INFO] Valor de atributo encontrado: {attr_value} para el atributo {attr_name} con ID {value_id}")
        value_ids.append(value_id)  # Se añade el ID del valor a la lista de IDs de valores

    return attr_id, value_ids  # Retorna el ID del atributo y los IDs de sus valores

# Leer atributos desde Excel y crearlos en Odoo
workbook_atr = openpyxl.load_workbook("atributos2.xlsx")  # Cargar archivo Excel con atributos
sheet_atr = workbook_atr.active  # Obtener la hoja activa
for row in sheet_atr.iter_rows(min_row=2):  # Empieza en la segunda fila para omitir encabezados
    attr_name, attr_value = row[0].value, row[1].value  # Leer nombre del atributo y su valor
    crear_atributo_valores(attr_name, [attr_value]) if attr_value else None  # Crear atributo si el valor no es None
workbook_atr.close()  # Cerrar archivo Excel

print('[INFO] Atributos y Valores Creados Correctamente')
#!/usr/bin/python3

from xmlrpc import client
import openpyxl

# Configuración de conexión
url = ''
dbname = ''
user = ''
pwd = ''

# Autenticación con Odoo
common = client.ServerProxy(f'{url}/xmlrpc/2/common')
uid = common.authenticate(dbname, user, pwd, {})
models = client.ServerProxy(f'{url}/xmlrpc/2/object', allow_none=True)

# Función para crear atributos y valores en Odoo
def crear_atributo_valores(attr_name, valores):
    if not attr_name or not valores:  # Si el atributo o los valores son None, se ignoran
        return None, []
    
    attr_id = models.execute_kw(dbname, uid, pwd, 'product.attribute', 'search', [[['name', '=', attr_name]]])
    if not attr_id:
        attr_id = models.execute_kw(dbname, uid, pwd, 'product.attribute', 'create', [{'name': attr_name}])
        print(f"[INFO] Atributo creado: {attr_name} con ID {attr_id}")
    else:
        attr_id = attr_id[0]
        print(f"[INFO] Atributo encontrado: {attr_name} con ID {attr_id}")

    value_ids = []
    for attr_value in valores:
        if not attr_value:
            continue  # Evitar valores vacíos
        
        value_id = models.execute_kw(dbname, uid, pwd, 'product.attribute.value', 'search', [[['attribute_id', '=', attr_id], ['name', '=', attr_value]]])
        if not value_id:
            value_id = models.execute_kw(dbname, uid, pwd, 'product.attribute.value', 'create', [{'name': attr_value, 'attribute_id': attr_id}])
        else:
            value_id = value_id[0]
        value_ids.append(value_id)

    return attr_id, value_ids

# Leer productos desde Excel y crear variantes en Odoo
workbook_prod = openpyxl.load_workbook("Productos-moda.xlsx")
sheet_prod = workbook_prod.active

productos = {}  # Diccionario para agrupar productos y sus variantes
nombre_producto_anterior = None  # Inicializar variable para evitar errores

# Recorrer todas las filas del archivo
for row in sheet_prod.iter_rows(min_row=2):
    nombre_producto = row[0].value if row[0].value else nombre_producto_anterior  # Si la celda está vacía, usar el último nombre registrado
    if not nombre_producto:
        continue  # Evita errores si la primera fila no tiene nombre
    
    codigo_barras = row[1].value
    referencia = row[2].value
    tipo_producto = row[3].value
    categoria_producto = row[4].value
    categoria_pdv = row[5].value
    precio_venta = row[6].value
    attr_name = row[8].value
    attr_value = row[9].value  # Solo un valor por fila

    if nombre_producto not in productos:
        productos[nombre_producto] = {
            "referencia": referencia,
            "tipo_producto": tipo_producto,
            "categoria_producto": categoria_producto,
            "categoria_pdv": categoria_pdv,
            "precio_venta": precio_venta,
            "atributo": attr_name if attr_name else None,  # Evita agregar atributos vacíos
            "valores": set(),  # Usamos un set para evitar duplicados
            "codigos_barras": []  # Lista de códigos de barras únicos
        }

    if attr_value:
        productos[nombre_producto]["valores"].add(attr_value)

    if codigo_barras:
        productos[nombre_producto]["codigos_barras"].append(codigo_barras)

    nombre_producto_anterior = nombre_producto  # Guardamos el último nombre leído

# Ahora procesamos todos los productos correctamente agrupados
for nombre_producto, data in productos.items():
    try:
        print(f"[INFO] Procesando producto: {nombre_producto}")
        
        # Crear el producto principal si no existe
        template_id = models.execute_kw(dbname, uid, pwd, 'product.template', 'search', [[['name', '=', nombre_producto]]])
        if not template_id:
            template_id = models.execute_kw(dbname, uid, pwd, 'product.template', 'create', [{
                'name': nombre_producto,
                'default_code': data["referencia"],
                'list_price': data["precio_venta"],
                'type': data["tipo_producto"]
            }])
            print(f"[INFO] Producto creado: {nombre_producto} (ID: {template_id})")
        else:
            template_id = template_id[0]
            print(f"[INFO] Producto encontrado: {nombre_producto} (ID: {template_id})")

        # Crear atributos y valores SOLO si existen
        if data["atributo"] and data["valores"]:
            attr_id, value_ids = crear_atributo_valores(data["atributo"], list(data["valores"]))

            # Agregar el atributo al producto
            existing_line = models.execute_kw(dbname, uid, pwd, 'product.template.attribute.line', 'search', [[['product_tmpl_id', '=', template_id], ['attribute_id', '=', attr_id]]])
            if existing_line:
                models.execute_kw(dbname, uid, pwd, 'product.template.attribute.line', 'write', [existing_line, {'value_ids': [(6, 0, value_ids)]}])
            else:
                models.execute_kw(dbname, uid, pwd, 'product.template.attribute.line', 'create', [{
                    'product_tmpl_id': template_id,
                    'attribute_id': attr_id,
                    'value_ids': [(6, 0, value_ids)]
                }])

        # Asignar los códigos de barras a las variantes
        variant_ids = models.execute_kw(dbname, uid, pwd, 'product.product', 'search', [[['product_tmpl_id', '=', template_id]]])
        for idx, variant_id in enumerate(variant_ids):
            if idx < len(data["codigos_barras"]):
                models.execute_kw(dbname, uid, pwd, 'product.product', 'write', [[variant_id], {'barcode': data["codigos_barras"][idx]}])

    except Exception as e:
        print(f"[ERROR] Error al procesar {nombre_producto}: {e}")

workbook_prod.close()
print("[INFO] Importación completada.")

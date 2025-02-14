#!/usr/bin/python3

from xmlrpc import client
import openpyxl

# Configuración de conexión
url = ""
dbname = ""
user = ""
pwd = ""

# Autenticación con Odoo
common = client.ServerProxy(f"{url}/xmlrpc/2/common")
uid = common.authenticate(dbname, user, pwd, {})
models = client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)


# Función para obtener o crear una categoría de producto
def obtener_o_crear_categoria(nombre_categoria):
    if not nombre_categoria:
        return None
    categoria_id = models.execute_kw(
        dbname,
        uid,
        pwd,
        "product.category",
        "search",
        [[["name", "=", nombre_categoria]]],
    )
    if not categoria_id:
        categoria_id = models.execute_kw(
            dbname, uid, pwd, "product.category", "create", [{"name": nombre_categoria}]
        )
        print(
            f"[INFO] Categoría de producto creada: {nombre_categoria} (ID: {categoria_id})"
        )
    else:
        categoria_id = categoria_id[0]
    return categoria_id


# Función para obtener o crear una categoría del PDV
def obtener_o_crear_categoria_pdv(nombre_categoria_pdv):
    if not nombre_categoria_pdv:
        return None
    categoria_pdv_id = models.execute_kw(
        dbname,
        uid,
        pwd,
        "pos.category",
        "search",
        [[["name", "=", nombre_categoria_pdv]]],
    )
    if not categoria_pdv_id:
        categoria_pdv_id = models.execute_kw(
            dbname, uid, pwd, "pos.category", "create", [{"name": nombre_categoria_pdv}]
        )
        print(
            f"[INFO] Categoría PDV creada: {nombre_categoria_pdv} (ID: {categoria_pdv_id})"
        )
    else:
        categoria_pdv_id = categoria_pdv_id[0]
    return categoria_pdv_id


# Función para obtener o crear una etiqueta (product.tag)
def obtener_o_crear_etiqueta(nombre_etiqueta):
    if not nombre_etiqueta:
        return None

    tag_id = models.execute_kw(
        dbname, uid, pwd, "product.tag", "search", [[["name", "=", nombre_etiqueta]]]
    )
    if not tag_id:
        tag_id = models.execute_kw(
            dbname, uid, pwd, "product.tag", "create", [{"name": nombre_etiqueta}]
        )
        print(f"[INFO] Etiqueta creada: {nombre_etiqueta} (ID: {tag_id})")
    else:
        tag_id = tag_id[0]

    return tag_id


# Función para crear atributos y valores en Odoo
def crear_atributo_valores(attr_name, valores):
    if not attr_name or not valores:
        return None, []

    attr_id = models.execute_kw(
        dbname, uid, pwd, "product.attribute", "search", [[["name", "=", attr_name]]]
    )
    if not attr_id:
        attr_id = models.execute_kw(
            dbname, uid, pwd, "product.attribute", "create", [{"name": attr_name}]
        )
        print(f"[INFO] Atributo creado: {attr_name} con ID {attr_id}")
    else:
        attr_id = attr_id[0]

    value_ids = []
    for attr_value in valores:
        if not attr_value:
            continue

        value_id = models.execute_kw(
            dbname,
            uid,
            pwd,
            "product.attribute.value",
            "search",
            [[["attribute_id", "=", attr_id], ["name", "=", attr_value]]],
        )
        if not value_id:
            value_id = models.execute_kw(
                dbname,
                uid,
                pwd,
                "product.attribute.value",
                "create",
                [{"name": attr_value, "attribute_id": attr_id}],
            )
        else:
            value_id = value_id[0]
        value_ids.append(value_id)

    return attr_id, value_ids


# Leer productos desde Excel y crear variantes en Odoo
workbook_prod = openpyxl.load_workbook("Productos.xlsx")
sheet_prod = workbook_prod.active

productos = {}
nombre_producto_anterior = None

# Recorrer todas las filas del archivo
for row in sheet_prod.iter_rows(min_row=2):
    nombre_producto = row[0].value if row[0].value else nombre_producto_anterior
    if not nombre_producto:
        continue

    codigo_barras = row[1].value
    referencia = row[2].value
    tipo_producto = row[3].value
    categoria_producto = row[4].value
    categoria_pdv = row[5].value
    precio_venta = row[6].value
    costo = row[7].value
    attr_name = row[9].value
    attr_value = row[10].value
    categoria_web = row[11].value
    etiqueta = row[12].value
    disponible_pdv = row[13].value

    if nombre_producto not in productos:
        productos[nombre_producto] = {
            "referencia": referencia,
            "tipo_producto": tipo_producto,
            "categoria_producto": categoria_producto,
            "categoria_pdv": categoria_pdv,
            "precio_venta": precio_venta,
            "costo": costo,
            "attr_name": attr_name,
            "attr_value": attr_value,
            "categoria_web": categoria_web,
            "etiqueta": etiqueta,
            "disponible_pdv": disponible_pdv,
            "atributo": attr_name if attr_name else None,
            "valores": set(),
            "codigos_barras": [],
        }

    if attr_value:
        productos[nombre_producto]["valores"].add(attr_value)

    if codigo_barras:
        productos[nombre_producto]["codigos_barras"].append(codigo_barras)

    nombre_producto_anterior = nombre_producto


# Procesar todos los productos
for nombre_producto, data in productos.items():
    try:
        print(f"[INFO] Procesando producto: {nombre_producto}")

        # Obtener o crear categorías
        categoria_id = obtener_o_crear_categoria(data["categoria_producto"])
        categoria_pdv_id = obtener_o_crear_categoria_pdv(data["categoria_pdv"])
        etiqueta_id = obtener_o_crear_etiqueta(data["etiqueta"])

        # Crear o buscar el producto
        template_id = models.execute_kw(
            dbname,
            uid,
            pwd,
            "product.template",
            "search",
            [[["name", "=", nombre_producto]]],
        )
        if not template_id:
            template_id = models.execute_kw(
                dbname,
                uid,
                pwd,
                "product.template",
                "create",
                [
                    {
                        "name": nombre_producto,
                        "default_code": data["referencia"],
                        "list_price": data["precio_venta"],
                        "standard_price": data["costo"],
                        "type": data["tipo_producto"],
                        "categ_id": categoria_id,  # Asignar categoría de producto
                        "pos_categ_id": categoria_pdv_id,  # Asignar categoría PDV
                        "product_tag_ids": data["etiqueta"],
                        "public_categ_ids": (
                            [(6, 0, [etiqueta_id])] if etiqueta_id else []
                        ),
                    }
                ],
            )
            print(f"[INFO] Producto creado: {nombre_producto} (ID: {template_id})")
        else:
            template_id = template_id[0]
        # Crear atributos y valores
        if data["atributo"] and data["valores"]:
            attr_id, value_ids = crear_atributo_valores(
                data["atributo"], list(data["valores"])
            )

            existing_line = models.execute_kw(
                dbname,
                uid,
                pwd,
                "product.template.attribute.line",
                "search",
                [
                    [
                        ["product_tmpl_id", "=", template_id],
                        ["attribute_id", "=", attr_id],
                    ]
                ],
            )
            if existing_line:
                models.execute_kw(
                    dbname,
                    uid,
                    pwd,
                    "product.template.attribute.line",
                    "write",
                    [existing_line, {"value_ids": [(6, 0, value_ids)]}],
                )
            else:
                models.execute_kw(
                    dbname,
                    uid,
                    pwd,
                    "product.template.attribute.line",
                    "create",
                    [
                        {
                            "product_tmpl_id": template_id,
                            "attribute_id": attr_id,
                            "value_ids": [(6, 0, value_ids)],
                        }
                    ],
                )

        # Asignar códigos de barras a variantes
        variant_ids = models.execute_kw(
            dbname,
            uid,
            pwd,
            "product.product",
            "search",
            [[["product_tmpl_id", "=", template_id]]],
        )
        for idx, variant_id in enumerate(variant_ids):
            if idx < len(data["codigos_barras"]):
                models.execute_kw(
                    dbname,
                    uid,
                    pwd,
                    "product.product",
                    "write",
                    [[variant_id], {"barcode": data["codigos_barras"][idx]}],
                )

    except Exception as e:
        print(f"[ERROR] Error al procesar {nombre_producto}: {e}")

workbook_prod.close()
print("[INFO] Importación completada.")

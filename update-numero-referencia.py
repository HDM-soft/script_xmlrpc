import openpyxl
import xmlrpc.client

# Configuración de conexión con Odoo
destiny_URL = "http://localhost:8069"  # Cambia por tu URL de Odoo
destiny_DB = "solaro-replica-prod_DEV"
destiny_USER = "admin"
destiny_PASSWORD = "solaro@betech1234"

# PRODUCTOS DE DESTINO
destiny_common = xmlrpc.client.ServerProxy(f"{destiny_URL}/xmlrpc/2/common")
destiny_uid = destiny_common.authenticate(
    destiny_DB, destiny_USER, destiny_PASSWORD, {}
)

destiny_models = xmlrpc.client.ServerProxy(f"{destiny_URL}/xmlrpc/2/object")

destiny_products = destiny_models.execute_kw(
    destiny_DB,
    destiny_uid,
    destiny_PASSWORD,
    "product.product",  # Still using product.product
    "search_read",  # Use search_read instead of search
    [],  # Your search domain
    {"fields": ["id", "display_name"]},  # Specify attributes to fetch
)
for product in destiny_products:
    if "] " in product["display_name"]:
        product["display_name"] = " ".join(
            product["display_name"].split("] ", 1)[1].split()
        ).strip()


# PRODUCTOS DE ORIGEN
origin_URL = "https://acerosolaro-acerosolaro-16ee-test3-17208540.dev.odoo.com"  # Cambia por tu URL de Odoo
origin_DB = "acerosolaro-acerosolaro-16ee-test3-17208540"
origin_USER = "admin"
origin_PASSWORD = "solaro@betech1234"

# Conectar con Odoo
origin_common = xmlrpc.client.ServerProxy(f"{origin_URL}/xmlrpc/2/common")
origin_uid = origin_common.authenticate(origin_DB, origin_USER, origin_PASSWORD, {})

origin_models = xmlrpc.client.ServerProxy(f"{origin_URL}/xmlrpc/2/object")

origin_products = origin_models.execute_kw(
    origin_DB,
    origin_uid,
    origin_PASSWORD,
    "product.product",  # Still using product.product
    "search_read",  # Use search_read instead of search
    [],  # Your search domain
    {"fields": ["display_name", "default_code"]},  # Specify attributes to fetch
)
for product in origin_products:
    if "] " in product["display_name"]:
        product["display_name"] = " ".join(
            product["display_name"].split("] ", 1)[1].split()
        ).strip()


missing_products = [["display_name", "default_code"]]
updated_products = 0
# Recorrer las filas del archivo Excel de origen
for product in origin_products:
    if (
        not product["display_name"] or not product["default_code"]
    ):  # variante o número de referencia saltamos la fila
        continue

    product_variant = product["display_name"]
    internal_ref = product["default_code"]  # Columna B: Referencia interna

    # Find the matching product in the second file
    matching_product = next(
        (p for p in destiny_products if p["display_name"] == product_variant), None
    )

    if matching_product:
        # print(f"✅ Found match for {product_variant}: ID = {matching_product['id']}")
        updated_products += 1
        variant_id = int(matching_product["id"])
    else:
        # print(f"❌ No match found for {product_variant}")
        missing_products.append([product_variant, internal_ref])
        print(f"❌ No se encontró la variante en Odoo: ({product_variant})")
        continue

    if variant_id:
        print(
            f"✅ Producto encontrado: ID {matching_product['id']} -> Asignando referencia {internal_ref}"
        )
        # Actualizar el default_code del producto encontrado
        destiny_models.execute_kw(
            destiny_DB,
            destiny_uid,
            destiny_PASSWORD,
            "product.product",
            "write",
            [[variant_id], {"default_code": internal_ref}],
        )
        print(
            f"✔ Producto actualizado correctamente: ({product_variant}) -> {internal_ref}"
        )


print("✅ Proceso finalizado.")
print(f"Productos actualizados: {updated_products}")
print(f"Productos no encontrados: {len(missing_products)}")
wb = openpyxl.Workbook()
ws = wb.active
for missing_product in missing_products:
    ws.append(missing_product)

wb.save("missing.xlsx")

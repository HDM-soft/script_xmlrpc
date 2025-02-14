import pandas as pd

try:
    # Cargar el archivo Excel
    df = pd.read_excel("productos2.xlsx", engine="openpyxl")

    # Eliminar espacios en los nombres de las columnas
    df.columns = df.columns.str.strip()

    # Mostrar los nombres de las columnas detectadas para depuración
    print("🔍 Columnas detectadas:", df.columns.tolist())

    # Renombrar columnas según las detectadas en la imagen
    columnas_mapeo = {
        "Nombre": "Nombre",
        "Referencia interna": "Referencia interna",
        "Atributos del producto/Nombre mostrado": "Atributos",
        "Atributos del producto/Valores/Nombre mostrado": "Valores"
    }
    df.rename(columns=columnas_mapeo, inplace=True)

    # Verificar si las columnas necesarias existen después del renombrado
    columnas_requeridas = {"Nombre", "Referencia interna", "Atributos", "Valores"}
    if not columnas_requeridas.issubset(df.columns):
        raise ValueError(f"El archivo debe contener las columnas exactas: {', '.join(columnas_requeridas)}")

    # 🚨 Depuración: Ver primeras filas antes de procesar
    print("🔹 Primeras filas antes de llenar vacíos:")
    print(df.head(10))

    # Aplicar ffill() para rellenar las filas vacías en 'Nombre', 'Referencia interna' y 'Atributos'
    df[["Nombre", "Referencia interna", "Atributos"]] = df[["Nombre", "Referencia interna", "Atributos"]].fillna(method="ffill")

    # 🚨 Depuración: Verificar si ALGODÓN C/LYCRA sigue en el DataFrame
    print("\n🔹 Registros después de ffill():")
    print(df[df["Nombre"] == "ALGODÓN C/LYCRA"])

    # Rellenar valores nulos en "Valores" con una cadena vacía
    df["Valores"] = df["Valores"].fillna("")

    # Agrupar por 'Nombre', 'Referencia interna' y 'Atributos', concatenando los valores de 'Valores' con comas
    df_modificado = df.groupby(["Nombre", "Referencia interna", "Atributos"])["Valores"]\
                      .apply(lambda x: ", ".join(x.astype(str))).reset_index()

    # 🚨 Depuración: Verificar si ALGODÓN C/LYCRA está en el archivo final
    print("\n🔹 Registros en df_modificado:")
    print(df_modificado[df_modificado["Nombre"] == "ALGODÓN C/LYCRA"])

    # Guardar el resultado en un nuevo archivo Excel
    df_modificado.to_excel("productos_modificados4.xlsx", index=False, engine="openpyxl")

    print("\n✅ Archivo 'productos_modificados4.xlsx' generado exitosamente.")

except Exception as e:
    print(f"\n❌ Ocurrió un error: {e}")

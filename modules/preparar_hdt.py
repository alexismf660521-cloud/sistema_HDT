import pandas as pd
import re
import os
import shutil

# =====================================================
# 🧹 LIMPIEZA DE ARCHIVOS TEMPORALES
# =====================================================
def limpiar_temporales():

    import shutil

    archivos = [
        "temp_datos.xlsx",
        "temp_dane.xlsx",
        "temp_preparado.xlsx",
        "plantilla.xlsx",
        "resultado.xlsx",
        "formularios.zip"
    ]

    for f in archivos:
        try:
            if os.path.exists(f):
                os.remove(f)
        except:
            pass

    # 🔴 limpiar carpeta salida completamente
    if os.path.exists("salida"):
        try:
            shutil.rmtree("salida")
        except:
            pass

    # 🔵 recrear carpeta vacía (importante)
    os.makedirs("salida", exist_ok=True)

# =====================================================
# 🧪 FUNCIÓN PRINCIPAL
# =====================================================
def preparar_hdt(archivo_datos, archivo_dane):

    # =====================================================
    # 📥 LEER ARCHIVOS
    # =====================================================
    df = pd.read_excel(archivo_datos, dtype={"Caso_Numero": str})
    municipios_df = pd.read_excel(archivo_dane)

    df.fillna("", inplace=True)

    # =====================================================
    # ✅ VALIDACIÓN
    # =====================================================
    columnas_requeridas = [
        "Caso_Numero",
        "Tipo de EMP",
        "Descripción EMP",
        "Caso_Destino"
    ]

    faltantes = [c for c in columnas_requeridas if c not in df.columns]

    if faltantes:
        raise Exception(f"❌ Faltan columnas: {faltantes}")

    # =====================================================
    # 1️⃣ MOVER Caso_Numero AL INICIO
    # =====================================================
    col = df.pop("Caso_Numero")
    df.insert(0, "Caso_Numero", col)

    # =====================================================
    # 2️⃣ LIMPIAR Tipo de EMP
    # =====================================================
    df["Tipo de EMP"] = df["Tipo de EMP"].apply(
        lambda x: re.sub(r'^.*?:\s*', '', str(x))
    )

    # =====================================================
    # 🔧 LIMPIEZA Caso_Destino
    # =====================================================
    df["Caso_Destino"] = (
        df["Caso_Destino"]
        .astype(str)
        .str.strip()
        .str.lstrip("'")
    )

    # =====================================================
    # ✅ FUNCIONES DE EXTRACCIÓN
    # =====================================================
    def extraer_cantidad(texto):
        texto = str(texto)
        match = re.search(
            r'(\d+\s*(?:ml|mL))',
            texto,
            re.IGNORECASE
        )
        return match.group(1) if match else ''

    def extraer_recipiente(texto):
        texto = str(texto)

        match = re.search(
            r'^(.*?)(?:\s+con|\s+conteniendo|\s+contenido)',
            texto,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

        return texto.strip()

    # =====================================================
    # 3️⃣ CREAR COLUMNAS DERIVADAS
    # =====================================================
    df["Cantidad"] = df["Descripción EMP"].apply(extraer_cantidad)
    df["Recipiente"] = df["Descripción EMP"].apply(extraer_recipiente)

    # Insertar en posiciones correctas
    df.insert(4, "Cantidad_tmp", df.pop("Cantidad"))
    df.insert(5, "Recipiente_tmp", df.pop("Recipiente"))

    df.rename(columns={
        "Cantidad_tmp": "Cantidad",
        "Recipiente_tmp": "Recipiente"
    }, inplace=True)

    # =====================================================
    # 4️⃣ COLUMNAS NUEVAS
    # =====================================================
    columnas_nuevas = [
        "ALCOHOLEMIA",
        "PSICOFÁRMACOS-PLAGUICIDAS",
        "CONTEXTO DEL CASO",
        "NOTAS PARA EL INFORME PERICIAL",
        "COMUNICACIONES CON EL CLIENTE"
    ]

    for i, nombre in enumerate(columnas_nuevas, start=6):
        df.insert(i, nombre, '')

    # =====================================================
    # 5️⃣ DANE
    # =====================================================
    def obtener_dane(valor):
        if pd.notna(valor) and len(str(valor)) >= 5:
            return str(valor)[:5]
        return ''

    df["DANE"] = df["Caso_Destino"].apply(obtener_dane)
    df["DANE"] = df["DANE"].astype(str).str.zfill(5)

    # =====================================================
    # ✅ PROCESAR CODIGOS DANE
    # =====================================================
    municipios_df.columns = municipios_df.columns.str.strip().str.lower()

    col_codigo = None
    col_nombre = None

    for col in municipios_df.columns:
        if "cod" in col:
            col_codigo = col
        if "nom" in col:
            col_nombre = col

    if col_codigo is None or col_nombre is None:
        raise Exception("❌ No se encontraron columnas de Código o Nombre")

    municipios_df.rename(columns={
        col_codigo: "Codigo",
        col_nombre: "Nombre"
    }, inplace=True)

    municipios_df["Codigo"] = municipios_df["Codigo"].astype(str).str.strip().str.zfill(5)

    # =====================================================
    # ✅ FORMATEAR NOMBRE DANE
    # =====================================================
    def formatear_nombre(texto):
        partes = str(texto).split()
        if len(partes) >= 2:
            return partes[0] + " - " + " ".join(partes[1:])
        return texto

    municipios_df["Lugar de Origen"] = municipios_df["Nombre"].apply(formatear_nombre)

    # =====================================================
    # 6️⃣ MERGE FINAL
    # =====================================================
    df = df.merge(
        municipios_df[["Codigo", "Lugar de Origen"]],
        left_on="DANE",
        right_on="Codigo",
        how="left"
    )

    df.drop(columns=["Codigo"], inplace=True)

    # =====================================================
    # 💾 GUARDAR
    # =====================================================
    output_file = "resultado.xlsx"
    df.to_excel(output_file, index=False)

    return output_file
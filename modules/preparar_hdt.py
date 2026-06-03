import pandas as pd
import re
import os
import shutil
import time

# =====================================================
# 🧹 LIMPIEZA
# =====================================================
def limpiar_temporales():

    archivos = [
        "temp_datos.xlsx",
        "temp_preparado.xlsx",
        "resultado.xlsx",
        "formularios.zip"
    ]

    for f in archivos:
        if os.path.exists(f):
            for _ in range(3):
                try:
                    os.remove(f)
                    break
                except:
                    time.sleep(0.5)

    if os.path.exists("salida"):
        for _ in range(3):
            try:
                shutil.rmtree("salida")
                break
            except:
                time.sleep(0.5)

    os.makedirs("salida", exist_ok=True)


# =====================================================
# 🧪 FUNCIÓN PRINCIPAL
# =====================================================
def preparar_hdt(archivo_datos, archivo_dane):

    # =============================
    # ✅ CARGA SEGURA
    # =============================
    df = pd.read_excel(archivo_datos, dtype=str, keep_default_na=False)
    municipios_df = pd.read_excel(archivo_dane, dtype=str, keep_default_na=False)

    df = df.fillna("").astype(str)
    municipios_df = municipios_df.fillna("").astype(str)

    # =============================
    # ✅ NORMALIZAR COLUMNAS
    # =============================
    df.columns = df.columns.str.strip().str.lower()
    municipios_df.columns = municipios_df.columns.str.strip().str.lower()

    # =============================
    # 🔍 DETECCIÓN AUTOMÁTICA DE COLUMNAS
    # =============================
    columnas_detectadas = {}

    for col in df.columns:

        if "caso" in col and ("numero" in col or "n°" in col):
            columnas_detectadas["caso_numero"] = col

        elif "tipo" in col and "emp" in col:
            columnas_detectadas["tipo_emp"] = col

        elif "descrip" in col and "emp" in col:
            columnas_detectadas["descripcion"] = col

        elif "destino" in col:
            columnas_detectadas["destino"] = col

    requeridas = ["caso_numero", "tipo_emp", "descripcion", "destino"]
    faltantes = [c for c in requeridas if c not in columnas_detectadas]

    if faltantes:
        raise Exception(f"❌ No se pudieron detectar columnas: {faltantes}")

    # =============================
    # ✅ RENOMBRAR
    # =============================
    df.rename(columns={
        columnas_detectadas["caso_numero"]: "Caso_Numero",
        columnas_detectadas["tipo_emp"]: "Tipo de EMP",
        columnas_detectadas["descripcion"]: "Descripción EMP",
        columnas_detectadas["destino"]: "Caso_Destino"
    }, inplace=True)

    # =============================
    # ✅ ORDENAR
    # =============================
    col = df.pop("Caso_Numero")
    df.insert(0, "Caso_Numero", col)

    # =============================
    # ✅ LIMPIAR TIPO EMP
    # =============================
    df["Tipo de EMP"] = df["Tipo de EMP"].apply(
        lambda x: re.sub(r'^.*?:\s*', '', str(x))
    )

    # =============================
    # ✅ FUNCIONES
    # =============================
    def extraer_cantidad(texto):
        match = re.search(r'(\d+\s*ml)', str(texto), re.IGNORECASE)
        return match.group(1) if match else ''

    def extraer_recipiente(texto):
        texto = str(texto)

        match = re.search(
            r'^(.*?)(?:\s+con|\s+conteniendo|\s+contenido)',
            texto,
            re.IGNORECASE
        )

        return match.group(1).strip() if match else texto.strip()

    def extraer_dane(valor):
        valor = str(valor).strip()
        match = re.search(r'\d{5}', valor)
        return match.group(0) if match else ""

    # =============================
    # ✅ COLUMNAS DERIVADAS
    # =============================
    df["Cantidad"] = df["Descripción EMP"].apply(extraer_cantidad)
    df["Recipiente"] = df["Descripción EMP"].apply(extraer_recipiente)

    df.insert(4, "Cantidad_tmp", df.pop("Cantidad"))
    df.insert(5, "Recipiente_tmp", df.pop("Recipiente"))

    df.rename(columns={
        "Cantidad_tmp": "Cantidad",
        "Recipiente_tmp": "Recipiente"
    }, inplace=True)

    # =============================
    # ✅ COLUMNAS NUEVAS
    # =============================
    nuevas = [
        "ALCOHOLEMIA",
        "PSICOFÁRMACOS-PLAGUICIDAS",
        "CONTEXTO DEL CASO",
        "NOTAS PARA EL INFORME PERICIAL",
        "COMUNICACIONES CON EL CLIENTE"
    ]

    for i, col in enumerate(nuevas, start=6):
        df.insert(i, col, "")

    # =============================
    # ✅ DANE (CORREGIDO)
    # =============================
    df["DANE"] = df["Caso_Destino"].apply(extraer_dane)

    # =============================
    # ✅ PROCESAR ARCHIVO DANE
    # =============================
    col_codigo = None
    col_nombre = None

    for col in municipios_df.columns:
        if "cod" in col:
            col_codigo = col
        elif "nom" in col:
            col_nombre = col

    if not col_codigo or not col_nombre:
        raise Exception("❌ No se encuentran columnas en archivo DANE")

    municipios_df.rename(columns={
        col_codigo: "Codigo",
        col_nombre: "Nombre"
    }, inplace=True)

    municipios_df["Codigo"] = (
        municipios_df["Codigo"]
        .astype(str)
        .str.strip()
        .str.zfill(5)
    )

    municipios_df["Lugar de Origen"] = municipios_df["Nombre"]

    # =============================
    # ✅ MERGE FINAL
    # =============================
    df = df.merge(
        municipios_df[["Codigo", "Lugar de Origen"]],
        left_on="DANE",
        right_on="Codigo",
        how="left"
    )

    df.drop(columns=["Codigo"], inplace=True)

    # =============================
    # 💾 GUARDAR
    # =============================
    salida = "resultado.xlsx"
    df.to_excel(salida, index=False)

    return salida
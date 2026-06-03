import streamlit as st
import os

# Importar funciones
from modules.preparar_hdt import preparar_hdt, limpiar_temporales
from modules.generar_hdt import generar_formularios

# =====================================================
# ⚙️ CONFIGURACIÓN
# =====================================================
st.set_page_config(
    page_title="Sistema HdT - Toxicología",
    layout="wide",
    page_icon="🧪"
)

st.markdown("""
# 🧪 Sistema Institucional de Hojas de Trabajo  
### Laboratorio de Toxicología Forense
---
""")

# =====================================================
# 🧪 🔍 DEBUG (TEMPORAL)
# =====================================================
st.markdown("### 🔍 Debug de rutas (temporal)")

st.write("📍 Ruta actual:", os.getcwd())
st.write("📂 Archivos en carpeta raíz:", os.listdir())

if os.path.exists("assets"):
    st.write("📂 Contenido carpeta assets:", os.listdir("assets"))
else:
    st.error("❌ No existe la carpeta 'assets'")

# =====================================================
# 📂 RUTAS INTERNAS
# =====================================================
ruta_dane = "assets/codigos_dane.xlsx"
ruta_plantilla = "assets/plantilla_formulario.xlsx"

# =====================================================
# 🧾 TABS
# =====================================================
tab1, tab2 = st.tabs(["📄 Preparar HdT", "🧪 Generar Hojas"])

# =====================================================
# 📄 TAB 1
# =====================================================
with tab1:

    st.subheader("Preparación de datos SAILFO")

    archivo_datos = st.file_uploader("📂 Archivo SAILFO", type=["xlsx"])

    if st.button("🚀 Procesar HdT"):

        limpiar_temporales()

        if archivo_datos:

            with open("temp_datos.xlsx", "wb") as f:
                f.write(archivo_datos.getbuffer())

            try:
                salida = preparar_hdt("temp_datos.xlsx", ruta_dane)

                st.success("✅ Archivo procesado correctamente")

                with open(salida, "rb") as f:
                    st.download_button(
                        label="⬇ Descargar archivo preparado",
                        data=f,
                        file_name="Datos_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ Error en procesamiento: {str(e)}")

        else:
            st.error("⚠️ Debe cargar el archivo SAILFO")

# =====================================================
# 🧪 TAB 2
# =====================================================
with tab2:

    st.subheader("Generación de hojas de trabajo")

    archivo_hdt = st.file_uploader("📂 Archivo preparado", type=["xlsx"])

    if st.button("📑 Generar formularios"):

        limpiar_temporales()

        if archivo_hdt:

            with open("temp_preparado.xlsx", "wb") as f:
                f.write(archivo_hdt.getbuffer())

            try:
                zip_file = generar_formularios("temp_preparado.xlsx", ruta_plantilla)

                st.success("✅ Formularios generados correctamente")

                with open(zip_file, "rb") as f:
                    st.download_button(
                        label="⬇ Descargar ZIP",
                        data=f,
                        file_name="formularios.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"❌ Error generando formularios: {str(e)}")

        else:
            st.error("⚠️ Debe cargar el archivo preparado")
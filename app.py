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
# 🧾 TABS
# =====================================================
tab1, tab2 = st.tabs(["📄 Preparar HdT", "🧪 Generar Hojas"])

# =====================================================
# 📄 TAB 1: PREPARAR HDT
# =====================================================
with tab1:

    st.subheader("Preparación de datos SAILFO")

    archivo_datos = st.file_uploader("📂 Archivo SAILFO", type=["xlsx"])
    archivo_dane = st.file_uploader("📂 Codigos DANE", type=["xlsx"])

    if st.button("🚀 Procesar HdT"):

        limpiar_temporales()

        if archivo_datos and archivo_dane:

            # Guardar archivos temporales
            with open("temp_datos.xlsx", "wb") as f:
                f.write(archivo_datos.getbuffer())

            with open("temp_dane.xlsx", "wb") as f:
                f.write(archivo_dane.getbuffer())

            try:
                salida = preparar_hdt("temp_datos.xlsx", "temp_dane.xlsx")

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
            st.error("⚠️ Debe cargar ambos archivos")

# =====================================================
# 🧪 TAB 2: GENERAR HOJAS DE TRABAJO
# =====================================================
with tab2:

    st.subheader("Generación de hojas de trabajo")

    archivo_hdt = st.file_uploader("📂 Archivo preparado", type=["xlsx"])
    plantilla = st.file_uploader("📂 Plantilla formulario", type=["xlsx"])

    if st.button("📑 Generar formularios"):

        limpiar_temporales()

        if archivo_hdt and plantilla:

            # Guardar archivos temporales
            with open("temp_preparado.xlsx", "wb") as f:
                f.write(archivo_hdt.getbuffer())

            with open("plantilla.xlsx", "wb") as f:
                f.write(plantilla.getbuffer())

            try:
                zip_file = generar_formularios("temp_preparado.xlsx", "plantilla.xlsx")

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
            st.error("⚠️ Debe cargar archivo preparado y plantilla")
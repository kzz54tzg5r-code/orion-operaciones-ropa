import streamlit as st
from modules.ui import setup_page, render_header
from modules.data_loader import load_excel_file

setup_page()
render_header()

st.sidebar.header("📂 Fuente de Datos")

uploaded_file = st.sidebar.file_uploader(
    "Cargar archivo Excel",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("Por favor cargue un archivo Excel para iniciar ORION.")
    st.stop()

data = load_excel_file(uploaded_file)

st.success("Archivo cargado correctamente.")

st.subheader("📋 Diagnóstico inicial")
st.write("Hojas detectadas:")
st.write(data["sheet_names"])

st.write("Hoja operativa detectada:")
st.write(data["operativa_name"])

st.write("Hojas mensuales detectadas:")
st.write(data["mensuales_names"])

st.write("Registros operativos:")
st.write(len(data["df_operativa"]))

st.write("Registros recuperación:")
st.write(len(data["df_recuperacion"]))

tab1, tab2, tab3 = st.tabs([
    "Panel Ejecutivo",
    "Diagnóstico de Datos",
    "Compartir ORION"
])

with tab1:
    st.title("Panel Ejecutivo")
    st.info("Bloque 1 listo. En el Bloque 2 agregaremos KPIs, rankings y filtros.")

with tab2:
    st.subheader("Columnas Operativas")
    st.write(data["df_operativa"].columns.tolist())

    st.subheader("Columnas Recuperación")
    st.write(data["df_recuperacion"].columns.tolist())

with tab3:
    st.subheader("🔗 Compartir ORION")
    st.write("Cuando publiques esta app en Streamlit Cloud, aquí se visualizará la URL compartible.")

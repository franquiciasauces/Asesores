import streamlit as st

st.set_page_config(
    page_title="Aplicativo Asesor",
    layout="wide"
)

st.title("Aplicativo Asesor")
st.write("Bloque 1.1 funcionando correctamente.")
import streamlit as st

st.set_page_config(
    page_title="Aplicativo Asesor",
    layout="wide"
)

st.title("Aplicativo Asesor")

st.sidebar.header("Menú principal")

modulo = st.sidebar.radio(
    "Seleccione una opción:",
    [
        "Aprendizaje",
        "Asesoría",
        "Evaluación"
    ]
)

st.write(f"### {modulo}")

if modulo == "Aprendizaje":
    st.write("Módulo de Aprendizaje")

elif modulo == "Asesoría":
    st.write("Módulo de Asesoría")

elif modulo == "Evaluación":
    st.write("Módulo de Evaluación")
    import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="Aplicativo Asesor",
    layout="wide"
)

st.title("Aplicativo Asesor")

st.sidebar.header("Menú principal")

modulo = st.sidebar.radio(
    "Seleccione una opción:",
    [
        "Aprendizaje",
        "Asesoría",
        "Evaluación"
    ]
)

if modulo == "Aprendizaje":

    st.header("Aprendizaje")

    archivo = Path(__file__).parent / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"

    if not archivo.exists():
        st.error("No se encontró la matriz de productos.")
        st.write("Archivo buscado:")
        st.code(str(archivo))
        st.stop()

    try:
        productos = pd.read_excel(
            archivo,
            sheet_name="Base_Productos"
        )

        st.success("Matriz de productos cargada correctamente.")

        st.write("Cantidad de registros:", len(productos))

    except Exception as e:
        st.error("No fue posible leer la matriz de productos.")
        st.exception(e)

else:
    st.write(f"### {modulo}")
    archivo = Path(__file__).parent / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"

if not archivo.exists():
    st.error("No se encontró la matriz de productos.")
    st.write("Archivo buscado:")
    st.code(str(archivo))
    st.stop()

try:
    productos = pd.read_excel(
        archivo,
        sheet_name="Base_Productos"
    )

    st.success("Matriz de productos cargada correctamente.")

    st.write("Cantidad de registros:", len(productos))

except Exception as e:
    st.error("No fue posible leer la matriz de productos.")
    st.exception(e)

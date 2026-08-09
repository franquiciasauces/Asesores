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

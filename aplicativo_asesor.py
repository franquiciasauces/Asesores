import streamlit as st
import pandas as pd
from pathlib import Path

# ============================================================
# 1.1 - CONFIGURACIÓN INICIAL
# ============================================================

st.set_page_config(
    page_title="Aplicativo Asesor",
    layout="wide"
)

st.title("Aplicativo Asesor")


# ============================================================
# 1.2 - MENÚ PRINCIPAL
# ============================================================

st.sidebar.header("Menú principal")

modulo = st.sidebar.radio(
    "Seleccione una opción:",
    [
        "Aprendizaje",
        "Asesoría",
        "Evaluación"
    ]
)


# ============================================================
# 1.3 - CARGA DE LA MATRIZ DE PRODUCTOS
# ============================================================

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

        st.write(
            "Cantidad de registros:",
            len(productos)
        )


        # ====================================================
        # 1.4 - LISTADO COMPLETO DE PRODUCTOS
        # ====================================================

        st.subheader("Listado completo de productos")

        st.dataframe(
            productos,
            use_container_width=True,
            hide_index=True
        )


        # ====================================================
        # 1.5 - REVISIÓN DE LA ESTRUCTURA DE PRODUCTOS
        # ====================================================

        st.subheader("Estructura de la matriz")

        st.write(
            "Columnas disponibles en Base_Productos:"
        )

        for columna in productos.columns:
            st.write(f"• {columna}")

    except Exception as e:

        st.error(
            "No fue posible leer la matriz de productos."
        )

        st.exception(e)


# ============================================================
# OTROS MÓDULOS — TODAVÍA NO SE DESARROLLAN
# ============================================================

elif modulo == "Asesoría":

    st.header("Asesoría")

    st.info(
        "El módulo de Asesoría se desarrollará posteriormente."
    )


elif modulo == "Evaluación":

    st.header("Evaluación")

    st.info(
        "El módulo de Evaluación se desarrollará posteriormente."
    )

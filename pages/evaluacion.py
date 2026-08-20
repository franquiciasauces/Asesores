# ============================================================
# FITOASISTE
# APLICATIVO DE EVALUACIÓN
# NORMALIZACIÓN DE PRODUCTO - COMPONENTE - ACCIÓN GENERAL
# ============================================================

from pathlib import Path
import pandas as pd
import streamlit as st


# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="FITOASISTE - Evaluación",
    page_icon="📝",
    layout="wide"
)


# ============================================================
# 2. RUTAS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

ARCHIVO_MATRIZ = (
    BASE_DIR /
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)

ARCHIVO_NORMALIZADO = (
    BASE_DIR /
    "DATAFRAME_PRODUCTO_COMPONENTE_ACCION.csv"
)


# ============================================================
# 3. SESIÓN
# ============================================================

USUARIO = st.session_state.get(
    "usuario_actual",
    ""
)

ROL = st.session_state.get(
    "rol_usuario",
    ""
)

if not USUARIO:
    st.warning(
        "Debe ingresar primero al Aplicativo Asesor."
    )
    st.stop()

if ROL.upper() != "ADMINISTRADOR":
    st.error(
        "Este módulo está disponible únicamente "
        "para el administrador."
    )
    st.stop()


# ============================================================
# 4. ENCABEZADO
# ============================================================

st.title("📝 FITOASISTE — EVALUACIÓN")

st.write(
    f"Administrador: **{USUARIO}**"
)

# ============================================================
# 5. VALIDAR MATRIZ
# ============================================================

if not ARCHIVO_MATRIZ.exists():

    st.error(
        "No se encontró el archivo "
        "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
    )

    st.write(
        f"Ruta buscada: {ARCHIVO_MATRIZ}"
    )

    st.stop()


st.success(
    "✓ MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx encontrada."
)

# ============================================================
# 5.1 — CARGAR BASE_PRODUCTOS
# ============================================================

@st.cache_data
def cargar_base_productos(ruta_matriz):

    return pd.read_excel(
        ruta_matriz,
        sheet_name="Base_Productos"
    )


try:

    df_base_productos = cargar_base_productos(
        ARCHIVO_MATRIZ
    )

except Exception as error:

    st.error(
        "No fue posible cargar la hoja "
        "Base_Productos."
    )

    st.exception(error)

    st.stop()


st.success(
    "✓ Base_Productos cargada correctamente."
)

# ============================================================
# 5.2 — COLUMNAS DE BASE_PRODUCTOS
# ============================================================

COL_PRODUCTO = "Producto"

COL_CATEGORIA_PRINCIPAL = (
    "Categoría principal"
)

COL_CATEGORIAS_COMPLEMENTARIAS = (
    "Categorías complementarias"
)

COL_COMPONENTES = "componentes"

COL_ACCIONES_GENERALES = (
    "Acciones generales"
)

COL_PRECIO = "Precio público"

COL_FOTO = "Foto"

# ============================================================
# 5.3 — BASE DE TRABAJO
# ============================================================

df_trabajo = df_base_productos[
    [
        COL_PRODUCTO,
        COL_COMPONENTES,
        COL_ACCIONES_GENERALES
    ]
].copy()


df_trabajo["Producto"] = (
    df_trabajo[COL_PRODUCTO]
    .fillna("")
    .astype(str)
    .str.strip()
)


df_trabajo["Componentes"] = (
    df_trabajo[COL_COMPONENTES]
    .fillna("")
    .astype(str)
    .str.strip()
)


df_trabajo["Acciones_generales"] = (
    df_trabajo[COL_ACCIONES_GENERALES]
    .fillna("")
    .astype(str)
    .str.strip()
)


df_trabajo = df_trabajo[
    df_trabajo["Producto"].ne("")
].copy()


# ============================================================
# 5.4 — COMPROBACIÓN DE BASE DE TRABAJO
# ============================================================

st.subheader(
    "Base de trabajo"
)

st.write(
    f"Registros: **{len(df_trabajo)}**"
)

st.dataframe(
    df_trabajo[
        [
            "Producto",
            "Componentes",
            "Acciones_generales"
        ]
    ],
    use_container_width=True
)


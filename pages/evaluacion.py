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
        "No se encontró la matriz."
    )

    st.stop()

st.success(
    "✓ Matriz encontrada."
)
# ============================================================
# 5.1 — COLUMNAS BASE_PRODUCTOS
# ============================================================

COL_PRODUCTO = "Producto"
COL_CATEGORIA_PRINCIPAL = "Categoría principal"
COL_CATEGORIAS_COMPLEMENTARIAS = "Categorías complementarias"
COL_COMPONENTES = "componentes"
COL_ACCIONES_GENERALES = "Acciones generales"
COL_PRECIO = "Precio público"
COL_FOTO = "Foto"

# ============================================================
# 6. CARGAR BASE_PRODUCTOS
# ============================================================

@st.cache_data
def cargar_base_productos(ruta):

    return pd.read_excel(
        ruta,
        sheet_name="Base_Productos"
    )


try:

    df_base = cargar_base_productos(
        ARCHIVO_MATRIZ
    )

except Exception as error:

    st.error(
        f"No fue posible leer Base_Productos: {error}"
    )

    st.stop()


# ============================================================
# 7. INFORMACIÓN DE LA MATRIZ
# ============================================================

st.subheader(
    "Base_Productos"
)

st.write(
    f"Registros: **{len(df_base)}**"
)

st.write(
    "Columnas disponibles:"
)

st.write(
    list(df_base.columns)
)


# ============================================================
# 8. SELECCIÓN DE COLUMNAS
# ============================================================

st.divider()

st.subheader(
    "Configuración de normalización"
)

st.info(
    "Seleccione las columnas reales de la matriz "
    "que corresponden a Producto, Componentes y "
    "Acción General. El sistema no modifica el "
    "contenido original."
)


columnas = list(df_base.columns)

col_producto = st.selectbox(
    "Columna PRODUCTO",
    columnas,
    key="normalizacion_producto"
)

col_componentes = st.selectbox(
    "Columna Componentes",
    ["(No existe)"] + columnas,
    key="normalizacion_componentes"
)

col_accion = st.selectbox(
    "Columna Acciones generales",
    ["(No existe)"] + columnas,
    key="normalizacion_accion"
)


# ============================================================
# 9. PREPARAR DATAFRAME
# ============================================================

if st.button(
    "Preparar normalización",
    type="primary"
):

    datos = pd.DataFrame()

    datos["Producto"] = (
        df_base[col_producto]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    if col_componentes == "(No existe)":

        datos["Componentes"] = ""

    else:

        datos["Componentes"] = (
            df_base[col_componentes]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    if col_accion == "(No existe)":

        datos["Accion_generales"] = ""

    else:

        datos["Acciones_generales"] = (
            df_base[col_accion]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # Eliminar filas completamente vacías
    datos = datos[
        datos["Producto"].ne("")
    ].copy()

    # Eliminar duplicados exactos
    datos = datos.drop_duplicates(
        subset=[
            "Producto",
            "Componentes",
            "Acciones_generales"
        ]
    )

    # Identificador estable del registro
    datos["ID_Normalizado"] = (
        datos[
            [
                "Producto",
                "Componentes",
                "Acciones_generales"
            ]
        ]
        .astype(str)
        .agg(" | ".join, axis=1)
        .str.strip()
    )

    # Orden
    datos = datos[
        [
            "ID_Normalizado",
            "Producto",
            "Componentes",
            "Acciones_generales"
        ]
    ]

    # Guardar
    datos.to_csv(
        ARCHIVO_NORMALIZADO,
        index=False,
        encoding="utf-8-sig"
    )

    st.success(
        "✓ DataFrame normalizado creado/actualizado."
    )

    st.write(
        f"Registros normalizados: **{len(datos)}**"
    )

    st.dataframe(
        datos,
        use_container_width=True
    )

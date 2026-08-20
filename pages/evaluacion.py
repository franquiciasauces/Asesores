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
        "No se encontró "
        "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
    )

    st.stop()

st.success(
    "✓ MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx encontrada."
)


# ============================================================
# 5.1 LEER HOJA BASE_PRODUCTOS DE LA MATRIZ
# ============================================================

df_base_productos = pd.read_excel(
    ARCHIVO_MATRIZ,
    sheet_name="Base_Productos"
)

st.success(
    "✓ Hoja Base_Productos cargada desde la matriz."
)


# ============================================================
# 5.2 COLUMNAS REALES DE BASE_PRODUCTOS
# ============================================================

COL_PRODUCTO = "Producto"
COL_CATEGORIA_PRINCIPAL = "Categoría principal"
COL_CATEGORIAS_COMPLEMENTARIAS = "Categorías complementarias"
COL_COMPONENTES = "Componentes"
COL_ACCIONES_GENERALES = "Acciones generales"
COL_PRECIO = "Precio público"
COL_FOTO = "Foto"


# ============================================================
# 5.3 VALIDAR LAS COLUMNAS DE LA HOJA
# ============================================================

COLUMNAS_REQUERIDAS = [
    COL_PRODUCTO,
    COL_CATEGORIA_PRINCIPAL,
    COL_CATEGORIAS_COMPLEMENTARIAS,
    COL_COMPONENTES,
    COL_ACCIONES_GENERALES,
    COL_PRECIO,
    COL_FOTO
]

faltantes = [
    columna
    for columna in COLUMNAS_REQUERIDAS
    if columna not in df_base_productos.columns
]

if faltantes:

    st.error(
        "Faltan columnas en la hoja Base_Productos:"
    )

    for columna in faltantes:
        st.write(f"- {columna}")

    st.stop()


st.success(
    "✓ Estructura de Base_Productos validada."
)


# ============================================================
# 5.4 CREAR BASE DE TRABAJO
# ============================================================

df_trabajo = df_base_productos[
    [
        "Producto",
        "Componentes",
        "Acciones generales"
    ]
].copy()


# ============================================================
# 5.5 LIMPIEZA TÉCNICA, SIN CAMBIAR EL CONTENIDO
# ============================================================

df_trabajo["Producto"] = (
    df_trabajo["Producto"]
    .fillna("")
    .astype(str)
    .str.strip()
)

df_trabajo["Componentes"] = (
    df_trabajo["Componentes"]
    .fillna("")
    .astype(str)
    .str.strip()
)

df_trabajo["Acciones generales"] = (
    df_trabajo["Acciones generales"]
    .fillna("")
    .astype(str)
    .str.strip()
)


# ============================================================
# 5.6 ELIMINAR FILAS SIN PRODUCTO
# ============================================================

df_trabajo = df_trabajo[
    df_trabajo["Producto"] != ""
].copy()


# ============================================================
# 5.7 RESULTADO DE LA BASE DE TRABAJO
# ============================================================

st.subheader(
    "Base de trabajo para normalización"
)

st.write(
    f"Registros encontrados: **{len(df_trabajo)}**"
)

st.dataframe(
    df_trabajo[
        [
            "Producto",
            "Componentes",
            "Acciones generales"
        ]
    ],
    use_container_width=True
)
# ============================================================
# 5.8 — PREPARAR ACCIONES GENERALES
# ============================================================

import re


def separar_acciones(valor):

    if pd.isna(valor):
        return []

    texto = str(valor).strip()

    if not texto:
        return []

    # Unificar saltos de línea para poder analizarlos
    texto = texto.replace("\r\n", "\n")
    texto = texto.replace("\r", "\n")

    # Separar únicamente cuando exista una línea independiente.
    partes = re.split(
        r"\n+",
        texto
    )

    acciones = []

    for parte in partes:

        parte = parte.strip()

        if not parte:
            continue

        acciones.append(
            parte
        )

    return acciones


df_trabajo["Lista_acciones"] = (
    df_trabajo["Acciones generales"]
    .apply(separar_acciones)
)


# ============================================================
# 5.9 — VISTA PREVIA
# ============================================================

st.subheader(
    "Acciones generales identificadas"
)

df_vista_acciones = df_trabajo[
    [
        "Producto",
        "Componentes",
        "Acciones generales",
        "Lista_acciones"
    ]
].copy()

st.dataframe(
    df_vista_acciones,
    use_container_width=True
)

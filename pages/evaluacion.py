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
# 5.8 — LIMPIEZA DE ACCIONES GENERALES
# ============================================================

import re


def limpiar_texto(valor):

    if pd.isna(valor):
        return ""

    return str(valor).strip()


def limpiar_acciones_generales(valor):

    texto = limpiar_texto(valor)

    if not texto:
        return []

    # Normalizar saltos de línea
    texto = texto.replace("\r\n", "\n")
    texto = texto.replace("\r", "\n")

    # --------------------------------------------------------
    # ELIMINAR SECCIONES QUE NO SON ACCIONES
    # --------------------------------------------------------

    # MODO DE ACCIÓN no es una acción
    texto = re.sub(
        r"(?i)\bMODO\s+DE\s+ACCI[ÓO]N\s*:?",
        "",
        texto
    )

    # COMBINAR CON y todo lo que sigue no pertenece
    # al DataFrame de acciones generales
    texto = re.split(
        r"(?i)\bCOMBINAR\s+CON\s*:?",
        texto,
        maxsplit=1
    )[0]

    # FRASE DE VENTA y todo lo que sigue no pertenece
    # al DataFrame de acciones generales
    texto = re.split(
        r"(?i)\bFRASE\s+DE\s+VENTA\s*:?",
        texto,
        maxsplit=1
    )[0]

    # --------------------------------------------------------
    # SEPARAR REGISTROS
    # --------------------------------------------------------

    partes = re.split(
        r"\n+",
        texto
    )

    acciones = []

    for parte in partes:

        parte = parte.strip()

        if not parte:
            continue

        # ----------------------------------------------------
        # DESCARTAR ENCABEZADOS AISLADOS
        # ----------------------------------------------------

        if re.fullmatch(
            r"(?i)MODO\s+DE\s+ACCI[ÓO]N\s*:?",
            parte
        ):
            continue

        if re.fullmatch(
            r"(?i)COMBINAR\s+CON\s*:?",
            parte
        ):
            continue

        if re.fullmatch(
            r"(?i)FRASE\s+DE\s+VENTA\s*:?",
            parte
        ):
            continue

        acciones.append(parte)

    return acciones


# ============================================================
# 5.9 — CONSTRUIR DATAFRAME DE ACCIONES GENERALES
# ============================================================

registros_acciones_generales = []


for _, fila in df_trabajo.iterrows():

    producto = limpiar_texto(
        fila["Producto"]
    )

    if not producto:
        continue

    acciones = limpiar_acciones_generales(
        fila["Acciones generales"]
    )

    for accion in acciones:

        registros_acciones_generales.append(
            {
                "Producto": producto,
                "Acción general": accion
            }
        )


df_acciones_generales = pd.DataFrame(
    registros_acciones_generales,
    columns=[
        "Producto",
        "Acción general"
    ]
)


# ============================================================
# 5.10 — ELIMINAR DUPLICADOS EXACTOS
# ============================================================

df_acciones_generales = (
    df_acciones_generales
    .drop_duplicates(
        subset=[
            "Producto",
            "Acción general"
        ]
    )
    .reset_index(drop=True)
)


# ============================================================
# 5.11 — MOSTRAR DATAFRAME
# ============================================================

st.subheader(
    "DataFrame — Acciones Generales"
)

st.write(
    f"Registros: **{len(df_acciones_generales)}**"
)

st.dataframe(
    df_acciones_generales,
    use_container_width=True
)


# ============================================================
# 5.12 — DESCARGAR
# ============================================================

csv_acciones_generales = (
    df_acciones_generales
    .to_csv(
        index=False,
        encoding="utf-8-sig"
    )
)

st.download_button(
    label="⬇️ Descargar Acciones Generales",
    data=csv_acciones_generales,
    file_name="ACCIONES_GENERALES.csv",
    mime="text/csv",
    key="descargar_acciones_generales"
)

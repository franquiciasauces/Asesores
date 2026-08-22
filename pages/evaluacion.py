# ============================================================
# FITOASISTE
# APLICATIVO DE EVALUACIÓN
# NORMALIZACIÓN DE PRODUCTO - COMPONENTE - ACCIÓN GENERAL
# ============================================================

from pathlib import Path
from unidecode import unidecode
from rapidfuzz import fuzz
import streamlit as st
import pandas as pd
import numpy as np

import base64
import urllib.request
import urllib.error
import json

GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]


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


# ============================================================
# 5.1 DIAGNÓSTICO Y LECTURA DE LA MATRIZ FUENTE
# ============================================================

st.markdown("### 5.1 Diagnóstico de la matriz fuente")

try:
    if not ARCHIVO_MATRIZ.exists():
        st.error(
            "❌ 5.1 ERROR: No se encontró el archivo de la matriz."
        )
    else:
        libro = pd.ExcelFile(ARCHIVO_MATRIZ)

        st.success(
            f"✅ 5.1 OK: Archivo de matriz encontrado. "
            f"Hojas disponibles: {len(libro.sheet_names)}"
        )

        st.write("**Hojas encontradas en la matriz:**")
        st.write(libro.sheet_names)

        hoja_fuente = st.selectbox(
            "Seleccione la hoja de la matriz que contiene la información:",
            libro.sheet_names,
            key="hoja_matriz_normalizacion"
        )

        df_fuente = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name=hoja_fuente
        )

        df_fuente = df_fuente.dropna(
            axis=1,
            how="all"
        )

        st.success(
            f"✅ Hoja cargada correctamente: **{hoja_fuente}**"
        )

        st.info(
            f"Registros encontrados: **{len(df_fuente)}** | "
            f"Columnas encontradas: **{len(df_fuente.columns)}**"
        )

        st.write("### Columnas REALES encontradas en la hoja")

        columnas_reales = pd.DataFrame({
            "N.º": range(1, len(df_fuente.columns) + 1),
            "Nombre real de la columna": [
                str(col)
                for col in df_fuente.columns
            ]
        })

        st.dataframe(
            columnas_reales,
            use_container_width=True,
            hide_index=True
        )

        st.write("### Primeros registros de la matriz original")

        st.dataframe(
            df_fuente.head(10),
            use_container_width=True,
            hide_index=True
        )

        st.success(
            "🟢 5.1 TERMINADO: La matriz fue leída correctamente. "
            "Todavía NO se ha realizado ninguna normalización."
        )

except Exception as e:
    st.error(
        f"🔴 5.1 ERROR al leer la matriz: "
        f"{type(e).__name__}: {e}"
    )

# ============================================================
# 5.2 SEPARAR ACCIONES GENERALES
# ============================================================

st.markdown("### 5.2 Separación de acciones generales")

try:
    import re

    requeridas_52 = [
        "Producto",
        "Acciones generales"
    ]

    faltantes_52 = [
        columna
        for columna in requeridas_52
        if columna not in df_fuente.columns
    ]

    if faltantes_52:
        st.error(
            "❌ 5.2 ERROR: Faltan columnas: "
            + ", ".join(faltantes_52)
        )
        st.stop()

    df_acciones_52 = df_fuente[
        requeridas_52
    ].copy()

    df_acciones_52 = df_acciones_52.fillna("")

    def separar_acciones_52(texto):
        texto = str(texto).strip()

        if not texto:
            return []

        partes = re.split(
            r"\s*;\s*|\s*,\s*|(?<=\.)\s+(?=[A-ZÁÉÍÓÚÑ])",
            texto
        )

        return [
            parte.strip(" ;,.")
            for parte in partes
            if parte.strip(" ;,.")
        ]

    df_acciones_52[
        "Acción general"
    ] = df_acciones_52[
        "Acciones generales"
    ].apply(separar_acciones_52)

    df_acciones_52 = (
        df_acciones_52
        .explode("Acción general")
        .reset_index(drop=True)
    )

    df_acciones_52["Producto"] = (
        df_acciones_52["Producto"]
        .astype(str)
        .str.strip()
    )

    df_acciones_52["Acción general"] = (
        df_acciones_52["Acción general"]
        .astype(str)
        .str.strip()
    )

    df_acciones_52 = df_acciones_52[
        (df_acciones_52["Producto"] != "")
        &
        (df_acciones_52["Acción general"] != "")
    ].copy()

    df_acciones_52 = df_acciones_52[
        [
            "Producto",
            "Acción general"
        ]
    ]

    st.session_state[
        "df_acciones_52"
    ] = df_acciones_52.copy()

    st.success(
        f"🟢 5.2 TERMINADO: "
        f"{len(df_fuente)} registros originales → "
        f"{len(df_acciones_52)} relaciones Producto–Acción."
    )

    st.dataframe(
        df_acciones_52,
        use_container_width=True,
        hide_index=True
    )

except Exception as e:
    st.error(
        f"🔴 5.2 ERROR: "
        f"{type(e).__name__}: {e}"
    )
# ============================================================
# 5.3 CLASIFICAR ACCIÓN SEGÚN COMPONENTES DEL PRODUCTO
# ============================================================

st.markdown("### 5.3 Clasificación de acciones")

try:
    requeridas_53 = [
        "Producto",
        "Acción general"
    ]

    faltantes_53 = [
        columna
        for columna in requeridas_53
        if columna not in df_acciones_52.columns
    ]

    if faltantes_53:
        st.error(
            "❌ 5.3 ERROR: Faltan columnas en 5.2: "
            + ", ".join(faltantes_53)
        )
        st.stop()

    if "Componentes" not in df_fuente.columns:
        st.error(
            "❌ 5.3 ERROR: La matriz no contiene "
            "la columna real 'Componentes'."
        )
        st.stop()

    df_componentes_53 = df_fuente[
        ["Producto", "Componentes"]
    ].copy()

    df_componentes_53["Producto"] = (
        df_componentes_53["Producto"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_componentes_53["Componentes"] = (
        df_componentes_53["Componentes"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_componentes_53 = (
        df_componentes_53
        .drop_duplicates("Producto")
    )

    df_53 = df_acciones_52.merge(
        df_componentes_53,
        on="Producto",
        how="left"
    )

    def normalizar_53(texto):
        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

    def separar_componentes_53(texto):
        texto = str(texto).strip()

        if not texto:
            return []

        return [
            parte.strip()
            for parte in re.split(
                r";|,",
                texto
            )
            if parte.strip()
        ]

    def buscar_componente_53(fila):
        accion = normalizar_53(
            fila["Acción general"]
        )

        componentes = separar_componentes_53(
            fila["Componentes"]
        )

        encontrados = []

        for componente in componentes:
            componente_n = normalizar_53(
                componente
            )

            if (
                componente_n
                and componente_n in accion
            ):
                encontrados.append(
                    componente
                )

        return "; ".join(encontrados)

    df_53["Componente identificado"] = (
        df_53.apply(
            buscar_componente_53,
            axis=1
        )
    )

    df_53["Clasificación"] = (
        df_53["Componente identificado"]
        .apply(
            lambda x:
                "COMPONENTE + FUNCIÓN"
                if x
                else "ACCIÓN GENERAL"
        )
    )

    df_53 = df_53[
        [
            "Producto",
            "Acción general",
            "Componentes",
            "Componente identificado",
            "Clasificación"
        ]
    ].copy()

    st.session_state[
        "df_clasificado_53"
    ] = df_53.copy()

    st.success(
        f"🟢 5.3 TERMINADO: "
        f"{len(df_53)} relaciones clasificadas."
    )

    st.dataframe(
        df_53,
        use_container_width=True,
        hide_index=True
    )

except Exception as e:
    st.error(
        f"🔴 5.3 ERROR: "
        f"{type(e).__name__}: {e}"
    )

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

        # Por ahora NO suponemos el nombre de la hoja.
        # El usuario selecciona la hoja real.
        hoja_fuente = st.selectbox(
            "Seleccione la hoja de la matriz que contiene la información:",
            libro.sheet_names,
            key="hoja_matriz_normalizacion"
        )

        df_fuente = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name=hoja_fuente
        )

        # Eliminamos únicamente columnas completamente vacías.
        df_fuente = df_fuente.dropna(
            axis=1,
            how="all"
        )

        st.success(
            f"✅ Hoja cargada correctamente: **{hoja_fuente}**"
        )

        st.info(
            f"Registros encontrados: **{len(df_fuente)}**  |  "
            f"Columnas encontradas: **{len(df_fuente.columns)}**"
        )

        st.write("### Columnas REALES encontradas en la hoja")

        columnas_reales = pd.DataFrame({
            "N.º": range(1, len(df_fuente.columns) + 1),
            "Nombre real de la columna": [
                str(col) for col in df_fuente.columns
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
        f"🔴 5.1 ERROR al leer la matriz: {type(e).__name__}: {e}"
    )
# ============================================================
# 5.2 IDENTIFICACIÓN AUTOMÁTICA DE PRODUCTO Y ACCIONES
# ============================================================

st.markdown("### 5.2 Identificación de producto y acciones")

try:
    columnas = list(df_fuente.columns)

    if len(columnas) == 0:
        st.error(
            "🔴 5.2 ERROR: La hoja no contiene columnas."
        )
    else:

        # ----------------------------------------------------
        # Normalización temporal SOLO para analizar encabezados
        # ----------------------------------------------------
        def normalizar_encabezado(valor):
            texto = str(valor).strip().lower()

            reemplazos = {
                "á": "a",
                "é": "e",
                "í": "i",
                "ó": "o",
                "ú": "u",
                "ü": "u",
                "ñ": "n"
            }

            for origen, destino in reemplazos.items():
                texto = texto.replace(origen, destino)

            texto = " ".join(texto.split())

            return texto

        columnas_normalizadas = {
            columna: normalizar_encabezado(columna)
            for columna in columnas
        }

        # ----------------------------------------------------
        # Palabras orientadoras.
        # NO cambian los nombres reales de las columnas.
        # ----------------------------------------------------
        indicadores_producto = [
            "producto",
            "nombre producto",
            "nombre del producto",
            "producto nombre"
        ]

        indicadores_accion = [
            "accion",
            "acciones",
            "accion general",
            "acciones generales",
            "accion_general",
            "acciones_generales"
        ]

        columna_producto = None
        columna_accion = None

        # ----------------------------------------------------
        # Buscar coincidencia exacta primero
        # ----------------------------------------------------
        for columna, nombre_normalizado in columnas_normalizadas.items():

            if nombre_normalizado in indicadores_producto:
                columna_producto = columna
                break

        for columna, nombre_normalizado in columnas_normalizadas.items():

            if nombre_normalizado in indicadores_accion:
                columna_accion = columna
                break

        # ----------------------------------------------------
        # Mostrar resultado del análisis
        # ----------------------------------------------------
        if columna_producto is not None:
            st.success(
                f"🟢 Producto identificado: **{columna_producto}**"
            )
        else:
            st.warning(
                "🟡 No se identificó automáticamente una columna "
                "de producto."
            )

        if columna_accion is not None:
            st.success(
                f"🟢 Acciones identificadas: **{columna_accion}**"
            )
        else:
            st.warning(
                "🟡 No se identificó automáticamente una columna "
                "de acciones."
            )

        # ----------------------------------------------------
        # Si ambas fueron identificadas, mostrar ejemplos
        # ----------------------------------------------------
        if columna_producto is not None and columna_accion is not None:

            df_ejemplo = df_fuente[
                [columna_producto, columna_accion]
            ].copy()

            df_ejemplo = df_ejemplo.dropna(
                how="all"
            )

            st.write("### Ejemplo de información identificada")

            st.dataframe(
                df_ejemplo.head(15),
                use_container_width=True,
                hide_index=True
            )

            st.success(
                "🟢 5.2 TERMINADO: Producto y acciones fueron "
                "identificados. Todavía NO se ha separado ni "
                "normalizado ninguna acción."
            )

        else:

            st.error(
                "🔴 5.2 DETENIDO: No se continuará con la "
                "normalización hasta identificar correctamente "
                "las columnas necesarias."
            )

except Exception as e:

    st.error(
        f"🔴 5.2 ERROR: {type(e).__name__}: {e}"
    )

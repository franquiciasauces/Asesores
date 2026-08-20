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
# 5.2 IDENTIFICACIÓN DE COLUMNAS REALES
# ============================================================

st.markdown("### 5.2 Identificación de columnas")

try:

    if "df_fuente" not in locals():
        st.error("🔴 5.2 ERROR: No está disponible la matriz cargada en 5.1.")
    else:

        columnas = list(df_fuente.columns)

        st.write("**Columnas que está utilizando el código:**")

        for i, columna in enumerate(columnas, start=1):
            st.write(f"{i}. `{columna}`")

        # ----------------------------------------------------
        # Buscar producto y acción usando SOLO los encabezados
        # reales que ya existen en la matriz.
        # ----------------------------------------------------

        producto_encontrado = None
        accion_encontrada = None

        for columna in columnas:

            nombre = str(columna).strip().lower()

            if (
                producto_encontrado is None
                and "producto" in nombre
            ):
                producto_encontrado = columna

            if (
                accion_encontrada is None
                and "accion" in nombre
            ):
                accion_encontrada = columna

        # ----------------------------------------------------
        # Mostrar resultado
        # ----------------------------------------------------

        if producto_encontrado is not None:
            st.success(
                f"🟢 Producto detectado: **{producto_encontrado}**"
            )
        else:
            st.error(
                "🔴 No se encontró automáticamente una columna "
                "relacionada con producto."
            )

        if accion_encontrada is not None:
            st.success(
                f"🟢 Acción detectada: **{accion_encontrada}**"
            )
        else:
            st.error(
                "🔴 No se encontró automáticamente una columna "
                "relacionada con acción."
            )

        # ----------------------------------------------------
        # Mostrar ejemplo REAL
        # ----------------------------------------------------

        if producto_encontrado is not None and accion_encontrada is not None:

            ejemplo = df_fuente[
                [producto_encontrado, accion_encontrada]
            ].dropna(how="all").head(10)

            st.write("### Ejemplo detectado")

            st.dataframe(
                ejemplo,
                use_container_width=True,
                hide_index=True
            )

            st.success(
                "✅ 5.2 TERMINADO: Las columnas fueron identificadas. "
                "No se ha modificado ni normalizado ninguna información."
            )

except Exception as e:

    st.error(
        f"🔴 5.2 ERROR: {type(e).__name__}: {e}"
    )

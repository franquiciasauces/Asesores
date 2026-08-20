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
# 5.2 NORMALIZACIÓN INICIAL: PRODUCTO + ACCIONES GENERALES
# FUENTE EXCLUSIVA: COLUMNA A Y COLUMNA E
# ============================================================

st.markdown("### 5.2 Normalización de acciones generales")

try:

    if "df_fuente" not in locals():

        st.error(
            "🔴 5.2 ERROR: No está disponible la matriz "
            "cargada en 5.1."
        )

    elif len(df_fuente.columns) < 5:

        st.error(
            "🔴 5.2 ERROR: La hoja seleccionada no tiene "
            "al menos 5 columnas. Se necesitan A y E."
        )

    else:

        # ----------------------------------------------------
        # FUENTES FIJAS
        # A = Producto
        # E = Acciones generales
        # ----------------------------------------------------

        columna_producto = df_fuente.columns[0]
        columna_acciones = df_fuente.columns[4]

        st.info(
            f"Fuente producto: **columna A → {columna_producto}**"
        )

        st.info(
            f"Fuente acciones: **columna E → {columna_acciones}**"
        )

        # ----------------------------------------------------
        # TOMAR EXCLUSIVAMENTE A Y E
        # ----------------------------------------------------

        df_trabajo = df_fuente[
            [columna_producto, columna_acciones]
        ].copy()

        # ----------------------------------------------------
        # GENERAR REGISTROS
        # ----------------------------------------------------

        registros = []

        for _, fila in df_trabajo.iterrows():

            producto = fila[columna_producto]
            acciones = fila[columna_acciones]

            if pd.isna(producto):
                continue

            producto = str(producto).strip()

            if not producto:
                continue

            if pd.isna(acciones):
                continue

            acciones = str(acciones).strip()

            if not acciones:
                continue

            # ------------------------------------------------
            # Separadores posibles dentro de ACCIONES GENERALES
            # ------------------------------------------------

            partes = re.split(
                r"\s*(?:\n|;|\||•|\u2022)\s*",
                acciones
            )

            for parte in partes:

                accion = str(parte).strip()

                if not accion:
                    continue

                registros.append(
                    {
                        "Nombre del producto": producto,
                        "Acción": accion
                    }
                )

        # ----------------------------------------------------
        # CREAR DATAFRAME
        # ----------------------------------------------------

        df_normalizado = pd.DataFrame(registros)

        # ----------------------------------------------------
        # CÓDIGO AUTOMÁTICO
        # ----------------------------------------------------

        if not df_normalizado.empty:

            df_normalizado.insert(
                0,
                "Código",
                [
                    f"AG{numero:06d}"
                    for numero in range(
                        1,
                        len(df_normalizado) + 1
                    )
                ]
            )

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        if df_normalizado.empty:

            st.error(
                "🔴 5.2 ERROR: No se encontraron acciones "
                "generales en la columna E."
            )

        else:

            st.success(
                f"🟢 5.2 OK: Se generaron "
                f"**{len(df_normalizado)} acciones** "
                f"para **{df_normalizado['Nombre del producto'].nunique()} productos**."
            )

            st.write(
                "### Resultado provisional"
            )

            st.dataframe(
                df_normalizado.head(100),
                use_container_width=True,
                hide_index=True
            )

            # ------------------------------------------------
            # VALIDACIÓN ESPECÍFICA DE PROSTENFIT
            # ------------------------------------------------

            ejemplo = df_normalizado[
                df_normalizado[
                    "Nombre del producto"
                ].str.contains(
                    "PROSTENFIT",
                    case=False,
                    na=False
                )
            ]

            if not ejemplo.empty:

                st.success(
                    f"🟢 PROSTENFIT: "
                    f"**{len(ejemplo)} acciones independientes** detectadas."
                )

                st.dataframe(
                    ejemplo,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "ℹ️ No se encontró PROSTENFIT en el resultado."
                )

            st.success(
                "✅ 5.2 TERMINADO: La normalización provisional "
                "utilizó EXCLUSIVAMENTE las columnas A y E. "
                "Las demás columnas fueron ignoradas."
            )

except Exception as e:

    st.error(
        f"🔴 5.2 ERROR: {type(e).__name__}: {e}"
    )

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
# 5.2 NORMALIZACIÓN INICIAL DE ACCIONES GENERALES
# FUENTE EXCLUSIVA:
# A = Producto
# E = Acciones generales
# ============================================================

st.markdown("### 5.2 Normalización inicial de acciones generales")

try:

    if "df_fuente" not in locals():

        st.error(
            "🔴 5.2 ERROR: No está disponible la matriz "
            "cargada en 5.1."
        )

    elif len(df_fuente.columns) < 5:

        st.error(
            "🔴 5.2 ERROR: La hoja seleccionada tiene menos "
            "de 5 columnas. Se necesitan las columnas A y E."
        )

    else:

        # ----------------------------------------------------
        # IDENTIFICACIÓN FIJA DE LAS COLUMNAS DE ORIGEN
        # ----------------------------------------------------

        columna_producto = df_fuente.columns[0]
        columna_acciones = df_fuente.columns[4]

        st.info(
            f"📌 Columna A utilizada como PRODUCTO: "
            f"**{columna_producto}**"
        )

        st.info(
            f"📌 Columna E utilizada como ACCIONES GENERALES: "
            f"**{columna_acciones}**"
        )

        # ----------------------------------------------------
        # COPIA DE TRABAJO:
        # SOLO A Y E
        # ----------------------------------------------------

        df_trabajo = df_fuente[
            [columna_producto, columna_acciones]
        ].copy()

        registros = []

        # ----------------------------------------------------
        # PROCESAR CADA PRODUCTO
        # ----------------------------------------------------

        for _, fila in df_trabajo.iterrows():

            producto = fila[columna_producto]
            acciones = fila[columna_acciones]

            # -----------------------------------------------
            # Validar producto
            # -----------------------------------------------

            if pd.isna(producto):
                continue

            producto = str(producto).strip()

            if not producto:
                continue

            # -----------------------------------------------
            # Validar acciones
            # -----------------------------------------------

            if pd.isna(acciones):
                continue

            acciones = str(acciones).strip()

            if not acciones:
                continue

            # -----------------------------------------------
            # Separadores de acciones
            #
            # NO se utiliza ninguna otra columna.
            # -----------------------------------------------

            texto_acciones = acciones

            for separador in [
                ";",
                "|",
                "•",
                "\n"
            ]:

                texto_acciones = texto_acciones.replace(
                    separador,
                    "\n"
                )

            partes = texto_acciones.split("\n")

            # -----------------------------------------------
            # UNA ACCIÓN = UNA FILA
            # -----------------------------------------------

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
        # CREAR DATAFRAME NORMALIZADO
        # ----------------------------------------------------

        df_normalizado = pd.DataFrame(
            registros,
            columns=[
                "Nombre del producto",
                "Acción"
            ]
        )

        # ----------------------------------------------------
        # GENERAR CÓDIGO AUTOMÁTICO
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
        # VALIDAR RESULTADO
        # ----------------------------------------------------

        if df_normalizado.empty:

            st.error(
                "🔴 5.2 ERROR: No se encontraron registros "
                "para normalizar utilizando A y E."
            )

        else:

            cantidad_productos = (
                df_normalizado[
                    "Nombre del producto"
                ]
                .nunique()
            )

            cantidad_acciones = len(
                df_normalizado
            )

            st.success(
                f"🟢 5.2 OK: "
                f"{cantidad_productos} productos | "
                f"{cantidad_acciones} acciones independientes."
            )

            # ------------------------------------------------
            # MOSTRAR SOLO LAS 3 COLUMNAS DEFINIDAS
            # ------------------------------------------------

            st.write(
                "### Matriz de normalización provisional"
            )

            st.dataframe(
                df_normalizado,
                use_container_width=True,
                hide_index=True
            )

            # ------------------------------------------------
            # VALIDACIÓN DE PROSTENFIT
            # ------------------------------------------------

            ejemplo_prostenfit = df_normalizado[
                df_normalizado[
                    "Nombre del producto"
                ]
                .str.contains(
                    "PROSTENFIT",
                    case=False,
                    na=False
                )
            ]

            if not ejemplo_prostenfit.empty:

                st.success(
                    "🟢 VALIDACIÓN PROSTENFIT: "
                    f"{len(ejemplo_prostenfit)} "
                    "acciones independientes."
                )

                st.dataframe(
                    ejemplo_prostenfit,
                    use_container_width=True,
                    hide_index=True
                )

            # ------------------------------------------------
            # CONFIRMACIÓN DE ESTRUCTURA
            # ------------------------------------------------

            columnas_salida = list(
                df_normalizado.columns
            )

            if columnas_salida == [
                "Código",
                "Nombre del producto",
                "Acción"
            ]:

                st.success(
                    "✅ Estructura correcta: "
                    "Código | Nombre del producto | Acción"
                )

            else:

                st.error(
                    "🔴 ERROR: La estructura de salida "
                    "no coincide con las tres columnas requeridas."
                )

            st.info(
                "ℹ️ 5.2 todavía NO realiza la depuración "
                "semántica de las acciones. En esta etapa "
                "solo se extraen A y E, se separan las acciones "
                "y se genera una fila y un código por acción."
            )

except Exception as e:

    st.error(
        f"🔴 5.2 ERROR: {type(e).__name__}: {e}"
    )

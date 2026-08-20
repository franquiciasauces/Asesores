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
# 5.2 SEPARACIÓN DE PRODUCTO, ACCIONES Y GENERACIÓN DE CÓDIGO
# ============================================================

st.markdown("### 5.2 Separación de productos y acciones")

try:

    if "df_fuente" not in locals():

        st.error(
            "🔴 5.2 ERROR: No se encuentra la matriz cargada "
            "por el bloque 5.1."
        )

    else:

        registros_normalizados = []

        # ----------------------------------------------------
        # Recorrer cada fila de la matriz original
        # ----------------------------------------------------

        for _, fila in df_fuente.iterrows():

            valores = []

            for valor in fila.tolist():

                if pd.isna(valor):
                    continue

                texto = str(valor).strip()

                if texto == "":
                    continue

                valores.append(texto)

            if not valores:
                continue

            # ------------------------------------------------
            # La primera celda puede contener:
            #
            # PRODUCTO,PRIMERA ACCIÓN
            #
            # ------------------------------------------------

            primera_celda = valores[0]

            if "," in primera_celda:

                producto, primera_accion = (
                    primera_celda.split(",", 1)
                )

                producto = producto.strip()
                primera_accion = primera_accion.strip()

                if primera_accion:
                    acciones = [primera_accion]
                else:
                    acciones = []

            else:

                producto = primera_celda.strip()
                acciones = []

            # ------------------------------------------------
            # Las demás columnas contienen acciones adicionales
            # ------------------------------------------------

            for valor in valores[1:]:

                accion = str(valor).strip()

                if accion:
                    acciones.append(accion)

            # ------------------------------------------------
            # Una acción = una fila
            # ------------------------------------------------

            for accion in acciones:

                if accion.strip():

                    registros_normalizados.append(
                        {
                            "Nombre del producto": producto,
                            "Acción": accion.strip()
                        }
                    )

        # ----------------------------------------------------
        # Crear DataFrame temporal
        # ----------------------------------------------------

        df_normalizado = pd.DataFrame(
            registros_normalizados
        )

        # ----------------------------------------------------
        # Generar código automático
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
        # Mostrar resultado
        # ----------------------------------------------------

        if df_normalizado.empty:

            st.error(
                "🔴 5.2 ERROR: No se pudieron extraer acciones "
                "de la matriz."
            )

        else:

            st.success(
                f"🟢 5.2 OK: Se generaron "
                f"**{len(df_normalizado)} registros de acciones** "
                f"a partir de **{df_normalizado['Nombre del producto'].nunique()} productos**."
            )

            st.write(
                "### Matriz normalizada provisional"
            )

            st.dataframe(
                df_normalizado.head(50),
                use_container_width=True,
                hide_index=True
            )

            # ------------------------------------------------
            # Ejemplo específico para comprobar múltiples
            # acciones de un mismo producto
            # ------------------------------------------------

            producto_ejemplo = (
                "FITO PROSTENFIT x 60 CAP"
            )

            ejemplo = df_normalizado[
                df_normalizado[
                    "Nombre del producto"
                ].str.upper()
                == producto_ejemplo.upper()
            ]

            if not ejemplo.empty:

                st.success(
                    "🟢 Ejemplo PROSTENFIT: "
                    f"se encontraron **{len(ejemplo)} acciones independientes**."
                )

                st.dataframe(
                    ejemplo,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "ℹ️ No se encontró PROSTENFIT con ese nombre "
                    "exacto en los registros procesados."
                )

except Exception as e:

    st.error(
        f"🔴 5.2 ERROR: {type(e).__name__}: {e}"
    )

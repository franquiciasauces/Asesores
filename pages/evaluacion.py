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
# 5.3 CLASIFICAR CADA ACCIÓN INDIVIDUAL
# ============================================================

st.markdown("### 5.3 Clasificación de acciones")

try:

    import re

    if "df_acciones_52" not in st.session_state:
        st.error("❌ 5.3 ERROR: No existe df_acciones_52.")
        st.stop()

    df_acciones_53 = st.session_state[
        "df_acciones_52"
    ].copy()

    if "Componentes" not in df_fuente.columns:
        st.error(
            "❌ 5.3 ERROR: No existe la columna real "
            "'Componentes' en la matriz."
        )
        st.stop()

    componentes_producto_53 = (
        df_fuente[
            ["Producto", "Componentes"]
        ]
        .drop_duplicates("Producto")
        .copy()
    )

    df_acciones_53 = df_acciones_53.merge(
        componentes_producto_53,
        on="Producto",
        how="left"
    )

    categorias_53 = [
        "ACCIÓN GENERAL",
        "COMPONENTE + FUNCIÓN",
        "RECOMENDACIÓN / COMPLEMENTO",
        "USO / POSOLOGÍA / PRECAUCIÓN",
        "RESTRICCIÓN / CONTRAINDICACIÓN",
        "COMERCIAL"
    ]

    def clasificar_accion_53(fila):

        accion = str(
            fila["Acción general"]
        ).strip()

        componentes = str(
            fila["Componentes"]
        ).strip()

        texto = accion.lower()

        if (
            "frase comercial" in texto
            or any(
                x in texto
                for x in [
                    "compra",
                    "ideal para ti",
                    "excelente opción",
                    "lleva una vida",
                    "tu mejor opción"
                ]
            )
        ):
            return "COMERCIAL"

        if any(
            x in texto
            for x in [
                "contraindicado",
                "contraindicación",
                "no usar",
                "no se recomienda",
                "evitar en",
                "precaución"
            ]
        ):
            return "RESTRICCIÓN / CONTRAINDICACIÓN"

        if any(
            x in texto
            for x in [
                "tomar",
                "consumir",
                "ingerir",
                "cápsula al día",
                "cápsulas al día",
                "ml al día",
                "por día",
                "dos veces al día"
            ]
        ):
            return "USO / POSOLOGÍA / PRECAUCIÓN"

        if any(
            x in texto
            for x in [
                "recomendado como complemento",
                "complemento",
                "se recomienda",
                "acompañar con",
                "acompañado de"
            ]
        ):
            return "RECOMENDACIÓN / COMPLEMENTO"

        componentes_lista = [
            x.strip().lower()
            for x in re.split(
                r";|,",
                componentes
            )
            if x.strip()
        ]

        if componentes_lista and any(
            componente in texto
            for componente in componentes_lista
        ):
            return "COMPONENTE + FUNCIÓN"

        return "ACCIÓN GENERAL"

    df_acciones_53[
        "Clasificación"
    ] = df_acciones_53.apply(
        clasificar_accion_53,
        axis=1
    )

    st.session_state[
        "df_acciones_53"
    ] = df_acciones_53.copy()

    st.success(
        f"🟢 5.3 TERMINADO: "
        f"{len(df_acciones_53)} acciones clasificadas."
    )

    st.dataframe(
        df_acciones_53[
            [
                "Producto",
                "Acción general",
                "Clasificación"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

except Exception as e:

    st.error(
        f"🔴 5.3 ERROR: "
        f"{type(e).__name__}: {e}"
    )

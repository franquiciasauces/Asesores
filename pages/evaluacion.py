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
# 5.3 DEPURAR Y CLASIFICAR ACCIONES INDIVIDUALES
# ============================================================

st.markdown("### 5.3 Depuración y clasificación de acciones")

try:
    import re
    import unicodedata

    df_base_53 = st.session_state.get(
        "df_acciones_52"
    )

    if df_base_53 is None or df_base_53.empty:
        st.error(
            "❌ 5.3 ERROR: No existe df_acciones_52."
        )
        st.stop()

    requeridas_53 = [
        "Producto",
        "Acción general"
    ]

    faltantes_53 = [
        c
        for c in requeridas_53
        if c not in df_base_53.columns
    ]

    if faltantes_53:
        st.error(
            "❌ 5.3 ERROR: Faltan columnas: "
            + ", ".join(faltantes_53)
        )
        st.stop()

    df_53 = df_base_53.copy()

    # --------------------------------------------------------
    # NORMALIZACIÓN DE TEXTO
    # --------------------------------------------------------

    def normalizar_53(texto):
        texto = str(texto).strip().lower()
        texto = unicodedata.normalize(
            "NFKD",
            texto
        )
        texto = "".join(
            c
            for c in texto
            if not unicodedata.combining(c)
        )
        texto = re.sub(
            r"\s+",
            " ",
            texto
        )
        return texto

    # --------------------------------------------------------
    # COLUMNAS REALES DE LA MATRIZ
    # --------------------------------------------------------

    columnas_matriz_53 = {
        normalizar_53(c): c
        for c in df_fuente.columns
    }

    col_componentes_53 = None

    for nombre in [
        "componentes",
        "componente",
        "componentes del producto"
    ]:
        if nombre in columnas_matriz_53:
            col_componentes_53 = columnas_matriz_53[
                nombre
            ]
            break

    if col_componentes_53 is None:
        st.error(
            "❌ 5.3 ERROR: No se encontró en la matriz "
            "una columna real de Componentes."
        )
        st.stop()

    col_producto_53 = None

    for nombre in [
        "producto",
        "nombre del producto"
    ]:
        if nombre in columnas_matriz_53:
            col_producto_53 = columnas_matriz_53[
                nombre
            ]
            break

    if col_producto_53 is None:
        st.error(
            "❌ 5.3 ERROR: No se encontró la columna "
            "real de Producto."
        )
        st.stop()

    # --------------------------------------------------------
    # COMPONENTES REALES POR PRODUCTO
    # --------------------------------------------------------

    componentes_producto_53 = {}

    for _, fila in df_fuente.iterrows():

        producto = str(
            fila[col_producto_53]
        ).strip()

        if not producto:
            continue

        componentes = str(
            fila[col_componentes_53]
        ).strip()

        lista = re.split(
            r"\s*;\s*|\s*,\s*|\s*\n\s*",
            componentes
        )

        componentes_producto_53[
            normalizar_53(producto)
        ] = [
            normalizar_53(x)
            for x in lista
            if normalizar_53(x)
        ]

    # --------------------------------------------------------
    # MARCADORES QUE INVALIDAN LA ACCIÓN
    # --------------------------------------------------------

    marcadores_eliminar_53 = [
        "frase comercial",
        "combinaciones",
        "recomendacion",
        "recomendación",
        "complementario",
        "complementarios",
        "contraindicacion",
        "contraindicación",
        "restriccion",
        "restricción",
        "posologia",
        "posología",
        "modo de accion",
        "modo de acción"
    ]

    marcadores_uso_53 = [
        "dosis",
        "tomar",
        "consumir",
        "uso",
        "posologia",
        "posología",
        "precaucion",
        "precaución",
        "advertencia"
    ]

    # --------------------------------------------------------
    # COMPONENTE REAL MENCIONADO EN LA ACCIÓN
    # --------------------------------------------------------

    def componente_mencionado_53(
        accion,
        componentes
    ):
        accion_n = normalizar_53(
            accion
        )

        encontrados = []

        for componente in componentes:

            if not componente:
                continue

            if componente in accion_n:
                encontrados.append(
                    componente
                )

        return encontrados

    # --------------------------------------------------------
    # CLASIFICACIÓN
    # --------------------------------------------------------

    def clasificar_53(fila):

        producto = str(
            fila["Producto"]
        ).strip()

        accion = str(
            fila["Acción general"]
        ).strip()

        accion_n = normalizar_53(
            accion
        )

        if not accion_n:
            return "ELIMINAR"

        # --------------------------------------------
        # 1. FRASE COMERCIAL / COMBINACIONES
        # --------------------------------------------

        for marcador in marcadores_eliminar_53:

            if normalizar_53(
                marcador
            ) in accion_n:

                return "ELIMINAR"

        # --------------------------------------------
        # 2. PRODUCTO Y COMPONENTES REALES
        # --------------------------------------------

        componentes = componentes_producto_53.get(
            normalizar_53(producto),
            []
        )

        encontrados = componente_mencionado_53(
            accion,
            componentes
        )

        # --------------------------------------------
        # 3. COMPONENTE + FUNCIÓN
        #
        # SOLO SI EL COMPONENTE APARECE
        # REALMENTE EN LOS COMPONENTES DEL PRODUCTO
        # --------------------------------------------

        if encontrados:
            return "COMPONENTE + FUNCIÓN"

        # --------------------------------------------
        # 4. USO / POSOLOGÍA / PRECAUCIÓN
        # --------------------------------------------

        for marcador in marcadores_uso_53:

            if normalizar_53(
                marcador
            ) in accion_n:

                return (
                    "USO / POSOLOGÍA / PRECAUCIÓN"
                )

        # --------------------------------------------
        # 5. RESTO
        # --------------------------------------------

        return "ACCIÓN GENERAL"

    df_53[
        "Clasificación"
    ] = df_53.apply(
        clasificar_53,
        axis=1
    )

    # --------------------------------------------------------
    # ELIMINAR LO QUE NO DEBE ENTRAR EN LA MATRIZ DEPURADA
    # --------------------------------------------------------

    df_depurado_53 = df_53[
        df_53["Clasificación"]
        != "ELIMINAR"
    ].copy()

    # --------------------------------------------------------
    # GUARDAR
    # --------------------------------------------------------

    st.session_state[
        "df_depurado_53"
    ] = df_depurado_53.copy()

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    st.success(
        f"🟢 5.3 TERMINADO: "
        f"{len(df_53)} acciones revisadas → "
        f"{len(df_depurado_53)} acciones depuradas."
    )

    st.dataframe(
        df_depurado_53,
        use_container_width=True,
        hide_index=True
    )

except Exception as e:

    st.error(
        f"🔴 5.3 ERROR: "
        f"{type(e).__name__}: {e}"
    )

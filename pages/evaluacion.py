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
# 5.3 CLASIFICACIÓN DE ACCIONES INDIVIDUALES

```python
# ============================================================
# 5.3 CLASIFICACIÓN DE ACCIONES INDIVIDUALES
# ============================================================

st.markdown("### 5.3 Clasificación de acciones individuales")

try:
    import re

    # --------------------------------------------------------
    # 1. VERIFICAR ENTRADAS REALES DE 5.1 Y 5.2
    # --------------------------------------------------------

    if "df_acciones_52" not in st.session_state:
        st.error(
            "❌ 5.3 ERROR: No existe df_acciones_52. "
            "Debe ejecutarse primero el 5.2."
        )
        st.stop()

    df_acciones_52 = st.session_state[
        "df_acciones_52"
    ].copy()

    if "df_fuente" not in locals():
        st.error(
            "❌ 5.3 ERROR: No existe la matriz fuente "
            "cargada por 5.1."
        )
        st.stop()

    # --------------------------------------------------------
    # 2. VALIDAR COLUMNAS REALES NECESARIAS
    # --------------------------------------------------------

    columnas_53 = [
        "Producto",
        "Acciones generales",
        "Componentes"
    ]

    faltantes_53 = [
        columna
        for columna in columnas_53
        if columna not in df_fuente.columns
    ]

    if faltantes_53:
        st.error(
            "❌ 5.3 ERROR: Faltan columnas reales en la matriz: "
            + ", ".join(faltantes_53)
        )
        st.stop()

    if "Producto" not in df_acciones_52.columns:
        st.error(
            "❌ 5.3 ERROR: df_acciones_52 no contiene "
            "la columna Producto."
        )
        st.stop()

    if "Acción general" not in df_acciones_52.columns:
        st.error(
            "❌ 5.3 ERROR: df_acciones_52 no contiene "
            "la columna Acción general."
        )
        st.stop()

    # --------------------------------------------------------
    # 3. NORMALIZACIÓN SOLO PARA COMPARAR
    #
    # No modifica los textos originales.
    # --------------------------------------------------------

    def normalizar_53(texto):
        texto = str(texto).lower().strip()

        texto = re.sub(
            r"[^a-záéíóúüñ0-9\s]",
            " ",
            texto
        )

        texto = re.sub(
            r"\s+",
            " ",
            texto
        )

        return texto.strip()

    # --------------------------------------------------------
    # 4. CATÁLOGO REAL DE PRODUCTOS
    # --------------------------------------------------------

    productos_53 = set()

    for valor in df_fuente["Producto"].fillna(""):
        valor = str(valor).strip()

        if valor:
            productos_53.add(
                normalizar_53(valor)
            )

    # --------------------------------------------------------
    # 5. COMPONENTES PROPIOS DE CADA PRODUCTO
    # --------------------------------------------------------

    componentes_producto_53 = {}

    for _, fila in df_fuente.iterrows():

        producto = str(
            fila["Producto"]
        ).strip()

        if not producto:
            continue

        clave_producto = normalizar_53(
            producto
        )

        texto_componentes = str(
            fila["Componentes"]
        ).strip()

        componentes = []

        if texto_componentes:
            componentes = re.split(
                r"\s*;\s*|\s*,\s*",
                texto_componentes
            )

        componentes_producto_53[
            clave_producto
        ] = {
            normalizar_53(c)
            for c in componentes
            if normalizar_53(c)
        }

    # --------------------------------------------------------
    # 6. CATÁLOGO GLOBAL DE COMPONENTES
    #
    # Se obtiene de la matriz.
    # --------------------------------------------------------

    componentes_globales_53 = set()

    for componentes in (
        componentes_producto_53.values()
    ):
        componentes_globales_53.update(
            componentes
        )

    # --------------------------------------------------------
    # 7. ORDEN DE CLASIFICACIÓN
    # --------------------------------------------------------

    def clasificar_53(producto, accion):

        producto_n = normalizar_53(
            producto
        )

        accion_n = normalizar_53(
            accion
        )

        if not accion_n:
            return "ELIMINAR"

        # ----------------------------------------------------
        # FRASES COMERCIALES
        # ----------------------------------------------------

        if (
            "frase comercial" in accion_n
            or "frase de venta" in accion_n
        ):
            return "ELIMINAR"

        # ----------------------------------------------------
        # COMBINACIONES
        # ----------------------------------------------------

        if (
            "combinaciones" in accion_n
            or "combinar con" in accion_n
            or "combinado con" in accion_n
        ):
            return "RECOMENDACIÓN / COMPLEMENTO"

        # ----------------------------------------------------
        # POSOLOGÍA / USO / PRECAUCIÓN
        # ----------------------------------------------------

        patrones_uso = [
            "dosis",
            "dosificación",
            "posología",
            "tomar",
            "consumir",
            "ingerir",
            "uso diario",
            "modo de uso",
            "precaución",
            "precauciones"
        ]

        if any(
            patron in accion_n
            for patron in patrones_uso
        ):
            return "USO / POSOLOGÍA / PRECAUCIÓN"

        # ----------------------------------------------------
        # RESTRICCIONES / CONTRAINDICACIONES
        # ----------------------------------------------------

        patrones_restriccion = [
            "contraindicación",
            "contraindicaciones",
            "restricción",
            "restricciones",
            "no usar",
            "no recomendado",
            "embarazo",
            "lactancia"
        ]

        if any(
            patron in accion_n
            for patron in patrones_restriccion
        ):
            return "RESTRICCIÓN / CONTRAINDICACIÓN"

        # ----------------------------------------------------
        # COMPONENTES DEL PRODUCTO ACTUAL
        # ----------------------------------------------------

        propios = componentes_producto_53.get(
            producto_n,
            set()
        )

        componentes_propios_mencionados = [
            componente
            for componente in propios
            if componente in accion_n
        ]

        # ----------------------------------------------------
        # COMPONENTE DE OTRO PRODUCTO
        #
        # Si aparece un componente conocido pero NO
        # pertenece al producto actual, es referencia externa.
        # ----------------------------------------------------

        componentes_externos = [
            componente
            for componente in componentes_globales_53
            if componente not in propios
            and len(componente) >= 4
            and componente in accion_n
        ]

        if componentes_externos:
            return "RECOMENDACIÓN / COMPLEMENTO"

        # ----------------------------------------------------
        # OTRO PRODUCTO DE LA MATRIZ
        #
        # No debe confundirse el producto consigo mismo.
        # ----------------------------------------------------

        for otro_producto in productos_53:

            if (
                otro_producto == producto_n
                or len(otro_producto) < 4
            ):
                continue

            if otro_producto in accion_n:
                return "RECOMENDACIÓN / COMPLEMENTO"

        # ----------------------------------------------------
        # COMPONENTE + FUNCIÓN
        #
        # SOLO si el componente aparece explícitamente
        # Y pertenece al producto actual.
        # ----------------------------------------------------

        if componentes_propios_mencionados:
            return "COMPONENTE + FUNCIÓN"

        # ----------------------------------------------------
        # RESTO
        # ----------------------------------------------------

        return "ACCIÓN GENERAL"

    # --------------------------------------------------------
    # 8. CLASIFICAR SIN MODIFICAR 5.2
    # --------------------------------------------------------

    df_clasificacion_53 = (
        df_acciones_52.copy()
    )

    df_clasificacion_53[
        "Clasificación"
    ] = df_clasificacion_53.apply(
        lambda fila: clasificar_53(
            fila["Producto"],
            fila["Acción general"]
        ),
        axis=1
    )

    # --------------------------------------------------------
    # 9. GUARDAR RESULTADO PARA 5.4
    # --------------------------------------------------------

    st.session_state[
        "df_clasificacion_53"
    ] = df_clasificacion_53.copy()

    # --------------------------------------------------------
    # 10. MOSTRAR RESULTADO
    # --------------------------------------------------------

    st.success(
        f"🟢 5.3 TERMINADO: "
        f"{len(df_clasificacion_53)} relaciones clasificadas."
    )

    st.dataframe(
        df_clasificacion_53,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # 11. RESUMEN
    # --------------------------------------------------------

    resumen_53 = (
        df_clasificacion_53[
            "Clasificación"
        ]
        .value_counts()
        .rename_axis(
            "Clasificación"
        )
        .reset_index(
            name="Cantidad"
        )
    )

    st.write(
        "### Resumen de clasificación"
    )

    st.dataframe(
        resumen_53,
        use_container_width=True,
        hide_index=True
    )

except Exception as e:
    st.error(
        f"🔴 5.3 ERROR: "
        f"{type(e).__name__}: {e}"
    )




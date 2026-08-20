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
```python
# ============================================================
# 5.8 — NORMALIZACIÓN UNIVERSAL DE ACCIONES GENERALES
# ============================================================

import re
import io


# ============================================================
# ARCHIVO PERMANENTE
# ============================================================

ARCHIVO_ACCIONES_GENERALES = (
    BASE_DIR / "ACCIONES_GENERALES.xlsx"
)


# ============================================================
# 5.8.1 — LIMPIEZA BÁSICA
# ============================================================

def limpiar_texto(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor)

    texto = texto.replace("\r\n", "\n")
    texto = texto.replace("\r", "\n")

    return texto.strip()


def normalizar_espacios(texto):

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


# ============================================================
# 5.8.2 — SECCIONES QUE NO PERTENECEN A ACCIONES
# ============================================================

MARCADORES_NO_ACCION = [
    r"MODO\s+DE\s+ACCI[ÓO]N\s*:?",
    r"COMBINACIONES?\s*:?",
    r"COMBINAR\s+CON\s*:?",
    r"FRASE\s+DE\s+VENTA\s*:?",
    r"RECOMENDACIONES?\s*:?",
    r"RECOMENDACI[ÓO]N\s+DE\s+USO\s*:?",
    r"PRECAUCIONES?\s*:?",
    r"CONTRAINDICACIONES?\s*:?",
    r"ADVERTENCIAS?\s*:?",
    r"COMPLEMENTAR\s+CON\s*:?",
    r"COMPLEMENTOS?\s*:?"
]


# ============================================================
# 5.8.3 — CORTAR SECCIONES POSTERIORES
# ============================================================

def cortar_secciones_no_accion(texto):

    posiciones = []

    for patron in MARCADORES_NO_ACCION:

        coincidencias = re.finditer(
            patron,
            texto,
            flags=re.IGNORECASE
        )

        for coincidencia in coincidencias:

            posiciones.append(
                coincidencia.start()
            )

    if posiciones:

        texto = texto[
            :min(posiciones)
        ]

    return texto.strip()


# ============================================================
# 5.8.4 — SEPARAR ELEMENTOS
# ============================================================

def separar_elementos(texto):

    texto = limpiar_texto(
        texto
    )

    if not texto:
        return []

    # Separadores estructurales
    partes = re.split(
        r"\s*;\s*|\n+|[•▪●]",
        texto
    )

    resultado = []

    for parte in partes:

        parte = normalizar_espacios(
            parte
        )

        if parte:
            resultado.append(
                parte
            )

    return resultado


# ============================================================
# 5.8.5 — IDENTIFICAR COMPONENTES
# ============================================================

def obtener_componentes(valor):

    return separar_elementos(
        valor
    )


# ============================================================
# 5.8.6 — NORMALIZAR NOMBRE DE COMPONENTE
# ============================================================

def normalizar_componente(
    componente
):

    componente = normalizar_espacios(
        componente
    ).lower()

    # Quitar información entre paréntesis
    componente = re.sub(
        r"\([^)]*\)",
        "",
        componente
    )

    return componente.strip()


# ============================================================
# 5.8.7 — DETECTAR SI EL TEXTO ES CLARAMENTE
# UNA REFERENCIA A COMPONENTE
# ============================================================

def contiene_referencia_explicita_componente(
    texto,
    componentes
):

    texto_lower = normalizar_espacios(
        texto
    ).lower()

    if not texto_lower:
        return False

    for componente in componentes:

        componente_norm = normalizar_componente(
            componente
        )

        if len(componente_norm) < 4:
            continue

        # Coincidencia de la expresión completa
        if componente_norm in texto_lower:
            return True

    return False


# ============================================================
# 5.8.8 — DETECTAR FRASES QUE SON RECOMENDACIONES,
# COMPLEMENTOS, PRECAUCIONES, ETC.
# ============================================================

def es_contenido_no_general(
    texto
):

    texto = normalizar_espacios(
        texto
    )

    if not texto:
        return True

    patrones = [

        r"(?i)^recomendad[oa]?\b",

        r"(?i)^recomendaci[óo]n\b",

        r"(?i)^se\s+recomienda\b",

        r"(?i)^complementar\s+con\b",

        r"(?i)^complemento\b",

        r"(?i)^combinar\s+con\b",

        r"(?i)^combinaci[óo]n\b",

        r"(?i)^precauci[óo]n\b",

        r"(?i)^contraindicaci[óo]n\b",

        r"(?i)^advertencia\b",

        r"(?i)^frase\s+de\s+venta\b",

        r"(?i)^modo\s+de\s+acci[óo]n\b"
    ]

    for patron in patrones:

        if re.search(
            patron,
            texto
        ):
            return True

    return False


# ============================================================
# 5.8.9 — OBTENER ACCIONES GENERALES
# ============================================================

def obtener_acciones_generales(
    valor_acciones,
    valor_componentes
):

    texto = limpiar_texto(
        valor_acciones
    )

    if not texto:
        return []

    # --------------------------------------------------------
    # Cortar únicamente cuando aparece un encabezado real
    # de una sección posterior.
    # --------------------------------------------------------

    texto = cortar_secciones_no_accion(
        texto
    )

    if not texto:
        return []

    componentes = obtener_componentes(
        valor_componentes
    )

    elementos = separar_elementos(
        texto
    )

    acciones = []

    for elemento in elementos:

        elemento = normalizar_espacios(
            elemento
        )

        if not elemento:
            continue

        # Encabezados / contenido no pertinente
        if es_contenido_no_general(
            elemento
        ):
            continue

        # ----------------------------------------------------
        # Si la propia frase identifica explícitamente un
        # componente, no se clasifica como acción general.
        # ----------------------------------------------------

        if contiene_referencia_explicita_componente(
            elemento,
            componentes
        ):
            continue

        acciones.append(
            elemento
        )

    # --------------------------------------------------------
    # DEDUPLICAR SIN ALTERAR EL TEXTO
    # --------------------------------------------------------

    resultado = []

    vistos = set()

    for accion in acciones:

        clave = accion.casefold()

        if clave in vistos:
            continue

        vistos.add(
            clave
        )

        resultado.append(
            accion
        )

    return resultado


# ============================================================
# 5.8.10 — GENERAR REGISTROS
# ============================================================

registros_nuevos = []


for _, fila in df_trabajo.iterrows():

    producto = limpiar_texto(
        fila["Producto"]
    )

    if not producto:
        continue

    acciones = obtener_acciones_generales(
        fila["Acciones generales"],
        fila["Componentes"]
    )

    for accion in acciones:

        registros_nuevos.append(
            {
                "Producto": producto,
                "Accion_general": accion
            }
        )


df_nuevo = pd.DataFrame(
    registros_nuevos,
    columns=[
        "Producto",
        "Accion_general"
    ]
)


# ============================================================
# 5.8.11 — ELIMINAR DUPLICADOS
# ============================================================

df_nuevo = (
    df_nuevo
    .drop_duplicates(
        subset=[
            "Producto",
            "Accion_general"
        ]
    )
    .reset_index(drop=True)
)


# ============================================================
# 5.8.12 — LEER ARCHIVO PERMANENTE
# ============================================================

if ARCHIVO_ACCIONES_GENERALES.exists():

    try:

        df_anterior = pd.read_excel(
            ARCHIVO_ACCIONES_GENERALES,
            sheet_name="Acciones_Generales"
        )

    except Exception:

        df_anterior = pd.DataFrame(
            columns=[
                "ID_Accion",
                "Producto",
                "Accion_general"
            ]
        )

else:

    df_anterior = pd.DataFrame(
        columns=[
            "ID_Accion",
            "Producto",
            "Accion_general"
        ]
    )


# ============================================================
# 5.8.13 — VALIDAR ESTRUCTURA PERMANENTE
# ============================================================

columnas_finales = [
    "ID_Accion",
    "Producto",
    "Accion_general"
]

if not set(columnas_finales).issubset(
    df_anterior.columns
):

    df_anterior = pd.DataFrame(
        columns=columnas_finales
    )
else:

    df_anterior = df_anterior[
        columnas_finales
    ].copy()


# ============================================================
# 5.8.14 — ÍNDICE DE REGISTROS EXISTENTES
# ============================================================

claves_existentes = {}

ids_utilizados = set()


for _, fila in df_anterior.iterrows():

    identificador = limpiar_texto(
        fila["ID_Accion"]
    )

    producto = limpiar_texto(
        fila["Producto"]
    )

    accion = limpiar_texto(
        fila["Accion_general"]
    )

    if identificador:

        ids_utilizados.add(
            identificador
        )

    if (
        producto
        and accion
        and identificador
    ):

        claves_existentes[
            (
                producto.casefold(),
                accion.casefold()
            )
        ] = identificador


# ============================================================
# 5.8.15 — SIGUIENTE ID ESTABLE
# ============================================================

def siguiente_id_accion():

    numeros = []

    for identificador in ids_utilizados:

        resultado = re.fullmatch(
            r"AG(\d+)",
            identificador
        )

        if resultado:

            numeros.append(
                int(
                    resultado.group(1)
                )
            )

    siguiente = (
        max(numeros) + 1
        if numeros
        else 1
    )

    nuevo_id = (
        f"AG{siguiente:06d}"
    )

    ids_utilizados.add(
        nuevo_id
    )

    return nuevo_id


# ============================================================
# 5.8.16 — CONSERVAR EXISTENTES Y AGREGAR NUEVOS
# ============================================================

registros_finales = []

claves_finales = set()


# Primero conservar el archivo permanente
for _, fila in df_anterior.iterrows():

    identificador = limpiar_texto(
        fila["ID_Accion"]
    )

    producto = limpiar_texto(
        fila["Producto"]
    )

    accion = limpiar_texto(
        fila["Accion_general"]
    )

    if not (
        identificador
        and producto
        and accion
    ):
        continue

    clave = (
        producto.casefold(),
        accion.casefold()
    )

    if clave in claves_finales:
        continue

    registros_finales.append(
        {
            "ID_Accion": identificador,
            "Producto": producto,
            "Accion_general": accion
        }
    )

    claves_finales.add(
        clave
    )


# Después incorporar únicamente lo nuevo
for _, fila in df_nuevo.iterrows():

    producto = limpiar_texto(
        fila["Producto"]
    )

    accion = limpiar_texto(
        fila["Accion_general"]
    )

    if not (
        producto
        and accion
    ):
        continue

    clave = (
        producto.casefold(),
        accion.casefold()
    )

    if clave in claves_finales:
        continue

    identificador = (
        siguiente_id_accion()
    )

    registros_finales.append(
        {
            "ID_Accion": identificador,
            "Producto": producto,
            "Accion_general": accion
        }
    )

    claves_finales.add(
        clave
    )


# ============================================================
# 5.8.17 — DATAFRAME FINAL
# ============================================================

df_acciones_generales = pd.DataFrame(
    registros_finales,
    columns=[
        "ID_Accion",
        "Producto",
        "Accion_general"
    ]
)


# ============================================================
# 5.8.18 — GUARDAR ARCHIVO PERMANENTE
# ============================================================

try:

    with pd.ExcelWriter(
        ARCHIVO_ACCIONES_GENERALES,
        engine="openpyxl"
    ) as writer:

        df_acciones_generales.to_excel(
            writer,
            sheet_name="Acciones_Generales",
            index=False
        )

    st.success(
        "✓ ACCIONES_GENERALES.xlsx "
        "actualizado correctamente."
    )

except Exception as error:

    st.error(
        f"No fue posible guardar "
        f"ACCIONES_GENERALES.xlsx: {error}"
    )


# ============================================================
# 5.8.19 — MOSTRAR
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
# 5.8.20 — DESCARGAR EXCEL
# ============================================================

buffer_excel = io.BytesIO()

with pd.ExcelWriter(
    buffer_excel,
    engine="openpyxl"
) as writer:

    df_acciones_generales.to_excel(
        writer,
        sheet_name="Acciones_Generales",
        index=False
    )


st.download_button(
    label="⬇️ Descargar ACCIONES_GENERALES.xlsx",
    data=buffer_excel.getvalue(),
    file_name="ACCIONES_GENERALES.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    key="descargar_acciones_generales"
)
```

# ============================================================

# ============================================================
# APLICATIVO ASESORES
# PAQUETE 1 - DIAGNÓSTICO Y CARGA DE ARCHIVOS
# ============================================================

from pathlib import Path
from unidecode import unidecode
from rapidfuzz import fuzz
import streamlit as st
import pandas as pd
import numpy as np


# ============================================================
# 1.1 CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Asesores",
    page_icon="📋",
    layout="wide"
)


# ============================================================
# 2. UBICACIÓN DEL PROYECTO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent


# ============================================================
# 3. ARCHIVOS PRINCIPALES
# ============================================================

ARCHIVO_MATRIZ = (
    BASE_DIR / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)

ARCHIVO_SEMANTICA = (
    BASE_DIR / "base_sintomas_semantica.csv"
)

ARCHIVO_EMBEDDINGS = (
    BASE_DIR / "embeddings_sintomas.npy"
)


# ============================================================
# 4. ENCABEZADO
# ============================================================

st.title("Aplicativo Asesores")

st.subheader(
    "Paquete 1 — Diagnóstico y carga de información"
)

st.write(
    "Verificación inicial de los archivos necesarios "
    "para el funcionamiento del aplicativo."
)


# ============================================================
# 5. FUNCIÓN DE VERIFICACIÓN
# ============================================================

def verificar_archivo(nombre, ruta):

    if ruta.exists():

        tamaño_mb = ruta.stat().st_size / (1024 * 1024)

        st.success(
            f"✓ {nombre} encontrado: "
            f"{ruta.name} "
            f"({tamaño_mb:.2f} MB)"
        )

        return True

    else:

        st.error(
            f"✗ No se encontró: {ruta.name}"
        )

        return False


# ============================================================
# 6. VERIFICAR ARCHIVOS
# ============================================================

st.header("1. Verificación de archivos")

matriz_ok = verificar_archivo(
    "Matriz",
    ARCHIVO_MATRIZ
)

semantica_ok = verificar_archivo(
    "Base semántica",
    ARCHIVO_SEMANTICA
)

embeddings_ok = verificar_archivo(
    "Embeddings",
    ARCHIVO_EMBEDDINGS
)


# ============================================================
# 7. DIAGNÓSTICO DEL EXCEL
# ============================================================

st.header("2. Diagnóstico de la matriz Excel")

if matriz_ok:

    try:

        libro = pd.ExcelFile(
            ARCHIVO_MATRIZ
        )

        hojas = libro.sheet_names

        st.success(
            f"Excel cargado correctamente. "
            f"Número de hojas: {len(hojas)}"
        )

        st.write("### Hojas encontradas")

        for numero, nombre_hoja in enumerate(
            hojas,
            start=1
        ):

            st.write(
                f"**{numero}. {nombre_hoja}**"
            )

            try:

                df = pd.read_excel(
                    ARCHIVO_MATRIZ,
                    sheet_name=nombre_hoja
                )

                filas = df.shape[0]
                columnas = df.shape[1]

                st.write(
                    f"Filas: **{filas:,}** | "
                    f"Columnas: **{columnas}**"
                )

                # --------------------------------------------
                # Información de columnas
                # --------------------------------------------

                informacion = []

                for columna in df.columns:

                    informacion.append({
                        "Columna": str(columna),
                        "Tipo de dato": str(
                            df[columna].dtype
                        ),
                        "No nulos": int(
                            df[columna].notna().sum()
                        ),
                        "Nulos": int(
                            df[columna].isna().sum()
                        )
                    })

                tabla_columnas = pd.DataFrame(
                    informacion
                )

                st.write(
                    "**Estructura de columnas:**"
                )

                st.dataframe(
                    tabla_columnas,
                    use_container_width=True,
                    hide_index=True
                )

                # --------------------------------------------
                # Muestra de datos
                # --------------------------------------------

                st.write(
                    "**Primeros 5 registros:**"
                )

                st.dataframe(
                    df.head(5),
                    use_container_width=True,
                    hide_index=True
                )

            except Exception as error_hoja:

                st.error(
                    f"Error leyendo la hoja "
                    f"'{nombre_hoja}': {error_hoja}"
                )

    except Exception as error_excel:

        st.error(
            f"Error cargando el archivo Excel: "
            f"{error_excel}"
        )

else:

    st.warning(
        "No se puede diagnosticar el Excel "
        "porque el archivo no fue encontrado."
    )

# ============================================================
# 7.1 CARGA DE HOJAS DE LA MATRIZ
# ============================================================

Base_Productos = None
Patologias = None
Condiciones = None
Restricciones = None
Complementarios = None
Reglas_Paquetes = None
Entrevista = None

if matriz_ok:

    try:

        Base_Productos = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Base_Productos"
        )

        Patologias = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Patologias"
        )

        Condiciones = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Condiciones"
        )

        Restricciones = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Restricciones"
        )

        Complementarios = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Complementarios"
        )

        Reglas_Paquetes = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Reglas_Paquetes"
        )

        Entrevista = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Entrevista"
        )

        st.success(
            "✓ Hojas de la matriz cargadas "
            "correctamente para el aplicativo."
        )

    except Exception as error_matriz:

        st.error(
            f"Error cargando las hojas de la matriz: "
            f"{error_matriz}"
        )
# ============================================================
# 8. DIAGNÓSTICO DE BASE SEMÁNTICA
# ============================================================

st.header("3. Diagnóstico de la base semántica")

base_semantica = None

if semantica_ok:

    try:

        base_semantica = pd.read_csv(
            ARCHIVO_SEMANTICA
        )

        filas = base_semantica.shape[0]
        columnas = base_semantica.shape[1]

        st.success(
            f"Base semántica cargada: "
            f"{filas:,} registros y "
            f"{columnas} columnas."
        )

        informacion = []

        for columna in base_semantica.columns:

            informacion.append({
                "Columna": str(columna),
                "Tipo de dato": str(
                    base_semantica[columna].dtype
                ),
                "No nulos": int(
                    base_semantica[columna].notna().sum()
                ),
                "Nulos": int(
                    base_semantica[columna].isna().sum()
                )
            })

        tabla_semantica = pd.DataFrame(
            informacion
        )

        st.write(
            "**Estructura de la base:**"
        )

        st.dataframe(
            tabla_semantica,
            use_container_width=True,
            hide_index=True
        )

        st.write(
            "**Primeros 10 registros:**"
        )

        st.dataframe(
            base_semantica.head(10),
            use_container_width=True,
            hide_index=True
        )

    except Exception as error_semantica:

        st.error(
            f"Error cargando la base semántica: "
            f"{error_semantica}"
        )

else:

    st.warning(
        "No se puede diagnosticar la base semántica "
        "porque el archivo no fue encontrado."
    )


# ============================================================
# 9. DIAGNÓSTICO DE EMBEDDINGS
# ============================================================

st.header("4. Diagnóstico de embeddings")

embeddings = None

if embeddings_ok:

    try:

        embeddings = np.load(
            ARCHIVO_EMBEDDINGS,
            allow_pickle=False
        )

        st.success(
            "Embeddings cargados correctamente."
        )

        st.write(
            f"Tipo: `{type(embeddings).__name__}`"
        )

        st.write(
            f"Tipo de dato: `{embeddings.dtype}`"
        )

        st.write(
            f"Dimensiones: `{embeddings.shape}`"
        )

        st.write(
            f"Cantidad total de elementos: "
            f"`{embeddings.size:,}`"
        )

        # --------------------------------------------
        # Comparación con base semántica
        # --------------------------------------------

        if base_semantica is not None:

            cantidad_base = len(
                base_semantica
            )

            if embeddings.ndim >= 1:

                cantidad_embeddings = (
                    embeddings.shape[0]
                )

                if (
                    cantidad_embeddings
                    == cantidad_base
                ):

                    st.success(
                        "✓ La cantidad de embeddings "
                        "coincide con la cantidad de "
                        "registros de la base semántica."
                    )

                else:

                    st.warning(
                        "⚠ La cantidad de embeddings "
                        "NO coincide con la cantidad "
                        "de registros semánticos."
                    )

                    st.write(
                        f"Registros semánticos: "
                        f"**{cantidad_base:,}**"
                    )

                    st.write(
                        f"Embeddings: "
                        f"**{cantidad_embeddings:,}**"
                    )

    except Exception as error_embeddings:

        st.error(
            f"Error cargando embeddings: "
            f"{error_embeddings}"
        )

else:

    st.warning(
        "No se puede diagnosticar embeddings "
        "porque el archivo no fue encontrado."
    )


# ============================================================
# 10. DETECCIÓN DE IMÁGENES
# ============================================================

st.header("5. Imágenes disponibles")

EXTENSIONES_IMAGEN = {
    ".png",
    ".jpg",
    ".jpeg"
}

imagenes = []

for archivo in BASE_DIR.iterdir():

    if archivo.is_file():

        if archivo.suffix.lower() in EXTENSIONES_IMAGEN:

            imagenes.append(archivo)


imagenes.sort(
    key=lambda archivo: archivo.name.lower()
)


if imagenes:

    st.success(
        f"Se encontraron "
        f"{len(imagenes)} imágenes."
    )

    informacion_imagenes = []

    for imagen in imagenes:

        tamaño_kb = (
            imagen.stat().st_size / 1024
        )

        informacion_imagenes.append({
            "Nombre": imagen.name,
            "Extensión": imagen.suffix.lower(),
            "Tamaño (KB)": round(
                tamaño_kb,
                2
            )
        })

    tabla_imagenes = pd.DataFrame(
        informacion_imagenes
    )

    st.dataframe(
        tabla_imagenes,
        use_container_width=True,
        hide_index=True
    )

else:

    st.info(
        "No se encontraron imágenes "
        "PNG, JPG o JPEG en la carpeta "
        "principal del proyecto."
    )


# ============================================================
# 11. RESUMEN
# ============================================================

st.header("6. Resumen del diagnóstico")

archivos_encontrados = sum([
    matriz_ok,
    semantica_ok,
    embeddings_ok
])

st.write(
    f"Archivos principales encontrados: "
    f"**{archivos_encontrados} de 3**"
)

st.write(
    f"Imágenes encontradas: "
    f"**{len(imagenes)}**"
)


if (
    matriz_ok
    and semantica_ok
    and embeddings_ok
):

    st.success(
        "✓ Los archivos principales están disponibles. "
        "El diagnóstico inicial terminó correctamente."
    )

else:

    st.warning(
        "⚠ Faltan archivos principales. "
        "Debemos corregirlos antes de continuar."
    )
# ============================================================
# FITOASISTE
# BLOQUE 1.2 — CONFIGURACIÓN GENERAL Y NAVEGACIÓN
# ============================================================

import streamlit as st


# ============================================================
# 1. CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="FITOASISTE",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# 2. ESTILOS BÁSICOS
# ============================================================

st.markdown(
    """
    <style>

    .titulo-principal {
        text-align: center;
        font-size: 42px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitulo-principal {
        text-align: center;
        font-size: 18px;
        margin-bottom: 35px;
    }

    .seccion-titulo {
        font-size: 25px;
        font-weight: 600;
        margin-top: 15px;
        margin-bottom: 10px;
    }

    .descripcion {
        font-size: 16px;
        margin-bottom: 15px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 3. TÍTULO DE LA APLICACIÓN
# ============================================================

st.markdown(
    '<div class="titulo-principal">FITOASISTE</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="subtitulo-principal">
    Herramienta de apoyo para tu proceso de aprendizaje y asesoría
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 4. MENÚ PRINCIPAL
# ============================================================

st.markdown(
    '<div class="seccion-titulo">Menú principal</div>',
    unsafe_allow_html=True
)

opcion_principal = st.radio(
    "Seleccione una sección:",
    [
        "CONSULTA",
        "ASESORÍA",
        "EVALUACIÓN"
    ],
    horizontal=True
)


# ============================================================
# 5. SECCIÓN CONSULTA
# ============================================================

if opcion_principal == "CONSULTA":

    st.header("CONSULTA")

    st.write(
        "Consulta información de Productos, Patologías, Complementarios "
        "y Restricciones."
    )

    opcion_consulta = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Productos",
            "Patologías",
            "Complementarios",
            "Restricciones"
        ]
    )

    if opcion_consulta == "Productos":

        st.info(
            "Módulo de consulta de Productos. "
            "Se incorporará en el siguiente bloque."
        )

    elif opcion_consulta == "Patologías":

        st.info(
            "Módulo de consulta de Patologías. "
            "Se incorporará posteriormente."
        )

    elif opcion_consulta == "Complementarios":

        st.info(
            "Módulo de consulta de Complementarios. "
            "Se incorporará posteriormente."
        )
    elif opcion_consulta == "Restricciones":

        st.info(
            "Módulo de consulta de Restricciones. "
            "Se incorporará posteriormente."
        )

# ============================================================
# 6. SECCIÓN ASESORÍA
# ============================================================

elif opcion_principal == "ASESORÍA":

    st.header("ASESORÍA")

    st.write(
        "La asesoría se desarrolla mediante un flujo "
        "encadenado de entrevista, direccionamiento y "
        "recomendación de paquetes."
    )

    st.subheader("Iniciar asesoría")

    if st.button(
        "Iniciar entrevista",
        use_container_width=True
    ):

        st.info(
            "El módulo de entrevista se incorporará "
            "posteriormente."
        )

    st.write("")

    st.caption(
        "Flujo previsto: "
        "Entrevista → Direccionamiento → "
        "Recomendación de paquetes → Cotización"
    )


# ============================================================
# 7. SECCIÓN EVALUACIÓN
# ============================================================

elif opcion_principal == "EVALUACIÓN":

    st.header("EVALUACIÓN")

    st.write(
        "Espacio para el aprendizaje, la evaluación "
        "y el seguimiento de resultados."
    )

    opcion_evaluacion = st.selectbox(
        "Seleccione una opción:",
        [
            "Seleccione una opción",
            "Autoevaluación",
            "Evaluación controlada",
            "Historial de evaluaciones"
        ]
    )

    if opcion_evaluacion == "Autoevaluación":

        st.info(
            "Módulo de autoevaluación. "
            "Se incorporará posteriormente."
        )

    elif opcion_evaluacion == "Evaluación controlada":

        st.info(
            "Módulo de evaluación controlada. "
            "Se incorporará posteriormente."
        )

    elif opcion_evaluacion == "Historial de evaluaciones":

        st.info(
            "Módulo de historial de evaluaciones. "
            "Se incorporará posteriormente."
        )


# ============================================================
# 8. PIE DE APLICACIÓN
# ============================================================

st.divider()

st.caption(
    "FITOASISTE — Herramienta de apoyo para tu proceso "
    "de aprendizaje y asesoría"
)
# ============================================================
# BLOQUE 2.1.1 — ESTRUCTURA Y MENÚ DE CONSULTA DE PRODUCTOS
# ============================================================

if opcion_consulta == "Productos":

    st.subheader("Consulta de productos")

    tipo_consulta_producto = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Ver todos los productos",
            "Buscar producto",
            "Componente → productos",
            "Categoría → productos",
            "Producto → acciones generales"
        ],
        key="menu_consulta_productos"
    )

    st.session_state["tipo_consulta_producto"] = (
        tipo_consulta_producto
    )
# ============================================================
# BLOQUE 2.1.2 — LISTADO ALFABÉTICO Y SELECCIÓN
# ============================================================

if opcion_consulta == "Productos" and st.session_state.get("tipo_consulta_producto") == "Ver todos los productos":

    st.subheader("Listado de productos")

    productos_ordenados = sorted(
        Base_Productos["Producto"]
        .dropna()
        .astype(str)
        .unique()
    )

    producto_seleccionado = st.selectbox(
        "Seleccione el producto que desea consultar:",
        productos_ordenados,
        key="producto_seleccionado_listado"
    )

    st.write(
        "Producto seleccionado:",
        producto_seleccionado
    )
# ============================================================
# BLOQUE 2.1.3 — FICHA DEL PRODUCTO SELECCIONADO
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto") == "Ver todos los productos"
    and "producto_seleccionado" in locals()
):

    producto_data = Base_Productos[
        Base_Productos["Producto"].astype(str) == str(producto_seleccionado)
    ]

    if not producto_data.empty:

        st.subheader("Ficha del producto")

        producto_info = producto_data.iloc[0]

        for columna in Base_Productos.columns:

            valor = producto_info[columna]

            if pd.notna(valor):
                st.write(
                    f"**{columna}:**",
                    valor
                )

    else:

        st.warning(
            "No se encontró información para el producto seleccionado."
        )
# ============================================================
# BLOQUE 2.1.4 — COMPONENTE → PRODUCTOS
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Componente → productos"
):

    st.subheader("Buscar productos por componente")

    componente_buscado = st.text_input(
        "Escriba el componente que desea buscar:",
        key="componente_buscado"
    )

    if componente_buscado.strip():

        texto_buscado = (
            componente_buscado
            .strip()
            .lower()
        )

        resultados_componentes = []

        for _, fila in Base_Productos.iterrows():

            componente = fila.iloc[3]

            if pd.isna(componente):
                continue

            texto_componente = str(componente).lower()

            if texto_buscado in texto_componente:

                resultados_componentes.append(
                    fila["Producto"]
                )

        productos_encontrados = sorted(
            set(resultados_componentes)
        )

        if len(productos_encontrados) == 0:

            st.warning(
                "No se encontraron productos que "
                "contengan el componente buscado."
            )

        elif len(productos_encontrados) == 1:

            producto_seleccionado_componente = (
                productos_encontrados[0]
            )

            st.success(
                f"Producto encontrado: "
                f"{producto_seleccionado_componente}"
            )

        else:

            st.write(
                f"Se encontraron "
                f"{len(productos_encontrados)} productos:"
            )

            producto_seleccionado_componente = st.selectbox(
                "Seleccione el producto que desea consultar:",
                productos_encontrados,
                key="producto_seleccionado_componente"
            )

        if (
            "producto_seleccionado_componente" in locals()
        ):

            producto_data_componente = Base_Productos[
                Base_Productos["Producto"].astype(str)
                == str(producto_seleccionado_componente)
            ]

            if not producto_data_componente.empty:

                st.subheader("Ficha del producto")

                producto_info = (
                    producto_data_componente.iloc[0]
                )

                for columna in Base_Productos.columns:

                    valor = producto_info[columna]

                    if pd.notna(valor):

                        st.write(
                            f"**{columna}:**",
                            valor
                        )
# ============================================================
# BLOQUE 2.1.5 — BÚSQUEDA DE PRODUCTO POR NOMBRE
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Buscar producto"
):

    st.subheader("Buscar producto")

    nombre_buscado = st.text_input(
        "Escriba el nombre del producto:",
        key="nombre_buscado_producto"
    )

    if nombre_buscado.strip():

        texto_buscado = (
            unidecode(nombre_buscado)
            .lower()
            .strip()
        )

        productos_disponibles = (
            Base_Productos["Producto"]
            .dropna()
            .astype(str)
            .unique()
        )

        coincidencias = []

        for producto in productos_disponibles:

            producto_normalizado = (
                unidecode(str(producto))
                .lower()
                .strip()
            )

            puntaje = fuzz.ratio(
                texto_buscado,
                producto_normalizado
            )

            if (
                texto_buscado in producto_normalizado
                or puntaje >= 60
            ):

                coincidencias.append(
                    (str(producto), puntaje)
                )

        coincidencias = sorted(
            coincidencias,
            key=lambda x: x[1],
            reverse=True
        )

        productos_encontrados = [
            producto
            for producto, puntaje in coincidencias
        ]

        if len(productos_encontrados) == 0:

            st.warning(
                "No se encontraron productos "
                "que coincidan con la búsqueda."
            )

        elif len(productos_encontrados) == 1:

            producto_seleccionado_nombre = (
                productos_encontrados[0]
            )

        else:

            st.write(
                f"Se encontraron "
                f"{len(productos_encontrados)} "
                f"posibles coincidencias:"
            )

            producto_seleccionado_nombre = st.selectbox(
                "Seleccione el producto que desea consultar:",
                productos_encontrados,
                key="producto_seleccionado_nombre"
            )

        if "producto_seleccionado_nombre" in locals():

            producto_data_nombre = Base_Productos[
                Base_Productos["Producto"].astype(str)
                == str(producto_seleccionado_nombre)
            ]

            if not producto_data_nombre.empty:

                st.subheader("Ficha del producto")

                producto_info = producto_data_nombre.iloc[0]

                for columna in Base_Productos.columns:

                    valor = producto_info[columna]

                    if pd.notna(valor):

                        st.write(
                            f"**{columna}:**",
                            valor
                        )
# ============================================================

# ============================================================
# BLOQUE — CATEGORÍA → PRODUCTOS
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Categoría → productos"
):

    st.subheader("Categoría → productos")

    # --------------------------------------------------------
    # FUNCIÓN PARA NORMALIZAR TEXTO
    # --------------------------------------------------------

    def normalizar_categoria(texto):

        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

    # --------------------------------------------------------
    # FUNCIÓN PARA SEPARAR CATEGORÍAS COMPLEMENTARIAS
    # --------------------------------------------------------

    def separar_categorias_complementarias(valor):

        if pd.isna(valor):
            return []

        texto = str(valor)

        for separador in [";", "|", "/", "\n"]:

            texto = texto.replace(
                separador,
                ","
            )

        categorias = []

        for parte in texto.split(","):

            categoria = parte.strip()

            if categoria:
                categorias.append(
                    categoria
                )

        return categorias

    # ========================================================
    # OPCIONES INTERNAS DE CONSULTA
    # ========================================================

    tipo_busqueda_categoria = st.radio(
        "¿Cómo desea realizar la consulta?",
        [
            "Seleccionar del listado de categorías",
            "Ingresar categoría o patología"
        ],
        key="tipo_busqueda_categoria"
    )

    # ========================================================
    # OPCIÓN 1
    # SELECCIONAR DEL LISTADO DE CATEGORÍAS
    # ========================================================

    if (
        tipo_busqueda_categoria
        == "Seleccionar del listado de categorías"
    ):

        st.write(
            "Seleccione una categoría principal:"
        )

        # ----------------------------------------------------
        # OBTENER CATEGORÍAS DE LA COLUMNA 2
        # ----------------------------------------------------

        categorias_principales = []

        for valor in Base_Productos.iloc[:, 1]:

            if pd.isna(valor):
                continue

            categoria = str(valor).strip()

            if categoria:
                categorias_principales.append(
                    categoria
                )

        # ----------------------------------------------------
        # ELIMINAR DUPLICADOS
        # ----------------------------------------------------

        categorias_unicas = {}

        for categoria in categorias_principales:

            clave = normalizar_categoria(
                categoria
            )

            if clave not in categorias_unicas:

                categorias_unicas[clave] = (
                    categoria
                )

        # ----------------------------------------------------
        # ORDEN ALFABÉTICO
        # ----------------------------------------------------

        categorias_finales = sorted(
            categorias_unicas.values(),
            key=lambda x: normalizar_categoria(x)
        )

        # ----------------------------------------------------
        # LISTADO DESPLEGABLE
        # ----------------------------------------------------

        categoria_seleccionada = st.selectbox(
            "Categoría principal:",
            [
                "Seleccione una categoría"
            ] + categorias_finales,
            key="categoria_principal_seleccionada"
        )

        # ----------------------------------------------------
        # SI SELECCIONÓ UNA CATEGORÍA
        # ----------------------------------------------------

        if (
            categoria_seleccionada
            != "Seleccione una categoría"
        ):

            categoria_buscada = normalizar_categoria(
                categoria_seleccionada
            )

            productos_directos = []

            # ------------------------------------------------
            # BUSCAR EN COLUMNA 2
            # ------------------------------------------------

            for _, fila in Base_Productos.iterrows():

                producto = fila.iloc[0]
                categoria_principal = fila.iloc[1]

                if pd.isna(producto):
                    continue

                if pd.isna(categoria_principal):
                    continue

                categoria_principal_normalizada = (
                    normalizar_categoria(
                        categoria_principal
                    )
                )

                if (
                    categoria_principal_normalizada
                    == categoria_buscada
                ):

                    productos_directos.append(
                        str(producto).strip()
                    )

            # ------------------------------------------------
            # ELIMINAR DUPLICADOS
            # ------------------------------------------------

            productos_directos = list(
                dict.fromkeys(
                    productos_directos
                )
            )

            productos_directos = sorted(
                productos_directos,
                key=lambda x: normalizar_categoria(x)
            )

            # ------------------------------------------------
            # MENSAJE
            # ------------------------------------------------

            st.success(
                f"Categoría seleccionada: "
                f"{categoria_seleccionada}"
            )

            st.write(
                "A continuación se presenta el listado "
                "de productos que tienen una acción directa "
                "sobre la patología:"
            )

            # ------------------------------------------------
            # LISTADO DE PRODUCTOS
            # ------------------------------------------------

            if not productos_directos:

                st.warning(
                    "No se encontraron productos "
                    "de acción directa para esta categoría."
                )

            else:

                producto_seleccionado = st.selectbox(
                    "Seleccione el producto que desea consultar:",
                    [
                        "Seleccione un producto"
                    ] + productos_directos,
                    key="producto_categoria_principal"
                )

                # ------------------------------------------------
                # FICHA DEL PRODUCTO
                # ------------------------------------------------

                if (
                    producto_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_ficha = Base_Productos[
                        Base_Productos.iloc[:, 0]
                        .astype(str)
                        .str.strip()
                        == str(
                            producto_seleccionado
                        ).strip()
                    ]

                    if not producto_ficha.empty:

                        st.divider()

                        st.subheader(
                            "Ficha completa del producto"
                        )

                        datos_producto = (
                            producto_ficha.iloc[0]
                        )

                        for columna in Base_Productos.columns:

                            valor = datos_producto[
                                columna
                            ]

                            if pd.notna(valor):

                                st.write(
                                    f"**{columna}:**",
                                    valor
                                )

    # ========================================================
    # OPCIÓN 2
    # INGRESAR CATEGORÍA O PATOLOGÍA
    # ========================================================

    else:

        st.write(
            "Registre el nombre de la patología "
            "o categoría que desea consultar:"
        )

        categoria_ingresada = st.text_input(
            "Nombre de la patología o categoría:",
            key="categoria_ingresada"
        )

        if categoria_ingresada.strip():

            texto_buscado = normalizar_categoria(
                categoria_ingresada
            )

            productos_directos = []
            productos_complementarios = []

            # ------------------------------------------------
            # RECORRER TODA LA BASE DE PRODUCTOS
            # ------------------------------------------------

            for _, fila in Base_Productos.iterrows():

                producto = fila.iloc[0]

                if pd.isna(producto):
                    continue

                producto = str(producto).strip()

                # ============================================
                # COLUMNA 2
                # CATEGORÍA PRINCIPAL
                # ============================================

                categoria_principal = fila.iloc[1]

                if not pd.isna(categoria_principal):

                    texto_principal = (
                        normalizar_categoria(
                            categoria_principal
                        )
                    )

                    coincidencia_principal = False

                    # Coincidencia directa
                    if (
                        texto_buscado
                        in texto_principal
                    ):

                        coincidencia_principal = True

                    # Coincidencia aproximada
                    else:

                        similitud = fuzz.ratio(
                            texto_buscado,
                            texto_principal
                        )

                        if similitud >= 70:

                            coincidencia_principal = True

                    if coincidencia_principal:

                        productos_directos.append(
                            producto
                        )

                # ============================================
                # COLUMNA 3
                # CATEGORÍAS COMPLEMENTARIAS
                # ============================================

                valor_complementarias = fila.iloc[2]

                categorias_complementarias = (
                    separar_categorias_complementarias(
                        valor_complementarias
                    )
                )

                for categoria_complementaria in (
                    categorias_complementarias
                ):

                    texto_complementario = (
                        normalizar_categoria(
                            categoria_complementaria
                        )
                    )

                    coincidencia_complementaria = False

                    # Coincidencia directa
                    if (
                        texto_buscado
                        in texto_complementario
                    ):

                        coincidencia_complementaria = True

                    # Coincidencia aproximada
                    else:

                        similitud = fuzz.ratio(
                            texto_buscado,
                            texto_complementario
                        )

                        if similitud >= 70:

                            coincidencia_complementaria = True

                    if coincidencia_complementaria:

                        productos_complementarios.append(
                            producto
                        )

                        break

            # ------------------------------------------------
            # ELIMINAR DUPLICADOS
            # ------------------------------------------------

            productos_directos = list(
                dict.fromkeys(
                    productos_directos
                )
            )

            productos_complementarios = list(
                dict.fromkeys(
                    productos_complementarios
                )
            )

            # ------------------------------------------------
            # SI UN PRODUCTO ES PRINCIPAL,
            # NO REPETIRLO COMO COMPLEMENTARIO
            # ------------------------------------------------

            productos_complementarios = [
                producto
                for producto in productos_complementarios
                if producto not in productos_directos
            ]

            # ------------------------------------------------
            # ORDENAR PRODUCTOS
            # ------------------------------------------------

            productos_directos = sorted(
                productos_directos,
                key=lambda x: normalizar_categoria(x)
            )

            productos_complementarios = sorted(
                productos_complementarios,
                key=lambda x: normalizar_categoria(x)
            )

            # =================================================
            # RESULTADOS
            # =================================================

            if (
                not productos_directos
                and not productos_complementarios
            ):

                st.warning(
                    "No se encontraron productos "
                    "relacionados con la categoría o "
                    "patología ingresada."
                )

            else:

                st.success(
                    f"Resultados para: "
                    f"{categoria_ingresada}"
                )

                # =============================================
                # DESPLEGABLE 1
                # PRODUCTOS CON ACCIÓN DIRECTA
                # =============================================

                st.markdown(
                    "### Productos con acción directa"
                )

                if productos_directos:

                    producto_directo_seleccionado = (
                        st.selectbox(
                            "Seleccione un producto principal:",
                            [
                                "Seleccione un producto"
                            ] + productos_directos,
                            key="producto_directo_categoria"
                        )
                    )

                else:

                    st.info(
                        "No se encontraron productos "
                        "con acción directa."
                    )

                    producto_directo_seleccionado = (
                        "Seleccione un producto"
                    )

                # =============================================
                # DESPLEGABLE 2
                # PRODUCTOS COMPLEMENTARIOS
                # =============================================

                st.markdown(
                    "### Productos complementarios"
                )

                if productos_complementarios:

                    producto_complementario_seleccionado = (
                        st.selectbox(
                            "Seleccione un producto complementario:",
                            [
                                "Seleccione un producto"
                            ] + productos_complementarios,
                            key="producto_complementario_categoria"
                        )
                    )

                else:

                    st.info(
                        "No se encontraron productos "
                        "complementarios."
                    )

                    producto_complementario_seleccionado = (
                        "Seleccione un producto"
                    )

                # =============================================
                # DETERMINAR PRODUCTO SELECCIONADO
                # =============================================

                producto_para_ficha = None

                if (
                    producto_directo_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_para_ficha = (
                        producto_directo_seleccionado
                    )

                elif (
                    producto_complementario_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_para_ficha = (
                        producto_complementario_seleccionado
                    )

                # =============================================
                # MOSTRAR FICHA
                # =============================================

                if producto_para_ficha:

                    producto_ficha = Base_Productos[
                        Base_Productos.iloc[:, 0]
                        .astype(str)
                        .str.strip()
                        == str(
                            producto_para_ficha
                        ).strip()
                    ]

                    if not producto_ficha.empty:

                        st.divider()

                        st.subheader(
                            "Ficha completa del producto"
                        )

                        datos_producto = (
                            producto_ficha.iloc[0]
                        )

                        for columna in Base_Productos.columns:

                            valor = datos_producto[
                                columna
                            ]

                            if pd.notna(valor):

                                st.write(
                                    f"**{columna}:**",
                                    valor
                                )

    # ========================================================
    # NAVEGACIÓN
    # ========================================================

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "🔎 Nueva consulta",
            key="nueva_consulta_categoria",
            use_container_width=True
        ):

            st.session_state.pop(
                "categoria_principal_seleccionada",
                None
            )

            st.session_state.pop(
                "categoria_ingresada",
                None
            )

            st.session_state.pop(
                "producto_categoria_principal",
                None
            )

            st.session_state.pop(
                "producto_directo_categoria",
                None
            )

            st.session_state.pop(
                "producto_complementario_categoria",
                None
            )

            st.rerun()

    with col2:

        if st.button(
            "← Menú Productos",
            key="volver_menu_productos_categoria",
            use_container_width=True
        ):

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()

    with col3:

        if st.button(
            "🏠 Menú principal",
            key="volver_menu_principal_categoria",
            use_container_width=True
        ):

            st.session_state[
                "opcion_consulta"
            ] = "Seleccione una opción"

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()



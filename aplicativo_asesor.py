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
        "Consulta información de Productos, Patologias, Complementarios "
        "y Restricciones."
    )

    opcion_consulta = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Productos",
            "Patologias",
            "Complementarios",
            "Restricciones"
        ]
    )

    if opcion_consulta == "Productos":

        st.info(
            "Módulo de consulta de Productos. "
            "Se incorporará en el siguiente bloque."
        )

    elif opcion_consulta == "Patologias":

        st.info(
            "Módulo de consulta de Patologias. "
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
            "Ingresar categoría o patologia"
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
                "sobre la patologia:"
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
    # INGRESAR CATEGORÍA O PATOLOGIA
    # ========================================================

    else:

        st.write(
            "Registre el nombre de la patologia "
            "o categoría que desea consultar:"
        )

        categoria_ingresada = st.text_input(
            "Nombre de la patologia o categoría:",
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
                    "patologia ingresada."
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


# ============================================================
# BLOQUE — PRODUCTO → ACCIONES GENERALES
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Producto → acciones generales"
):

    st.subheader("Producto → acciones generales")

    # ========================================================
    # FUNCIONES AUXILIARES
    # ========================================================

    def normalizar_accion_general(texto):

        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

    def separar_acciones_generales(valor):

        if pd.isna(valor):
            return []

        texto = str(valor)

        partes = texto.split(";")

        acciones = []

        for parte in partes:

            accion = parte.strip()

            if accion:

                acciones.append(
                    accion
                )

        return acciones

    # ========================================================
    # TIPO DE CONSULTA
    # ========================================================

    tipo_busqueda_accion = st.radio(
        "¿Cómo desea realizar la consulta?",
        [
            "Buscar por producto",
            "Buscar por acción general"
        ],
        key="tipo_busqueda_accion_general"
    )

    # ========================================================
    # 1. BUSCAR POR PRODUCTO
    # ========================================================

    if (
        tipo_busqueda_accion
        == "Buscar por producto"
    ):

        st.write(
            "Seleccione el producto para consultar "
            "sus acciones generales:"
        )

        # ----------------------------------------------------
        # LISTADO DE PRODUCTOS
        # ----------------------------------------------------

        productos = []

        for valor in Base_Productos.iloc[:, 0]:

            if pd.isna(valor):
                continue

            producto = str(valor).strip()

            if producto:

                productos.append(
                    producto
                )

        # ----------------------------------------------------
        # ELIMINAR DUPLICADOS
        # ----------------------------------------------------

        productos_unicos = {}

        for producto in productos:

            clave = (
                normalizar_accion_general(
                    producto
                )
            )

            if clave not in productos_unicos:

                productos_unicos[clave] = (
                    producto
                )

        productos_finales = sorted(
            productos_unicos.values(),
            key=lambda x:
                normalizar_accion_general(x)
        )

        # ----------------------------------------------------
        # SELECCIÓN DEL PRODUCTO
        # ----------------------------------------------------

        producto_seleccionado = st.selectbox(
            "Producto:",
            [
                "Seleccione un producto"
            ] + productos_finales,
            key="producto_acciones_generales"
        )

        # ----------------------------------------------------
        # MOSTRAR ACCIONES
        # ----------------------------------------------------

        if (
            producto_seleccionado
            != "Seleccione un producto"
        ):

            producto_fila = Base_Productos[
                Base_Productos.iloc[:, 0]
                .astype(str)
                .str.strip()
                == str(
                    producto_seleccionado
                ).strip()
            ]

            if not producto_fila.empty:

                datos_producto = (
                    producto_fila.iloc[0]
                )

                # COLUMNA 5
                acciones = (
                    separar_acciones_generales(
                        datos_producto.iloc[4]
                    )
                )

                st.divider()

                st.subheader(
                    "Acciones generales"
                )

                if acciones:

                    for numero, accion in enumerate(
                        acciones,
                        start=1
                    ):

                        st.write(
                            f"{numero}. {accion}"
                        )

                else:

                    st.info(
                        "Este producto no tiene "
                        "acciones generales registradas."
                    )

    # ========================================================
    # 2. BUSCAR POR ACCIÓN GENERAL
    # ========================================================

    else:

        st.write(
            "Registre la acción general que desea buscar:"
        )

        # ----------------------------------------------------
        # CAMPO DE BÚSQUEDA
        # ----------------------------------------------------

        accion_ingresada = st.text_input(
            "Escriba la acción general:",
            key="accion_general_ingresada"
        )

        # ----------------------------------------------------
        # SOLO BUSCAR SI HAY TEXTO
        # ----------------------------------------------------

        if accion_ingresada.strip():

            texto_buscado = (
                normalizar_accion_general(
                    accion_ingresada
                )
            )

            resultados = []

            # =================================================
            # RECORRER TODA LA BASE
            # =================================================

            for _, fila in Base_Productos.iterrows():

                producto = fila.iloc[0]

                if pd.isna(producto):
                    continue

                producto = str(producto).strip()

                # ------------------------------------------------
                # COLUMNA 5 — ACCIONES GENERALES
                # ------------------------------------------------

                valor_acciones = fila.iloc[4]

                acciones = (
                    separar_acciones_generales(
                        valor_acciones
                    )
                )

                # ------------------------------------------------
                # REVISAR CADA ACCIÓN
                # ------------------------------------------------

                for accion in acciones:

                    accion_normalizada = (
                        normalizar_accion_general(
                            accion
                        )
                    )

                    coincidencia = False

                    # ==========================================
                    # NIVEL 1 — FRASE COMPLETA
                    # ==========================================

                    if (
                        texto_buscado
                        in accion_normalizada
                    ):

                        coincidencia = True

                    # ==========================================
                    # NIVEL 2 — PALABRAS
                    # ==========================================

                    if not coincidencia:

                        palabras_buscadas = (
                            texto_buscado.split()
                        )

                        palabras_accion = (
                            accion_normalizada.split()
                        )

                        total_encontradas = 0

                        for palabra_buscada in (
                            palabras_buscadas
                        ):

                            mejor_puntaje = 0

                            for palabra_accion in (
                                palabras_accion
                            ):

                                puntaje = fuzz.ratio(
                                    palabra_buscada,
                                    palabra_accion
                                )

                                if (
                                    puntaje
                                    > mejor_puntaje
                                ):

                                    mejor_puntaje = (
                                        puntaje
                                    )

                            if mejor_puntaje >= 70:

                                total_encontradas += 1

                        if palabras_buscadas:

                            porcentaje = (
                                total_encontradas
                                / len(
                                    palabras_buscadas
                                )
                            ) * 100

                            if porcentaje >= 70:

                                coincidencia = True

                    # ==========================================
                    # SI ENCUENTRA LA ACCIÓN
                    # ==========================================

                    if coincidencia:

                        resultados.append(
                            {
                                "Producto": producto,
                                "Accion": accion
                            }
                        )

                        # Un producto solo aparece una vez
                        break

            # =================================================
            # ELIMINAR PRODUCTOS DUPLICADOS
            # =================================================

            resultados_unicos = {}

            for resultado in resultados:

                clave = (
                    normalizar_accion_general(
                        resultado["Producto"]
                    )
                )

                if clave not in resultados_unicos:

                    resultados_unicos[clave] = (
                        resultado
                    )

            resultados = list(
                resultados_unicos.values()
            )

            # =================================================
            # ORDENAR RESULTADOS
            # =================================================

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_accion_general(
                        x["Producto"]
                    )
            )

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not resultados:

                st.warning(
                    "No se encontraron productos "
                    "con esa acción general."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"productos relacionados."
                )

                # ------------------------------------------------
                # LISTADO DESPLEGABLE
                # ------------------------------------------------

                productos_encontrados = [
                    resultado["Producto"]
                    for resultado in resultados
                ]

                producto_seleccionado = st.selectbox(
                    "Seleccione el producto que desea consultar:",
                    [
                        "Seleccione un producto"
                    ] + productos_encontrados,
                    key="producto_resultado_accion"
                )

                # =================================================
                # MOSTRAR FICHA
                # =================================================

                if (
                    producto_seleccionado
                    != "Seleccione un producto"
                ):

                    accion_encontrada = None

                    for resultado in resultados:

                        if (
                            resultado["Producto"]
                            == producto_seleccionado
                        ):

                            accion_encontrada = (
                                resultado["Accion"]
                            )

                            break

                    if accion_encontrada:

                        st.write(
                            "**Acción general encontrada:**"
                        )

                        st.info(
                            accion_encontrada
                        )

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
    # NAVEGACIÓN
    # ========================================================

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "🔎 Nueva consulta",
            key="nueva_consulta_accion_general",
            use_container_width=True
        ):

            st.session_state.pop(
                "producto_acciones_generales",
                None
            )

            st.session_state.pop(
                "accion_general_ingresada",
                None
            )

            st.session_state.pop(
                "producto_resultado_accion",
                None
            )

            st.rerun()

    with col2:

        if st.button(
            "← Menú Productos",
            key="volver_menu_productos_accion_general",
            use_container_width=True
        ):

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()

    with col3:

        if st.button(
            "🏠 Menú principal",
            key="volver_menu_principal_accion_general",
            use_container_width=True
        ):

            st.session_state[
                "opcion_consulta"
            ] = "Seleccione una opción"

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()


# ============================================================
# BLOQUE — PATOLOGIAS
# CONSULTA GENERAL
# ============================================================

if (
    opcion_consulta == "Patologias"
):

    st.subheader("Consulta de patologias")

    # ========================================================
    # FUNCIONES AUXILIARES
    # ========================================================

    def normalizar_patologia(texto):

        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

        # ========================================================
    # MENÚ DE CONSULTA
    # ========================================================

    tipo_busqueda_patologia = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Ver todas las patologias",
            "Ingresar código o nombre de la patologia",
            "Patologia → causa y síntoma"
        ],
        key="tipo_busqueda_patologia"
    )

    # ========================================================
    # 1. VER TODAS LAS PATOLOGIAS
    # ========================================================

    if (
        tipo_busqueda_patologia
        == "Ver todas las patologias"
    ):

        st.write(
            "Seleccione la patologia que desea consultar:"
        )
        # ----------------------------------------------------
        # OBTENER NOMBRES DE patologiaS
        # COLUMNA 2
        # ----------------------------------------------------

        patologias = []

        for valor in Patologias.iloc[:, 1]:

            if pd.isna(valor):
                continue

            patologia = str(valor).strip()

            if patologia:

                patologias.append(
                    patologia
                )

        # ----------------------------------------------------
        # ELIMINAR DUPLICADOS
        # ----------------------------------------------------

        patologias_unicas = {}

        for patologia in patologias:

            clave = (
                normalizar_patologia(
                    patologia
                )
            )

            if clave not in patologias_unicas:

                patologias_unicas[clave] = (
                    patologia
                )

        # ----------------------------------------------------
        # ORDEN ALFABÉTICO
        # ----------------------------------------------------

        patologias_finales = sorted(
            patologias_unicas.values(),
            key=lambda x:
                normalizar_patologia(x)
        )

        # ----------------------------------------------------
        # LISTADO DESPLEGABLE
        # ----------------------------------------------------

        patologia_seleccionada = st.selectbox(
            "patologia:",
            [
                "Seleccione una patologia"
            ] + patologias_finales,
            key="patologia_listado_general"
        )

        # ----------------------------------------------------
        # MOSTRAR FICHA
        # ----------------------------------------------------

        if (
            patologia_seleccionada
            != "Seleccione una patologia"
        ):

            patologia_ficha = Patologias[
                Patologias.iloc[:, 1]
                .astype(str)
                .str.strip()
                == str(
                    patologia_seleccionada
                ).strip()
            ]

            if not patologia_ficha.empty:

                datos = (
                    patologia_ficha.iloc[0]
                )

                st.divider()

                st.subheader(
                    "Ficha completa de la patologia"
                )

                st.write(
                    f"**Código:** "
                    f"{datos.iloc[0]}"
                )

                st.write(
                    f"**patologia:** "
                    f"{datos.iloc[1]}"
                )

                st.write(
                    f"**Descripción breve:** "
                    f"{datos.iloc[2]}"
                )

                st.write(
                    f"**Causas frecuentes:** "
                    f"{datos.iloc[3]}"
                )

                st.write(
                    f"**Síntomas / Señales clave:** "
                    f"{datos.iloc[4]}"
                )

                st.write(
                    f"**Objetivo del paquete:** "
                    f"{datos.iloc[5]}"
                )

                st.write(
                    f"**Notas para el asesor:** "
                    f"{datos.iloc[6]}"
                )

    # ========================================================
    # 2. INGRESAR CÓDIGO O NOMBRE
    # ========================================================

    else:

        st.write(
            "Ingrese el código o nombre de la patologia:"
        )

        texto_ingresado = st.text_input(
            "Código o nombre:",
            key="buscar_patologia_general"
        )

        if texto_ingresado.strip():

            texto_buscado = (
                normalizar_patologia(
                    texto_ingresado
                )
            )

            resultados = []

            # =================================================
            # RECORRER LA BASE DE patologiaS
            # =================================================

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]

                if pd.isna(codigo):
                    codigo = ""

                if pd.isna(nombre):
                    nombre = ""

                codigo = str(
                    codigo
                ).strip()

                nombre = str(
                    nombre
                ).strip()

                codigo_normalizado = (
                    normalizar_patologia(
                        codigo
                    )
                )

                nombre_normalizado = (
                    normalizar_patologia(
                        nombre
                    )
                )

                coincidencia = False

                # =============================================
                # COINCIDENCIA DIRECTA CON CÓDIGO
                # =============================================

                if (
                    texto_buscado
                    in codigo_normalizado
                ):

                    coincidencia = True

                # =============================================
                # COINCIDENCIA DIRECTA CON NOMBRE
                # =============================================

                if (
                    texto_buscado
                    in nombre_normalizado
                ):

                    coincidencia = True

                # =============================================
                # COINCIDENCIA APROXIMADA CON NOMBRE
                # =============================================

                if not coincidencia:

                    similitud = fuzz.ratio(
                        texto_buscado,
                        nombre_normalizado
                    )

                    if similitud >= 65:

                        coincidencia = True

                # =============================================
                # COINCIDENCIA POR PALABRAS
                # =============================================

                if not coincidencia:

                    palabras_buscadas = (
                        texto_buscado.split()
                    )

                    palabras_nombre = (
                        nombre_normalizado.split()
                    )

                    palabras_encontradas = 0

                    for palabra_buscada in (
                        palabras_buscadas
                    ):

                        for palabra_nombre in (
                            palabras_nombre
                        ):

                            similitud_palabra = (
                                fuzz.ratio(
                                    palabra_buscada,
                                    palabra_nombre
                                )
                            )

                            if (
                                similitud_palabra
                                >= 70
                            ):

                                palabras_encontradas += 1
                                break

                    if palabras_buscadas:

                        porcentaje = (
                            palabras_encontradas
                            / len(
                                palabras_buscadas
                            )
                        ) * 100

                        if porcentaje >= 70:

                            coincidencia = True

                # =============================================
                # GUARDAR RESULTADO
                # =============================================

                if coincidencia:

                    resultados.append(
                        {
                            "Codigo": codigo,
                            "Patologia": nombre
                        }
                    )

            # =================================================
            # ELIMINAR DUPLICADOS
            # =================================================

            resultados_unicos = {}

            for resultado in resultados:

                clave = (
                    resultado["Codigo"]
                    + "|"
                    + normalizar_patologia(
                        resultado["Patologia"]
                    )
                )

                if clave not in resultados_unicos:

                    resultados_unicos[clave] = (
                        resultado
                    )

            resultados = list(
                resultados_unicos.values()
            )

            # =================================================
            # ORDENAR RESULTADOS
            # =================================================

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_patologia(
                        x["Patologia"]
                    )
            )

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not resultados:

                st.warning(
                    "No se encontraron patologias "
                    "relacionadas con la búsqueda."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"posibles coincidencias."
                )

                opciones = [
                    "Seleccione una patologia"
                ]

                for resultado in resultados:

                    opciones.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion = st.selectbox(
                    "Seleccione la patologia que desea consultar:",
                    opciones,
                    key="resultado_busqueda_patologia"
                )

                # =================================================
                # MOSTRAR FICHA DE LA SELECCIÓN
                # =================================================

                if (
                    seleccion
                    != "Seleccione una patologia"
                ):

                    codigo_seleccionado = (
                        seleccion.split(
                            " — "
                        )[0]
                    )

                    patologia_ficha = Patologias[
                        Patologias.iloc[:, 0]
                        .astype(str)
                        .str.strip()
                        == codigo_seleccionado.strip()
                    ]

                    if not patologia_ficha.empty:

                        datos = (
                            patologia_ficha.iloc[0]
                        )

                        st.divider()

                        st.subheader(
                            "Ficha completa de la patologia"
                        )

                        st.write(
                            f"**Código:** "
                            f"{datos.iloc[0]}"
                        )

                        st.write(
                            f"**patologia:** "
                            f"{datos.iloc[1]}"
                        )

                        st.write(
                            f"**Descripción breve:** "
                            f"{datos.iloc[2]}"
                        )

                        st.write(
                            f"**Causas frecuentes:** "
                            f"{datos.iloc[3]}"
                        )

                        st.write(
                            f"**Síntomas / Señales clave:** "
                            f"{datos.iloc[4]}"
                        )

                        st.write(
                            f"**Objetivo del paquete:** "
                            f"{datos.iloc[5]}"
                        )

                        st.write(
                            f"**Notas para el asesor:** "
                            f"{datos.iloc[6]}"
                        )
    # ============================================================
    # BLOQUE — PATOLOGIA → CAUSA Y SÍNTOMA
    # ============================================================

    if (
        tipo_busqueda_patologia
        == "Patologia → causa y síntoma"
    ):

        st.subheader(
            "Patología → causa y síntoma"
        )

        tipo_busqueda_causa_sintoma = st.selectbox(
            "¿Qué desea buscar?",
            [
                "Seleccione una opción",
                "Patología",
                "Causa",
                "Síntoma"
            ],
            key="tipo_busqueda_causa_sintoma"
        )

    # ========================================================
    # FUNCIÓN PARA COMPARAR TEXTO CON TOLERANCIA
    # ========================================================

    def coincide_texto_patologia(
        texto_buscado,
        texto_base,
        umbral=70
    ):

        buscado = normalizar_patologia(
            texto_buscado
        )

        base = normalizar_patologia(
            texto_base
        )

        if not buscado or not base:
            return False

        # ----------------------------------------
        # COINCIDENCIA DIRECTA
        # ----------------------------------------

        if buscado in base:
            return True

        # ----------------------------------------
        # COMPARACIÓN POR PALABRAS
        # ----------------------------------------

        palabras_buscadas = buscado.split()
        palabras_base = base.split()

        if not palabras_buscadas:
            return False

        palabras_encontradas = 0

        for palabra_buscada in palabras_buscadas:

            mejor_puntaje = 0

            for palabra_base in palabras_base:

                puntaje = fuzz.ratio(
                    palabra_buscada,
                    palabra_base
                )

                if puntaje > mejor_puntaje:
                    mejor_puntaje = puntaje

            if mejor_puntaje >= umbral:
                palabras_encontradas += 1

        porcentaje = (
            palabras_encontradas
            / len(palabras_buscadas)
        ) * 100

        return porcentaje >= 70

    # ========================================================
    # FUNCIÓN PARA MOSTRAR LA FICHA COMPLETA
    # ========================================================

    def mostrar_ficha_patologia(
        codigo_seleccionado
    ):

        ficha = Patologias[
            Patologias.iloc[:, 0]
            .astype(str)
            .str.strip()
            == str(
                codigo_seleccionado
            ).strip()
        ]

        if ficha.empty:

            st.warning(
                "No fue posible encontrar "
                "la ficha de la patología."
            )

            return

        datos = ficha.iloc[0]

        st.divider()

        st.subheader(
            "Ficha completa de la patología"
        )

        st.write(
            f"**Código:** {datos.iloc[0]}"
        )

        st.write(
            f"**Patología:** {datos.iloc[1]}"
        )

        st.write(
            f"**Descripción breve:** {datos.iloc[2]}"
        )

        st.write(
            f"**Causas frecuentes:** {datos.iloc[3]}"
        )

        st.write(
            f"**Síntomas / Señales clave:** {datos.iloc[4]}"
        )

        st.write(
            f"**Objetivo del paquete:** {datos.iloc[5]}"
        )

        st.write(
            f"**Notas para el asesor:** {datos.iloc[6]}"
        )

    # ========================================================
    # 1. BUSCAR POR PATOLOGÍA
    # ========================================================

    if (
        tipo_busqueda_causa_sintoma
        == "Patología"
    ):

        texto_buscado = st.text_input(
            "Ingrese el nombre de la patología:",
            key="texto_busqueda_patologia_causa"
        )

        if texto_buscado.strip():

            resultados = []

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]

                if pd.isna(codigo):
                    continue

                if pd.isna(nombre):
                    continue

                if coincide_texto_patologia(
                    texto_buscado,
                    nombre,
                    umbral=70
                ):

                    resultados.append(
                        {
                            "Codigo": str(
                                codigo
                            ).strip(),
                            "Patologia": str(
                                nombre
                            ).strip()
                        }
                    )

            # --------------------------------------------
            # ELIMINAR DUPLICADOS
            # --------------------------------------------

            resultados_unicos = {}

            for resultado in resultados:

                clave = resultado["Codigo"]

                if clave not in resultados_unicos:

                    resultados_unicos[
                        clave
                    ] = resultado

            resultados = list(
                resultados_unicos.values()
            )

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_patologia(
                        x["Patologia"]
                    )
            )

            # --------------------------------------------
            # RESULTADOS
            # --------------------------------------------

            if not resultados:

                st.warning(
                    "No se encontraron patologías "
                    "relacionadas con la búsqueda."
                )

            else:

                opciones = [
                    "Seleccione una patología"
                ]

                for resultado in resultados:

                    opciones.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion = st.selectbox(
                    "Seleccione la patología:",
                    opciones,
                    key="seleccion_patologia_causa"
                )

                if (
                    seleccion
                    != "Seleccione una patología"
                ):

                    codigo = (
                        seleccion
                        .split(" — ")[0]
                        .strip()
                    )

                    mostrar_ficha_patologia(
                        codigo
                    )

    # ========================================================
    # 2. BUSCAR POR CAUSA
    # ========================================================

    elif (
        tipo_busqueda_causa_sintoma
        == "Causa"
    ):

        texto_buscado = st.text_input(
            "Ingrese la causa que desea buscar:",
            key="texto_busqueda_causa_patologia"
        )

        if texto_buscado.strip():

            resultados = []

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]
                causas = fila.iloc[3]

                if pd.isna(codigo):
                    continue

                if pd.isna(nombre):
                    continue

                if pd.isna(causas):
                    continue

                causas = str(causas)

                # ----------------------------------------
                # LAS CAUSAS PUEDEN ESTAR SEPARADAS POR ;
                # ----------------------------------------

                elementos = causas.split(";")

                coincide = False
                causa_encontrada = ""

                for elemento in elementos:

                    elemento = elemento.strip()

                    if not elemento:
                        continue

                    if coincide_texto_patologia(
                        texto_buscado,
                        elemento,
                        umbral=70
                    ):

                        coincide = True
                        causa_encontrada = elemento
                        break

                if coincide:

                    resultados.append(
                        {
                            "Codigo": str(
                                codigo
                            ).strip(),
                            "Patologia": str(
                                nombre
                            ).strip(),
                            "Coincidencia":
                                causa_encontrada
                        }
                    )

            # --------------------------------------------
            # ELIMINAR DUPLICADOS
            # --------------------------------------------

            resultados_unicos = {}

            for resultado in resultados:

                clave = resultado["Codigo"]

                if clave not in resultados_unicos:

                    resultados_unicos[
                        clave
                    ] = resultado

            resultados = list(
                resultados_unicos.values()
            )

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_patologia(
                        x["Patologia"]
                    )
            )

            # --------------------------------------------
            # MOSTRAR RESULTADOS
            # --------------------------------------------

            if not resultados:

                st.warning(
                    "No se encontraron patologías "
                    "relacionadas con esa causa."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"patologías relacionadas."
                )

                opciones = [
                    "Seleccione una patología"
                ]

                for resultado in resultados:

                    opciones.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion = st.selectbox(
                    "Seleccione la patología:",
                    opciones,
                    key="seleccion_patologia_causa_busqueda"
                )

                if (
                    seleccion
                    != "Seleccione una patología"
                ):

                    codigo = (
                        seleccion
                        .split(" — ")[0]
                        .strip()
                    )

                    mostrar_ficha_patologia(
                        codigo
                    )

    # ========================================================
    # 3. BUSCAR POR SÍNTOMA
    # ========================================================

    elif (
        tipo_busqueda_causa_sintoma
        == "Síntoma"
    ):

        texto_buscado = st.text_input(
            "Ingrese el síntoma o señal que desea buscar:",
            key="texto_busqueda_sintoma_patologia"
        )

        if texto_buscado.strip():

            resultados = []

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]
                sintomas = fila.iloc[4]

                if pd.isna(codigo):
                    continue

                if pd.isna(nombre):
                    continue

                if pd.isna(sintomas):
                    continue

                sintomas = str(sintomas)

                # ----------------------------------------
                # LOS SÍNTOMAS PUEDEN ESTAR SEPARADOS POR ;
                # ----------------------------------------

                elementos = sintomas.split(";")

                coincide = False
                sintoma_encontrado = ""

                for elemento in elementos:

                    elemento = elemento.strip()

                    if not elemento:
                        continue

                    if coincide_texto_patologia(
                        texto_buscado,
                        elemento,
                        umbral=70
                    ):

                        coincide = True
                        sintoma_encontrado = elemento
                        break

                if coincide:

                    resultados.append(
                        {
                            "Codigo": str(
                                codigo
                            ).strip(),
                            "Patologia": str(
                                nombre
                            ).strip(),
                            "Coincidencia":
                                sintoma_encontrado
                        }
                    )

            # --------------------------------------------
            # ELIMINAR DUPLICADOS
            # --------------------------------------------

            resultados_unicos = {}

            for resultado in resultados:

                clave = resultado["Codigo"]

                if clave not in resultados_unicos:

                    resultados_unicos[
                        clave
                    ] = resultado

            resultados = list(
                resultados_unicos.values()
            )

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_patologia(
                        x["Patologia"]
                    )
            )

            # --------------------------------------------
            # MOSTRAR RESULTADOS
            # --------------------------------------------

            if not resultados:

                st.warning(
                    "No se encontraron patologías "
                    "relacionadas con ese síntoma."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"patologías relacionadas."
                )

                opciones = [
                    "Seleccione una patología"
                ]

                for resultado in resultados:

                    opciones.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion = st.selectbox(
                    "Seleccione la patología:",
                    opciones,
                    key="seleccion_patologia_sintoma_busqueda"
                )

                if (
                    seleccion
                    != "Seleccione una patología"
                ):

                    codigo = (
                        seleccion
                        .split(" — ")[0]
                        .strip()
                    )

                    mostrar_ficha_patologia(
                        codigo
                    )
    # ========================================================
    # NAVEGACIÓN
    # ========================================================

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "🔎 Nueva consulta",
            key="nueva_consulta_patologia_general",
            use_container_width=True
        ):

            st.session_state.pop(
                "patologia_listado_general",
                None
            )

            st.session_state.pop(
                "buscar_patologia_general",
                None
            )

            st.session_state.pop(
                "resultado_busqueda_patologia",
                None
            )

            st.rerun()

    with col2:

        if st.button(
            "← Menú principal",
            key="volver_menu_principal_patologia",
            use_container_width=True
        ):

            st.session_state[
                "opcion_consulta"
            ] = "Seleccione una opción"

            st.rerun()

    with col3:

        if st.button(
            "🏠 Inicio",
            key="inicio_patologia_general",
            use_container_width=True
        ):

            st.session_state[
                "opcion_consulta"
            ] = "Seleccione una opción"

            st.rerun()


# ============================================================
# BLOQUE — RESTRICCIONES
# MENÚ Y CONSULTAS
# ============================================================

if opcion_consulta == "Restricciones":

    st.subheader("Consulta de restricciones")

    tipo_consulta_restriccion = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Ver todas las restricciones",
            "Ingresar código o nombre del producto",
            "Restricción → precaución / contraindicación",
            "Producto → motivo y alternativa"
        ],
        key="menu_consulta_restricciones"
    )

    st.session_state["tipo_consulta_restriccion"] = (
        tipo_consulta_restriccion
    )

    # ========================================================
    # CONSULTA 1 — VER TODAS LAS RESTRICCIONES
    # ========================================================

    if (
        tipo_consulta_restriccion
        == "Ver todas las restricciones"
    ):

        st.write(
            "Seleccione el producto que desea consultar:"
        )

        # ====================================================
        # OBTENER PRODUCTOS ÚNICOS
        # ====================================================

        productos_unicos = {}

        for _, fila in Restricciones.iterrows():

            producto = fila.iloc[1]

            if pd.isna(producto):
                continue

            producto = str(producto).strip()

            if not producto:
                continue

            # Normalización local para evitar depender
            # de funciones del módulo de Patologías
            clave = (
                producto
                .lower()
                .strip()
            )

            if clave not in productos_unicos:

                productos_unicos[clave] = producto

        # ====================================================
        # ORDEN ALFABÉTICO
        # ====================================================

        productos_ordenados = sorted(
            productos_unicos.values(),
            key=lambda x: (
                str(x)
                .lower()
                .strip()
            )
        )

        # ====================================================
        # LISTADO DESPLEGABLE
        # ====================================================

        opciones_productos = [
            "Seleccione un producto"
        ]

        opciones_productos.extend(
            productos_ordenados
        )

        producto_seleccionado = st.selectbox(
            "Productos con restricciones:",
            opciones_productos,
            key="producto_restriccion_general"
        )

        # ====================================================
        # MOSTRAR FICHA
        # ====================================================

        if (
            producto_seleccionado
            != "Seleccione un producto"
        ):

            producto_normalizado = (
                str(producto_seleccionado)
                .lower()
                .strip()
            )

            # Buscar todas las filas correspondientes
            # al producto seleccionado
            ficha = Restricciones[
                Restricciones.iloc[:, 1]
                .astype(str)
                .str.lower()
                .str.strip()
                == producto_normalizado
            ]

            if ficha.empty:

                st.warning(
                    "No se encontraron restricciones "
                    "para este producto."
                )

            else:

                st.divider()

                st.subheader(
                    "Ficha de restricciones del producto"
                )

                st.write(
                    f"**Producto:** "
                    f"{producto_seleccionado}"
                )

                # =================================================
                # MOSTRAR TODAS LAS RESTRICCIONES AGRUPADAS
                # =================================================

                for _, datos in ficha.iterrows():

                    st.markdown("---")

                    st.write(
                        f"**Restricción ID:** "
                        f"{datos.iloc[0]}"
                    )

                    st.write(
                        f"**Tipo:** "
                        f"{datos.iloc[2]}"
                    )

                    st.write(
                        f"**Precaución / "
                        f"Contraindicación:** "
                        f"{datos.iloc[3]}"
                    )

                    st.write(
                        f"**Motivo:** "
                        f"{datos.iloc[4]}"
                    )

                    st.write(
                        f"**Alternativas seguras:** "
                        f"{datos.iloc[5]}"
                    )

                # =================================================
                # NAVEGACIÓN
                # =================================================

                st.divider()

                siguiente_accion = st.selectbox(
                    "¿Qué desea hacer ahora?",
                    [
                        "Seleccione una opción",
                        "Seleccionar otro producto de la lista",
                        "Realizar otra búsqueda",
                        "Ir al menú principal"
                    ],
                    key="navegacion_restricciones_general"
                )

                if (
                    siguiente_accion
                    == "Seleccionar otro producto de la lista"
                ):

                    st.info(
                        "Seleccione otro producto "
                        "en el listado."
                    )

                elif (
                    siguiente_accion
                    == "Realizar otra búsqueda"
                ):

                    st.info(
                        "Seleccione otra consulta "
                        "en el menú de Restricciones."
                    )

                elif (
                    siguiente_accion
                    == "Ir al menú principal"
                ):

                    st.session_state[
                        "volver_menu_principal"
                    ] = True

                    st.rerun()


    # ========================================================
    # CONSULTA 2 — INGRESAR CÓDIGO O NOMBRE DEL PRODUCTO
    # ========================================================

    if (
        tipo_consulta_restriccion
        == "Ingresar código o nombre del producto"
    ):

        st.write(
            "Ingrese el código o nombre del producto:"
        )

        texto_busqueda_restriccion = st.text_input(
            "Buscar producto:",
            key="busqueda_producto_restriccion"
        )

        if texto_busqueda_restriccion.strip():

            consulta = (
                texto_busqueda_restriccion
                .strip()
                .lower()
            )

            resultados_restriccion = []

            # =================================================
            # BUSCAR PRODUCTOS
            # =================================================

            productos_encontrados = {}

            for _, fila in Restricciones.iterrows():

                codigo = fila.iloc[0]
                producto = fila.iloc[1]

                if pd.isna(codigo) or pd.isna(producto):
                    continue

                codigo = str(codigo).strip()
                producto = str(producto).strip()

                if not codigo or not producto:
                    continue

                codigo_normalizado = (
                    codigo.lower()
                )

                producto_normalizado = (
                    producto.lower()
                )

                # ---------------------------------------------
                # COINCIDENCIA DIRECTA
                # ---------------------------------------------

                if (
                    consulta in codigo_normalizado
                    or consulta in producto_normalizado
                ):

                    clave = producto_normalizado

                    productos_encontrados[
                        clave
                    ] = producto

            # =================================================
            # BÚSQUEDA TOLERANTE A ERRORES
            # =================================================

            if not productos_encontrados:

                candidatos = {}

                for _, fila in Restricciones.iterrows():

                    codigo = fila.iloc[0]
                    producto = fila.iloc[1]

                    if pd.isna(codigo) or pd.isna(producto):
                        continue

                    codigo = str(codigo).strip()
                    producto = str(producto).strip()

                    if not codigo or not producto:
                        continue

                    # Buscar similitud contra producto
                    puntuacion_producto = fuzz.partial_ratio(
                        consulta,
                        producto.lower()
                    )

                    # Buscar similitud contra código
                    puntuacion_codigo = fuzz.partial_ratio(
                        consulta,
                        codigo.lower()
                    )

                    puntuacion = max(
                        puntuacion_producto,
                        puntuacion_codigo
                    )

                    if puntuacion >= 60:

                        clave = producto.lower()

                        if (
                            clave not in candidatos
                            or puntuacion
                            > candidatos[clave][1]
                        ):

                            candidatos[clave] = (
                                producto,
                                puntuacion
                            )

                # ---------------------------------------------
                # ORDENAR POR MAYOR SIMILITUD
                # ---------------------------------------------

                candidatos_ordenados = sorted(
                    candidatos.values(),
                    key=lambda x: x[1],
                    reverse=True
                )

                for producto, puntuacion in (
                    candidatos_ordenados[:10]
                ):

                    productos_encontrados[
                        producto.lower()
                    ] = producto

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not productos_encontrados:

                st.warning(
                    "No se encontraron productos "
                    "que coincidan con la búsqueda."
                )

            else:

                st.write(
                    "Seleccione el producto que desea consultar:"
                )

                opciones_resultados = [
                    "Seleccione un producto"
                ]

                opciones_resultados.extend(
                    sorted(
                        productos_encontrados.values(),
                        key=lambda x: x.lower()
                    )
                )

                producto_resultado = st.selectbox(
                    "Productos encontrados:",
                    opciones_resultados,
                    key="resultado_busqueda_restriccion"
                )

                # =================================================
                # MOSTRAR FICHA DEL PRODUCTO
                # =================================================

                if (
                    producto_resultado
                    != "Seleccione un producto"
                ):

                    producto_normalizado = (
                        producto_resultado
                        .lower()
                        .strip()
                    )

                    ficha = Restricciones[
                        Restricciones.iloc[:, 1]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        == producto_normalizado
                    ]

                    if ficha.empty:

                        st.warning(
                            "No se encontraron restricciones "
                            "para este producto."
                        )

                    else:

                        st.divider()

                        st.subheader(
                            "Ficha de restricciones del producto"
                        )

                        st.write(
                            f"**Producto:** "
                            f"{producto_resultado}"
                        )

                        # =====================================
                        # MOSTRAR TODAS LAS RESTRICCIONES
                        # =====================================

                        for _, datos in ficha.iterrows():

                            st.markdown("---")

                            st.write(
                                f"**Restricción ID:** "
                                f"{datos.iloc[0]}"
                            )

                            st.write(
                                f"**Tipo:** "
                                f"{datos.iloc[2]}"
                            )

                            st.write(
                                f"**Precaución / "
                                f"Contraindicación:** "
                                f"{datos.iloc[3]}"
                            )

                            st.write(
                                f"**Motivo:** "
                                f"{datos.iloc[4]}"
                            )

                            st.write(
                                f"**Alternativas seguras:** "
                                f"{datos.iloc[5]}"
                            )

                        # =====================================
                        # NAVEGACIÓN
                        # =====================================

                        st.divider()

                        siguiente_accion_2 = st.selectbox(
                            "¿Qué desea hacer ahora?",
                            [
                                "Seleccione una opción",
                                "Seleccionar otro producto",
                                "Realizar otra búsqueda",
                                "Ir al menú principal"
                            ],
                            key="navegacion_restricciones_busqueda"
                        )

                        if (
                            siguiente_accion_2
                            == "Seleccionar otro producto"
                        ):

                            st.info(
                                "Puede seleccionar otro "
                                "producto de los resultados."
                            )

                        elif (
                            siguiente_accion_2
                            == "Realizar otra búsqueda"
                        ):

                            st.info(
                                "Ingrese un nuevo código "
                                "o nombre de producto."
                            )

                        elif (
                            siguiente_accion_2
                            == "Ir al menú principal"
                        ):

                            st.session_state[
                                "volver_menu_principal"
                            ] = True

                            st.rerun()
    # ========================================================
    # CONSULTA 3 — RESTRICCIÓN → PRECAUCIÓN / CONTRAINDICACIÓN
    # ========================================================

    if (
        tipo_consulta_restriccion
        == "Restricción → precaución / contraindicación"
    ):

        st.write(
            "Ingrese la precaución o contraindicación que desea buscar:"
        )

        texto_busqueda_restriccion = st.text_input(
            "Buscar restricción:",
            key="busqueda_restriccion_precaucion"
        )

        if texto_busqueda_restriccion.strip():

            consulta = (
                texto_busqueda_restriccion
                .strip()
                .lower()
            )

            resultados_restricciones = []

            # =================================================
            # BUSCAR EN LA COLUMNA
            # PRECAUCIÓN / CONTRAINDICACIÓN
            # =================================================

            for _, fila in Restricciones.iterrows():

                codigo = fila.iloc[0]
                producto = fila.iloc[1]
                tipo = fila.iloc[2]
                restriccion = fila.iloc[3]

                if pd.isna(restriccion):
                    continue

                restriccion_texto = (
                    str(restriccion)
                    .strip()
                )

                if not restriccion_texto:
                    continue

                # ---------------------------------------------
                # COINCIDENCIA DIRECTA
                # ---------------------------------------------

                if consulta in restriccion_texto.lower():

                    resultados_restricciones.append(
                        {
                            "codigo": codigo,
                            "producto": producto,
                            "tipo": tipo,
                            "restriccion": restriccion_texto
                        }
                    )

            # =================================================
            # BÚSQUEDA TOLERANTE A ERRORES
            # =================================================

            if not resultados_restricciones:

                candidatos_restricciones = []

                for _, fila in Restricciones.iterrows():

                    codigo = fila.iloc[0]
                    producto = fila.iloc[1]
                    tipo = fila.iloc[2]
                    restriccion = fila.iloc[3]

                    if pd.isna(restriccion):
                        continue

                    restriccion_texto = (
                        str(restriccion)
                        .strip()
                    )

                    if not restriccion_texto:
                        continue

                    puntuacion = fuzz.partial_ratio(
                        consulta,
                        restriccion_texto.lower()
                    )

                    if puntuacion >= 60:

                        candidatos_restricciones.append(
                            (
                                codigo,
                                producto,
                                tipo,
                                restriccion_texto,
                                puntuacion
                            )
                        )

                candidatos_restricciones.sort(
                    key=lambda x: x[4],
                    reverse=True
                )

                for (
                    codigo,
                    producto,
                    tipo,
                    restriccion_texto,
                    puntuacion
                ) in candidatos_restricciones[:20]:

                    resultados_restricciones.append(
                        {
                            "codigo": codigo,
                            "producto": producto,
                            "tipo": tipo,
                            "restriccion": restriccion_texto
                        }
                    )

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not resultados_restricciones:

                st.warning(
                    "No se encontraron productos relacionados "
                    "con esa precaución o contraindicación."
                )

            else:

                # ---------------------------------------------
                # OBTENER PRODUCTOS ÚNICOS
                # ---------------------------------------------

                productos_encontrados = {}

                for resultado in resultados_restricciones:

                    producto = resultado["producto"]

                    if pd.isna(producto):
                        continue

                    producto = str(producto).strip()

                    if not producto:
                        continue

                    clave = producto.lower()

                    if clave not in productos_encontrados:

                        productos_encontrados[
                            clave
                        ] = producto

                productos_ordenados = sorted(
                    productos_encontrados.values(),
                    key=lambda x: x.lower()
                )

                # ---------------------------------------------
                # LISTADO DE PRODUCTOS
                # ---------------------------------------------

                st.write(
                    "Productos relacionados encontrados:"
                )

                opciones_productos = [
                    "Seleccione un producto"
                ]

                opciones_productos.extend(
                    productos_ordenados
                )

                producto_seleccionado = st.selectbox(
                    "Seleccione el producto que desea consultar:",
                    opciones_productos,
                    key="producto_resultado_restriccion"
                )

                # =================================================
                # MOSTRAR FICHA COMPLETA DEL PRODUCTO
                # =================================================

                if (
                    producto_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_normalizado = (
                        producto_seleccionado
                        .lower()
                        .strip()
                    )

                    ficha = Restricciones[
                        Restricciones.iloc[:, 1]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        == producto_normalizado
                    ]

                    if ficha.empty:

                        st.warning(
                            "No se encontraron restricciones "
                            "para este producto."
                        )

                    else:

                        st.divider()

                        st.subheader(
                            "Ficha de restricciones del producto"
                        )

                        st.write(
                            f"**Producto:** "
                            f"{producto_seleccionado}"
                        )

                        # =========================================
                        # MOSTRAR TODAS LAS RESTRICCIONES DEL PRODUCTO
                        # =========================================

                        for _, datos in ficha.iterrows():

                            st.markdown("---")

                            st.write(
                                f"**Restricción ID:** "
                                f"{datos.iloc[0]}"
                            )

                            st.write(
                                f"**Tipo:** "
                                f"{datos.iloc[2]}"
                            )

                            st.write(
                                f"**Precaución / "
                                f"Contraindicación:** "
                                f"{datos.iloc[3]}"
                            )

                            st.write(
                                f"**Motivo:** "
                                f"{datos.iloc[4]}"
                            )

                            st.write(
                                f"**Alternativas seguras:** "
                                f"{datos.iloc[5]}"
                            )

                        # =========================================
                        # NAVEGACIÓN
                        # =========================================

                        st.divider()

                        siguiente_accion_3 = st.selectbox(
                            "¿Qué desea hacer ahora?",
                            [
                                "Seleccione una opción",
                                "Seleccionar otro producto",
                                "Realizar otra búsqueda",
                                "Ir al menú principal"
                            ],
                            key="navegacion_restricciones_busqueda_3"
                        )

                        if (
                            siguiente_accion_3
                            == "Seleccionar otro producto"
                        ):

                            st.info(
                                "Puede seleccionar otro "
                                "producto de los resultados."
                            )

                        elif (
                            siguiente_accion_3
                            == "Realizar otra búsqueda"
                        ):

                            st.info(
                                "Ingrese una nueva "
                                "precaución o contraindicación."
                            )

                        elif (
                            siguiente_accion_3
                            == "Ir al menú principal"
                        ):

                            st.session_state[
                                "volver_menu_principal"
                            ] = True

                            st.rerun()
    # ========================================================


    # ========================================================
    # CONSULTA 4 — PRODUCTO → MOTIVO Y ALTERNATIVA
    # ========================================================

    if (
        tipo_consulta_restriccion
        == "Producto → motivo y alternativa"
    ):

        st.write(
            "Ingrese el nombre o código del producto:"
        )

        texto_busqueda_producto_4 = st.text_input(
            "Buscar producto:",
            key="busqueda_producto_motivo_alternativa"
        )

        if texto_busqueda_producto_4.strip():

            consulta = (
                texto_busqueda_producto_4
                .strip()
                .lower()
            )

            # =================================================
            # BUSCAR PRODUCTOS
            # =================================================

            productos_encontrados_4 = {}

            for _, fila in Restricciones.iterrows():

                codigo = fila.iloc[0]
                producto = fila.iloc[1]

                if pd.isna(codigo) or pd.isna(producto):
                    continue

                codigo = str(codigo).strip()
                producto = str(producto).strip()

                if not codigo or not producto:
                    continue

                codigo_normalizado = codigo.lower()
                producto_normalizado = producto.lower()

                # ---------------------------------------------
                # COINCIDENCIA DIRECTA
                # ---------------------------------------------

                if (
                    consulta in codigo_normalizado
                    or consulta in producto_normalizado
                ):

                    productos_encontrados_4[
                        producto_normalizado
                    ] = producto

            # =================================================
            # BÚSQUEDA TOLERANTE A ERRORES
            # =================================================

            if not productos_encontrados_4:

                candidatos_4 = {}

                for _, fila in Restricciones.iterrows():

                    codigo = fila.iloc[0]
                    producto = fila.iloc[1]

                    if pd.isna(codigo) or pd.isna(producto):
                        continue

                    codigo = str(codigo).strip()
                    producto = str(producto).strip()

                    if not codigo or not producto:
                        continue

                    puntuacion_producto = fuzz.partial_ratio(
                        consulta,
                        producto.lower()
                    )

                    puntuacion_codigo = fuzz.partial_ratio(
                        consulta,
                        codigo.lower()
                    )

                    puntuacion = max(
                        puntuacion_producto,
                        puntuacion_codigo
                    )

                    if puntuacion >= 60:

                        clave = producto.lower()

                        if (
                            clave not in candidatos_4
                            or puntuacion
                            > candidatos_4[clave][1]
                        ):

                            candidatos_4[clave] = (
                                producto,
                                puntuacion
                            )

                candidatos_ordenados_4 = sorted(
                    candidatos_4.values(),
                    key=lambda x: x[1],
                    reverse=True
                )

                for producto, puntuacion in (
                    candidatos_ordenados_4[:10]
                ):

                    productos_encontrados_4[
                        producto.lower()
                    ] = producto

            # =================================================
            # MOSTRAR PRODUCTOS ENCONTRADOS
            # =================================================

            if not productos_encontrados_4:

                st.warning(
                    "No se encontraron productos "
                    "relacionados con la búsqueda."
                )

            else:

                st.write(
                    "Seleccione el producto que desea consultar:"
                )

                opciones_productos_4 = [
                    "Seleccione un producto"
                ]

                opciones_productos_4.extend(
                    sorted(
                        productos_encontrados_4.values(),
                        key=lambda x: x.lower()
                    )
                )

                producto_seleccionado_4 = st.selectbox(
                    "Productos encontrados:",
                    opciones_productos_4,
                    key="resultado_producto_motivo_alternativa"
                )

                # =================================================
                # MOSTRAR MOTIVO Y ALTERNATIVAS
                # =================================================

                if (
                    producto_seleccionado_4
                    != "Seleccione un producto"
                ):

                    producto_normalizado_4 = (
                        producto_seleccionado_4
                        .lower()
                        .strip()
                    )

                    ficha_4 = Restricciones[
                        Restricciones.iloc[:, 1]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        == producto_normalizado_4
                    ]

                    if ficha_4.empty:

                        st.warning(
                            "No se encontraron restricciones "
                            "para este producto."
                        )

                    else:

                        st.divider()

                        st.subheader(
                            "Motivos y alternativas del producto"
                        )

                        st.write(
                            f"**Producto:** "
                            f"{producto_seleccionado_4}"
                        )

                        # =========================================
                        # MOSTRAR CADA RESTRICCIÓN
                        # =========================================

                        for _, datos in ficha_4.iterrows():

                            st.markdown("---")

                            st.write(
                                f"**Restricción ID:** "
                                f"{datos.iloc[0]}"
                            )

                            st.write(
                                f"**Tipo:** "
                                f"{datos.iloc[2]}"
                            )

                            st.write(
                                f"**Precaución / "
                                f"Contraindicación:** "
                                f"{datos.iloc[3]}"
                            )

                            st.write(
                                f"**Motivo:** "
                                f"{datos.iloc[4]}"
                            )

                            st.write(
                                f"**Alternativas seguras:** "
                                f"{datos.iloc[5]}"
                            )

                        # =========================================
                        # NAVEGACIÓN
                        # =========================================

                        st.divider()

                        siguiente_accion_4 = st.selectbox(
                            "¿Qué desea hacer ahora?",
                            [
                                "Seleccione una opción",
                                "Seleccionar otro producto",
                                "Realizar otra búsqueda",
                                "Ir al menú principal"
                            ],
                            key="navegacion_restricciones_4"
                        )

                        if (
                            siguiente_accion_4
                            == "Seleccionar otro producto"
                        ):

                            st.info(
                                "Puede seleccionar otro "
                                "producto de los resultados."
                            )

                        elif (
                            siguiente_accion_4
                            == "Realizar otra búsqueda"
                        ):

                            st.info(
                                "Ingrese un nuevo "
                                "producto o código."
                            )

                        elif (
                            siguiente_accion_4
                            == "Ir al menú principal"
                        ):

                            st.session_state[
                                "volver_menu_principal"
                            ] = True

                            st.rerun()
    # ========================================================


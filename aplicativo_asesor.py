# ============================================================
# APLICATIVO ASESORES
# PAQUETE 1 - DIAGNÓSTICO Y CARGA DE ARCHIVOS
# ============================================================

from pathlib import Path

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
# FITOASISTE
# BLOQUE 2.1 — CONSULTA DE PRODUCTOS
# ============================================================

import difflib
import re


# ============================================================
# 2.1.1 — FUNCIONES AUXILIARES
# ============================================================

def normalizar_texto(texto):
    """
    Normaliza un texto para facilitar búsquedas tolerantes
    a mayúsculas, minúsculas, tildes y espacios.
    """

    if texto is None:
        return ""

    texto = str(texto).strip().lower()

    reemplazos = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ü": "u",
        "ñ": "n"
    }

    for original, reemplazo in reemplazos.items():
        texto = texto.replace(original, reemplazo)

    texto = re.sub(r"\s+", " ", texto)

    return texto


def dividir_valores(texto):
    """
    Divide una celda que contiene varios valores separados
    por punto y coma.
    """

    if texto is None:
        return []

    if isinstance(texto, float) and texto != texto:
        return []

    texto = str(texto).strip()

    if not texto:
        return []

    valores = texto.split(";")

    return [
        valor.strip()
        for valor in valores
        if valor.strip()
    ]


def coincidencia_tolerante(texto_busqueda, texto_objetivo):
    """
    Determina si existe una coincidencia razonable entre
    la búsqueda y un texto.

    Permite:
    - coincidencia exacta;
    - coincidencia parcial;
    - errores de digitación.
    """

    busqueda = normalizar_texto(texto_busqueda)
    objetivo = normalizar_texto(texto_objetivo)

    if not busqueda or not objetivo:
        return False

    # Coincidencia directa
    if busqueda in objetivo:
        return True

    # Comparación de palabras individuales
    palabras = objetivo.split()

    for palabra in palabras:

        if len(busqueda) >= 3 and len(palabra) >= 3:

            similitud = difflib.SequenceMatcher(
                None,
                busqueda,
                palabra
            ).ratio()

            if similitud >= 0.78:
                return True

    # Comparación contra el texto completo
    similitud = difflib.SequenceMatcher(
        None,
        busqueda,
        objetivo
    ).ratio()

    return similitud >= 0.78


def buscar_productos_por_texto(df, texto):
    """
    Busca productos por nombre utilizando búsqueda
    tolerante a errores de digitación.
    """

    resultados = []

    for indice, fila in df.iterrows():

        producto = fila.get("Producto", "")

        if coincidencia_tolerante(texto, producto):

            resultados.append(indice)

    return resultados


def mostrar_selector_productos(df, indices, clave):
    """
    Si hay varios productos, permite seleccionar uno.
    """

    if not indices:
        st.warning(
            "No se encontraron productos que coincidan "
            "con la búsqueda."
        )
        return None

    if len(indices) == 1:

        return indices[0]

    opciones = []

    for indice in indices:

        producto = str(
            df.loc[indice, "Producto"]
        )

        opciones.append(producto)

    seleccion = st.selectbox(
        "Se encontraron varios productos. Seleccione el de su interés:",
        opciones,
        key=clave
    )

    for indice in indices:

        if str(df.loc[indice, "Producto"]) == seleccion:
            return indice

    return None


def mostrar_ficha_producto(df, indice):
    """
    Muestra la ficha completa del producto seleccionado.
    """

    if indice is None:
        return

    producto = df.loc[indice]

    st.subheader("Ficha del producto")

    st.markdown(
        f"### {producto.get('Producto', 'Sin información')}"
    )

    # --------------------------------------------------------
    # Categoría principal
    # --------------------------------------------------------

    st.markdown("**Categoría principal**")

    categoria = producto.get(
        "Categoría principal",
        ""
    )

    if str(categoria).strip():
        st.write(categoria)
    else:
        st.write("Sin información registrada.")

    # --------------------------------------------------------
    # Categorías complementarias
    # --------------------------------------------------------

    st.markdown("**Categorías complementarias**")

    categorias = producto.get(
        "Categorías complementarias",
        ""
    )

    valores = dividir_valores(categorias)

    if valores:

        for valor in valores:
            st.write(f"• {valor}")

    else:
        st.write("Sin información registrada.")

    # --------------------------------------------------------
    # Componentes
    # --------------------------------------------------------

    st.markdown("**Componentes**")

    componentes = producto.get(
        "Componentes",
        ""
    )

    valores = dividir_valores(componentes)

    if valores:

        for valor in valores:
            st.write(f"• {valor}")

    else:
        st.write("Sin información registrada.")

    # --------------------------------------------------------
    # Acciones generales
    # --------------------------------------------------------

    st.markdown("**Acciones generales**")

    acciones = producto.get(
        "Acciones generales",
        ""
    )

    valores = dividir_valores(acciones)

    if valores:

        for valor in valores:
            st.write(f"• {valor}")

    else:
        st.write("Sin información registrada.")

    # --------------------------------------------------------
    # Precio
    # --------------------------------------------------------

    st.markdown("**Precio público**")

    precio = producto.get(
        "Precio público",
        ""
    )

    if str(precio).strip():
        st.write(precio)
    else:
        st.write("Precio no registrado.")

    # --------------------------------------------------------
    # Imagen
    # --------------------------------------------------------

    foto = producto.get(
        "Foto",
        ""
    )

    if str(foto).strip():

        nombre_foto = str(foto).strip()

        # Buscar la imagen dentro de IMAGENESPRODUCTOS
        rutas_posibles = [
            f"IMAGENESPRODUCTOS/{nombre_foto}",
            f"IMAGENESPRODUCTOS/{nombre_foto.lower()}",
            nombre_foto
        ]

        imagen_encontrada = None

        for ruta in rutas_posibles:

            try:

                import os

                if os.path.exists(ruta):

                    imagen_encontrada = ruta
                    break

            except Exception:
                pass

        if imagen_encontrada:

            st.image(
                imagen_encontrada,
                caption=producto.get(
                    "Producto",
                    ""
                ),
                use_container_width=True
            )

        else:

            st.info(
                "La ficha tiene registrada una imagen "
                f"({nombre_foto}), pero todavía no se "
                "encuentra disponible en la ruta de la aplicación."
            )

    # --------------------------------------------------------
    # Opciones posteriores
    # --------------------------------------------------------

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        nueva_consulta = st.button(
            "Nueva consulta",
            key=f"nueva_consulta_{indice}",
            use_container_width=True
        )

    with col2:

        volver_menu = st.button(
            "Volver al menú de CONSULTA",
            key=f"volver_consulta_{indice}",
            use_container_width=True
        )

    if nueva_consulta:

        st.rerun()

    if volver_menu:

        st.rerun()


# ============================================================
# 2.1.2 — VERIFICACIÓN DE LA BASE DE PRODUCTOS
# ============================================================

st.header("Consulta de productos")

# La variable Base_Productos debe haber sido cargada
# previamente por el Bloque 1.

if "Base_Productos" not in globals():

    st.error(
        "No se encontró la base de productos cargada "
        "por el Bloque 1."
    )

else:

    df_productos = Base_Productos.copy()

    columnas_requeridas = [
        "Producto",
        "Categoría principal",
        "Categorías complementarias",
        "Componentes",
        "Acciones generales",
        "Precio público",
        "Foto"
    ]

    columnas_faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_productos.columns
    ]

    if columnas_faltantes:

        st.error(
            "La base de productos no contiene las siguientes "
            "columnas requeridas:"
        )

        for columna in columnas_faltantes:
            st.write(f"• {columna}")

    else:

        # ====================================================
        # 2.1.3 — TIPO DE CONSULTA
        # ====================================================

        tipo_consulta_producto = st.radio(
            "Seleccione el tipo de consulta:",
            [
                "Ver todos los productos",
                "Buscar producto",
                "Componente → productos",
                "Categoría → productos",
                "Producto → acciones generales"
            ],
            key="tipo_consulta_producto"
        )

        # ====================================================
        # 2.1.4 — VER TODOS LOS PRODUCTOS
        # ====================================================

        if tipo_consulta_producto == "Ver todos los productos":

            st.subheader("Listado de productos")

            df_listado = df_productos.copy()

            df_listado["_orden"] = (
                df_listado["Producto"]
                .astype(str)
                .apply(normalizar_texto)
            )

            df_listado = df_listado.sort_values(
                "_orden"
            )

            productos = df_listado[
                "Producto"
            ].astype(str).tolist()

            producto_seleccionado = st.selectbox(
                "Seleccione el producto que desea consultar:",
                productos,
                key="producto_listado"
            )

            indice_seleccionado = None

            for indice in df_listado.index:

                if (
                    str(
                        df_listado.loc[
                            indice,
                            "Producto"
                        ]
                    )
                    == producto_seleccionado
                ):

                    indice_seleccionado = indice
                    break

            mostrar_ficha_producto(
                df_productos,
                indice_seleccionado
            )

        # ====================================================
        # 2.1.5 — BUSCAR PRODUCTO
        # ====================================================

        elif tipo_consulta_producto == "Buscar producto":

            st.subheader("Buscar producto")

            texto_busqueda = st.text_input(
                "Ingrese el nombre del producto:",
                placeholder=(
                    "Ejemplo: prost... / prostenfit"
                ),
                key="busqueda_producto"
            )

            if texto_busqueda.strip():

                indices = buscar_productos_por_texto(
                    df_productos,
                    texto_busqueda
                )

                indice_seleccionado = (
                    mostrar_selector_productos(
                        df_productos,
                        indices,
                        "selector_busqueda_producto"
                    )
                )

                if indice_seleccionado is not None:

                    mostrar_ficha_producto(
                        df_productos,
                        indice_seleccionado
                    )

        # ====================================================
        # 2.1.6 — COMPONENTE → PRODUCTOS
        # ====================================================

        elif tipo_consulta_producto == "Componente → productos":

            st.subheader(
                "Buscar productos por componente"
            )

            texto_componente = st.text_input(
                "Ingrese el componente:",
                placeholder=(
                    "Ejemplo: biotina, zinc, vitamina..."
                ),
                key="busqueda_componente"
            )

            if texto_componente.strip():

                indices = []

                for indice, fila in df_productos.iterrows():

                    componentes = dividir_valores(
                        fila.get(
                            "Componentes",
                            ""
                        )
                    )

                    for componente in componentes:

                        if coincidencia_tolerante(
                            texto_componente,
                            componente
                        ):

                            indices.append(indice)
                            break

                indice_seleccionado = (
                    mostrar_selector_productos(
                        df_productos,
                        indices,
                        "selector_componente"
                    )
                )

                if indice_seleccionado is not None:

                    mostrar_ficha_producto(
                        df_productos,
                        indice_seleccionado
                    )

        # ====================================================
        # 2.1.7 — CATEGORÍA → PRODUCTOS
        # ====================================================

        elif tipo_consulta_producto == "Categoría → productos":

            st.subheader(
                "Buscar productos por categoría"
            )

            texto_categoria = st.text_input(
                "Ingrese la categoría:",
                placeholder=(
                    "Ejemplo: cabello, próstata..."
                ),
                key="busqueda_categoria"
            )

            if texto_categoria.strip():

                indices = []

                for indice, fila in df_productos.iterrows():

                    # Categoría principal
                    categoria_principal = str(
                        fila.get(
                            "Categoría principal",
                            ""
                        )
                    )

                    if coincidencia_tolerante(
                        texto_categoria,
                        categoria_principal
                    ):

                        indices.append(indice)
                        continue

                    # Categorías complementarias
                    categorias = dividir_valores(
                        fila.get(
                            "Categorías complementarias",
                            ""
                        )
                    )

                    encontrado = False

                    for categoria in categorias:

                        if coincidencia_tolerante(
                            texto_categoria,
                            categoria
                        ):

                            encontrado = True
                            break

                    if encontrado:
                        indices.append(indice)

                indice_seleccionado = (
                    mostrar_selector_productos(
                        df_productos,
                        indices,
                        "selector_categoria"
                    )
                )

                if indice_seleccionado is not None:

                    mostrar_ficha_producto(
                        df_productos,
                        indice_seleccionado
                    )

        # ====================================================
        # 2.1.8 — PRODUCTO → ACCIONES GENERALES
        # ====================================================

        elif (
            tipo_consulta_producto
            == "Producto → acciones generales"
        ):

            st.subheader(
                "Consultar acciones generales por producto"
            )

            texto_producto_accion = st.text_input(
                "Ingrese el nombre del producto:",
                placeholder=(
                    "Ejemplo: Prostenfit"
                ),
                key="busqueda_producto_accion"
            )

            if texto_producto_accion.strip():

                indices = buscar_productos_por_texto(
                    df_productos,
                    texto_producto_accion
                )

                indice_seleccionado = (
                    mostrar_selector_productos(
                        df_productos,
                        indices,
                        "selector_producto_accion"
                    )
                )

                if indice_seleccionado is not None:

                    producto = df_productos.loc[
                        indice_seleccionado
                    ]

                    st.subheader(
                        str(
                            producto.get(
                                "Producto",
                                ""
                            )
                        )
                    )

                    st.markdown(
                        "**Acciones generales**"
                    )

                    acciones = dividir_valores(
                        producto.get(
                            "Acciones generales",
                            ""
                        )
                    )

                    if acciones:

                        for accion in acciones:
                            st.write(f"• {accion}")

                    else:

                        st.info(
                            "No hay acciones generales "
                            "registradas para este producto."
                        )

                    st.divider()

                    col1, col2 = st.columns(2)

                    with col1:

                        st.button(
                            "Nueva consulta",
                            key=(
                                f"nueva_accion_"
                                f"{indice_seleccionado}"
                            ),
                            use_container_width=True
                        )

                    with col2:

                        st.button(
                            "Volver al menú de CONSULTA",
                            key=(
                                f"volver_accion_"
                                f"{indice_seleccionado}"
                            ),
                            use_container_width=True
                        )

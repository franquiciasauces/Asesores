# ============================================================
# APLICATIVO ASESORES
# PAQUETE 1 - DIAGNÓSTICO Y CARGA DE ARCHIVOS
# ============================================================

import os
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np


# ============================================================
# 1. CONFIGURACIÓN DE LA APLICACIÓN
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

ARCHIVO_MATRIZ = BASE_DIR / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"

ARCHIVO_SEMANTICA = BASE_DIR / "base_sintomas_semantica.csv"

ARCHIVO_EMBEDDINGS = BASE_DIR / "embeddings_sintomas.npy"


# ============================================================
# 4. TÍTULO
# ============================================================

st.title("Aplicativo Asesores")

st.subheader("Paquete 1 — Diagnóstico y carga de información")

st.write(
    "Esta etapa verifica los archivos disponibles en el proyecto "
    "antes de desarrollar las consultas, la asesoría y las evaluaciones."
)


# ============================================================
# 5. FUNCIÓN PARA MOSTRAR ESTADO DE ARCHIVOS
# ============================================================

def mostrar_estado_archivo(nombre, ruta):

    existe = ruta.exists()

    if existe:

        tamaño_mb = ruta.stat().st_size / (1024 * 1024)

        st.success(
            f"✓ {nombre} encontrado — "
            f"{ruta.name} — {tamaño_mb:.2f} MB"
        )

    else:

        st.error(
            f"✗ {nombre} NO encontrado — "
            f"{ruta.name}"
        )

    return existe


# ============================================================
# 6. VERIFICACIÓN DE ARCHIVOS PRINCIPALES
# ============================================================

st.header("1. Verificación de archivos")

matriz_ok = mostrar_estado_archivo(
    "Matriz de productos, patologías y paquetes",
    ARCHIVO_MATRIZ
)

semantica_ok = mostrar_estado_archivo(
    "Base semántica",
    ARCHIVO_SEMANTICA
)

embeddings_ok = mostrar_estado_archivo(
    "Embeddings de síntomas",
    ARCHIVO_EMBEDDINGS
)


# ============================================================
# 7. DIAGNÓSTICO DEL ARCHIVO EXCEL
# ============================================================

st.header("2. Diagnóstico de la matriz Excel")

if matriz_ok:

    try:

        libro = pd.ExcelFile(ARCHIVO_MATRIZ)

        st.success(
            f"Archivo Excel cargado correctamente. "
            f"Número de hojas: {len(libro.sheet_names)}"
        )

        st.write("### Hojas encontradas")

        for numero, nombre_hoja in enumerate(
            libro.sheet_names,
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

                filas, columnas = df.shape

                st.write(
                    f"Filas: **{filas:,}**  |  "
                    f"Columnas: **{columnas}**"
                )

                # ------------------------------------------------
                # Columnas y tipos de datos
                # ------------------------------------------------

                informacion_columnas = pd.DataFrame({
                    "Columna": df.columns.astype(str),
                    "Tipo de dato": [
                        str(tipo)
                        for tipo in df.dtypes
                    ],
                    "Valores no nulos": [
                        int(df[col].notna().sum())
                        for col in df.columns
                    ],
                    "Valores nulos": [
                        int(df[col].isna().sum())
                        for col in df.columns
                    ]
                })

                st.write("**Estructura de columnas:**")

                st.dataframe(
                    informacion_columnas,
                    use_container_width=True,
                    hide_index=True
                )

                # ------------------------------------------------
                # Muestra de datos
                # ------------------------------------------------

                st.write("**Muestra de registros:**")

                st.dataframe(
                    df.head(5),
                    use_container_width=True,
                    hide_index=True
                )

            except Exception as error_hoja:

                st.error(
                    f"No fue posible leer la hoja "
                    f"'{nombre_hoja}': {error_hoja}"
                )

else:

    st.warning(
        "La matriz Excel no está disponible. "
        "El diagnóstico de las hojas no puede realizarse."
    )


# ============================================================
# 8. DIAGNÓSTICO DE LA BASE SEMÁNTICA
# ============================================================

st.header("3. Diagnóstico de la base semántica")

base_semantica = None

if semantica_ok:

    try:

        base_semantica = pd.read_csv(
            ARCHIVO_SEMANTICA
        )

        filas, columnas = base_semantica.shape

        st.success(
            f"Base semántica cargada correctamente: "
            f"{filas:,} registros y {columnas} columnas."
        )

        informacion_semantica = pd.DataFrame({
            "Columna": base_semantica.columns.astype(str),
            "Tipo de dato": [
                str(tipo)
                for tipo in base_semantica.dtypes
            ],
            "Valores no nulos": [
                int(base_semantica[col].notna().sum())
                for col in base_semantica.columns
            ],
            "Valores nulos": [
                int(base_semantica[col].isna().sum())
                for col in base_semantica.columns
            ]
        })

        st.write("**Estructura de la base:**")

        st.dataframe(
            informacion_semantica,
            use_container_width=True,
            hide_index=True
        )

        st.write("**Muestra de registros:**")

        st.dataframe(
            base_semantica.head(10),
            use_container_width=True,
            hide_index=True
        )

    except Exception as error_semantica:

        st.error(
            f"No fue posible cargar la base semántica: "
            f"{error_semantica}"
        )

else:

    st.warning(
        "La base semántica no está disponible."
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
            "Archivo de embeddings cargado correctamente."
        )

        st.write(
            f"**Tipo de objeto:** `{type(embeddings).__name__}`"
        )

        st.write(
            f"**Tipo de dato:** `{embeddings.dtype}`"
        )

        st.write(
            f"**Dimensiones:** `{embeddings.shape}`"
        )

        st.write(
            f"**Número de elementos:** `{embeddings.size:,}`"
        )

        # --------------------------------------------------------
        # Comparación con la base semántica
        # --------------------------------------------------------

        if base_semantica is not None:

            numero_registros_semantica = len(
                base_semantica
            )

            if embeddings.ndim >= 1:

                numero_embeddings = embeddings.shape[0]

                if numero_embeddings == numero_registros_semantica:

                    st.success(
                        "✓ La cantidad de embeddings coincide "
                        "con la cantidad de registros de la "
                        "base semántica."
                    )

                else:

                    st.warning(
                        "⚠ La cantidad de embeddings NO coincide "
                        "con la cantidad de registros de la "
                        "base semántica."
                    )

                    st.write(
                        f"Registros base semántica: "
                        f"**{numero_registros_semantica:,}**"
                    )

                    st.write(
                        f"Embeddings: "
                        f"**{numero_embeddings:,}**"
                    )

    except Exception as error_embeddings:

        st.error(
            f"No fue posible cargar los embeddings: "
            f"{error_embeddings}"
        )

else:

    st.warning(
        "El archivo de embeddings no está disponible."
    )


# ============================================================
# 10. DETECCIÓN DE IMÁGENES
# ============================================================

st.header("5. Imágenes disponibles")

extensiones_imagen = {
    ".png",
    ".jpg",
    ".jpeg"
}

imagenes = sorted([
    archivo
    for archivo in BASE_DIR.iterdir()
    if archivo.is_file()
    and archivo.suffix.lower() in extensiones_imagen
])


if len(imagenes) > 0:

    st.success(
        f"Se encontraron {len(imagenes)} imagen(es)."
    )

    tabla_imagenes = pd.DataFrame({
        "Nombre del archivo": [
            imagen.name
            for imagen in imagenes
        ],
        "Extensión": [
            imagen.suffix.lower()
            for imagen in imagenes
        ],
        "Tamaño (KB)": [
            round(
                imagen.stat().st_size / 1024,
                2
            )
            for imagen in imagenes
        ]
    })

    st.dataframe(
        tabla_imagenes,
        use_container_width=True,
        hide_index=True
    )

else:

    st.info(
        "No se encontraron imágenes PNG, JPG o JPEG "
        "en la raíz del proyecto."
    )


# ============================================================
# 11. RESUMEN DEL DIAGNÓSTICO
# ============================================================

st.header("6. Resumen")

archivos_correctos = sum([
    matriz_ok,
    semantica_ok,
    embeddings_ok
])

st.write(
    f"Archivos principales encontrados: "
    f"**{archivos_correctos} de 3**"
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
        "✓ El paquete básico de archivos está disponible. "
        "Podemos continuar con el siguiente bloque."
    )

else:

    st.warning(
        "⚠ Hay archivos pendientes o con problemas. "
        "Debemos corregirlos antes de continuar."
    )

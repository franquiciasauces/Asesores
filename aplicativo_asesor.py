# ============================================================
# APLICATIVO ASESORES
# PAQUETE 1 - DIAGNÓSTICO Y CARGA DE ARCHIVOS
# ============================================================

from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np


# ============================================================
# 1. CONFIGURACIÓN
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

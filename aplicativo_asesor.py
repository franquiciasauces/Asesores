from pathlib import Path

code = r'''# -*- coding: utf-8 -*-
"""
BLOQUE 1 DE 8 — BASE DEL APLICATIVO ASESOR

Este bloque:
- Configura Streamlit.
- Localiza la matriz principal.
- Lee las hojas de datos necesarias.
- Localiza y valida la base semántica y los embeddings ya generados.
- No genera embeddings nuevamente.
- Deja preparada la información para los bloques siguientes.

No incluye todavía la interfaz funcional de Productos, Patologías,
Restricciones ni Asesoría.
"""

from pathlib import Path
import os

import numpy as np
import pandas as pd
import streamlit as st


# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Aplicativo Asesor",
    page_icon="🩺",
    layout="wide",
)


# ============================================================
# 2. RUTAS DEL PROYECTO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

# Se buscan primero las ubicaciones que corresponden a la
# estructura que ya veníamos utilizando.
CARPETAS_DATOS = [
    BASE_DIR / "PORTAFOLIO" / "DATOS_MATRIZ",
    BASE_DIR / "DATOS_MATRIZ",
    BASE_DIR,
]

CARPETAS_SEMANTICA = [
    BASE_DIR / "PORTAFOLIO" / "DATOS_MATRIZ" / "SEMANTICA",
    BASE_DIR / "DATOS_MATRIZ" / "SEMANTICA",
    BASE_DIR / "SEMANTICA",
    BASE_DIR,
]

NOMBRE_MATRIZ = "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
NOMBRE_BASE_SEMANTICA = "base_sintomas_semantica.csv"
NOMBRE_EMBEDDINGS = "embeddings_sintomas.npy"


def localizar_archivo(nombre, carpetas):
    """Busca un archivo dentro de las carpetas previstas."""
    for carpeta in carpetas:
        if not carpeta.exists():
            continue

        directo = carpeta / nombre
        if directo.is_file():
            return directo

        try:
            encontrados = list(carpeta.rglob(nombre))
            if encontrados:
                return encontrados[0]
        except Exception:
            pass

    return None


RUTA_MATRIZ = localizar_archivo(NOMBRE_MATRIZ, CARPETAS_DATOS)
RUTA_BASE_SEMANTICA = localizar_archivo(
    NOMBRE_BASE_SEMANTICA, CARPETAS_SEMANTICA
)
RUTA_EMBEDDINGS = localizar_archivo(
    NOMBRE_EMBEDDINGS, CARPETAS_SEMANTICA
)


# ============================================================
# 3. CARGA DE LA MATRIZ
# ============================================================

HOJAS_REQUERIDAS = {
    "productos": "Base_Productos",
    "patologias": "Patologias",
    "condiciones": "Condiciones",
    "reglas": "Reglas_Paquetes",
    "restricciones": "Restricciones",
    "complementarios": "Complementarios",
    "entrevista": "Entrevista",
}


@st.cache_data(show_spinner="Cargando matriz...")
def cargar_matriz(ruta):
    excel = pd.ExcelFile(ruta)

    faltantes = [
        hoja
        for hoja in HOJAS_REQUERIDAS.values()
        if hoja not in excel.sheet_names
    ]

    if faltantes:
        raise ValueError(
            "La matriz no contiene las hojas requeridas: "
            + ", ".join(faltantes)
        )

    datos = {}

    for clave, hoja in HOJAS_REQUERIDAS.items():
        datos[clave] = pd.read_excel(ruta, sheet_name=hoja)

    return datos


# ============================================================
# 4. CARGA DE LA BASE SEMÁNTICA Y EMBEDDINGS
# ============================================================

@st.cache_data(show_spinner="Cargando base semántica...")
def cargar_semantica(ruta_base, ruta_embeddings):
    base = pd.read_csv(ruta_base, encoding="utf-8-sig")
    embeddings = np.load(ruta_embeddings)

    if len(base) != len(embeddings):
        raise ValueError(
            "La cantidad de registros de la base semántica "
            "no coincide con la cantidad de embeddings."
        )

    return base, embeddings


# ============================================================
# 5. VALIDACIÓN INICIAL
# ============================================================

errores = []

if RUTA_MATRIZ is None:
    errores.append(
        f"No se encontró {NOMBRE_MATRIZ}."
    )

if RUTA_BASE_SEMANTICA is None:
    errores.append(
        f"No se encontró {NOMBRE_BASE_SEMANTICA}."
    )

if RUTA_EMBEDDINGS is None:
    errores.append(
        f"No se encontró {NOMBRE_EMBEDDINGS}."
    )

if errores:
    st.error("La aplicación no puede iniciar porque faltan archivos.")
    for error in errores:
        st.write(f"• {error}")

    st.info(
        "Coloque los archivos en el repositorio. "
        "El bloque busca automáticamente dentro de "
        "PORTAFOLIO/DATOS_MATRIZ, DATOS_MATRIZ y sus subcarpetas."
    )

    st.stop()


try:
    datos = cargar_matriz(RUTA_MATRIZ)

    productos = datos["productos"]
    patologias = datos["patologias"]
    condiciones = datos["condiciones"]
    reglas = datos["reglas"]
    restricciones = datos["restricciones"]
    complementarios = datos["complementarios"]
    entrevista = datos["entrevista"]

except Exception as exc:
    st.error("No fue posible cargar la matriz.")
    st.exception(exc)
    st.stop()


try:
    base_semantica, embeddings = cargar_semantica(
        RUTA_BASE_SEMANTICA,
        RUTA_EMBEDDINGS,
    )

except Exception as exc:
    st.error("No fue posible cargar la base semántica.")
    st.exception(exc)
    st.stop()


# ============================================================
# 6. ESTADO INICIAL DE LA APLICACIÓN
# ============================================================

if "patologia_seleccionada" not in st.session_state:
    st.session_state.patologia_seleccionada = None

if "producto_seleccionado" not in st.session_state:
    st.session_state.producto_seleccionado = None

if "productos_cotizacion" not in st.session_state:
    st.session_state.productos_cotizacion = []

if "historial" not in st.session_state:
    st.session_state.historial = []


# ============================================================
# 7. INTERFAZ DE VALIDACIÓN DEL BLOQUE 1
# ============================================================

st.title("🩺 Aplicativo Asesor")

st.subheader("Bloque 1 de 8 — Configuración y carga de datos")

st.success("La aplicación inició correctamente y los archivos fueron cargados.")


col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Productos", f"{len(productos):,}")

with col2:
    st.metric("Patologías", f"{len(patologias):,}")

with col3:
    st.metric("Restricciones", f"{len(restricciones):,}")


st.divider()

st.subheader("Archivos cargados")

st.write(f"**Matriz:** `{RUTA_MATRIZ.name}`")
st.write(f"**Base semántica:** `{RUTA_BASE_SEMANTICA.name}`")
st.write(f"**Embeddings:** `{RUTA_EMBEDDINGS.name}`")

st.write(f"Registros semánticos: **{len(base_semantica):,}**")
st.write(f"Embeddings cargados: **{len(embeddings):,}**")

st.divider()

st.subheader("Hojas disponibles")

for clave, hoja in HOJAS_REQUERIDAS.items():
    st.write(f"✓ {hoja}: **{len(datos[clave]):,} registros**")

st.divider()

st.info(
    "Este bloque solamente valida la base de la aplicación. "
    "El siguiente bloque será Aprendizaje → Productos."
)
'''

path = Path("/mnt/data/bloque_1_asesor.py")
path.write_text(code, encoding="utf-8")
compile(code, str(path), "exec")

print(f"Archivo creado: {path}")
print("Sintaxis Python: OK")

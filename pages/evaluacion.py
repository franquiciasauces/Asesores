# ============================================================
# FITOASISTE
# APLICATIVO INDEPENDIENTE DE EVALUACIÓN
# ============================================================

from pathlib import Path
import base64
import json
import urllib.request
import urllib.error

import streamlit as st
import pandas as pd


# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="FITOASISTE — Evaluación",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# 2. UBICACIÓN DEL PROYECTO
# ============================================================

from pathlib import Path

# evaluacion.py está dentro de /pages.
# La carpeta principal del proyecto es su carpeta padre.
BASE_DIR = Path(__file__).resolve().parent.parent


# ============================================================
# 3. ARCHIVOS DEL APLICATIVO
# ============================================================

ARCHIVO_USUARIOS = BASE_DIR / "USUARIOS.xlsx"


# ============================================================
# 4. ARCHIVO DE MATRIZ
# ============================================================

ARCHIVO_MATRIZ = (
    BASE_DIR / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)


# ============================================================
# 5. AUTENTICACIÓN COMPARTIDA CON APLICATIVO ASESOR
# ============================================================

st.session_state.setdefault(
    "usuario_autenticado",
    False
)

st.session_state.setdefault(
    "usuario_actual",
    ""
)

st.session_state.setdefault(
    "rol_usuario",
    ""
)

# ------------------------------------------------------------
# EVALUACIÓN UTILIZA LA SESIÓN DEL APLICATIVO ASESOR
# ------------------------------------------------------------

if st.session_state.get("usuario_autenticado", False):

    USUARIO_ACTUAL = (
        st.session_state.get(
            "usuario_actual",
            ""
        )
        .strip()
        .upper()
    )

    ROL_ACTUAL = (
        st.session_state.get(
            "rol_usuario",
            ""
        )
        .strip()
        .upper()
    )

else:

    USUARIO_ACTUAL = ""
    ROL_ACTUAL = ""

# ------------------------------------------------------------
# SI NO EXISTE SESIÓN, EVALUACIÓN NO PIDE OTRA CLAVE
# ------------------------------------------------------------

if not USUARIO_ACTUAL:

    st.warning(
        "Debe ingresar primero al Aplicativo Asesor "
        "con un usuario autorizado."
    )

    st.stop()

# ------------------------------------------------------------
# VALIDAR ADMINISTRADOR
# ------------------------------------------------------------

ES_ADMINISTRADOR = (
    ROL_ACTUAL == "ADMINISTRADOR"
)

# ============================================================
# 9. ENCABEZADO
# ============================================================

st.title(
    "📝 FITOASISTE — EVALUACIÓN"
)

st.write(
    f"Usuario: **{USUARIO_ACTUAL}**"
)

st.write(
    f"Rol: **{ROL_ACTUAL}**"
)


# ============================================================
# 10. VALIDACIÓN DE ADMINISTRADOR
# ============================================================

ES_ADMINISTRADOR = (
    ROL_ACTUAL == "ADMINISTRADOR"
)


if ES_ADMINISTRADOR:

    st.success(
        "Acceso de ADMINISTRADOR habilitado."
    )

else:

    st.info(
        "Acceso de usuario asesor habilitado."
    )


# ============================================================
# 11. MENÚ PRINCIPAL
# ============================================================

st.divider()

st.header(
    "Menú de Evaluación"
)


if ES_ADMINISTRADOR:

    opciones_menu = [
        "Inicio",
        "Banco General de Preguntas",
        "Banco de Preguntas Especiales",
        "Generador de Evaluaciones",
        "Evaluaciones",
        "Historial"
    ]

else:

    opciones_menu = [
        "Inicio",
        "Evaluaciones",
        "Historial"
    ]


opcion = st.selectbox(
    "Seleccione una opción",
    opciones_menu,
    key="evaluacion_menu_principal"
)


# ============================================================
# 12. INICIO
# ============================================================

if opcion == "Inicio":

    st.subheader(
        "Aplicativo independiente de EVALUACIÓN"
    )

    st.write(
        "Este aplicativo está separado de "
        "aplicativo_asesor.py."
    )

    st.write(
        "Aquí se construirá y administrará "
        "el sistema de evaluación."
    )

    if ES_ADMINISTRADOR:

        st.success(
            "El administrador tiene habilitadas "
            "las herramientas de construcción."
        )

    else:

        st.info(
            "Las herramientas de construcción "
            "están reservadas para el administrador."
        )


# ============================================================
# 13. BANCO GENERAL
# ============================================================

elif opcion == "Banco General de Preguntas":

    if not ES_ADMINISTRADOR:

        st.error(
            "Esta sección está disponible "
            "únicamente para el ADMINISTRADOR."
        )

        st.stop()

    st.header(
        "Banco General de Preguntas"
    )

    st.info(
        "Módulo preparado para construir "
        "el Banco General de Preguntas."
    )

    if ARCHIVO_MATRIZ.exists():

        st.success(
            "✓ MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx "
            "encontrado."
        )

    else:

        st.error(
            "✗ No se encontró "
            "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
        )


# ============================================================
# 14. BANCO DE PREGUNTAS ESPECIALES
# ============================================================

elif opcion == "Banco de Preguntas Especiales":

    if not ES_ADMINISTRADOR:

        st.error(
            "Esta sección está disponible "
            "únicamente para el ADMINISTRADOR."
        )

    else:

        st.header(
            "Banco de Preguntas Especiales"
        )

        st.info(
            "Módulo reservado para preguntas especiales."
        )


# ============================================================
# 15. GENERADOR DE EVALUACIONES
# ============================================================

elif opcion == "Generador de Evaluaciones":

    if not ES_ADMINISTRADOR:

        st.error(
            "Esta sección está disponible "
            "únicamente para el ADMINISTRADOR."
        )

    else:

        st.header(
            "Generador de Evaluaciones"
        )

        st.info(
            "Aquí se construirá el generador "
            "de evaluaciones."
        )


# ============================================================
# 16. EVALUACIONES
# ============================================================

elif opcion == "Evaluaciones":

    st.header(
        "Evaluaciones"
    )

    st.info(
        "Aquí aparecerán las evaluaciones "
        "disponibles para los usuarios."
    )


# ============================================================
# 17. HISTORIAL
# ============================================================

elif opcion == "Historial":

    st.header(
        "Historial de Evaluaciones"
    )

    st.info(
        "Aquí se visualizará el historial "
        "de evaluaciones."
    )


# ============================================================
# 18. INFORMACIÓN DE CONEXIÓN
# ============================================================

st.sidebar.divider()

st.sidebar.subheader(
    "Estado del aplicativo"
)

if GITHUB_TOKEN:

    st.sidebar.success(
        "✓ GITHUB_TOKEN disponible"
    )

else:

    st.sidebar.error(
        "✗ GITHUB_TOKEN no encontrado"
    )

st.sidebar.write(
    f"Repositorio: "
    f"{GITHUB_USUARIO}/{GITHUB_REPOSITORIO}"
)


# ============================================================
# 19. CERRAR SESIÓN
# ============================================================

st.sidebar.divider()

if st.sidebar.button(
    "Cerrar sesión",
    key="evaluacion_cerrar_sesion"
):

    st.session_state[
        "evaluacion_autenticado"
    ] = False

    st.session_state[
        "evaluacion_usuario"
    ] = ""

    st.session_state[
        "evaluacion_rol"
    ] = ""

    st.rerun()

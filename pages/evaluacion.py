# ============================================================
# FITOASISTE
# APLICATIVO EVALUACIÓN
# ============================================================

from pathlib import Path
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
# 2. RUTA DEL PROYECTO
# ============================================================

# Este archivo está en /pages.
# La matriz está en la carpeta principal.
BASE_DIR = Path(__file__).resolve().parent.parent


# ============================================================
# 3. ARCHIVO DE MATRIZ
# ============================================================

ARCHIVO_MATRIZ = (
    BASE_DIR /
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)


# ============================================================
# 4. ACCESO
# ============================================================

USUARIO = st.session_state.get(
    "usuario_actual",
    ""
)

ROL = st.session_state.get(
    "rol_usuario",
    ""
)


# ============================================================
# 5. ENCABEZADO
# ============================================================

st.title("📝 FITOASISTE — EVALUACIÓN")

if USUARIO:
    st.write(f"Usuario: **{USUARIO}**")

if ROL:
    st.write(f"Rol: **{ROL}**")


# ============================================================
# 6. VALIDACIÓN DE ACCESO
# ============================================================

if not USUARIO:

    st.warning(
        "No se encontró una sesión activa. "
        "Ingrese primero al Aplicativo Asesor."
    )

    st.stop()


if ROL.upper() != "ADMINISTRADOR":

    st.error(
        "Esta aplicación está disponible "
        "únicamente para el administrador."
    )

    st.stop()


st.success("Acceso de administrador habilitado.")


# ============================================================
# 7. VALIDACIÓN DE MATRIZ
# ============================================================

st.subheader("Fuente de información")

if ARCHIVO_MATRIZ.exists():

    st.success(
        "✓ MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx encontrada."
    )

else:

    st.error(
        "✗ No se encontró "
        "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
    )

    st.stop()


# ============================================================
# 8. MENÚ DE EVALUACIÓN
# ============================================================

st.divider()

st.header("Gestión de Evaluación")

opcion = st.radio(
    "Seleccione el módulo:",
    [
        "Banco General de Preguntas",
        "Banco de Preguntas Especiales"
    ],
    key="menu_evaluacion"
)


# ============================================================
# 9. BANCO GENERAL
# ============================================================

if opcion == "Banco General de Preguntas":

    st.subheader(
        "Banco General de Preguntas"
    )

    st.info(
        "Este módulo será construido por etapas."
    )

    st.write(
        "Primero validaremos la lectura de la matriz "
        "y posteriormente la generación de preguntas."
    )


# ============================================================
# 10. BANCO DE PREGUNTAS ESPECIALES
# ============================================================

if opcion == "Banco de Preguntas Especiales":

    st.subheader(
        "Banco de Preguntas Especiales"
    )

    st.info(
        "Este módulo se construirá después "
        "del Banco General."
    )

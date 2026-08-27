# ============================================================
# FITOASISTE
# EVALUACIONES ESPECIALES Y CONTROLADAS
# ============================================================

from pathlib import Path
import streamlit as st
import pandas as pd


# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="FITOASISTE - Evaluaciones Especiales",
    page_icon="📝",
    layout="wide"
)


# ============================================================
# 2. RUTAS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

# ============================================================
# LA RUTA DEL ARCHIVO DE PERMANENCIA DE 11.B
# SE DEFINIRÁ CUANDO 11.B ESTÉ FUNCIONANDO.
# ============================================================


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

st.title(
    "📝 FITOASISTE — EVALUACIONES ESPECIALES Y CONTROLADAS"
)

st.write(
    f"Administrador: **{USUARIO}**"
)


# ============================================================
# 5. CARGA DE PERMANENCIA
# ============================================================

st.markdown("### Evaluaciones especiales y controladas")

st.info(
    "Este aplicativo utilizará únicamente las preguntas "
    "disponibles provenientes del archivo de permanencia "
    "generado por 11.B."
)

# ============================================================
# AQUÍ SE INCORPORARÁ LA LECTURA DE 11.B
# CUANDO EL ARCHIVO ESTÉ GENERADO Y CONFIRMADO.
# ============================================================

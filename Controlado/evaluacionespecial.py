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

# ------------------------------------------------------------
# La ruta del archivo de permanencia generado por 11.B
# se incorporará cuando 11.B esté funcionando.
# ------------------------------------------------------------

ARCHIVO_PERMANENCIA = None


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


# ============================================================
# 4. CONTROL DE ACCESO
# ============================================================

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
# 5. ENCABEZADO
# ============================================================

st.title(
    "📝 FITOASISTE — EVALUACIONES ESPECIALES Y CONTROLADAS"
)

st.write(
    f"Administrador: **{USUARIO}**"
)


# ============================================================
# 6. CONTENIDO INICIAL
# ============================================================

st.markdown(
    "### Evaluaciones especiales y controladas"
)

st.info(
    "Este aplicativo utilizará únicamente las preguntas "
    "disponibles provenientes del archivo de permanencia "
    "generado por 11.B."
)


# ============================================================
# 7. ESTADO DEL ARCHIVO DE PERMANENCIA
# ============================================================

st.markdown("### Estado de la fuente")

if ARCHIVO_PERMANENCIA is None:

    st.warning(
        "La permanencia de 11.B todavía no está disponible. "
        "La fuente se conectará cuando el módulo 11.B "
        "esté funcionando."
    )

else:

    archivo_permanencia = Path(
        ARCHIVO_PERMANENCIA
    )

    if archivo_permanencia.exists():

        st.success(
            "🟢 Archivo de permanencia encontrado."
        )

    else:

        st.error(
            "🔴 No se encontró el archivo de permanencia."
        )


# ============================================================
# 8. ESTADO DEL APLICATIVO
# ============================================================

st.divider()

st.caption(
    "FITOASISTE — Módulo de evaluaciones especiales "
    "y controladas."
)

st.caption(
    "La construcción de este módulo se realizará "
    "a partir de la permanencia generada por 11.B."
)

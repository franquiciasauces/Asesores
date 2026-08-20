# ============================================================
# FITOASISTE
# APLICATIVO EVALUACIÓN
# ============================================================

from pathlib import Path
import urllib.request
import urllib.error
import json

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
# 2. UBICACIÓN
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent


# ============================================================
# 3. MATRIZ
# ============================================================

ARCHIVO_MATRIZ = (
    BASE_DIR /
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)


# ============================================================
# 4. REPOSITORIO
# ============================================================

GITHUB_USUARIO = "franquiciasauces"
GITHUB_REPOSITORIO = "Asesores"


# ============================================================
# 5. AUTENTICACIÓN
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


# ============================================================
# 6. ENCABEZADO
# ============================================================

st.title("📝 FITOASISTE — EVALUACIÓN")

st.write(
    f"Administrador: **{USUARIO}**"
)


# ============================================================
# 7. VALIDAR MATRIZ
# ============================================================

st.subheader("Estado del sistema")

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
# 8. CONEXIÓN CON GITHUB
# ============================================================

def comprobar_github():

    try:

        token = st.secrets["GITHUB_TOKEN"]

    except Exception:

        return False, (
            "No se encontró GITHUB_TOKEN "
            "en los Secrets de Streamlit."
        )

    url = (
        "https://api.github.com/repos/"
        f"{GITHUB_USUARIO}/{GITHUB_REPOSITORIO}"
    )

    solicitud = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "FITOASISTE"
        },
        method="GET"
    )

    try:

        with urllib.request.urlopen(
            solicitud,
            timeout=15
        ) as respuesta:

            if respuesta.status == 200:

                return True, (
                    "Conexión con GitHub establecida."
                )

            return False, (
                f"GitHub respondió con código "
                f"{respuesta.status}."
            )

    except urllib.error.HTTPError as error:

        if error.code == 401:

            return False, (
                "GitHub rechazó el token."
            )

        if error.code == 403:

            return False, (
                "El token no tiene permisos suficientes."
            )

        if error.code == 404:

            return False, (
                "No se encontró el repositorio."
            )

        return False, (
            f"Error de GitHub: {error.code}"
        )

    except Exception as error:

        return False, (
            f"No fue posible conectar con GitHub: {error}"
        )


github_ok, mensaje_github = comprobar_github()


if github_ok:

    st.success(
        "✓ " + mensaje_github
    )

    st.success(
        "✓ Almacenamiento permanente disponible."
    )

else:

    st.error(
        "✗ " + mensaje_github
    )


# ============================================================
# 9. MENÚ
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
# 10. BANCO GENERAL
# ============================================================

if opcion == "Banco General de Preguntas":

    st.subheader(
        "Banco General de Preguntas"
    )

    if github_ok:

        st.info(
            "La conexión está lista. "
            "El siguiente paso será construir "
            "el Banco General."
        )

    else:

        st.warning(
            "Primero debe estar disponible "
            "la conexión con GitHub."
        )


# ============================================================
# 11. BANCO ESPECIAL
# ============================================================

elif opcion == "Banco de Preguntas Especiales":

    st.subheader(
        "Banco de Preguntas Especiales"
    )

    st.info(
        "Este banco se construirá después "
        "del Banco General."
    )

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

BASE_DIR = Path(__file__).resolve().parent


# ============================================================
# 3. CONFIGURACIÓN GITHUB
# ============================================================

GITHUB_USUARIO = "franquiciasauces"
GITHUB_REPOSITORIO = "Asesores"


try:

    GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]

except Exception:

    GITHUB_TOKEN = ""


# ============================================================
# 4. ARCHIVO DE MATRIZ
# ============================================================

ARCHIVO_MATRIZ = (
    BASE_DIR / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)


# ============================================================
# 5. ESTADO DE SESIÓN
# ============================================================

st.session_state.setdefault(
    "evaluacion_autenticado",
    False
)

st.session_state.setdefault(
    "evaluacion_usuario",
    ""
)

st.session_state.setdefault(
    "evaluacion_rol",
    ""
)


# ============================================================
# 6. FUNCIÓN PARA CARGAR USUARIOS
# ============================================================

def cargar_usuarios():

    ruta_usuarios = (
        BASE_DIR / "USUARIOS.xlsx"
    )

    if not ruta_usuarios.exists():

        return pd.DataFrame(
            columns=[
                "Usuario_ID",
                "Nombre",
                "Documento_ID",
                "Nombre_Usuario",
                "Clave",
                "Correo",
                "Rol",
                "Estado"
            ]
        )

    usuarios = pd.read_excel(
        ruta_usuarios,
        dtype=str
    )

    usuarios = usuarios.fillna("")

    return usuarios


# ============================================================
# 7. PANTALLA DE INGRESO
# ============================================================

if not st.session_state[
    "evaluacion_autenticado"
]:

    st.title(
        "📝 FITOASISTE — EVALUACIÓN"
    )

    st.subheader(
        "Aplicativo independiente de evaluación"
    )

    st.write(
        "Ingrese sus credenciales para acceder "
        "al módulo de evaluación."
    )

    usuario_ingresado = st.text_input(
        "Nombre de usuario",
        key="evaluacion_login_usuario"
    )

    clave_ingresada = st.text_input(
        "Contraseña",
        type="password",
        key="evaluacion_login_clave"
    )

    if st.button(
        "Ingresar",
        key="evaluacion_boton_ingresar"
    ):

        usuarios = cargar_usuarios()

        usuario_normalizado = (
            usuario_ingresado
            .strip()
            .upper()
        )

        clave_normalizada = (
            clave_ingresada
            .strip()
        )

        coincidencias = usuarios[
            usuarios[
                "Nombre_Usuario"
            ]
            .astype(str)
            .str.strip()
            .str.upper()
            ==
            usuario_normalizado
        ]

        if coincidencias.empty:

            st.error(
                "Usuario o contraseña incorrectos."
            )

        else:

            usuario = coincidencias.iloc[0]

            clave_guardada = str(
                usuario["Clave"]
            ).strip()

            estado = str(
                usuario["Estado"]
            ).strip().upper()

            if estado != "ACTIVO":

                st.error(
                    "El usuario se encuentra inactivo."
                )

            elif clave_guardada != clave_normalizada:

                st.error(
                    "Usuario o contraseña incorrectos."
                )

            else:

                st.session_state[
                    "evaluacion_autenticado"
                ] = True

                st.session_state[
                    "evaluacion_usuario"
                ] = usuario_normalizado

                st.session_state[
                    "evaluacion_rol"
                ] = str(
                    usuario["Rol"]
                ).strip().upper()

                st.rerun()

    st.stop()


# ============================================================
# 8. DATOS DEL USUARIO ACTUAL
# ============================================================

USUARIO_ACTUAL = (
    st.session_state.get(
        "evaluacion_usuario",
        ""
    )
)

ROL_ACTUAL = (
    st.session_state.get(
        "evaluacion_rol",
        ""
    )
    .strip()
    .upper()
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

# APLICATIVO ASESORES
# PAQUETE 1 - DIAGNÓSTICO Y CARGA DE ARCHIVOS

from pathlib import Path
from unidecode import unidecode
from rapidfuzz import fuzz
import streamlit as st
import pandas as pd
import numpy as np

import base64
import urllib.request
import urllib.error
import json

GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]

GITHUB_USUARIO = "franquiciasauces"
GITHUB_REPOSITORIO = "Asesores"

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
# ============================================================
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
# 3.1 — USUARIOS, AUTENTICACIÓN Y PERMISOS
# ============================================================

RUTA_USUARIOS = (
    BASE_DIR / "USUARIOS.xlsx"
)

COLUMNAS_USUARIOS = [
    "Usuario_ID",
    "Nombre",
    "Documento_ID",
    "Nombre_Usuario",
    "Clave",
    "Correo",
    "Rol",
    "Estado"
]


# ============================================================
# CARGAR USUARIOS
# ============================================================

def cargar_usuarios():

    if RUTA_USUARIOS.exists():

        usuarios = pd.read_excel(
            RUTA_USUARIOS,
            dtype=str
        )

        usuarios = usuarios.fillna("")

        for columna in COLUMNAS_USUARIOS:

            if columna not in usuarios.columns:

                usuarios[columna] = ""

        usuarios = usuarios[
            COLUMNAS_USUARIOS
        ]

        return usuarios


    # ========================================================
    # CREAR ADMINISTRADOR INICIAL
    # SOLO SI EL ARCHIVO NO EXISTE
    # ========================================================

    usuarios = pd.DataFrame(
        [
            {
                "Usuario_ID": "USR0001",
                "Nombre": "Administrador",
                "Documento_ID": "",
                "Nombre_Usuario": "FRANQUICIASAUCES",
                "Clave": "8810",
                "Correo": (
                    "FRANQUICIASAUCES"
                    "@FITOMEDICS.COM"
                ),
                "Rol": "ADMINISTRADOR",
                "Estado": "ACTIVO"
            }
        ],
        columns=COLUMNAS_USUARIOS
    )

    usuarios.to_excel(
        RUTA_USUARIOS,
        index=False
    )

    return usuarios


# ============================================================
# CARGAR REGISTRO PERMANENTE
# ============================================================

Usuarios = cargar_usuarios()


# ============================================================
# GUARDAR USUARIOS
# ============================================================

def guardar_usuarios(usuarios):

    # ========================================================
    # 1. ASEGURAR COLUMNAS CORRECTAS
    # ========================================================

    usuarios = usuarios[
        COLUMNAS_USUARIOS
    ].copy()

    # ========================================================
    # 2. GUARDAR EL DATAFRAME EN USUARIOS.XLSX
    # ========================================================

    usuarios.to_excel(
        RUTA_USUARIOS,
        index=False
    )

    # ========================================================
    # 3. PREPARAR CONEXIÓN CON GITHUB
    # ========================================================

    ruta_github = "USUARIOS.xlsx"

    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_USUARIO}/"
        f"{GITHUB_REPOSITORIO}/"
        f"contents/{ruta_github}"
    )

    headers = {
        "Authorization": (
            f"Bearer {GITHUB_TOKEN}"
        ),
        "Accept": (
            "application/vnd.github+json"
        ),
        "X-GitHub-Api-Version": (
            "2022-11-28"
        )
    }

    # ========================================================
    # 4. OBTENER SHA DEL ARCHIVO ACTUAL
    # ========================================================

    solicitud = urllib.request.Request(
        url,
        headers=headers,
        method="GET"
    )

    try:

        with urllib.request.urlopen(
            solicitud
        ) as respuesta:

            informacion = json.loads(
                respuesta.read().decode(
                    "utf-8"
                )
            )

        sha_actual = informacion["sha"]

    except urllib.error.HTTPError as error:

        detalle = error.read().decode(
            "utf-8",
            errors="ignore"
        )

        st.error(
            "No se pudo consultar "
            "USUARIOS.xlsx en GitHub."
        )

        st.code(detalle)

        return

    except Exception as error:

        st.error(
            "Error al conectar con GitHub."
        )

        st.code(str(error))

        return

    # ========================================================
    # 5. LEER EL EXCEL ACTUALIZADO
    # ========================================================

    try:

        with open(
            RUTA_USUARIOS,
            "rb"
        ) as archivo:

            contenido = archivo.read()

    except Exception as error:

        st.error(
            "No se pudo leer "
            "USUARIOS.xlsx."
        )

        st.code(str(error))

        return

    # ========================================================
    # 6. CODIFICAR ARCHIVO PARA GITHUB
    # ========================================================

    contenido_base64 = (
        base64.b64encode(
            contenido
        ).decode("utf-8")
    )

    # ========================================================
    # 7. PREPARAR ACTUALIZACIÓN
    # ========================================================

    datos = {
        "message": (
            "Actualizar USUARIOS.xlsx "
            "desde FITOASISTE"
        ),
        "content": contenido_base64,
        "sha": sha_actual
    }

    datos_json = json.dumps(
        datos
    ).encode("utf-8")

    # ========================================================
    # 8. ENVIAR ARCHIVO A GITHUB
    # ========================================================

    solicitud = urllib.request.Request(
        url,
        data=datos_json,
        headers={
            **headers,
            "Content-Type": (
                "application/json"
            )
        },
        method="PUT"
    )

    try:

        with urllib.request.urlopen(
            solicitud
        ) as respuesta:

            resultado = json.loads(
                respuesta.read().decode(
                    "utf-8"
                )
            )

        if resultado.get("content"):

            st.success(
                "Usuarios guardados "
                "correctamente en GitHub."
            )

    except urllib.error.HTTPError as error:

        detalle = error.read().decode(
            "utf-8",
            errors="ignore"
        )

        st.error(
            "El usuario se guardó "
            "localmente, pero no se pudo "
            "actualizar GitHub."
        )

        st.code(detalle)

    except Exception as error:

        st.error(
            "Error al guardar el archivo "
            "en GitHub."
        )

        st.code(str(error))

# ============================================================
# GENERAR ID AUTOMÁTICO
# ============================================================

def generar_usuario_id(usuarios):

    numeros = []

    for valor in usuarios[
        "Usuario_ID"
    ]:

        texto = str(
            valor
        ).strip().upper()

        if texto.startswith("USR"):

            try:

                numero = int(
                    texto.replace(
                        "USR",
                        ""
                    )
                )

                numeros.append(
                    numero
                )

            except ValueError:

                pass

    if not numeros:

        siguiente = 1

    else:

        siguiente = (
            max(numeros) + 1
        )

    return (
        f"USR{siguiente:04d}"
    )


# ============================================================
# INICIO DE SESIÓN
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


# ============================================================
# PANTALLA DE INGRESO
# ============================================================

if not st.session_state[
    "usuario_autenticado"
]:

    st.title(
        "Ingreso a FITOASISTE"
    )

    st.write(
        "Ingrese sus credenciales "
        "para acceder al sistema."
    )

    usuario_ingresado = st.text_input(
        "Nombre de usuario:",
        key="login_usuario"
    )

    clave_ingresada = st.text_input(
        "Contraseña:",
        type="password",
        key="login_clave"
    )

    if st.button(
        "Ingresar",
        key="boton_ingresar"
    ):

        usuario_normalizado = (
            usuario_ingresado
            .strip()
            .upper()
        )

        clave_normalizada = (
            clave_ingresada
            .strip()
        )

        usuarios_actuales = (
            cargar_usuarios()
        )

        coincidencias = (
            usuarios_actuales[
                usuarios_actuales[
                    "Nombre_Usuario"
                ]
                .astype(str)
                .str.strip()
                .str.upper()
                ==
                usuario_normalizado
            ]
        )

        if coincidencias.empty:

            st.error(
                "Usuario o contraseña "
                "incorrectos."
            )

        else:

            usuario = (
                coincidencias.iloc[0]
            )

            clave_guardada = str(
                usuario["Clave"]
            ).strip()

            estado_usuario = str(
                usuario["Estado"]
            ).strip().upper()

            if estado_usuario != "ACTIVO":

                st.error(
                    "El usuario se encuentra "
                    "INACTIVO. No puede "
                    "ingresar al sistema."
                )

            elif (
                clave_guardada
                != clave_normalizada
            ):

                st.error(
                    "Usuario o contraseña "
                    "incorrectos."
                )

            else:

                st.session_state[
                    "usuario_autenticado"
                ] = True

                st.session_state[
                    "usuario_actual"
                ] = (
                    usuario_normalizado
                )

                st.session_state[
                    "rol_usuario"
                ] = str(
                    usuario["Rol"]
                ).strip().upper()

                st.success(
                    "Ingreso exitoso."
                )

                st.rerun()

    st.stop()


# ============================================================
# ROL ACTUAL
# ============================================================

ROL_ACTUAL = (
    st.session_state.get(
        "rol_usuario",
        ""
    )
    .strip()
    .upper()
)


# ============================================================
# ADMINISTRACIÓN DE USUARIOS
# ============================================================

def mostrar_administracion_usuarios():

    global Usuarios

    if ROL_ACTUAL != "ADMINISTRADOR":

        st.error(
            "No tiene permisos para "
            "acceder a esta sección."
        )

        return


    st.header(
        "Administración de usuarios"
    )


    # ========================================================
    # RECARGAR INFORMACIÓN ACTUAL
    # ========================================================

    Usuarios = cargar_usuarios()


    # ========================================================
    # REGISTRAR NUEVO USUARIO
    # ========================================================

    st.subheader(
        "Registrar nuevo usuario"
    )

    with st.form(
        "formulario_nuevo_usuario"
    ):

        nombre = st.text_input(
            "Nombre completo"
        )

        documento = st.text_input(
            "Documento de identidad"
        )

        nombre_usuario = st.text_input(
            "Nombre de usuario"
        )

        clave = st.text_input(
            "Clave asignada",
            type="password"
        )

        correo = st.text_input(
            "Correo electrónico"
        )

        rol = st.selectbox(
            "Rol",
            [
                "ASESOR",
                "ADMINISTRADOR"
            ]
        )

        registrar = st.form_submit_button(
            "Registrar usuario"
        )


    if registrar:

        nombre_limpio = (
            nombre.strip()
        )

        documento_limpio = (
            documento.strip()
        )

        usuario_limpio = (
            nombre_usuario
            .strip()
            .upper()
        )

        clave_limpia = (
            clave.strip()
        )

        correo_limpio = (
            correo.strip()
            .lower()
        )

        # ====================================================
        # VALIDAR CAMPOS
        # ====================================================

        if not nombre_limpio:

            st.error(
                "Debe ingresar el nombre."
            )

        elif not documento_limpio:

            st.error(
                "Debe ingresar el documento."
            )

        elif not usuario_limpio:

            st.error(
                "Debe ingresar el nombre "
                "de usuario."
            )

        elif not clave_limpia:

            st.error(
                "Debe asignar una clave."
            )

        elif not correo_limpio:

            st.error(
                "Debe registrar un correo."
            )

        else:

            usuarios_actuales = (
                cargar_usuarios()
            )

            # ================================================
            # VALIDAR USUARIO REPETIDO
            # ================================================

            usuario_repetido = (
                usuarios_actuales[
                    usuarios_actuales[
                        "Nombre_Usuario"
                    ]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    ==
                    usuario_limpio
                ]
            )

            if not usuario_repetido.empty:

                st.error(
                    "Ese nombre de usuario "
                    "ya existe. "
                    "Debe utilizar otro."
                )

            else:

                # ============================================
                # VALIDAR DOCUMENTO REPETIDO
                # ============================================

                documento_repetido = (
                    usuarios_actuales[
                        usuarios_actuales[
                            "Documento_ID"
                        ]
                        .astype(str)
                        .str.strip()
                        ==
                        documento_limpio
                    ]
                )

                if not documento_repetido.empty:

                    st.error(
                        "Ese documento de "
                        "identidad ya está "
                        "registrado."
                    )

                else:

                    nuevo_id = (
                        generar_usuario_id(
                            usuarios_actuales
                        )
                    )

                    nuevo_usuario = {
                        "Usuario_ID": nuevo_id,
                        "Nombre": nombre_limpio,
                        "Documento_ID": (
                            documento_limpio
                        ),
                        "Nombre_Usuario": (
                            usuario_limpio
                        ),
                        "Clave": clave_limpia,
                        "Correo": correo_limpio,
                        "Rol": rol,
                        "Estado": "ACTIVO"
                    }

                    usuarios_actuales = pd.concat(
                        [
                            usuarios_actuales,
                            pd.DataFrame(
                                [nuevo_usuario]
                            )
                        ],
                        ignore_index=True
                    )

                    guardar_usuarios(
                        usuarios_actuales
                    )

                    Usuarios = (
                        usuarios_actuales
                    )

                    st.success(
                        f"Usuario {nuevo_id} "
                        f"registrado correctamente."
                    )

                    st.rerun()


    st.divider()


    # ========================================================
    # USUARIOS REGISTRADOS
    # ========================================================

    st.subheader(
        "Usuarios registrados"
    )

    if Usuarios.empty:

        st.info(
            "No existen usuarios registrados."
        )

        return


    for indice, fila in Usuarios.iterrows():

        usuario_id = str(
            fila["Usuario_ID"]
        )

        nombre_usuario = str(
            fila["Nombre_Usuario"]
        )

        nombre = str(
            fila["Nombre"]
        )

        estado = str(
            fila["Estado"]
        ).upper()

        rol_usuario = str(
            fila["Rol"]
        ).upper()


        with st.expander(
            f"{usuario_id} — "
            f"{nombre_usuario} — "
            f"{estado}"
        ):

            st.write(
                f"**Nombre:** {nombre}"
            )

            st.write(
                f"**Documento:** "
                f"{fila['Documento_ID']}"
            )

            st.write(
                f"**Usuario:** "
                f"{nombre_usuario}"
            )

            st.write(
                f"**Correo:** "
                f"{fila['Correo']}"
            )

            st.write(
                f"**Rol:** "
                f"{rol_usuario}"
            )

            st.write(
                f"**Estado:** "
                f"{estado}"
            )


            # =================================================
            # ACTIVAR / INACTIVAR
            # =================================================

            if usuario_id != "USR0001":

                nuevo_estado = (
                    "INACTIVO"
                    if estado == "ACTIVO"
                    else "ACTIVO"
                )

                texto_boton = (
                    "Inactivar usuario"
                    if estado == "ACTIVO"
                    else "Activar usuario"
                )

                if st.button(
                    texto_boton,
                    key=(
                        f"cambiar_estado_"
                        f"{usuario_id}"
                    )
                ):

                    Usuarios.loc[
                        indice,
                        "Estado"
                    ] = nuevo_estado

                    guardar_usuarios(
                        Usuarios
                    )

                    st.success(
                        f"Usuario {usuario_id} "
                        f"ahora está "
                        f"{nuevo_estado}."
                    )

                    st.rerun()
# ============================================================
# 3.2 — ADMINISTRACIÓN DE USUARIOS
# ============================================================

if ROL_ACTUAL == "ADMINISTRADOR":

    st.sidebar.divider()

    mostrar_administracion = st.sidebar.checkbox(
        "Administración de usuarios",
        key="mostrar_administracion_usuarios"
    )

    if mostrar_administracion:

        mostrar_administracion_usuarios()

        st.stop()
        # ============================================================

# ============================================================
# CONTROL DE VISIBILIDAD DEL DIAGNÓSTICO
# ============================================================

MOSTRAR_DIAGNOSTICO = (
    ROL_ACTUAL == "ADMINISTRADOR"
)
# ============================================================
# 4. ENCABEZADO
# ============================================================

if MOSTRAR_DIAGNOSTICO:

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

if ROL_ACTUAL == "ADMINISTRADOR":

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

else:

    matriz_ok = ARCHIVO_MATRIZ.exists()
    semantica_ok = ARCHIVO_SEMANTICA.exists()
    embeddings_ok = ARCHIVO_EMBEDDINGS.exists()

# ============================================================
# 7. DIAGNÓSTICO DEL EXCEL
# ============================================================

if ROL_ACTUAL == "ADMINISTRADOR":

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
# 7.1 CARGA DE HOJAS DE LA MATRIZ
# ============================================================

Base_Productos = None
Patologias = None
Condiciones = None
Restricciones = None
Complementarios = None
Reglas_Paquetes = None
Entrevista = None

if matriz_ok:

    try:

        Base_Productos = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Base_Productos"
        )

        Patologias = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Patologias"
        )

        Condiciones = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Condiciones"
        )

        Restricciones = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Restricciones"
        )

        Complementarios = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Complementarios"
        )

        Reglas_Paquetes = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Reglas_Paquetes"
        )

        Entrevista = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name="Entrevista"
        )

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.success(
                "✓ Hojas de la matriz cargadas "
                "correctamente para el aplicativo."
            )

    except Exception as error_matriz:

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.error(
                f"Error cargando las hojas de la matriz: "
                f"{error_matriz}"
            )
# ============================================================
# ============================================================
# 8. DIAGNÓSTICO DE BASE SEMÁNTICA
#    Y CARGA DEL MODELO BIOMÉDICO
# ============================================================

base_semantica = None


# ============================================================
# 8.1 CARGAR BASE SEMÁNTICA
# ============================================================

if semantica_ok:

    try:

        base_semantica = pd.read_csv(
            ARCHIVO_SEMANTICA
        )

        filas = base_semantica.shape[0]
        columnas = base_semantica.shape[1]

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.header(
                "3. Diagnóstico de la base semántica"
            )

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

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.error(
                f"Error cargando la base semántica: "
                f"{error_semantica}"
            )

else:

    if ROL_ACTUAL == "ADMINISTRADOR":

        st.warning(
            "No se puede diagnosticar la base semántica "
            "porque el archivo no fue encontrado."
        )

# ============================================================
# 8.2 CARGAR EMBEDDINGS PRECALCULADOS
# ============================================================

embeddings_sintomas = None


if embeddings_ok:

    try:

        embeddings_sintomas = np.load(
            ARCHIVO_EMBEDDINGS,
            allow_pickle=False
        )

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.success(
                "Embeddings de síntomas cargados "
                "correctamente."
            )

            st.write(
                f"Cantidad de embeddings: "
                f"{len(embeddings_sintomas):,}"
            )

            st.write(
                f"Dimensiones de los embeddings: "
                f"{embeddings_sintomas.shape}"
            )


        # ----------------------------------------------------
        # VALIDAR CORRESPONDENCIA
        # ----------------------------------------------------

        if (
            base_semantica is not None
            and
            len(embeddings_sintomas)
            !=
            len(base_semantica)
        ):

            if ROL_ACTUAL == "ADMINISTRADOR":

                st.error(
                    "ERROR: la cantidad de embeddings "
                    "no coincide con la cantidad de registros "
                    "de la base semántica."
                )

            embeddings_sintomas = None

        elif base_semantica is not None:

            if ROL_ACTUAL == "ADMINISTRADOR":

                st.success(
                    "Validación correcta: cada embedding "
                    "corresponde a un registro de la "
                    "base semántica."
                )


    except Exception as error_embeddings:

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.error(
                f"Error cargando los embeddings: "
                f"{error_embeddings}"
            )

else:

    if ROL_ACTUAL == "ADMINISTRADOR":

        st.warning(
            "No se puede cargar embeddings_sintomas.npy "
            "porque el archivo no fue encontrado."
        )
# ============================================================
# 8.3 CARGAR MODELO BIOMÉDICO
# ============================================================

modelo_biomedico = None


try:

    from sentence_transformers import SentenceTransformer


    @st.cache_resource
    def cargar_modelo_biomedico():

        modelo = SentenceTransformer(
            "SINAI/ALIA-MrBERT-es-biomedical-embeddings"
        )

        return modelo


    modelo_biomedico = (
        cargar_modelo_biomedico()
    )


    if ROL_ACTUAL == "ADMINISTRADOR":

        st.success(
            "Modelo biomédico cargado correctamente."
        )


except Exception as error_modelo:

    if ROL_ACTUAL == "ADMINISTRADOR":

        st.warning(
            "El modelo biomédico no pudo cargarse "
            "en este momento."
        )

        st.code(
            str(error_modelo)
        )


# ============================================================
# 8.4 VALIDACIÓN GENERAL DE LA INFRAESTRUCTURA
# ============================================================

estado_base = (
    base_semantica is not None
)

estado_embeddings = (
    embeddings_sintomas is not None
)

estado_modelo = (
    modelo_biomedico is not None
)


if ROL_ACTUAL == "ADMINISTRADOR":

    st.write(
        "**Estado de la búsqueda semántica:**"
    )


    if (
        estado_base
        and
        estado_embeddings
        and
        estado_modelo
    ):

        st.success(
            "✓ Infraestructura semántica preparada."
        )

        st.caption(
            "Base semántica + embeddings precalculados "
            "+ modelo biomédico disponibles."
        )

    else:

        st.warning(
            "La infraestructura semántica "
            "todavía no está completamente disponible."
        )

        if not estado_base:

            st.write(
                "• Base semántica: no disponible"
            )

        else:

            st.write(
                "• Base semántica: ✓ disponible"
            )


        if not estado_embeddings:

            st.write(
                "• Embeddings: no disponibles"
            )

        else:

            st.write(
                "• Embeddings: ✓ disponibles"
            )


        if not estado_modelo:

            st.write(
                "• Modelo biomédico: no disponible"
            )

        else:

            st.write(
                "• Modelo biomédico: ✓ disponible"
            )
# ============================================================
# 9. DIAGNÓSTICO DE EMBEDDINGS
# ============================================================

embeddings = None

if embeddings_ok:

    try:

        embeddings = np.load(
            ARCHIVO_EMBEDDINGS,
            allow_pickle=False
        )

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.header(
                "4. Diagnóstico de embeddings"
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

        if ROL_ACTUAL == "ADMINISTRADOR":

            st.error(
                f"Error cargando embeddings: "
                f"{error_embeddings}"
            )

else:

    if ROL_ACTUAL == "ADMINISTRADOR":

        st.warning(
            "No se puede diagnosticar embeddings "
            "porque el archivo no fue encontrado."
        )


# ============================================================
# 10. DETECCIÓN DE IMÁGENES
# ============================================================

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


if ROL_ACTUAL == "ADMINISTRADOR":

    st.header("5. Imágenes disponibles")

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

if ROL_ACTUAL == "ADMINISTRADOR":

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

opciones_menu = [
    "CONSULTA",
    "ASESORÍA",
    "EVALUACIÓN"
]

opcion_principal = st.radio(
    "Seleccione una sección:",
    opciones_menu,
    horizontal=True
)
# ============================================================
# 5. SECCIÓN CONSULTA
# ============================================================

opcion_consulta = None
if opcion_principal == "CONSULTA":

    st.header("CONSULTA")

    st.write(
        "Consulta información de Productos, Patologias, Complementarios "
        "y Restricciones."
    )

    opcion_consulta = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Productos",
            "Patologias",
            "Complementarios",
            "Restricciones"
        ]
    )

    if opcion_consulta == "Productos":

        st.info(
            "Módulo de consulta de Productos. "
            "Se incorporará en el siguiente bloque."
        )

    elif opcion_consulta == "Patologias":

        st.info(
            "Módulo de consulta de Patologias. "
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

    opcion_asesoria = st.selectbox(
        "¿Qué desea realizar?",
        [
            "Seleccione una opción",
            "Entrevista"
        ],
        key="menu_asesoria"
    )

# ============================================================
# ============================================================
# 7. SECCIÓN EVALUACIÓN
# ============================================================

elif opcion_principal == "EVALUACIÓN":

    st.header("EVALUACIÓN")

    st.write(
        "Espacio para el aprendizaje, la evaluación "
        "y el seguimiento de resultados."
    )

    if ROL_ACTUAL == "ADMINISTRADOR":

        opciones_evaluacion = [
            "Seleccione una opción",
            "Banco general de preguntas",
            "Banco de preguntas especiales",
            "Evaluaciones generales",
            "Evaluación controlada",
            "Historial de evaluaciones"
        ]

    else:

        opciones_evaluacion = [
            "Seleccione una opción",
            "Evaluaciones generales",
            "Evaluación controlada",
            "Historial de evaluaciones"
        ]

    opcion_evaluacion = st.selectbox(
        "Seleccione una opción:",
        opciones_evaluacion,
        key="menu_evaluacion"
    )
    # ========================================================
    # 7.1 BANCO GENERAL DE PREGUNTAS
    # SOLO ADMINISTRADOR
    # ========================================================

    if (
        ROL_ACTUAL == "ADMINISTRADOR"
        and
        opcion_evaluacion
        == "Banco general de preguntas"
    ):

        st.subheader(
            "Banco general de preguntas"
        )

        st.write(
            "Administración del banco permanente "
            "de preguntas generales."
        )

        st.info(
            "El banco general será alimentado "
            "automáticamente a partir de la matriz "
            "y las preguntas deberán ser revisadas "
            "por el Administrador antes de poder "
            "utilizarse en una evaluación."
        )

    # ========================================================
    # 7.2 BANCO GENERAL — CARGA Y CREACIÓN DEL ARCHIVO
    # SOLO ADMINISTRADOR
    # ========================================================

    if (
        ROL_ACTUAL == "ADMINISTRADOR"
        and
        opcion_evaluacion
        == "Banco general de preguntas"
    ):

        st.subheader(
            "Banco General de Preguntas"
        )

        st.write(
            "Archivo permanente de preguntas generales "
            "del aplicativo."
        )

        RUTA_BANCO_GENERAL = (
            BASE_DIR
            / "BANCO_PREGUNTAS_GENERALES.xlsx"
        )

        COLUMNAS_BANCO_GENERAL = [

            "Pregunta_ID",
            "Modulo",
            "Tema",
            "Nivel",
            "Tipo_Relacion",
            "Pregunta",
            "Respuesta_1",
            "Respuesta_2",
            "Respuesta_3",
            "Respuesta_4",
            "Respuesta_Correcta",
            "Estado",
            "Observacion_Administrador",
            "Fecha_Generacion",
            "Fuente_ID"

        ]

        # ----------------------------------------------------
        # CARGAR BANCO SI YA EXISTE
        # ----------------------------------------------------

        if RUTA_BANCO_GENERAL.exists():

            try:

                banco_general = pd.read_excel(
                    RUTA_BANCO_GENERAL,
                    dtype=str
                )

                banco_general = (
                    banco_general.fillna("")
                )

                for columna in (
                    COLUMNAS_BANCO_GENERAL
                ):

                    if columna not in banco_general.columns:

                        banco_general[columna] = ""

                banco_general = banco_general[
                    COLUMNAS_BANCO_GENERAL
                ]

                st.success(
                    "Banco General cargado correctamente."
                )

                st.write(
                    f"Preguntas almacenadas: "
                    f"**{len(banco_general):,}**"
                )

            except Exception as error_banco:

                st.error(
                    "Error cargando el Banco General."
                )

                st.code(
                    str(error_banco)
                )

        # ----------------------------------------------------
        # CREAR BANCO SI NO EXISTE
        # ----------------------------------------------------

        else:

            banco_general = pd.DataFrame(
                columns=COLUMNAS_BANCO_GENERAL
            )

            try:

                banco_general.to_excel(
                    RUTA_BANCO_GENERAL,
                    index=False,
                    sheet_name="Banco_General"
                )

                st.success(
                    "Banco General creado correctamente."
                )

                st.write(
                    "El archivo fue creado con la "
                    "estructura inicial y está listo "
                    "para recibir preguntas."
                )

            except Exception as error_creacion:

                st.error(
                    "No fue posible crear el "
                    "Banco General."
                )

                st.code(
                    str(error_creacion)
                )

    # ========================================================
    # 7.3 BANCO GENERAL — SINCRONIZACIÓN CON GITHUB
    # SOLO ADMINISTRADOR
    # ========================================================

    if (
        ROL_ACTUAL == "ADMINISTRADOR"
        and
        opcion_evaluacion
        == "Banco general de preguntas"
    ):

        st.subheader(
            "Sincronización del Banco General"
        )

        ruta_github_banco = (
            "BANCO_PREGUNTAS_GENERALES.xlsx"
        )

        url_github_banco = (
            f"https://api.github.com/repos/"
            f"{GITHUB_USUARIO}/"
            f"{GITHUB_REPOSITORIO}/"
            f"contents/{ruta_github_banco}"
        )

        headers_github_banco = {

            "Authorization":
                f"Bearer {GITHUB_TOKEN}",

            "Accept":
                "application/vnd.github+json",

            "X-GitHub-Api-Version":
                "2022-11-28"

        }

        if st.button(
            "Sincronizar Banco General con GitHub",
            key="sincronizar_banco_general"
        ):

            if not RUTA_BANCO_GENERAL.exists():

                st.error(
                    "No existe el archivo local "
                    "BANCO_PREGUNTAS_GENERALES.xlsx."
                )

            else:

                try:

                    # ----------------------------------------
                    # LEER ARCHIVO LOCAL
                    # ----------------------------------------

                    with open(
                        RUTA_BANCO_GENERAL,
                        "rb"
                    ) as archivo:

                        contenido_banco = (
                            archivo.read()
                        )


                    contenido_base64 = (
                        base64.b64encode(
                            contenido_banco
                        ).decode("utf-8")
                    )


                    # ----------------------------------------
                    # CONSULTAR SI YA EXISTE EN GITHUB
                    # ----------------------------------------

                    solicitud_get = (
                        urllib.request.Request(
                            url_github_banco,
                            headers=headers_github_banco,
                            method="GET"
                        )
                    )


                    sha_actual = None


                    try:

                        with urllib.request.urlopen(
                            solicitud_get
                        ) as respuesta:

                            informacion_github = (
                                json.loads(
                                    respuesta.read()
                                    .decode("utf-8")
                                )
                            )

                        sha_actual = (
                            informacion_github["sha"]
                        )


                    except urllib.error.HTTPError as error:

                        if error.code != 404:

                            raise


                    # ----------------------------------------
                    # PREPARAR GUARDADO
                    # ----------------------------------------

                    datos_banco = {

                        "message": (
                            "Actualizar Banco General "
                            "de Preguntas"
                        ),

                        "content":
                            contenido_base64

                    }


                    # Si ya existe, GitHub exige SHA.
                    if sha_actual:

                        datos_banco["sha"] = (
                            sha_actual
                        )


                    datos_json = (
                        json.dumps(
                            datos_banco
                        ).encode("utf-8")
                    )


                    # ----------------------------------------
                    # GUARDAR EN GITHUB
                    # ----------------------------------------

                    solicitud_put = (
                        urllib.request.Request(
                            url_github_banco,
                            data=datos_json,
                            headers={
                                **headers_github_banco,
                                "Content-Type":
                                    "application/json"
                            },
                            method="PUT"
                        )
                    )


                    with urllib.request.urlopen(
                        solicitud_put
                    ) as respuesta:

                        resultado_github = (
                            json.loads(
                                respuesta.read()
                                .decode("utf-8")
                            )
                        )


                    if resultado_github.get(
                        "content"
                    ):

                        st.success(
                            "✓ Banco General "
                            "sincronizado correctamente "
                            "con GitHub."
                        )

                    else:

                        st.warning(
                            "GitHub respondió, pero "
                            "no se confirmó el archivo."
                        )


                except urllib.error.HTTPError as error:

                    detalle = (
                        error.read().decode(
                            "utf-8",
                            errors="ignore"
                        )
                    )

                    st.error(
                        "No fue posible sincronizar "
                        "el Banco General con GitHub."
                    )

                    st.code(
                        detalle
                    )


                except Exception as error:

                    st.error(
                        "Error durante la sincronización "
                        "del Banco General."
                    )

                    st.code(
                        str(error)
                    )
# ========================================================
# ========================================================
# 7.4 BANCO GENERAL — GENERADOR Y VALIDACIÓN
# SOLO ADMINISTRADOR
# ========================================================

if (
    ROL_ACTUAL == "ADMINISTRADOR"
    and opcion_evaluacion == "Banco general de preguntas"
):

    st.subheader(
        "Banco General de Preguntas"
    )

    # ====================================================
    # 7.4.1 CARGAR BANCO
    # ====================================================

    if RUTA_BANCO_GENERAL.exists():

        try:

            banco_general = pd.read_excel(
                RUTA_BANCO_GENERAL,
                dtype=str
            ).fillna("")

        except Exception as error_banco:

            st.error(
                "No fue posible cargar el Banco General."
            )

            st.code(
                str(error_banco)
            )

            banco_general = pd.DataFrame()

    else:

        banco_general = pd.DataFrame(
            columns=COLUMNAS_BANCO_GENERAL
        )

    # ====================================================
    # 7.4.2 FUNCIONES AUXILIARES
    # ====================================================

    def limpiar_74(valor):

        if pd.isna(valor):
            return ""

        return str(valor).strip()


    def lista_74(valor):

        texto = limpiar_74(valor)

        if not texto:
            return []

        elementos = []

        for elemento in texto.split(";"):

            elemento = elemento.strip()

            if elemento:
                elementos.append(elemento)

        return elementos


    def siguiente_id_74(banco):

        numeros = []

        if not banco.empty:

            for valor in banco["Pregunta_ID"]:

                texto = limpiar_74(valor)

                if texto.startswith("PAT_"):

                    try:

                        numeros.append(
                            int(texto[4:])
                        )

                    except ValueError:
                        pass

        if numeros:
            return max(numeros) + 1

        return 1


    def distractores_patologias_74(
        patologias,
        patologia_actual,
        cantidad=3
    ):

        disponibles = [

            p
            for p in patologias
            if p.lower()
            != patologia_actual.lower()

        ]

        if len(disponibles) < cantidad:
            return []

        return list(
            np.random.choice(
                disponibles,
                size=cantidad,
                replace=False
            )
        )


    def opciones_nivel1_74(
        correcta,
        distractores
    ):

        opciones = [
            correcta
        ] + distractores

        np.random.shuffle(
            opciones
        )

        posicion_correcta = (
            opciones.index(correcta) + 1
        )

        return (
            opciones,
            posicion_correcta
        )


    def opciones_nivel2_74(
        correcta_1,
        correcta_2,
        distractor_1,
        distractor_2
    ):

        opciones = [

            correcta_1,
            correcta_2,
            distractor_1,
            distractor_2

        ]

        np.random.shuffle(
            opciones
        )

        posiciones_correctas = [

            indice + 1
            for indice, opcion
            in enumerate(opciones)
            if opcion in [
                correcta_1,
                correcta_2
            ]

        ]

        return (
            opciones,
            sorted(posiciones_correctas)
        )


    # ====================================================
    # 7.4.3 VALIDAR HOJA PATOLOGÍAS
    # ====================================================

    columnas_patologias = [

        "Patologia_ID",
        "Patología",
        "Descripción breve (para cliente)",
        "Causas frecuentes (resumen)",
        "Síntomas/Señales clave (checklist)"

    ]

    faltantes = [

        columna
        for columna in columnas_patologias
        if columna not in Patologias.columns

    ]

    if faltantes:

        st.error(
            "Faltan columnas necesarias en la hoja "
            "Patologias:"
        )

        for columna in faltantes:
            st.write(
                f"- {columna}"
            )

    else:

        patologias = (
            Patologias[
                columnas_patologias
            ]
            .copy()
            .fillna("")
        )

        for columna in columnas_patologias:

            patologias[columna] = (
                patologias[columna]
                .astype(str)
                .str.strip()
            )

        # ================================================
        # SOLO SE UTILIZAN REGISTROS COMPLETOS
        # ================================================

        patologias = patologias[
            (patologias["Patologia_ID"] != "")
            &
            (patologias["Patología"] != "")
            &
            (
                patologias[
                    "Descripción breve (para cliente)"
                ] != ""
            )
            &
            (
                patologias[
                    "Causas frecuentes (resumen)"
                ] != ""
            )
            &
            (
                patologias[
                    "Síntomas/Señales clave (checklist)"
                ] != ""
            )
        ].copy()

        st.info(
            f"Patologías completas disponibles: "
            f"{len(patologias)}"
        )

        # =================================================
        # 7.4.4 SELECCIÓN DE NIVELES
        # =================================================

        niveles_generar = st.multiselect(

            "Seleccione los niveles que desea generar:",
            [
                "Nivel 1",
                "Nivel 2"
            ],

            default=[
                "Nivel 1",
                "Nivel 2"
            ],

            key="niveles_generacion_patologias"

        )

        cantidad_por_patologia = st.number_input(

            "Cantidad máxima de preguntas por patología:",
            min_value=1,
            max_value=20,
            value=6,
            step=1,

            key="cantidad_preguntas_patologia"

        )

        # =================================================
        # 7.4.5 GENERAR
        # =================================================

        if st.button(
            "GENERAR PREGUNTAS",
            key="generar_banco_patologias"
        ):

            if not niveles_generar:

                st.warning(
                    "Seleccione al menos un nivel."
                )

            else:

                nuevas_preguntas = []

                siguiente_id = (
                    siguiente_id_74(
                        banco_general
                    )
                )

                relaciones_existentes = set()

                if not banco_general.empty:

                    for _, fila in (
                        banco_general.iterrows()
                    ):

                        estado = (
                            limpiar_74(
                                fila.get(
                                    "Estado",
                                    ""
                                )
                            )
                            .upper()
                        )

                        if estado in [
                            "PENDIENTE",
                            "APROBADA"
                        ]:

                            fuente = limpiar_74(
                                fila.get(
                                    "Fuente_ID",
                                    ""
                                )
                            )

                            nivel = limpiar_74(
                                fila.get(
                                    "Nivel",
                                    ""
                                )
                            )

                            relacion = limpiar_74(
                                fila.get(
                                    "Tipo_Relacion",
                                    ""
                                )
                            )

                            relaciones_existentes.add(
                                (
                                    fuente,
                                    nivel,
                                    relacion
                                )
                            )

                # =========================================
                # RECORRER PATOLOGÍAS
                # =========================================

                for _, fila in (
                    patologias.iterrows()
                ):

                    if (
                        len(nuevas_preguntas)
                        >=
                        (
                            len(patologias)
                            *
                            cantidad_por_patologia
                        )
                    ):
                        break

                    patologia_id = limpiar_74(
                        fila["Patologia_ID"]
                    )

                    patologia = limpiar_74(
                        fila["Patología"]
                    )

                    descripcion = limpiar_74(
                        fila[
                            "Descripción breve (para cliente)"
                        ]
                    )

                    causas = lista_74(
                        fila[
                            "Causas frecuentes (resumen)"
                        ]
                    )

                    sintomas = lista_74(
                        fila[
                            "Síntomas/Señales clave (checklist)"
                        ]
                    )

                    otras_patologias = [

                        limpiar_74(
                            valor
                        )

                        for valor in (
                            patologias["Patología"]
                        )

                        if limpiar_74(
                            valor
                        ).lower()
                        !=
                        patologia.lower()

                    ]

                    if len(
                        otras_patologias
                    ) < 3:

                        continue

                    # =====================================
                    # RELACIONES POSIBLES
                    # =====================================

                    relaciones = []

                    if "Nivel 1" in niveles_generar:

                        relaciones.extend([

                            (
                                "Nivel 1",
                                "Descripcion_Patologia"
                            ),

                            (
                                "Nivel 1",
                                "Patologia_Descripcion"
                            ),

                            (
                                "Nivel 1",
                                "Sintomas_Patologia"
                            ),

                            (
                                "Nivel 1",
                                "Causas_Patologia"
                            )

                        ])

                    if "Nivel 2" in niveles_generar:

                        relaciones.extend([

                            (
                                "Nivel 2",
                                "Descripcion_Sintomas"
                            ),

                            (
                                "Nivel 2",
                                "Descripcion_Causa"
                            ),

                            (
                                "Nivel 2",
                                "Sintomas_Causa"
                            )

                        ])

                    np.random.shuffle(
                        relaciones
                    )

                    # =====================================
                    # CREAR PREGUNTAS
                    # =====================================

                    for nivel, tipo in relaciones:

                        if (
                            len(nuevas_preguntas)
                            >=
                            (
                                len(patologias)
                                *
                                cantidad_por_patologia
                            )
                        ):
                            break

                        clave = (
                            patologia_id,
                            nivel,
                            tipo
                        )

                        if (
                            clave
                            in relaciones_existentes
                        ):
                            continue

                        distractores = (
                            distractores_patologias_74(
                                otras_patologias,
                                patologia,
                                3
                            )
                        )

                        if len(
                            distractores
                        ) < 3:

                            continue

                        pregunta = ""
                        opciones = []
                        correctas = []

                        # =================================
                        # NIVEL 1
                        # =================================

                        if (
                            nivel == "Nivel 1"
                            and
                            tipo
                            == "Descripcion_Patologia"
                        ):

                            pregunta = (
                                "Una persona presenta "
                                f"la siguiente descripción: "
                                f"{descripcion} "
                                "¿A cuál de las siguientes "
                                "patologías corresponde?"
                            )

                            opciones, correcta = (
                                opciones_nivel1_74(
                                    patologia,
                                    distractores
                                )
                            )

                            correctas = [
                                correcta
                            ]

                        elif (
                            nivel == "Nivel 1"
                            and
                            tipo
                            == "Patologia_Descripcion"
                        ):

                            pregunta = (
                                f"¿Cuál de las siguientes "
                                f"descripciones corresponde "
                                f"a {patologia}?"
                            )

                            descripciones = []

                            for _, otra in (
                                patologias.iterrows()
                            ):

                                otra_id = limpiar_74(
                                    otra[
                                        "Patologia_ID"
                                    ]
                                )

                                if (
                                    otra_id
                                    !=
                                    patologia_id
                                ):

                                    texto = limpiar_74(
                                        otra[
                                            "Descripción breve (para cliente)"
                                        ]
                                    )

                                    if texto:
                                        descripciones.append(
                                            texto
                                        )

                            if len(
                                descripciones
                            ) < 3:

                                continue

                            distractores_descripcion = (
                                list(
                                    np.random.choice(
                                        descripciones,
                                        size=3,
                                        replace=False
                                    )
                                )
                            )

                            opciones, correcta = (
                                opciones_nivel1_74(
                                    descripcion,
                                    distractores_descripcion
                                )
                            )

                            correctas = [
                                correcta
                            ]

                        elif (
                            nivel == "Nivel 1"
                            and
                            tipo
                            == "Sintomas_Patologia"
                        ):

                            if len(
                                sintomas
                            ) < 2:

                                continue

                            pregunta = (
                                "Una persona presenta "
                                f"señales como "
                                f"{'; '.join(sintomas)}. "
                                "¿Con cuál de las siguientes "
                                "patologías se relacionan?"
                            )

                            opciones, correcta = (
                                opciones_nivel1_74(
                                    patologia,
                                    distractores
                                )
                            )

                            correctas = [
                                correcta
                            ]

                        elif (
                            nivel == "Nivel 1"
                            and
                            tipo
                            == "Causas_Patologia"
                        ):

                            if len(
                                causas
                            ) < 1:

                                continue

                            pregunta = (
                                "Factores como "
                                f"{'; '.join(causas)} "
                                "pueden estar relacionados "
                                "con cuál de las siguientes "
                                "patologías?"
                            )

                            opciones, correcta = (
                                opciones_nivel1_74(
                                    patologia,
                                    distractores
                                )
                            )

                            correctas = [
                                correcta
                            ]

                        # =================================
                        # NIVEL 2
                        # =================================

                        elif (
                            nivel == "Nivel 2"
                            and
                            tipo
                            == "Descripcion_Sintomas"
                        ):

                            if len(
                                sintomas
                            ) < 1:

                                continue

                            correcta_1 = (
                                "La descripción corresponde "
                                f"a {patologia}."
                            )

                            correcta_2 = (
                                "Las señales indicadas son "
                                f"compatibles con {patologia}."
                            )

                            opciones, correctas = (
                                opciones_nivel2_74(
                                    correcta_1,
                                    correcta_2,
                                    (
                                        "La descripción corresponde "
                                        "a una condición diferente."
                                    ),
                                    (
                                        "Las señales indicadas "
                                        "no son compatibles "
                                        "con esta patología."
                                    )
                                )
                            )

                            pregunta = (
                                f"Considere la patología "
                                f"{patologia}. "
                                f"Su descripción es: "
                                f"{descripcion} "
                                f"y presenta señales como "
                                f"{'; '.join(sintomas)}. "
                                "Seleccione las DOS afirmaciones "
                                "correctas."
                            )

                        elif (
                            nivel == "Nivel 2"
                            and
                            tipo
                            == "Descripcion_Causa"
                        ):

                            if len(
                                causas
                            ) < 1:

                                continue

                            correcta_1 = (
                                f"La descripción corresponde "
                                f"a {patologia}."
                            )

                            correcta_2 = (
                                "Una causa o factor relacionado "
                                f"es {causas[0]}."
                            )

                            opciones, correctas = (
                                opciones_nivel2_74(
                                    correcta_1,
                                    correcta_2,
                                    (
                                        "La descripción corresponde "
                                        "principalmente a otra "
                                        "condición."
                                    ),
                                    (
                                        "El factor indicado no "
                                        "se relaciona con esta "
                                        "patología."
                                    )
                                )
                            )

                            pregunta = (
                                f"Considere {patologia}. "
                                f"La descripción es: "
                                f"{descripcion} "
                                f"y uno de los factores "
                                f"relacionados es "
                                f"{causas[0]}. "
                                "Seleccione las DOS afirmaciones "
                                "correctas."
                            )

                        elif (
                            nivel == "Nivel 2"
                            and
                            tipo
                            == "Sintomas_Causa"
                        ):

                            if (
                                len(sintomas) < 1
                                or
                                len(causas) < 1
                            ):

                                continue

                            correcta_1 = (
                                f"{sintomas[0]} puede "
                                f"presentarse en {patologia}."
                            )

                            correcta_2 = (
                                f"{causas[0]} puede estar "
                                f"relacionada con {patologia}."
                            )

                            opciones, correctas = (
                                opciones_nivel2_74(
                                    correcta_1,
                                    correcta_2,
                                    (
                                        f"{sintomas[0]} "
                                        "corresponde principalmente "
                                        "a otra condición."
                                    ),
                                    (
                                        f"{causas[0]} "
                                        "corresponde exclusivamente "
                                        "a otra patología."
                                    )
                                )
                            )

                            pregunta = (
                                f"En relación con "
                                f"{patologia}, considere "
                                f"la señal "
                                f"{sintomas[0]} "
                                f"y el factor "
                                f"{causas[0]}. "
                                "Seleccione las DOS afirmaciones "
                                "correctas."
                            )

                        # =================================
                        # VALIDAR PREGUNTA
                        # =================================

                        if (
                            not pregunta
                            or
                            len(opciones) != 4
                            or
                            len(correctas) == 0
                        ):

                            continue

                        # =================================
                        # CREAR REGISTRO
                        # =================================

                        nuevas_preguntas.append({

                            "Pregunta_ID":
                                f"PAT_{siguiente_id:05d}",

                            "Modulo":
                                "Patologias",

                            "Tema":
                                patologia,

                            "Nivel":
                                nivel,

                            "Tipo_Relacion":
                                tipo,

                            "Pregunta":
                                pregunta,

                            "Respuesta_1":
                                opciones[0],

                            "Respuesta_2":
                                opciones[1],

                            "Respuesta_3":
                                opciones[2],

                            "Respuesta_4":
                                opciones[3],

                            "Respuesta_Correcta":
                                ",".join(
                                    map(
                                        str,
                                        correctas
                                    )
                                ),

                            "Estado":
                                "PENDIENTE",

                            "Observacion_Administrador":
                                "",

                            "Fecha_Generacion":
                                pd.Timestamp.now().strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                ),

                            "Fuente_ID":
                                patologia_id

                        })

                        siguiente_id += 1

                        relaciones_existentes.add(
                            clave
                        )

                # =========================================
                # 7.4.6 GUARDAR EXCEL
                # =========================================

                if nuevas_preguntas:

                    nuevas_df = pd.DataFrame(
                        nuevas_preguntas
                    )

                    banco_general = pd.concat(
                        [
                            banco_general,
                            nuevas_df
                        ],
                        ignore_index=True
                    )

                    banco_general = (
                        banco_general[
                            COLUMNAS_BANCO_GENERAL
                        ]
                    )

                    banco_general.to_excel(
                        RUTA_BANCO_GENERAL,
                        index=False,
                        sheet_name="Banco_General"
                    )

                    st.success(
                        f"Se generaron "
                        f"{len(nuevas_df)} preguntas "
                        "y se guardaron en Excel."
                    )

                    # =====================================
                    # 7.4.7 SINCRONIZAR GITHUB
                    # =====================================

                    try:

                        ruta_github_banco = (
                            "BANCO_PREGUNTAS_GENERALES.xlsx"
                        )

                        url_github_banco = (
                            f"https://api.github.com/repos/"
                            f"{GITHUB_USUARIO}/"
                            f"{GITHUB_REPOSITORIO}/"
                            f"contents/"
                            f"{ruta_github_banco}"
                        )

                        headers_github_banco = {

                            "Authorization":
                                f"Bearer {GITHUB_TOKEN}",

                            "Accept":
                                "application/vnd.github+json",

                            "X-GitHub-Api-Version":
                                "2022-11-28"

                        }

                        with open(
                            RUTA_BANCO_GENERAL,
                            "rb"
                        ) as archivo:

                            contenido = (
                                archivo.read()
                            )

                        contenido_base64 = (
                            base64.b64encode(
                                contenido
                            ).decode(
                                "utf-8"
                            )
                        )

                        solicitud_get = (
                            urllib.request.Request(
                                url_github_banco,
                                headers=headers_github_banco,
                                method="GET"
                            )
                        )

                        sha_actual = None

                        try:

                            with urllib.request.urlopen(
                                solicitud_get
                            ) as respuesta:

                                informacion = json.loads(
                                    respuesta.read()
                                    .decode("utf-8")
                                )

                            sha_actual = (
                                informacion["sha"]
                            )

                        except urllib.error.HTTPError as error:

                            if error.code != 404:
                                raise

                        datos = {

                            "message":
                                "Actualizar Banco General",

                            "content":
                                contenido_base64

                        }

                        if sha_actual:
                            datos["sha"] = sha_actual

                        solicitud_put = (
                            urllib.request.Request(
                                url_github_banco,
                                data=json.dumps(
                                    datos
                                ).encode("utf-8"),
                                headers={
                                    **headers_github_banco,
                                    "Content-Type":
                                        "application/json"
                                },
                                method="PUT"
                            )
                        )

                        with urllib.request.urlopen(
                            solicitud_put
                        ) as respuesta:

                            resultado = json.loads(
                                respuesta.read()
                                .decode("utf-8")
                            )

                        if resultado.get("content"):

                            st.success(
                                "✓ Banco sincronizado "
                                "correctamente con GitHub."
                            )

                    except Exception as error_github:

                        st.warning(
                            "Las preguntas se guardaron "
                            "en Excel, pero no fue posible "
                            "sincronizar con GitHub."
                        )

                        st.code(
                            str(error_github)
                        )

                    # =====================================
                    # MOSTRAR RESULTADO
                    # =====================================

                    st.subheader(
                        "Preguntas generadas"
                    )

                    st.dataframe(
                        nuevas_df[
                            [
                                "Pregunta_ID",
                                "Modulo",
                                "Tema",
                                "Nivel",
                                "Tipo_Relacion",
                                "Pregunta",
                                "Respuesta_1",
                                "Respuesta_2",
                                "Respuesta_3",
                                "Respuesta_4",
                                "Respuesta_Correcta",
                                "Estado"
                            ]
                        ],
                        use_container_width=True
                    )

                else:

                    st.warning(
                        "No se generaron preguntas nuevas. "
                        "Las relaciones disponibles ya pueden "
                        "existir en el banco o la información "
                        "de la matriz puede ser insuficiente."
                    )

    # ====================================================
    # 7.4.8 VALIDACIÓN POR BLOQUES
    # ====================================================

    st.divider()

    st.subheader(
        "Validación de preguntas"
    )

    if RUTA_BANCO_GENERAL.exists():

        banco_validacion = pd.read_excel(
            RUTA_BANCO_GENERAL,
            dtype=str
        ).fillna("")

        pendientes = (
            banco_validacion[
                banco_validacion["Estado"]
                .astype(str)
                .str.strip()
                .str.upper()
                == "PENDIENTE"
            ]
            .copy()
        )

        st.write(
            f"Preguntas pendientes: "
            f"**{len(pendientes)}**"
        )

        if not pendientes.empty:

            # =============================================
            # FILTROS
            # =============================================

            modulos_disponibles = sorted(
                pendientes[
                    "Modulo"
                ]
                .dropna()
                .unique()
                .tolist()
            )

            niveles_disponibles = sorted(
                pendientes[
                    "Nivel"
                ]
                .dropna()
                .unique()
                .tolist()
            )

            modulo_filtro = st.selectbox(
                "Módulo:",
                [
                    "Todos"
                ] + modulos_disponibles,
                key="filtro_modulo_validacion"
            )

            nivel_filtro = st.selectbox(
                "Nivel:",
                [
                    "Todos"
                ] + niveles_disponibles,
                key="filtro_nivel_validacion"
            )

            bloque = pendientes.copy()

            if modulo_filtro != "Todos":

                bloque = bloque[
                    bloque["Modulo"]
                    == modulo_filtro
                ]

            if nivel_filtro != "Todos":

                bloque = bloque[
                    bloque["Nivel"]
                    == nivel_filtro
                ]

            st.info(
                f"Preguntas en este bloque: "
                f"{len(bloque)}"
            )

            # =============================================
            # OPCIÓN PARA ESTABLECER ESTADO DEL BLOQUE
            # =============================================

            accion_bloque = st.radio(

                "Acción para el bloque:",
                [
                    "Mantener estado individual",
                    "Aprobar todas",
                    "Rechazar todas"
                ],

                horizontal=True,

                key="accion_bloque_validacion"

            )

            cambios_bloque = {}

            # =============================================
            # MOSTRAR PREGUNTAS
            # =============================================

            for indice, fila in (
                bloque.iterrows()
            ):

                pregunta_id = (
                    limpiar_74(
                        fila["Pregunta_ID"]
                    )
                )

                st.markdown(
                    f"### {pregunta_id}"
                )

                st.write(
                    f"**Nivel:** {fila['Nivel']}"
                )

                st.write(
                    f"**Tipo:** "
                    f"{fila['Tipo_Relacion']}"
                )

                st.write(
                    f"**Pregunta:** "
                    f"{fila['Pregunta']}"
                )

                st.write(
                    f"1. {fila['Respuesta_1']}"
                )

                st.write(
                    f"2. {fila['Respuesta_2']}"
                )

                st.write(
                    f"3. {fila['Respuesta_3']}"
                )

                st.write(
                    f"4. {fila['Respuesta_4']}"
                )

                st.info(
                    "Respuesta(s) correcta(s): "
                    f"{fila['Respuesta_Correcta']}"
                )

                if accion_bloque == (
                    "Aprobar todas"
                ):

                    estado = "APROBADA"

                elif accion_bloque == (
                    "Rechazar todas"
                ):

                    estado = "RECHAZADA"

                else:

                    estado = st.selectbox(

                        "Estado:",
                        [
                            "PENDIENTE",
                            "APROBADA",
                            "RECHAZADA"
                        ],

                        index=0,

                        key=(
                            f"estado_validacion_"
                            f"{pregunta_id}"
                        )

                    )

                observacion = st.text_area(

                    "Observación:",
                    value=limpiar_74(
                        fila[
                            "Observacion_Administrador"
                        ]
                    ),

                    key=(
                        f"obs_validacion_"
                        f"{pregunta_id}"
                    ),

                    height=70

                )

                cambios_bloque[
                    pregunta_id
                ] = {

                    "Estado":
                        estado,

                    "Observacion":
                        observacion

                }

                st.divider()

            # =============================================
            # GUARDAR TODO EL BLOQUE
            # =============================================

            if st.button(
                "GUARDAR BLOQUE",
                key="guardar_bloque_validacion"
            ):

                try:

                    banco_guardar = pd.read_excel(
                        RUTA_BANCO_GENERAL,
                        dtype=str
                    ).fillna("")

                    for pregunta_id, cambios in (
                        cambios_bloque.items()
                    ):

                        posicion = (
                            banco_guardar[
                                "Pregunta_ID"
                            ]
                            .astype(str)
                            .str.strip()
                            ==
                            pregunta_id
                        )

                        banco_guardar.loc[
                            posicion,
                            "Estado"
                        ] = cambios["Estado"]

                        banco_guardar.loc[
                            posicion,
                            "Observacion_Administrador"
                        ] = cambios["Observacion"]

                    banco_guardar.to_excel(
                        RUTA_BANCO_GENERAL,
                        index=False,
                        sheet_name="Banco_General"
                    )

                    st.success(
                        f"Se guardaron "
                        f"{len(cambios_bloque)} "
                        "preguntas del bloque."
                    )

                    # =====================================
                    # SINCRONIZAR CAMBIOS CON GITHUB
                    # =====================================

                    try:

                        with open(
                            RUTA_BANCO_GENERAL,
                            "rb"
                        ) as archivo:

                            contenido = (
                                archivo.read()
                            )

                        contenido_base64 = (
                            base64.b64encode(
                                contenido
                            ).decode(
                                "utf-8"
                            )
                        )

                        url_github_banco = (
                            f"https://api.github.com/repos/"
                            f"{GITHUB_USUARIO}/"
                            f"{GITHUB_REPOSITORIO}/"
                            f"contents/"
                            f"BANCO_PREGUNTAS_GENERALES.xlsx"
                        )

                        headers_github_banco = {

                            "Authorization":
                                f"Bearer {GITHUB_TOKEN}",

                            "Accept":
                                "application/vnd.github+json",

                            "X-GitHub-Api-Version":
                                "2022-11-28"

                        }

                        solicitud_get = (
                            urllib.request.Request(
                                url_github_banco,
                                headers=headers_github_banco,
                                method="GET"
                            )
                        )

                        with urllib.request.urlopen(
                            solicitud_get
                        ) as respuesta:

                            informacion = json.loads(
                                respuesta.read()
                                .decode("utf-8")
                            )

                        datos = {

                            "message":
                                "Actualizar estados "
                                "Banco General",

                            "content":
                                contenido_base64,

                            "sha":
                                informacion["sha"]

                        }

                        solicitud_put = (
                            urllib.request.Request(
                                url_github_banco,
                                data=json.dumps(
                                    datos
                                ).encode("utf-8"),
                                headers={
                                    **headers_github_banco,
                                    "Content-Type":
                                        "application/json"
                                },
                                method="PUT"
                            )
                        )

                        with urllib.request.urlopen(
                            solicitud_put
                        ) as respuesta:

                            resultado = json.loads(
                                respuesta.read()
                                .decode("utf-8")
                            )

                        if resultado.get(
                            "content"
                        ):

                            st.success(
                                "✓ Bloque sincronizado "
                                "con GitHub."
                            )

                    except Exception as error_github:

                        st.warning(
                            "El bloque se guardó en Excel, "
                            "pero no fue posible sincronizar "
                            "con GitHub."
                        )

                        st.code(
                            str(error_github)
                        )

                    st.rerun()

                except Exception as error_guardado:

                    st.error(
                        "No fue posible guardar el bloque."
                    )

                    st.code(
                        str(error_guardado)
                    )

        else:

            st.success(
                "No hay preguntas pendientes de validación."
            )



# ========================================================
# RESTRICCIONES — GENERADOR + VALIDACIÓN + GUARDADO
# + SINCRONIZACIÓN CON GITHUB
# ========================================================

if (
    ROL_ACTUAL == "ADMINISTRADOR"
    and opcion_evaluacion == "Banco general de preguntas"
):

    st.subheader("Preguntas de Restricciones")

    RUTA_RESTRICCIONES = BASE_DIR / "BANCO_PREGUNTAS_GENERALES.xlsx"

    # ====================================================
    # FUNCIONES AUXILIARES
    # ====================================================

    def limpiar_restr(valor):
        if pd.isna(valor):
            return ""
        return str(valor).strip()

    def lista_restr(valor):
        texto = limpiar_restr(valor)

        if not texto:
            return []

        partes = []

        for elemento in texto.split(";"):
            elemento = elemento.strip()

            if elemento:
                partes.append(elemento)

        return partes

    def siguiente_id_restr(banco):

        numeros = []

        if not banco.empty and "Pregunta_ID" in banco.columns:

            for valor in banco["Pregunta_ID"]:

                texto = limpiar_restr(valor)

                if texto.startswith("RES_"):

                    try:
                        numeros.append(
                            int(texto[4:])
                        )
                    except ValueError:
                        pass

        if numeros:
            return max(numeros) + 1

        return 1

    def mezclar_opciones_restr(correcta, distractores):

        opciones = [correcta] + distractores

        opciones = list(dict.fromkeys(opciones))

        if len(opciones) != 4:
            return [], None

        np.random.shuffle(opciones)

        posicion_correcta = (
            opciones.index(correcta) + 1
        )

        return opciones, posicion_correcta

    def mezclar_nivel2_restr(
        correcta_1,
        correcta_2,
        distractor_1,
        distractor_2
    ):

        opciones = [
            correcta_1,
            correcta_2,
            distractor_1,
            distractor_2
        ]

        opciones = list(dict.fromkeys(opciones))

        if len(opciones) != 4:
            return [], []

        np.random.shuffle(opciones)

        posiciones_correctas = [
            indice + 1
            for indice, opcion in enumerate(opciones)
            if opcion in [
                correcta_1,
                correcta_2
            ]
        ]

        return (
            opciones,
            sorted(posiciones_correctas)
        )

    # ====================================================
    # CARGAR BANCO GENERAL
    # ====================================================

    if RUTA_RESTRICCIONES.exists():

        try:

            banco_restr = pd.read_excel(
                RUTA_RESTRICCIONES,
                dtype=str
            ).fillna("")

        except Exception as error_restr:

            st.error(
                "No fue posible cargar el Banco General."
            )

            st.code(str(error_restr))

            banco_restr = pd.DataFrame()

    else:

        banco_restr = pd.DataFrame(
            columns=COLUMNAS_BANCO_GENERAL
        )

    # ====================================================
    # CARGAR MATRIZ DE RESTRICCIONES
    # ====================================================

    columnas_restricciones = [

        "Restriccion_ID",
        "Producto",
        "Tipo",
        "Precaución / Contraindicación",
        "Motivo",
        "Alternativas seguras"

    ]

    faltantes_restr = [

        columna
        for columna in columnas_restricciones
        if columna not in Restricciones.columns

    ]

    if faltantes_restr:

        st.error(
            "Faltan columnas en la hoja Restricciones:"
        )

        for columna in faltantes_restr:
            st.write(f"- {columna}")

    else:

        matriz_restr = (
            Restricciones[
                columnas_restricciones
            ]
            .copy()
            .fillna("")
        )

        for columna in columnas_restricciones:

            matriz_restr[columna] = (
                matriz_restr[columna]
                .astype(str)
                .str.strip()
            )

        # =================================================
        # SOLO REGISTROS CON INFORMACIÓN REAL
        # =================================================

        matriz_restr = matriz_restr[
            (matriz_restr["Restriccion_ID"] != "")
            &
            (matriz_restr["Producto"] != "")
            &
            (matriz_restr["Tipo"] != "")
            &
            (matriz_restr["Precaución / Contraindicación"] != "")
        ].copy()

        st.info(
            f"Restricciones disponibles: "
            f"{len(matriz_restr)}"
        )

        # =================================================
        # CANTIDAD A GENERAR
        # =================================================

        col1, col2 = st.columns(2)

        with col1:

            cantidad_nivel1_restr = st.number_input(

                "Nivel 1 — cantidad máxima",

                min_value=0,
                max_value=5,
                value=5,
                step=1,

                key="cantidad_nivel1_restricciones"

            )

        with col2:

            cantidad_nivel2_restr = st.number_input(

                "Nivel 2 — cantidad máxima",

                min_value=0,
                max_value=5,
                value=5,
                step=1,

                key="cantidad_nivel2_restricciones"

            )

        # =================================================
        # GENERAR
        # =================================================

        if st.button(
            "GENERAR PREGUNTAS DE RESTRICCIONES",
            key="generar_restricciones_nuevo",
            type="primary"
        ):

            nuevas_restricciones = []

            # =================================================
            # RELACIONES YA UTILIZADAS
            #
            # IMPORTANTE:
            # APROBADAS Y RECHAZADAS NO SE PUEDEN REPETIR.
            # =================================================

            relaciones_usadas_restr = set()

            if not banco_restr.empty:

                for _, fila_banco in (
                    banco_restr.iterrows()
                ):

                    modulo = limpiar_restr(
                        fila_banco.get(
                            "Modulo",
                            ""
                        )
                    ).lower()

                    if modulo != "restricciones":
                        continue

                    estado = limpiar_restr(
                        fila_banco.get(
                            "Estado",
                            ""
                        )
                    ).upper()

                    if estado not in [
                        "PENDIENTE",
                        "APROBADA",
                        "RECHAZADA"
                    ]:
                        continue

                    fuente = limpiar_restr(
                        fila_banco.get(
                            "Fuente_ID",
                            ""
                        )
                    )

                    nivel = limpiar_restr(
                        fila_banco.get(
                            "Nivel",
                            ""
                        )
                    )

                    tipo_relacion = limpiar_restr(
                        fila_banco.get(
                            "Tipo_Relacion",
                            ""
                        )
                    )

                    if fuente and nivel and tipo_relacion:

                        relaciones_usadas_restr.add(
                            (
                                fuente,
                                nivel,
                                tipo_relacion
                            )
                        )

            # =================================================
            # CONTADORES
            # =================================================

            contador_nivel1 = 0
            contador_nivel2 = 0

            siguiente_id_restriccion = (
                siguiente_id_restr(
                    banco_restr
                )
            )

            # =================================================
            # ORDEN ALEATORIO DE LA MATRIZ
            # =================================================

            matriz_generacion = (
                matriz_restr
                .sample(
                    frac=1
                )
                .reset_index(
                    drop=True
                )
            )

            # =================================================
            # GENERAR NIVEL 1
            # =================================================

            for _, fila in matriz_generacion.iterrows():

                if contador_nivel1 >= int(
                    cantidad_nivel1_restr
                ):
                    break

                restriccion_id = limpiar_restr(
                    fila["Restriccion_ID"]
                )

                producto = limpiar_restr(
                    fila["Producto"]
                )

                tipo = limpiar_restr(
                    fila["Tipo"]
                )

                caracteristica = limpiar_restr(
                    fila[
                        "Precaución / Contraindicación"
                    ]
                )

                motivo = limpiar_restr(
                    fila["Motivo"]
                )

                alternativa = limpiar_restr(
                    fila["Alternativas seguras"]
                )

                # =============================================
                # NO USAR TEXTOS BASURA
                # =============================================

                textos_prohibidos = [

                    "no guarda relación",
                    "no corresponde",
                    "situación diferente",
                    "condición diferente",
                    "otra condición",
                    "otra patología",
                    "mostrar advertencia",
                    "no es una restricción",
                    "no se relaciona"

                ]

                def es_texto_valido_restr(texto):

                    texto_limpio = texto.lower().strip()

                    if not texto_limpio:
                        return False

                    for prohibido in textos_prohibidos:

                        if prohibido in texto_limpio:
                            return False

                    return True

                # =============================================
                # NIVEL 1
                #
                # Producto + situación
                # -> característica exacta de la matriz
                #
                # Los distractores salen únicamente de
                # OTRAS FILAS REALES de la matriz.
                # =============================================

                clave = (
                    restriccion_id,
                    "Nivel 1",
                    "Producto_Restriccion"
                )

                if clave in relaciones_usadas_restr:
                    continue

                if not (
                    es_texto_valido_restr(
                        tipo
                    )
                    and
                    es_texto_valido_restr(
                        caracteristica
                    )
                ):
                    continue

                candidatos = []

                for _, otra in matriz_generacion.iterrows():

                    otra_id = limpiar_restr(
                        otra[
                            "Restriccion_ID"
                        ]
                    )

                    if otra_id == restriccion_id:
                        continue

                    otra_caracteristica = limpiar_restr(
                        otra[
                            "Precaución / Contraindicación"
                        ]
                    )

                    if not es_texto_valido_restr(
                        otra_caracteristica
                    ):
                        continue

                    if (
                        otra_caracteristica
                        == caracteristica
                    ):
                        continue

                    candidatos.append(
                        otra_caracteristica
                    )

                candidatos = list(
                    dict.fromkeys(
                        candidatos
                    )
                )

                if len(candidatos) < 3:
                    continue

                distractores = list(
                    np.random.choice(
                        candidatos,
                        size=3,
                        replace=False
                    )
                )

                opciones, correcta = (
                    mezclar_opciones_restr(
                        caracteristica,
                        distractores
                    )
                )

                if len(opciones) != 4:
                    continue

                pregunta = (
                    f"En el producto {producto}, "
                    f"la situación indicada es: {tipo}. "
                    f"¿Cuál es la característica "
                    f"registrada para esta situación?"
                )

                nuevas_restricciones.append({

                    "Pregunta_ID":
                        f"RES_{siguiente_id_restriccion:05d}",

                    "Modulo":
                        "Restricciones",

                    "Tema":
                        producto,

                    "Nivel":
                        "Nivel 1",

                    "Tipo_Relacion":
                        "Producto_Restriccion",

                    "Pregunta":
                        pregunta,

                    "Respuesta_1":
                        opciones[0],

                    "Respuesta_2":
                        opciones[1],

                    "Respuesta_3":
                        opciones[2],

                    "Respuesta_4":
                        opciones[3],

                    "Respuesta_Correcta":
                        str(correcta),

                    "Estado":
                        "PENDIENTE",

                    "Observacion_Administrador":
                        "",

                    "Fecha_Generacion":
                        pd.Timestamp.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "Fuente_ID":
                        restriccion_id

                })

                relaciones_usadas_restr.add(
                    clave
                )

                siguiente_id_restriccion += 1
                contador_nivel1 += 1

            # =================================================
            # GENERAR NIVEL 2
            #
            # Producto + situación + motivo
            #
            # Las DOS respuestas correctas son información
            # REAL de la misma fila.
            #
            # Los distractores también salen de la matriz.
            # =================================================

            for _, fila in matriz_generacion.iterrows():

                if contador_nivel2 >= int(
                    cantidad_nivel2_restr
                ):
                    break

                restriccion_id = limpiar_restr(
                    fila["Restriccion_ID"]
                )

                producto = limpiar_restr(
                    fila["Producto"]
                )

                tipo = limpiar_restr(
                    fila["Tipo"]
                )

                caracteristica = limpiar_restr(
                    fila[
                        "Precaución / Contraindicación"
                    ]
                )

                motivo = limpiar_restr(
                    fila["Motivo"]
                )

                if not (
                    es_texto_valido_restr(
                        tipo
                    )
                    and
                    es_texto_valido_restr(
                        caracteristica
                    )
                    and
                    es_texto_valido_restr(
                        motivo
                    )
                ):
                    continue

                clave = (
                    restriccion_id,
                    "Nivel 2",
                    "Producto_Restriccion_Motivo"
                )

                if clave in relaciones_usadas_restr:
                    continue

                # =============================================
                # CANDIDATOS REALES DE OTRAS FILAS
                # =============================================

                candidatos_caracteristica = []
                candidatos_motivo = []

                for _, otra in matriz_generacion.iterrows():

                    otra_id = limpiar_restr(
                        otra[
                            "Restriccion_ID"
                        ]
                    )

                    if otra_id == restriccion_id:
                        continue

                    otra_caracteristica = limpiar_restr(
                        otra[
                            "Precaución / Contraindicación"
                        ]
                    )

                    otro_motivo = limpiar_restr(
                        otra[
                            "Motivo"
                        ]
                    )

                    if es_texto_valido_restr(
                        otra_caracteristica
                    ):

                        if (
                            otra_caracteristica
                            != caracteristica
                        ):

                            candidatos_caracteristica.append(
                                otra_caracteristica
                            )

                    if es_texto_valido_restr(
                        otro_motivo
                    ):

                        if otro_motivo != motivo:

                            candidatos_motivo.append(
                                otro_motivo
                            )

                candidatos_caracteristica = list(
                    dict.fromkeys(
                        candidatos_caracteristica
                    )
                )

                candidatos_motivo = list(
                    dict.fromkeys(
                        candidatos_motivo
                    )
                )

                if (
                    len(
                        candidatos_caracteristica
                    ) < 1
                    or
                    len(
                        candidatos_motivo
                    ) < 1
                ):
                    continue

                distractor_1 = (
                    np.random.choice(
                        candidatos_caracteristica
                    )
                )

                distractor_2 = (
                    np.random.choice(
                        candidatos_motivo
                    )
                )

                correcta_1 = (
                    f"La característica registrada "
                    f"es: {caracteristica}."
                )

                correcta_2 = (
                    f"El motivo registrado es: "
                    f"{motivo}."
                )

                distractor_1 = (
                    f"La característica registrada "
                    f"es: {distractor_1}."
                )

                distractor_2 = (
                    f"El motivo registrado es: "
                    f"{distractor_2}."
                )

                opciones, correctas = (
                    mezclar_nivel2_restr(
                        correcta_1,
                        correcta_2,
                        distractor_1,
                        distractor_2
                    )
                )

                if (
                    len(opciones) != 4
                    or
                    len(correctas) != 2
                ):
                    continue

                pregunta = (
                    f"En relación con el producto "
                    f"{producto}, considere la situación "
                    f"{tipo}. Seleccione las DOS afirmaciones "
                    f"que corresponden a la información "
                    f"registrada."
                )

                nuevas_restricciones.append({

                    "Pregunta_ID":
                        f"RES_{siguiente_id_restriccion:05d}",

                    "Modulo":
                        "Restricciones",

                    "Tema":
                        producto,

                    "Nivel":
                        "Nivel 2",

                    "Tipo_Relacion":
                        "Producto_Restriccion_Motivo",

                    "Pregunta":
                        pregunta,

                    "Respuesta_1":
                        opciones[0],

                    "Respuesta_2":
                        opciones[1],

                    "Respuesta_3":
                        opciones[2],

                    "Respuesta_4":
                        opciones[3],

                    "Respuesta_Correcta":
                        ",".join(
                            map(
                                str,
                                correctas
                            )
                        ),

                    "Estado":
                        "PENDIENTE",

                    "Observacion_Administrador":
                        "",

                    "Fecha_Generacion":
                        pd.Timestamp.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "Fuente_ID":
                        restriccion_id

                })

                relaciones_usadas_restr.add(
                    clave
                )

                siguiente_id_restriccion += 1
                contador_nivel2 += 1

            # =================================================
            # GUARDAR GENERACIÓN
            # =================================================

            if nuevas_restricciones:

                nuevas_df_restr = pd.DataFrame(
                    nuevas_restricciones
                )

                banco_restr = pd.concat(
                    [
                        banco_restr,
                        nuevas_df_restr
                    ],
                    ignore_index=True
                )

                banco_restr = banco_restr[
                    COLUMNAS_BANCO_GENERAL
                ]

                banco_restr.to_excel(
                    RUTA_RESTRICCIONES,
                    index=False,
                    sheet_name="Banco_General"
                )

                st.success(
                    f"Se generaron "
                    f"{len(nuevas_df_restr)} preguntas."
                )

                # =============================================
                # SINCRONIZAR GENERACIÓN CON GITHUB
                # =============================================

                try:

                    with open(
                        RUTA_RESTRICCIONES,
                        "rb"
                    ) as archivo:

                        contenido_restr = (
                            archivo.read()
                        )

                    contenido_base64_restr = (
                        base64.b64encode(
                            contenido_restr
                        ).decode(
                            "utf-8"
                        )
                    )

                    url_github_restr = (
                        f"https://api.github.com/repos/"
                        f"{GITHUB_USUARIO}/"
                        f"{GITHUB_REPOSITORIO}/"
                        f"contents/"
                        f"BANCO_PREGUNTAS_GENERALES.xlsx"
                    )

                    headers_github_restr = {

                        "Authorization":
                            f"Bearer {GITHUB_TOKEN}",

                        "Accept":
                            "application/vnd.github+json",

                        "X-GitHub-Api-Version":
                            "2022-11-28"

                    }

                    solicitud_get_restr = (
                        urllib.request.Request(
                            url_github_restr,
                            headers=headers_github_restr,
                            method="GET"
                        )
                    )

                    with urllib.request.urlopen(
                        solicitud_get_restr
                    ) as respuesta:

                        informacion_restr = json.loads(
                            respuesta.read()
                            .decode("utf-8")
                        )

                    datos_restr = {

                        "message":
                            "Generar preguntas "
                            "de restricciones",

                        "content":
                            contenido_base64_restr,

                        "sha":
                            informacion_restr["sha"]

                    }

                    solicitud_put_restr = (
                        urllib.request.Request(
                            url_github_restr,
                            data=json.dumps(
                                datos_restr
                            ).encode("utf-8"),
                            headers={
                                **headers_github_restr,
                                "Content-Type":
                                    "application/json"
                            },
                            method="PUT"
                        )
                    )

                    with urllib.request.urlopen(
                        solicitud_put_restr
                    ) as respuesta:

                        resultado_restr = json.loads(
                            respuesta.read()
                            .decode("utf-8")
                        )

                    if resultado_restr.get("content"):

                        st.success(
                            "✓ Generación sincronizada "
                            "con GitHub."
                        )

                except Exception as error_github_restr:

                    st.warning(
                        "Las preguntas se guardaron en Excel, "
                        "pero no fue posible sincronizarlas "
                        "con GitHub."
                    )

                    st.code(
                        str(error_github_restr)
                    )

            else:

                st.warning(
                    "No hay suficientes relaciones nuevas "
                    "válidas para generar más preguntas."
                )

    # ====================================================
    # VALIDACIÓN
    # ====================================================

    st.divider()

    st.subheader(
        "Validación de Restricciones"
    )

    if RUTA_RESTRICCIONES.exists():

        banco_validacion_restr = pd.read_excel(
            RUTA_RESTRICCIONES,
            dtype=str
        ).fillna("")

        pendientes_restr = (
            banco_validacion_restr[
                (
                    banco_validacion_restr[
                        "Modulo"
                    ]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    ==
                    "restricciones"
                )
                &
                (
                    banco_validacion_restr[
                        "Estado"
                    ]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    ==
                    "PENDIENTE"
                )
            ]
            .copy()
        )

        st.write(
            f"Preguntas de Restricciones pendientes: "
            f"**{len(pendientes_restr)}**"
        )

        if not pendientes_restr.empty:

            cantidad_bloque_restr = st.number_input(

                "Cantidad de preguntas por bloque",

                min_value=1,
                max_value=min(
                    50,
                    len(pendientes_restr)
                ),

                value=min(
                    5,
                    len(pendientes_restr)
                ),

                step=1,

                key="cantidad_bloque_restricciones_nuevo"

            )

            pendientes_restr = (
                pendientes_restr
                .reset_index(
                    drop=False
                )
            )

            cantidad_bloque_restr = int(
                cantidad_bloque_restr
            )

            total_bloques_restr = (
                (
                    len(pendientes_restr)
                    +
                    cantidad_bloque_restr
                    -
                    1
                )
                //
                cantidad_bloque_restr
            )

            bloque_actual_restr = st.number_input(

                "Bloque",

                min_value=1,
                max_value=max(
                    1,
                    total_bloques_restr
                ),

                value=1,
                step=1,

                key="bloque_restricciones_nuevo"

            )

            inicio_restr = (
                (
                    int(
                        bloque_actual_restr
                    )
                    - 1
                )
                *
                cantidad_bloque_restr
            )

            fin_restr = (
                inicio_restr
                +
                cantidad_bloque_restr
            )

            bloque_restr = (
                pendientes_restr
                .iloc[
                    inicio_restr:fin_restr
                ]
                .copy()
            )

            st.info(
                f"Bloque "
                f"{int(bloque_actual_restr)} "
                f"de {total_bloques_restr} — "
                f"{len(bloque_restr)} preguntas."
            )

            # =================================================
            # DECISIÓN DEL BLOQUE
            # =================================================

            decision_restr = st.radio(

                "Acción para este bloque:",

                [
                    "Validar individualmente",
                    "Aprobar todas",
                    "Rechazar todas"
                ],

                horizontal=True,

                key="decision_restricciones_nuevo"

            )

            cambios_restr = {}

            # =================================================
            # MOSTRAR PREGUNTAS
            # =================================================

            for numero_restr, (
                indice_original_restr,
                fila_restr
            ) in enumerate(

                bloque_restr.iterrows(),
                start=1

            ):

                pregunta_id_restr = limpiar_restr(
                    fila_restr[
                        "Pregunta_ID"
                    ]
                )

                st.markdown(
                    f"### Pregunta {numero_restr} "
                    f"— {pregunta_id_restr}"
                )

                st.caption(
                    f"Producto: "
                    f"{fila_restr['Tema']} | "
                    f"Nivel: "
                    f"{fila_restr['Nivel']} | "
                    f"Relación: "
                    f"{fila_restr['Tipo_Relacion']}"
                )

                st.write(
                    f"**{fila_restr['Pregunta']}**"
                )

                st.write(
                    f"1. {fila_restr['Respuesta_1']}"
                )

                st.write(
                    f"2. {fila_restr['Respuesta_2']}"
                )

                st.write(
                    f"3. {fila_restr['Respuesta_3']}"
                )

                st.write(
                    f"4. {fila_restr['Respuesta_4']}"
                )

                st.caption(
                    "Respuesta correcta registrada: "
                    f"{fila_restr['Respuesta_Correcta']}"
                )

                if decision_restr == "Aprobar todas":

                    estado_restr = "APROBADA"

                elif decision_restr == "Rechazar todas":

                    estado_restr = "RECHAZADA"

                else:

                    estado_restr = st.selectbox(

                        "Estado",

                        [
                            "PENDIENTE",
                            "APROBADA",
                            "RECHAZADA"
                        ],

                        index=0,

                        key=(
                            "estado_restr_"
                            f"{pregunta_id_restr}"
                        )

                    )

                observacion_restr = st.text_area(

                    "Observación",

                    value=limpiar_restr(
                        fila_restr[
                            "Observacion_Administrador"
                        ]
                    ),

                    key=(
                        "obs_restr_"
                        f"{pregunta_id_restr}"
                    ),

                    height=60

                )

                cambios_restr[
                    pregunta_id_restr
                ] = {

                    "Estado":
                        estado_restr,

                    "Observacion":
                        observacion_restr,

                    "indice":
                        indice_original_restr

                }

                st.divider()

            # =================================================
            # GUARDAR Y SINCRONIZAR
            # =================================================

            if st.button(
                "GUARDAR Y SINCRONIZAR BLOQUE",
                key="guardar_sincronizar_restricciones_nuevo",
                type="primary"
            ):

                try:

                    banco_guardar_restr = pd.read_excel(
                        RUTA_RESTRICCIONES,
                        dtype=str
                    ).fillna("")

                    for (
                        pregunta_id_restr,
                        cambios_restr_item
                    ) in cambios_restr.items():

                        posicion_restr = (
                            banco_guardar_restr[
                                "Pregunta_ID"
                            ]
                            .astype(str)
                            .str.strip()
                            ==
                            pregunta_id_restr
                        )

                        banco_guardar_restr.loc[
                            posicion_restr,
                            "Estado"
                        ] = cambios_restr_item[
                            "Estado"
                        ]

                        banco_guardar_restr.loc[
                            posicion_restr,
                            "Observacion_Administrador"
                        ] = cambios_restr_item[
                            "Observacion"
                        ]

                    banco_guardar_restr.to_excel(
                        RUTA_RESTRICCIONES,
                        index=False,
                        sheet_name="Banco_General"
                    )

                    st.success(
                        "✓ Validación guardada en Excel."
                    )

                    # =========================================
                    # SINCRONIZAR CON GITHUB
                    # =========================================

                    with open(
                        RUTA_RESTRICCIONES,
                        "rb"
                    ) as archivo:

                        contenido_restr = (
                            archivo.read()
                        )

                    contenido_base64_restr = (
                        base64.b64encode(
                            contenido_restr
                        ).decode(
                            "utf-8"
                        )
                    )

                    url_github_restr = (
                        f"https://api.github.com/repos/"
                        f"{GITHUB_USUARIO}/"
                        f"{GITHUB_REPOSITORIO}/"
                        f"contents/"
                        f"BANCO_PREGUNTAS_GENERALES.xlsx"
                    )

                    headers_github_restr = {

                        "Authorization":
                            f"Bearer {GITHUB_TOKEN}",

                        "Accept":
                            "application/vnd.github+json",

                        "X-GitHub-Api-Version":
                            "2022-11-28"

                    }

                    solicitud_get_restr = (
                        urllib.request.Request(
                            url_github_restr,
                            headers=headers_github_restr,
                            method="GET"
                        )
                    )

                    with urllib.request.urlopen(
                        solicitud_get_restr
                    ) as respuesta:

                        informacion_restr = json.loads(
                            respuesta.read()
                            .decode("utf-8")
                        )

                    datos_restr = {

                        "message":
                            "Validar preguntas "
                            "de restricciones",

                        "content":
                            contenido_base64_restr,

                        "sha":
                            informacion_restr["sha"]

                    }

                    solicitud_put_restr = (
                        urllib.request.Request(
                            url_github_restr,
                            data=json.dumps(
                                datos_restr
                            ).encode("utf-8"),
                            headers={
                                **headers_github_restr,
                                "Content-Type":
                                    "application/json"
                            },
                            method="PUT"
                        )
                    )

                    with urllib.request.urlopen(
                        solicitud_put_restr
                    ) as respuesta:

                        resultado_restr = json.loads(
                            respuesta.read()
                            .decode("utf-8")
                        )

                    if resultado_restr.get("content"):

                        st.success(
                            "✓ Validación sincronizada "
                            "correctamente con GitHub."
                        )

                    st.rerun()

                except Exception as error_guardado_restr:

                    st.error(
                        "No fue posible guardar o "
                        "sincronizar las preguntas."
                    )

                    st.code(
                        str(error_guardado_restr)
                    )

        else:

            st.success(
                "No hay preguntas de Restricciones "
                "pendientes de validación."
            )

# ========================================================

# ========================================================
# ========================================================
# 7.5 BANCO GENERAL — PRODUCTO + ACCIÓN GENERAL
# ========================================================
#
# 7.5.1 VERIFICACIÓN Y CARGA DEL BANCO GENERAL
#
# SOLO ADMINISTRADOR
# ========================================================

if (
    ROL_ACTUAL == "ADMINISTRADOR"
    and globals().get(
        "opcion_evaluacion"
    ) == "Banco general de preguntas"
):

    import re
    import base64
    import json
    import urllib.request

    st.subheader(
        "Banco General de Preguntas — Producto + Acción General"
    )

    # ====================================================
    # 7.5.1 VERIFICACIÓN Y CARGA DEL BANCO GENERAL
    # ====================================================

    if not RUTA_BANCO_GENERAL.exists():

        st.error(
            "No existe el archivo "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.stop()

    try:

        banco_general_75 = pd.read_excel(
            RUTA_BANCO_GENERAL,
            dtype=str
        ).fillna("")

    except Exception as error_banco_75:

        st.error(
            "No fue posible cargar el Banco General."
        )

        st.code(
            str(error_banco_75)
        )

        st.stop()

    st.success(
        "7.5.1 ✓ Banco General cargado correctamente."
    )
    # ====================================================
    # 7.5.2 FUNCIONES BÁSICAS DE LIMPIEZA Y NORMALIZACIÓN
    # ====================================================

    def limpiar_texto_75(valor):

        if pd.isna(valor):
            return ""

        return " ".join(
            str(valor)
            .replace("\r", " ")
            .replace("\n", " ")
            .replace("\t", " ")
            .split()
        ).strip()


    def firma_texto_75(valor):

        texto_75 = limpiar_texto_75(
            valor
        )

        return texto_75.casefold().strip(
            " .;,:"
        )


    def separar_lista_75(valor):

        texto_75 = limpiar_texto_75(
            valor
        )

        if not texto_75:
            return []

        resultado_75 = []

        for elemento_75 in texto_75.split(";"):

            elemento_75 = limpiar_texto_75(
                elemento_75
            )

            if elemento_75:
                resultado_75.append(
                    elemento_75
                )

        return resultado_75


    def quitar_duplicados_75(elementos_75):

        resultado_75 = []
        vistos_75 = set()

        for elemento_75 in elementos_75:

            firma_75 = firma_texto_75(
                elemento_75
            )

            if (
                firma_75
                and firma_75 not in vistos_75
            ):

                vistos_75.add(
                    firma_75
                )

                resultado_75.append(
                    limpiar_texto_75(
                        elemento_75
                    )
                )

        return resultado_75


    st.success(
        "7.5.2 ✓ Funciones de limpieza preparadas."
    )  

    # ====================================================
    # 7.5.3 DATAFRAME NUEVO — NORMALIZACIÓN Y ABSTRACCIÓN
    # NO MODIFICA LA MATRIZ ORIGINAL
    # ====================================================

    registros_normalizados_75 = []

    for _, fila_75 in fuente_75.iterrows():

        producto_75 = limpiar_texto_75(
            fila_75["Producto"]
        )

        if not producto_75:
            continue

        acciones_originales_75 = (
            separar_lista_75(
                fila_75["Acciones generales"]
            )
        )

        componentes_75 = []

        if "Componentes" in fuente_75.columns:

            componentes_75 = separar_lista_75(
                fila_75["Componentes"]
            )

        for accion_original_75 in acciones_originales_75:

            accion_original_75 = limpiar_texto_75(
                accion_original_75
            )

            if not accion_original_75:
                continue

            # --------------------------------------------
            # LA ABSTRACCIÓN SE HACE SOBRE UNA COPIA
            # --------------------------------------------

            accion_abstraida_75 = (
                abstraer_accion_75(
                    accion_original_75,
                    componentes_75
                )
            )

            if isinstance(
                accion_abstraida_75,
                str
            ):

                acciones_resultado_75 = [
                    accion_abstraida_75
                ]

            else:

                acciones_resultado_75 = (
                    accion_abstraida_75
                )

            for accion_general_75 in (
                acciones_resultado_75
            ):

                accion_general_75 = (
                    limpiar_texto_75(
                        accion_general_75
                    )
                )

                if not accion_general_75:
                    continue

                firma_75 = (
                    firma_texto_75(
                        producto_75
                    )
                    + "||"
                    + firma_texto_75(
                        accion_general_75
                    )
                )

                registros_normalizados_75.append({

                    "Producto":
                        producto_75,

                    "Accion_Original":
                        accion_original_75,

                    "Accion_General":
                        accion_general_75,

                    "Firma":
                        firma_75
                })


    # ====================================================
    # DATAFRAME NUEVO
    # ====================================================

    dataframe_normalizado_75 = pd.DataFrame(
        registros_normalizados_75,
        columns=[
            "Producto",
            "Accion_Original",
            "Accion_General",
            "Firma"
        ]
    )


    # ====================================================
    # ELIMINAR DUPLICADOS SOLO DEL DATAFRAME NUEVO
    # ====================================================

    if not dataframe_normalizado_75.empty:

        dataframe_normalizado_75 = (
            dataframe_normalizado_75
            .drop_duplicates(
                subset=["Firma"]
            )
            .reset_index(
                drop=True
            )
        )


    st.success(
        "7.5.3 ✓ DataFrame nuevo de "
        "normalización y abstracción creado. "
        f"Relaciones: "
        f"{len(dataframe_normalizado_75)}"
    )


    # ====================================================
    # MUESTRA PARA VERIFICAR LA ABSTRACCIÓN
    # ====================================================

    if not dataframe_normalizado_75.empty:

        st.subheader(
            "Ejemplo de normalización y abstracción"
        )

        st.dataframe(
            dataframe_normalizado_75[
                [
                    "Producto",
                    "Accion_Original",
                    "Accion_General"
                ]
            ].head(10),
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "La matriz original permanece intacta. "
            "La normalización y abstracción se realizan "
            "exclusivamente en el DataFrame nuevo."
        )


    # ====================================================
    # 7.5.4 VERIFICACIÓN
    # ====================================================

    if dataframe_normalizado_75.empty:

        st.warning(
            "7.5.4 — No existen relaciones disponibles."
        )

    else:

        productos_nivel_1_75 = (
            dataframe_normalizado_75[
                "Producto"
            ].nunique()
        )

        productos_nivel_2_75 = (
            dataframe_normalizado_75
            .groupby("Producto")
            .size()
            .ge(2)
            .sum()
        )

        st.info(
            f"Productos aptos para Nivel 1: "
            f"{productos_nivel_1_75}"
        )

        st.info(
            f"Productos aptos para Nivel 2: "
            f"{productos_nivel_2_75}"
        )

        st.success(
            "7.5.4 ✓ DataFrame normalizado verificado."
        )


    # ====================================================
    # 7.5.5 RELACIONES YA REGISTRADAS
    # ====================================================

    relaciones_bloqueadas_75 = set()

    if not banco_general_75.empty:

        for _, fila_75 in banco_general_75.iterrows():

            estado_75 = firma_texto_75(
                fila_75.get(
                    "Estado",
                    ""
                )
            ).upper()

            if estado_75 not in {
                "PENDIENTE",
                "APROBADA",
                "RECHAZADA"
            }:
                continue

            nivel_75 = limpiar_texto_75(
                fila_75.get(
                    "Nivel",
                    ""
                )
            )

            fuente_75 = limpiar_texto_75(
                fila_75.get(
                    "Fuente_ID",
                    ""
                )
            )

            if nivel_75 and fuente_75:

                relaciones_bloqueadas_75.add(
                    (
                        nivel_75.casefold(),
                        fuente_75.casefold()
                    )
                )


    st.success(
        "7.5.5 ✓ Relaciones existentes identificadas. "
        f"Bloqueadas: "
        f"{len(relaciones_bloqueadas_75)}"
    )


    # ====================================================
    # 7.5.5A MUESTRA DE NORMALIZACIÓN
    # ====================================================

    st.subheader(
        "Ejemplo de normalización y abstracción"
    )

    if dataframe_normalizado_75.empty:

        st.warning(
            "No hay registros normalizados."
        )

    else:

        muestra_normalizada_75 = (
            dataframe_normalizado_75[
                [
                    "Producto",
                    "Accion_Original",
                    "Accion_General"
                ]
            ]
            .head(10)
        )

        st.dataframe(
            muestra_normalizada_75,
            use_container_width=True,
            hide_index=True
        )

        st.success(
            "7.5.5A ✓ Muestra disponible."
        )


    # ====================================================
    # 7.5.6 SELECCIÓN DEL NIVEL
    # ====================================================

    st.divider()

    modo_generacion_75 = st.radio(
        "¿Qué desea generar?",
        [
            "Nivel 1",
            "Nivel 2",
            "Nivel 1 y Nivel 2"
        ],
        horizontal=True,
        key="modo_generacion_75"
    )


    # ====================================================
    # 7.5.7 CANTIDAD MÁXIMA
    # ====================================================

    cantidad_maxima_75 = st.number_input(
        "Cantidad máxima de preguntas",
        min_value=1,
        max_value=100,
        value=10,
        step=1,
        key="cantidad_maxima_75"
    )

    st.caption(
        "Es un máximo global. El sistema no obliga "
        "a ningún producto a generar una cantidad "
        "determinada de preguntas."
    )


    # ====================================================
    # 7.5.8 PREPARAR CANDIDATOS
    # ====================================================

    candidatos_75 = []

    acciones_por_producto_75 = (
        dataframe_normalizado_75
        .groupby("Producto")[
            "Accion_General"
        ]
        .apply(
            lambda valores: list(
                dict.fromkeys(
                    valores
                )
            )
        )
        .to_dict()
    )

    if modo_generacion_75 == "Nivel 1":

        niveles_75 = [
            "Nivel 1"
        ]

    elif modo_generacion_75 == "Nivel 2":

        niveles_75 = [
            "Nivel 2"
        ]

    else:

        niveles_75 = [
            "Nivel 1",
            "Nivel 2"
        ]


    for producto_75, acciones_75 in (
        acciones_por_producto_75.items()
    ):

        # -----------------------------------------------
        # NIVEL 1
        # UNA ACCIÓN GENERAL REAL
        # -----------------------------------------------

        if "Nivel 1" in niveles_75:

            for accion_75 in acciones_75:

                firma_75 = (
                    firma_texto_75(
                        producto_75
                    )
                    + "||"
                    + firma_texto_75(
                        accion_75
                    )
                )

                fuente_75 = (
                    "N1||"
                    + firma_75
                )

                relacion_75 = (
                    "Nivel 1".casefold(),
                    fuente_75.casefold()
                )

                if relacion_75 in (
                    relaciones_bloqueadas_75
                ):
                    continue

                candidatos_75.append({

                    "Producto":
                        producto_75,

                    "Accion_1":
                        accion_75,

                    "Accion_2":
                        "",

                    "Nivel":
                        "Nivel 1",

                    "Fuente_ID":
                        fuente_75
                })


        # -----------------------------------------------
        # NIVEL 2
        # DOS ACCIONES GENERALES DEL MISMO PRODUCTO
        # -----------------------------------------------

        if (
            "Nivel 2" in niveles_75
            and len(acciones_75) >= 2
        ):

            for indice_1_75 in range(
                len(acciones_75)
            ):

                for indice_2_75 in range(
                    indice_1_75 + 1,
                    len(acciones_75)
                ):

                    accion_1_75 = (
                        acciones_75[
                            indice_1_75
                        ]
                    )

                    accion_2_75 = (
                        acciones_75[
                            indice_2_75
                        ]
                    )

                    firma_75 = (
                        firma_texto_75(
                            producto_75
                        )
                        + "||"
                        + firma_texto_75(
                            accion_1_75
                        )
                        + "||"
                        + firma_texto_75(
                            accion_2_75
                        )
                    )

                    fuente_75 = (
                        "N2||"
                        + firma_75
                    )

                    relacion_75 = (
                        "Nivel 2".casefold(),
                        fuente_75.casefold()
                    )

                    if relacion_75 in (
                        relaciones_bloqueadas_75
                    ):
                        continue

                    candidatos_75.append({

                        "Producto":
                            producto_75,

                        "Accion_1":
                            accion_1_75,

                        "Accion_2":
                            accion_2_75,

                        "Nivel":
                            "Nivel 2",

                        "Fuente_ID":
                            fuente_75
                    })


    st.info(
        f"Candidatos disponibles: "
        f"{len(candidatos_75)}"
    )


    # ====================================================
    # 7.5.8 GENERAR
    # ====================================================

    generar_preguntas_75 = st.button(
        "GENERAR PREGUNTAS",
        type="primary",
        key="generar_preguntas_75"
    )


    if generar_preguntas_75:

        preguntas_generadas_75 = []

        candidatos_trabajo_75 = (
            candidatos_75.copy()
        )

        np.random.shuffle(
            candidatos_trabajo_75
        )

        # =================================================
        # FUNCIONES PARA DISTRACTORES
        # =================================================

        todas_acciones_75 = (
            dataframe_normalizado_75[
                "Accion_General"
            ]
            .dropna()
            .astype(str)
            .map(
                limpiar_texto_75
            )
            .tolist()
        )

        todas_acciones_75 = list(
            dict.fromkeys(
                todas_acciones_75
            )
        )


        for candidato_75 in (
            candidatos_trabajo_75
        ):

            if (
                len(preguntas_generadas_75)
                >= int(
                    cantidad_maxima_75
                )
            ):
                break


            producto_75 = candidato_75[
                "Producto"
            ]

            accion_1_75 = candidato_75[
                "Accion_1"
            ]

            accion_2_75 = candidato_75[
                "Accion_2"
            ]

            nivel_75 = candidato_75[
                "Nivel"
            ]


            # ---------------------------------------------
            # NIVEL 1
            # ---------------------------------------------

            if nivel_75 == "Nivel 1":

                distractores_75 = [
                    accion_75
                    for accion_75
                    in todas_acciones_75
                    if (
                        firma_texto_75(
                            accion_75
                        )
                        !=
                        firma_texto_75(
                            accion_1_75
                        )
                    )
                ]

                if len(
                    distractores_75
                ) < 3:

                    continue

                np.random.shuffle(
                    distractores_75
                )

                opciones_75 = [
                    accion_1_75,
                    distractores_75[0],
                    distractores_75[1],
                    distractores_75[2]
                ]

                np.random.shuffle(
                    opciones_75
                )

                posiciones_correctas_75 = [
                    str(
                        opciones_75.index(
                            accion_1_75
                        ) + 1
                    )
                ]

                pregunta_75 = (
                    f"¿Cuál de las siguientes "
                    f"acciones generales corresponde "
                    f"al producto {producto_75}?"
                )


            # ---------------------------------------------
            # NIVEL 2
            # DOS ACCIONES CORRECTAS
            # ---------------------------------------------

            else:

                distractores_75 = [
                    accion_75
                    for accion_75
                    in todas_acciones_75
                    if (
                        firma_texto_75(
                            accion_75
                        )
                        not in {
                            firma_texto_75(
                                accion_1_75
                            ),
                            firma_texto_75(
                                accion_2_75
                            )
                        }
                    )
                ]

                if len(
                    distractores_75
                ) < 2:

                    continue

                np.random.shuffle(
                    distractores_75
                )

                opciones_75 = [
                    accion_1_75,
                    accion_2_75,
                    distractores_75[0],
                    distractores_75[1]
                ]

                np.random.shuffle(
                    opciones_75
                )

                posiciones_correctas_75 = [

                    str(
                        indice_75 + 1
                    )

                    for indice_75,
                    opcion_75
                    in enumerate(
                        opciones_75
                    )

                    if (
                        firma_texto_75(
                            opcion_75
                        )
                        in {
                            firma_texto_75(
                                accion_1_75
                            ),
                            firma_texto_75(
                                accion_2_75
                            )
                        }
                    )
                ]

                pregunta_75 = (
                    f"¿Cuáles de las siguientes "
                    f"acciones generales corresponden "
                    f"al producto {producto_75}? "
                    f"Seleccione las dos opciones correctas."
                )


            preguntas_generadas_75.append({

                "Pregunta_ID":
                    None,

                "Modulo":
                    "Productos",

                "Tema":
                    producto_75,

                "Nivel":
                    nivel_75,

                "Tipo_Relacion":
                    (
                        "Producto_AccionGeneral"
                        if nivel_75 == "Nivel 1"
                        else
                        "Producto_AccionGeneral_Nivel2"
                    ),

                "Pregunta":
                    pregunta_75,

                "Respuesta_1":
                    opciones_75[0],

                "Respuesta_2":
                    opciones_75[1],

                "Respuesta_3":
                    opciones_75[2],

                "Respuesta_4":
                    opciones_75[3],

                "Respuesta_Correcta":
                    ",".join(
                        posiciones_correctas_75
                    ),

                "Estado":
                    "PENDIENTE",

                "Observacion_Administrador":
                    "",

                "Fecha_Generacion":
                    pd.Timestamp.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),

                "Fuente_ID":
                    candidato_75[
                        "Fuente_ID"
                    ]
            })


        st.success(
            "7.5.9 ✓ Preguntas generadas: "
            f"{len(preguntas_generadas_75)}"
        )


        if preguntas_generadas_75:

            dataframe_preguntas_75 = (
                pd.DataFrame(
                    preguntas_generadas_75
                )
            )

            st.dataframe(
                dataframe_preguntas_75,
                use_container_width=True,
                hide_index=True
            )
    # ====================================================
    # 7.5.10 VALIDACIÓN DE PREGUNTAS GENERADAS
    # ====================================================

    if (
        generar_preguntas_75
        and preguntas_generadas_75
    ):

        st.divider()

        st.subheader(
            "Validación de preguntas generadas"
        )

        st.write(
            "Revise cada pregunta y seleccione su estado."
        )

        validar_bloque_75 = st.radio(
            "Acción para este bloque:",
            [
                "Validar individualmente",
                "Aprobar todo el bloque",
                "Rechazar todo el bloque"
            ],
            horizontal=True,
            key="validar_bloque_75"
        )

        if validar_bloque_75 == (
            "Aprobar todo el bloque"
        ):

            for pregunta_75 in (
                preguntas_generadas_75
            ):

                pregunta_75[
                    "Estado"
                ] = "APROBADA"

            st.success(
                "7.5.10 ✓ Todo el bloque quedó "
                "marcado como APROBADA."
            )

        elif validar_bloque_75 == (
            "Rechazar todo el bloque"
        ):

            for pregunta_75 in (
                preguntas_generadas_75
            ):

                pregunta_75[
                    "Estado"
                ] = "RECHAZADA"

            st.warning(
                "7.5.10 ✓ Todo el bloque quedó "
                "marcado como RECHAZADA."
            )

        else:

            for indice_75, pregunta_75 in enumerate(
                preguntas_generadas_75
            ):

                st.markdown(
                    f"**Pregunta {indice_75 + 1} — "
                    f"{pregunta_75['Nivel']} — "
                    f"{pregunta_75['Tema']}**"
                )

                st.write(
                    pregunta_75["Pregunta"]
                )

                st.write(
                    f"1. {pregunta_75['Respuesta_1']}"
                )

                st.write(
                    f"2. {pregunta_75['Respuesta_2']}"
                )

                st.write(
                    f"3. {pregunta_75['Respuesta_3']}"
                )

                st.write(
                    f"4. {pregunta_75['Respuesta_4']}"
                )

                estado_75 = st.selectbox(
                    "Estado:",
                    [
                        "PENDIENTE",
                        "APROBADA",
                        "RECHAZADA"
                    ],
                    key=f"estado_75_{indice_75}"
                )

                pregunta_75[
                    "Estado"
                ] = estado_75

                st.divider()

        st.success(
            "7.5.10 ✓ Validación preparada."
        )

    # ====================================================
    # 7.5.11 GUARDAR VALIDACIÓN EN EL BANCO GENERAL
    # ====================================================

    if (
        generar_preguntas_75
        and preguntas_generadas_75
    ):

        try:

            banco_75 = pd.read_excel(
                RUTA_BANCO_GENERAL,
                dtype=str
            ).fillna("")

            columnas_banco_75 = [
                "Pregunta_ID",
                "Modulo",
                "Tema",
                "Nivel",
                "Tipo_Relacion",
                "Pregunta",
                "Respuesta_1",
                "Respuesta_2",
                "Respuesta_3",
                "Respuesta_4",
                "Respuesta_Correcta",
                "Estado",
                "Observacion_Administrador",
                "Fecha_Generacion",
                "Fuente_ID"
            ]

            for columna_75 in columnas_banco_75:

                if columna_75 not in banco_75.columns:

                    banco_75[columna_75] = ""

            banco_75 = banco_75[
                columnas_banco_75
            ]

            nuevas_75 = pd.DataFrame(
                preguntas_generadas_75
            )

            numeros_75 = []

            for valor_75 in banco_75[
                "Pregunta_ID"
            ]:

                texto_75 = str(
                    valor_75
                ).strip()

                if texto_75.startswith("PROD_"):

                    try:

                        numeros_75.append(
                            int(texto_75[5:])
                        )

                    except ValueError:

                        pass

            siguiente_id_75 = (
                max(numeros_75) + 1
                if numeros_75
                else 1
            )

            for indice_75 in nuevas_75.index:

                nuevas_75.at[
                    indice_75,
                    "Pregunta_ID"
                ] = (
                    f"PROD_{siguiente_id_75:05d}"
                )

                siguiente_id_75 += 1

            banco_75 = pd.concat(
                [
                    banco_75,
                    nuevas_75
                ],
                ignore_index=True
            )

            banco_75.to_excel(
                RUTA_BANCO_GENERAL,
                index=False
            )

            st.success(
                "7.5.11 ✓ Validación guardada "
                "en el Banco General."
            )

        except Exception as error_75:

            st.error(
                "7.5.11 ✗ Error al guardar "
                "el Banco General."
            )

            st.code(
                str(error_75)
            )

    # ====================================================
    # 7.5.12 SINCRONIZACIÓN FINAL
    # ====================================================

    if (
        generar_preguntas_75
        and preguntas_generadas_75
    ):

        st.divider()

        st.subheader(
            "Sincronización"
        )

        sincronizar_75 = st.button(
            "SINCRONIZAR CON GITHUB",
            type="primary",
            key="sincronizar_75"
        )

        if sincronizar_75:

            st.info(
                "7.5.12 ✓ Iniciando sincronización "
                "del Banco General."
            )
# ========================================================
# 7.10 PRODUCTOS — VALIDACIÓN
# PRODUCTO → CATEGORÍA PRINCIPAL + COMPLEMENTARIAS
# NIVEL 2 ....OJOOOOOOO
# ========================================================

if (
    ROL_ACTUAL == "ADMINISTRADOR"
    and opcion_evaluacion == "Banco general de preguntas"
):

    st.subheader(
        "Validación — Producto → Categorías — Nivel 2"
    )

    RUTA_BANCO_710 = (
        BASE_DIR
        / "BANCO_PREGUNTAS_GENERALES.xlsx"
    )

    if not RUTA_BANCO_710.exists():

        st.warning(
            "No existe el Banco General de Preguntas."
        )

    else:

        banco_710 = pd.read_excel(
            RUTA_BANCO_710,
            dtype=str
        ).fillna("")

        # ----------------------------------------------------
        # BUSCAR SOLO NIVEL 2 PENDIENTE
        # ----------------------------------------------------

        pendientes_710 = banco_710[
            (
                banco_710["Tipo_Relacion"]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                "producto_categoria_nivel2"
            )
            &
            (
                banco_710["Nivel"]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                "nivel 2"
            )
            &
            (
                banco_710["Estado"]
                .astype(str)
                .str.strip()
                .str.upper()
                ==
                "PENDIENTE"
            )
        ].copy()

        # ----------------------------------------------------
        # MÁXIMO 5
        # ----------------------------------------------------

        pendientes_710 = pendientes_710.head(5)

        if pendientes_710.empty:

            st.info(
                "No hay preguntas pendientes "
                "de Producto → Categorías Nivel 2."
            )

        else:

            st.write(
                f"Preguntas para validar: "
                f"**{len(pendientes_710)}**"
            )

            # ------------------------------------------------
            # DECISIÓN PARA TODO EL BLOQUE
            # ------------------------------------------------

            decision_bloque_710 = st.radio(
                "Decisión para todo el bloque",
                [
                    "Mantener pendiente",
                    "Aprobar todo",
                    "Rechazar todo"
                ],
                index=0,
                horizontal=True,
                key="decision_bloque_producto_710"
            )

            decisiones_710 = {}

            # ------------------------------------------------
            # MOSTRAR PREGUNTAS
            # ------------------------------------------------

            for numero_710, (
                indice_710,
                fila_710
            ) in enumerate(
                pendientes_710.iterrows(),
                start=1
            ):

                pregunta_id_710 = str(
                    fila_710["Pregunta_ID"]
                )

                st.markdown(
                    f"### Pregunta {numero_710}"
                )

                st.caption(
                    f"ID: {pregunta_id_710}"
                )

                st.write(
                    f"**Producto:** "
                    f"{fila_710['Fuente_ID']}"
                )

                st.write(
                    f"**Pregunta:** "
                    f"{fila_710['Pregunta']}"
                )

                st.write(
                    f"1. {fila_710['Respuesta_1']}"
                )

                st.write(
                    f"2. {fila_710['Respuesta_2']}"
                )

                st.write(
                    f"3. {fila_710['Respuesta_3']}"
                )

                st.write(
                    f"4. {fila_710['Respuesta_4']}"
                )

                st.write(
                    f"**Respuestas correctas:** "
                    f"{fila_710['Respuesta_Correcta']}"
                )

                decision_individual_710 = st.radio(
                    "Estado",
                    [
                        "PENDIENTE",
                        "APROBADA",
                        "RECHAZADA"
                    ],
                    index=0,
                    horizontal=True,
                    key=(
                        f"estado_producto_710_"
                        f"{pregunta_id_710}"
                    )
                )

                decisiones_710[indice_710] = (
                    decision_individual_710
                )

                st.divider()

            # ------------------------------------------------
            # GUARDAR
            # ------------------------------------------------

            if st.button(
                "GUARDAR VALIDACIÓN",
                type="primary",
                key="guardar_producto_710"
            ):

                for indice_710, decision_710 in (
                    decisiones_710.items()
                ):

                    if (
                        decision_bloque_710
                        ==
                        "Aprobar todo"
                    ):

                        banco_710.loc[
                            indice_710,
                            "Estado"
                        ] = "APROBADA"

                    elif (
                        decision_bloque_710
                        ==
                        "Rechazar todo"
                    ):

                        banco_710.loc[
                            indice_710,
                            "Estado"
                        ] = "RECHAZADA"

                    else:

                        banco_710.loc[
                            indice_710,
                            "Estado"
                        ] = decision_710

                # ------------------------------------------------
                # GUARDAR EXCEL
                # ------------------------------------------------

                banco_710.to_excel(
                    RUTA_BANCO_710,
                    index=False,
                    sheet_name="Banco_General"
                )

                st.success(
                    "✓ Estados guardados correctamente."
                )

                st.rerun()




# ============================================================
# 8. PIE DE APLICACIÓN
# ============================================================

st.divider()

st.caption(
    "FITOASISTE — Herramienta de apoyo para tu proceso "
    "de aprendizaje y asesoría"
)
# ============================================================
# ============================================================
# BLOQUE 2.1.1 — ESTRUCTURA Y MENÚ DE CONSULTA DE PRODUCTOS
# ============================================================

if opcion_consulta == "Productos":

    st.subheader("Consulta de productos")

    tipo_consulta_producto = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Ver todos los productos",
            "Buscar producto",
            "Componente → productos",
            "Categoría → productos",
            "Producto → acciones generales"
        ],
        key="menu_consulta_productos"
    )

    st.session_state["tipo_consulta_producto"] = (
        tipo_consulta_producto
    )
# ============================================================
# BLOQUE 2.1.2 — LISTADO ALFABÉTICO Y SELECCIÓN
# ============================================================

if opcion_consulta == "Productos" and st.session_state.get("tipo_consulta_producto") == "Ver todos los productos":

    st.subheader("Listado de productos")

    productos_ordenados = sorted(
        Base_Productos["Producto"]
        .dropna()
        .astype(str)
        .unique()
    )

    producto_seleccionado = st.selectbox(
        "Seleccione el producto que desea consultar:",
        productos_ordenados,
        key="producto_seleccionado_listado"
    )

    st.write(
        "Producto seleccionado:",
        producto_seleccionado
    )
# ============================================================
# BLOQUE 2.1.3 — FICHA DEL PRODUCTO SELECCIONADO
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto") == "Ver todos los productos"
    and "producto_seleccionado" in locals()
):

    producto_data = Base_Productos[
        Base_Productos["Producto"].astype(str) == str(producto_seleccionado)
    ]

    if not producto_data.empty:

        st.subheader("Ficha del producto")

        producto_info = producto_data.iloc[0]

        for columna in Base_Productos.columns:

            valor = producto_info[columna]

            if pd.notna(valor):
                st.write(
                    f"**{columna}:**",
                    valor
                )

    else:

        st.warning(
            "No se encontró información para el producto seleccionado."
        )
# ============================================================
# BLOQUE 2.1.4 — COMPONENTE → PRODUCTOS
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Componente → productos"
):

    st.subheader("Buscar productos por componente")

    componente_buscado = st.text_input(
        "Escriba el componente que desea buscar:",
        key="componente_buscado"
    )

    if componente_buscado.strip():

        texto_buscado = (
            unidecode(
                componente_buscado
            )
            .lower()
            .strip()
        )

        resultados_componentes = []

        for _, fila in Base_Productos.iterrows():

            componente = fila.iloc[3]

            if pd.isna(componente):
                continue

            texto_componente = (
                unidecode(
                    str(componente)
                )
                .lower()
                .strip()
            )

            # ------------------------------------------------
            # COINCIDENCIA DIRECTA
            # ------------------------------------------------

            if texto_buscado in texto_componente:

                resultados_componentes.append(
                    fila["Producto"]
                )

                continue

            # ------------------------------------------------
            # COINCIDENCIA TOLERANTE A ERRORES ORTOGRÁFICOS
            # ------------------------------------------------

            palabras_buscadas = (
                texto_buscado.split()
            )

            palabras_componente = (
                texto_componente.split()
            )

            coincidencias = 0

            for palabra_buscada in palabras_buscadas:

                mejor_puntaje = 0

                for palabra_componente in palabras_componente:

                    puntaje = fuzz.ratio(
                        palabra_buscada,
                        palabra_componente
                    )

                    if puntaje > mejor_puntaje:

                        mejor_puntaje = puntaje

                if mejor_puntaje >= 80:

                    coincidencias += 1

            if (
                palabras_buscadas
                and coincidencias
                == len(palabras_buscadas)
            ):

                resultados_componentes.append(
                    fila["Producto"]
                )

        productos_encontrados = sorted(
            set(resultados_componentes)
        )

        if len(productos_encontrados) == 0:

            st.warning(
                "No se encontraron productos que "
                "contengan el componente buscado."
            )

        elif len(productos_encontrados) == 1:

            producto_seleccionado_componente = (
                productos_encontrados[0]
            )

            st.success(
                f"Producto encontrado: "
                f"{producto_seleccionado_componente}"
            )

        else:

            st.write(
                f"Se encontraron "
                f"{len(productos_encontrados)} productos:"
            )

            producto_seleccionado_componente = st.selectbox(
                "Seleccione el producto que desea consultar:",
                productos_encontrados,
                key="producto_seleccionado_componente"
            )

        if (
            "producto_seleccionado_componente" in locals()
        ):

            producto_data_componente = Base_Productos[
                Base_Productos["Producto"].astype(str)
                == str(producto_seleccionado_componente)
            ]

            if not producto_data_componente.empty:

                st.subheader("Ficha del producto")

                producto_info = (
                    producto_data_componente.iloc[0]
                )

                for columna in Base_Productos.columns:

                    valor = producto_info[columna]

                    if pd.notna(valor):

                        st.write(
                            f"**{columna}:**",
                            valor
                        )
# ============================================================
# BLOQUE 2.1.5 — BÚSQUEDA DE PRODUCTO POR NOMBRE
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Buscar producto"
):

    st.subheader("Buscar producto")

    nombre_buscado = st.text_input(
        "Escriba el nombre del producto:",
        key="nombre_buscado_producto"
    )

    if nombre_buscado.strip():

        texto_buscado = (
            unidecode(nombre_buscado)
            .lower()
            .strip()
        )

        productos_disponibles = (
            Base_Productos["Producto"]
            .dropna()
            .astype(str)
            .unique()
        )

        coincidencias = []

        for producto in productos_disponibles:

            producto_normalizado = (
                unidecode(str(producto))
                .lower()
                .strip()
            )

            puntaje = fuzz.ratio(
                texto_buscado,
                producto_normalizado
            )

            if (
                texto_buscado in producto_normalizado
                or puntaje >= 60
            ):

                coincidencias.append(
                    (str(producto), puntaje)
                )

        coincidencias = sorted(
            coincidencias,
            key=lambda x: x[1],
            reverse=True
        )

        productos_encontrados = [
            producto
            for producto, puntaje in coincidencias
        ]

        if len(productos_encontrados) == 0:

            st.warning(
                "No se encontraron productos "
                "que coincidan con la búsqueda."
            )

        elif len(productos_encontrados) == 1:

            producto_seleccionado_nombre = (
                productos_encontrados[0]
            )

        else:

            st.write(
                f"Se encontraron "
                f"{len(productos_encontrados)} "
                f"posibles coincidencias:"
            )

            producto_seleccionado_nombre = st.selectbox(
                "Seleccione el producto que desea consultar:",
                productos_encontrados,
                key="producto_seleccionado_nombre"
            )

        if "producto_seleccionado_nombre" in locals():

            producto_data_nombre = Base_Productos[
                Base_Productos["Producto"].astype(str)
                == str(producto_seleccionado_nombre)
            ]

            if not producto_data_nombre.empty:

                st.subheader("Ficha del producto")

                producto_info = producto_data_nombre.iloc[0]

                for columna in Base_Productos.columns:

                    valor = producto_info[columna]

                    if pd.notna(valor):

                        st.write(
                            f"**{columna}:**",
                            valor
                        )
# ============================================================

# ============================================================
# BLOQUE — CATEGORÍA → PRODUCTOS
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Categoría → productos"
):

    st.subheader("Categoría → productos")

    # --------------------------------------------------------
    # FUNCIÓN PARA NORMALIZAR TEXTO
    # --------------------------------------------------------

    def normalizar_categoria(texto):

        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

    # --------------------------------------------------------
    # FUNCIÓN PARA SEPARAR CATEGORÍAS COMPLEMENTARIAS
    # --------------------------------------------------------

    def separar_categorias_complementarias(valor):

        if pd.isna(valor):
            return []

        texto = str(valor)

        for separador in [";", "|", "/", "\n"]:

            texto = texto.replace(
                separador,
                ","
            )

        categorias = []

        for parte in texto.split(","):

            categoria = parte.strip()

            if categoria:
                categorias.append(
                    categoria
                )

        return categorias

    # ========================================================
    # OPCIONES INTERNAS DE CONSULTA
    # ========================================================

    tipo_busqueda_categoria = st.radio(
        "¿Cómo desea realizar la consulta?",
        [
            "Seleccionar del listado de categorías",
            "Ingresar categoría o patologia"
        ],
        key="tipo_busqueda_categoria"
    )

    # ========================================================
    # OPCIÓN 1
    # SELECCIONAR DEL LISTADO DE CATEGORÍAS
    # ========================================================

    if (
        tipo_busqueda_categoria
        == "Seleccionar del listado de categorías"
    ):

        st.write(
            "Seleccione una categoría principal:"
        )

        # ----------------------------------------------------
        # OBTENER CATEGORÍAS DE LA COLUMNA 2
        # ----------------------------------------------------

        categorias_principales = []

        for valor in Base_Productos.iloc[:, 1]:

            if pd.isna(valor):
                continue

            categoria = str(valor).strip()

            if categoria:
                categorias_principales.append(
                    categoria
                )

        # ----------------------------------------------------
        # ELIMINAR DUPLICADOS
        # ----------------------------------------------------

        categorias_unicas = {}

        for categoria in categorias_principales:

            clave = normalizar_categoria(
                categoria
            )

            if clave not in categorias_unicas:

                categorias_unicas[clave] = (
                    categoria
                )

        # ----------------------------------------------------
        # ORDEN ALFABÉTICO
        # ----------------------------------------------------

        categorias_finales = sorted(
            categorias_unicas.values(),
            key=lambda x: normalizar_categoria(x)
        )

        # ----------------------------------------------------
        # LISTADO DESPLEGABLE
        # ----------------------------------------------------

        categoria_seleccionada = st.selectbox(
            "Categoría principal:",
            [
                "Seleccione una categoría"
            ] + categorias_finales,
            key="categoria_principal_seleccionada"
        )

        # ----------------------------------------------------
        # SI SELECCIONÓ UNA CATEGORÍA
        # ----------------------------------------------------

        if (
            categoria_seleccionada
            != "Seleccione una categoría"
        ):

            categoria_buscada = normalizar_categoria(
                categoria_seleccionada
            )

            productos_directos = []

            # ------------------------------------------------
            # BUSCAR EN COLUMNA 2
            # ------------------------------------------------

            for _, fila in Base_Productos.iterrows():

                producto = fila.iloc[0]
                categoria_principal = fila.iloc[1]

                if pd.isna(producto):
                    continue

                if pd.isna(categoria_principal):
                    continue

                categoria_principal_normalizada = (
                    normalizar_categoria(
                        categoria_principal
                    )
                )

                if (
                    categoria_principal_normalizada
                    == categoria_buscada
                ):

                    productos_directos.append(
                        str(producto).strip()
                    )

            # ------------------------------------------------
            # ELIMINAR DUPLICADOS
            # ------------------------------------------------

            productos_directos = list(
                dict.fromkeys(
                    productos_directos
                )
            )

            productos_directos = sorted(
                productos_directos,
                key=lambda x: normalizar_categoria(x)
            )

            # ------------------------------------------------
            # MENSAJE
            # ------------------------------------------------

            st.success(
                f"Categoría seleccionada: "
                f"{categoria_seleccionada}"
            )

            st.write(
                "A continuación se presenta el listado "
                "de productos que tienen una acción directa "
                "sobre la patologia:"
            )

            # ------------------------------------------------
            # LISTADO DE PRODUCTOS
            # ------------------------------------------------

            if not productos_directos:

                st.warning(
                    "No se encontraron productos "
                    "de acción directa para esta categoría."
                )

            else:

                producto_seleccionado = st.selectbox(
                    "Seleccione el producto que desea consultar:",
                    [
                        "Seleccione un producto"
                    ] + productos_directos,
                    key="producto_categoria_principal"
                )

                # ------------------------------------------------
                # FICHA DEL PRODUCTO
                # ------------------------------------------------

                if (
                    producto_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_ficha = Base_Productos[
                        Base_Productos.iloc[:, 0]
                        .astype(str)
                        .str.strip()
                        == str(
                            producto_seleccionado
                        ).strip()
                    ]

                    if not producto_ficha.empty:

                        st.divider()

                        st.subheader(
                            "Ficha completa del producto"
                        )

                        datos_producto = (
                            producto_ficha.iloc[0]
                        )

                        for columna in Base_Productos.columns:

                            valor = datos_producto[
                                columna
                            ]

                            if pd.notna(valor):

                                st.write(
                                    f"**{columna}:**",
                                    valor
                                )

    # ========================================================
    # OPCIÓN 2
    # INGRESAR CATEGORÍA O PATOLOGIA
    # ========================================================

    else:

        st.write(
            "Registre el nombre de la patologia "
            "o categoría que desea consultar:"
        )

        categoria_ingresada = st.text_input(
            "Nombre de la patologia o categoría:",
            key="categoria_ingresada"
        )

        if categoria_ingresada.strip():

            texto_buscado = normalizar_categoria(
                categoria_ingresada
            )

            productos_directos = []
            productos_complementarios = []

            # ------------------------------------------------
            # RECORRER TODA LA BASE DE PRODUCTOS
            # ------------------------------------------------

            for _, fila in Base_Productos.iterrows():

                producto = fila.iloc[0]

                if pd.isna(producto):
                    continue

                producto = str(producto).strip()

                # ============================================
                # COLUMNA 2
                # CATEGORÍA PRINCIPAL
                # ============================================

                categoria_principal = fila.iloc[1]

                if not pd.isna(categoria_principal):

                    texto_principal = (
                        normalizar_categoria(
                            categoria_principal
                        )
                    )

                    coincidencia_principal = False

                    # Coincidencia directa
                    if (
                        texto_buscado
                        in texto_principal
                    ):

                        coincidencia_principal = True

                    # Coincidencia aproximada
                    else:

                        similitud = fuzz.ratio(
                            texto_buscado,
                            texto_principal
                        )

                        if similitud >= 70:

                            coincidencia_principal = True

                    if coincidencia_principal:

                        productos_directos.append(
                            producto
                        )

                # ============================================
                # COLUMNA 3
                # CATEGORÍAS COMPLEMENTARIAS
                # ============================================

                valor_complementarias = fila.iloc[2]

                categorias_complementarias = (
                    separar_categorias_complementarias(
                        valor_complementarias
                    )
                )

                for categoria_complementaria in (
                    categorias_complementarias
                ):

                    texto_complementario = (
                        normalizar_categoria(
                            categoria_complementaria
                        )
                    )

                    coincidencia_complementaria = False

                    # Coincidencia directa
                    if (
                        texto_buscado
                        in texto_complementario
                    ):

                        coincidencia_complementaria = True

                    # Coincidencia aproximada
                    else:

                        similitud = fuzz.ratio(
                            texto_buscado,
                            texto_complementario
                        )

                        if similitud >= 70:

                            coincidencia_complementaria = True

                    if coincidencia_complementaria:

                        productos_complementarios.append(
                            producto
                        )

                        break

            # ------------------------------------------------
            # ELIMINAR DUPLICADOS
            # ------------------------------------------------

            productos_directos = list(
                dict.fromkeys(
                    productos_directos
                )
            )

            productos_complementarios = list(
                dict.fromkeys(
                    productos_complementarios
                )
            )

            # ------------------------------------------------
            # SI UN PRODUCTO ES PRINCIPAL,
            # NO REPETIRLO COMO COMPLEMENTARIO
            # ------------------------------------------------

            productos_complementarios = [
                producto
                for producto in productos_complementarios
                if producto not in productos_directos
            ]

            # ------------------------------------------------
            # ORDENAR PRODUCTOS
            # ------------------------------------------------

            productos_directos = sorted(
                productos_directos,
                key=lambda x: normalizar_categoria(x)
            )

            productos_complementarios = sorted(
                productos_complementarios,
                key=lambda x: normalizar_categoria(x)
            )

            # =================================================
            # RESULTADOS
            # =================================================

            if (
                not productos_directos
                and not productos_complementarios
            ):

                st.warning(
                    "No se encontraron productos "
                    "relacionados con la categoría o "
                    "patologia ingresada."
                )

            else:

                st.success(
                    f"Resultados para: "
                    f"{categoria_ingresada}"
                )

                # =============================================
                # DESPLEGABLE 1
                # PRODUCTOS CON ACCIÓN DIRECTA
                # =============================================

                st.markdown(
                    "### Productos con acción directa"
                )

                if productos_directos:

                    producto_directo_seleccionado = (
                        st.selectbox(
                            "Seleccione un producto principal:",
                            [
                                "Seleccione un producto"
                            ] + productos_directos,
                            key="producto_directo_categoria"
                        )
                    )

                else:

                    st.info(
                        "No se encontraron productos "
                        "con acción directa."
                    )

                    producto_directo_seleccionado = (
                        "Seleccione un producto"
                    )

                # =============================================
                # DESPLEGABLE 2
                # PRODUCTOS COMPLEMENTARIOS
                # =============================================

                st.markdown(
                    "### Productos complementarios"
                )

                if productos_complementarios:

                    producto_complementario_seleccionado = (
                        st.selectbox(
                            "Seleccione un producto complementario:",
                            [
                                "Seleccione un producto"
                            ] + productos_complementarios,
                            key="producto_complementario_categoria"
                        )
                    )

                else:

                    st.info(
                        "No se encontraron productos "
                        "complementarios."
                    )

                    producto_complementario_seleccionado = (
                        "Seleccione un producto"
                    )

                # =============================================
                # DETERMINAR PRODUCTO SELECCIONADO
                # =============================================

                producto_para_ficha = None

                if (
                    producto_directo_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_para_ficha = (
                        producto_directo_seleccionado
                    )

                elif (
                    producto_complementario_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_para_ficha = (
                        producto_complementario_seleccionado
                    )

                # =============================================
                # MOSTRAR FICHA
                # =============================================

                if producto_para_ficha:

                    producto_ficha = Base_Productos[
                        Base_Productos.iloc[:, 0]
                        .astype(str)
                        .str.strip()
                        == str(
                            producto_para_ficha
                        ).strip()
                    ]

                    if not producto_ficha.empty:

                        st.divider()

                        st.subheader(
                            "Ficha completa del producto"
                        )

                        datos_producto = (
                            producto_ficha.iloc[0]
                        )

                        for columna in Base_Productos.columns:

                            valor = datos_producto[
                                columna
                            ]

                            if pd.notna(valor):

                                st.write(
                                    f"**{columna}:**",
                                    valor
                                )

    # ========================================================
    # NAVEGACIÓN
    # ========================================================

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "🔎 Nueva consulta",
            key="nueva_consulta_categoria",
            use_container_width=True
        ):

            st.session_state.pop(
                "categoria_principal_seleccionada",
                None
            )

            st.session_state.pop(
                "categoria_ingresada",
                None
            )

            st.session_state.pop(
                "producto_categoria_principal",
                None
            )

            st.session_state.pop(
                "producto_directo_categoria",
                None
            )

            st.session_state.pop(
                "producto_complementario_categoria",
                None
            )

            st.rerun()

    with col2:

        if st.button(
            "← Menú Productos",
            key="volver_menu_productos_categoria",
            use_container_width=True
        ):

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()

    with col3:

        if st.button(
            "🏠 Menú principal",
            key="volver_menu_principal_categoria",
            use_container_width=True
        ):

            st.session_state[
                "opcion_consulta"
            ] = "Seleccione una opción"

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()


# ============================================================
# BLOQUE — PRODUCTO → ACCIONES GENERALES
# ============================================================

if (
    opcion_consulta == "Productos"
    and st.session_state.get("tipo_consulta_producto")
    == "Producto → acciones generales"
):

    st.subheader("Producto → acciones generales")

    # ========================================================
    # FUNCIONES AUXILIARES
    # ========================================================

    def normalizar_accion_general(texto):

        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

    def separar_acciones_generales(valor):

        if pd.isna(valor):
            return []

        texto = str(valor)

        partes = texto.split(";")

        acciones = []

        for parte in partes:

            accion = parte.strip()

            if accion:

                acciones.append(
                    accion
                )

        return acciones

    # ========================================================
    # TIPO DE CONSULTA
    # ========================================================

    tipo_busqueda_accion = st.radio(
        "¿Cómo desea realizar la consulta?",
        [
            "Buscar por producto",
            "Buscar por acción general"
        ],
        key="tipo_busqueda_accion_general"
    )

    # ========================================================
    # 1. BUSCAR POR PRODUCTO
    # ========================================================

    if (
        tipo_busqueda_accion
        == "Buscar por producto"
    ):

        st.write(
            "Seleccione el producto para consultar "
            "sus acciones generales:"
        )

        # ----------------------------------------------------
        # LISTADO DE PRODUCTOS
        # ----------------------------------------------------

        productos = []

        for valor in Base_Productos.iloc[:, 0]:

            if pd.isna(valor):
                continue

            producto = str(valor).strip()

            if producto:

                productos.append(
                    producto
                )

        # ----------------------------------------------------
        # ELIMINAR DUPLICADOS
        # ----------------------------------------------------

        productos_unicos = {}

        for producto in productos:

            clave = (
                normalizar_accion_general(
                    producto
                )
            )

            if clave not in productos_unicos:

                productos_unicos[clave] = (
                    producto
                )

        productos_finales = sorted(
            productos_unicos.values(),
            key=lambda x:
                normalizar_accion_general(x)
        )

        # ----------------------------------------------------
        # SELECCIÓN DEL PRODUCTO
        # ----------------------------------------------------

        producto_seleccionado = st.selectbox(
            "Producto:",
            [
                "Seleccione un producto"
            ] + productos_finales,
            key="producto_acciones_generales"
        )

        # ----------------------------------------------------
        # MOSTRAR ACCIONES
        # ----------------------------------------------------

        if (
            producto_seleccionado
            != "Seleccione un producto"
        ):

            producto_fila = Base_Productos[
                Base_Productos.iloc[:, 0]
                .astype(str)
                .str.strip()
                == str(
                    producto_seleccionado
                ).strip()
            ]

            if not producto_fila.empty:

                datos_producto = (
                    producto_fila.iloc[0]
                )

                # COLUMNA 5
                acciones = (
                    separar_acciones_generales(
                        datos_producto.iloc[4]
                    )
                )

                st.divider()

                st.subheader(
                    "Acciones generales"
                )

                if acciones:

                    for numero, accion in enumerate(
                        acciones,
                        start=1
                    ):

                        st.write(
                            f"{numero}. {accion}"
                        )

                else:

                    st.info(
                        "Este producto no tiene "
                        "acciones generales registradas."
                    )

    # ========================================================
    # 2. BUSCAR POR ACCIÓN GENERAL
    # ========================================================

    else:

        st.write(
            "Registre la acción general que desea buscar:"
        )

        # ----------------------------------------------------
        # CAMPO DE BÚSQUEDA
        # ----------------------------------------------------

        accion_ingresada = st.text_input(
            "Escriba la acción general:",
            key="accion_general_ingresada"
        )

        # ----------------------------------------------------
        # SOLO BUSCAR SI HAY TEXTO
        # ----------------------------------------------------

        if accion_ingresada.strip():

            texto_buscado = (
                normalizar_accion_general(
                    accion_ingresada
                )
            )

            resultados = []

            # =================================================
            # RECORRER TODA LA BASE
            # =================================================

            for _, fila in Base_Productos.iterrows():

                producto = fila.iloc[0]

                if pd.isna(producto):
                    continue

                producto = str(producto).strip()

                # ------------------------------------------------
                # COLUMNA 5 — ACCIONES GENERALES
                # ------------------------------------------------

                valor_acciones = fila.iloc[4]

                acciones = (
                    separar_acciones_generales(
                        valor_acciones
                    )
                )

                # ------------------------------------------------
                # REVISAR CADA ACCIÓN
                # ------------------------------------------------

                for accion in acciones:

                    accion_normalizada = (
                        normalizar_accion_general(
                            accion
                        )
                    )

                    coincidencia = False

                    # ==========================================
                    # NIVEL 1 — FRASE COMPLETA
                    # ==========================================

                    if (
                        texto_buscado
                        in accion_normalizada
                    ):

                        coincidencia = True

                    # ==========================================
                    # NIVEL 2 — PALABRAS
                    # ==========================================

                    if not coincidencia:

                        palabras_buscadas = (
                            texto_buscado.split()
                        )

                        palabras_accion = (
                            accion_normalizada.split()
                        )

                        total_encontradas = 0

                        for palabra_buscada in (
                            palabras_buscadas
                        ):

                            mejor_puntaje = 0

                            for palabra_accion in (
                                palabras_accion
                            ):

                                puntaje = fuzz.ratio(
                                    palabra_buscada,
                                    palabra_accion
                                )

                                if (
                                    puntaje
                                    > mejor_puntaje
                                ):

                                    mejor_puntaje = (
                                        puntaje
                                    )

                            if mejor_puntaje >= 70:

                                total_encontradas += 1

                        if palabras_buscadas:

                            porcentaje = (
                                total_encontradas
                                / len(
                                    palabras_buscadas
                                )
                            ) * 100

                            if porcentaje >= 70:

                                coincidencia = True

                    # ==========================================
                    # SI ENCUENTRA LA ACCIÓN
                    # ==========================================

                    if coincidencia:

                        resultados.append(
                            {
                                "Producto": producto,
                                "Accion": accion
                            }
                        )

                        # Un producto solo aparece una vez
                        break

            # =================================================
            # ELIMINAR PRODUCTOS DUPLICADOS
            # =================================================

            resultados_unicos = {}

            for resultado in resultados:

                clave = (
                    normalizar_accion_general(
                        resultado["Producto"]
                    )
                )

                if clave not in resultados_unicos:

                    resultados_unicos[clave] = (
                        resultado
                    )

            resultados = list(
                resultados_unicos.values()
            )

            # =================================================
            # ORDENAR RESULTADOS
            # =================================================

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_accion_general(
                        x["Producto"]
                    )
            )

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not resultados:

                st.warning(
                    "No se encontraron productos "
                    "con esa acción general."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"productos relacionados."
                )

                # ------------------------------------------------
                # LISTADO DESPLEGABLE
                # ------------------------------------------------

                productos_encontrados = [
                    resultado["Producto"]
                    for resultado in resultados
                ]

                producto_seleccionado = st.selectbox(
                    "Seleccione el producto que desea consultar:",
                    [
                        "Seleccione un producto"
                    ] + productos_encontrados,
                    key="producto_resultado_accion"
                )

                # =================================================
                # MOSTRAR FICHA
                # =================================================

                if (
                    producto_seleccionado
                    != "Seleccione un producto"
                ):

                    accion_encontrada = None

                    for resultado in resultados:

                        if (
                            resultado["Producto"]
                            == producto_seleccionado
                        ):

                            accion_encontrada = (
                                resultado["Accion"]
                            )

                            break

                    if accion_encontrada:

                        st.write(
                            "**Acción general encontrada:**"
                        )

                        st.info(
                            accion_encontrada
                        )

                    producto_ficha = Base_Productos[
                        Base_Productos.iloc[:, 0]
                        .astype(str)
                        .str.strip()
                        == str(
                            producto_seleccionado
                        ).strip()
                    ]

                    if not producto_ficha.empty:

                        st.divider()

                        st.subheader(
                            "Ficha completa del producto"
                        )

                        datos_producto = (
                            producto_ficha.iloc[0]
                        )

                        for columna in Base_Productos.columns:

                            valor = datos_producto[
                                columna
                            ]

                            if pd.notna(valor):

                                st.write(
                                    f"**{columna}:**",
                                    valor
                                )

    # ========================================================
    # NAVEGACIÓN
    # ========================================================

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "🔎 Nueva consulta",
            key="nueva_consulta_accion_general",
            use_container_width=True
        ):

            st.session_state.pop(
                "producto_acciones_generales",
                None
            )

            st.session_state.pop(
                "accion_general_ingresada",
                None
            )

            st.session_state.pop(
                "producto_resultado_accion",
                None
            )

            st.rerun()

    with col2:

        if st.button(
            "← Menú Productos",
            key="volver_menu_productos_accion_general",
            use_container_width=True
        ):

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()

    with col3:

        if st.button(
            "🏠 Menú principal",
            key="volver_menu_principal_accion_general",
            use_container_width=True
        ):

            st.session_state[
                "opcion_consulta"
            ] = "Seleccione una opción"

            st.session_state[
                "tipo_consulta_producto"
            ] = "Seleccione una opción"

            st.rerun()


# ============================================================
# BLOQUE — PATOLOGIAS
# CONSULTA GENERAL
# ============================================================

if (
    opcion_consulta == "Patologias"
):

    st.subheader("Consulta de patologias")

    # ========================================================
    # FUNCIONES AUXILIARES
    # ========================================================

    def normalizar_patologia(texto):

        return (
            unidecode(str(texto))
            .lower()
            .strip()
        )

    # ============================================================
    # MENÚ DE CONSULTA — PATOLOGÍAS
    # ============================================================

    tipo_busqueda_patologia = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Ver todas las patologias",
            "Ingresar código o nombre de la patologia",
            "Patologia → causa y síntoma"
        ],
        key="tipo_busqueda_patologia"
    )

    # ============================================================
    # 1. VER TODAS LAS PATOLOGÍAS
    # ============================================================

    if (
        tipo_busqueda_patologia
        == "Ver todas las patologias"
    ):

        st.write(
            "Seleccione la patologia que desea consultar:"
        )

        patologias = []

        for valor in Patologias.iloc[:, 1]:

            if pd.isna(valor):
                continue

            patologia = str(valor).strip()

            if patologia:
                patologias.append(patologia)

        # --------------------------------------------------------
        # ELIMINAR DUPLICADOS
        # --------------------------------------------------------

        patologias_unicas = {}

        for patologia in patologias:

            clave = normalizar_patologia(
                patologia
            )

            if clave not in patologias_unicas:

                patologias_unicas[clave] = patologia

        # --------------------------------------------------------
        # ORDEN ALFABÉTICO
        # --------------------------------------------------------

        patologias_finales = sorted(
            patologias_unicas.values(),
            key=lambda x:
                normalizar_patologia(x)
        )

        # --------------------------------------------------------
        # SELECCIÓN
        # --------------------------------------------------------

        patologia_seleccionada = st.selectbox(
            "Patologia:",
            [
                "Seleccione una patologia"
            ] + patologias_finales,
            key="patologia_listado_general"
        )

        # --------------------------------------------------------
        # MOSTRAR FICHA
        # --------------------------------------------------------

        if (
            patologia_seleccionada
            != "Seleccione una patologia"
        ):

            patologia_ficha = Patologias[
                Patologias.iloc[:, 1]
                .astype(str)
                .str.strip()
                ==
                str(
                    patologia_seleccionada
                ).strip()
            ]

            if not patologia_ficha.empty:

                datos = patologia_ficha.iloc[0]

                st.divider()

                st.subheader(
                    "Ficha completa de la patologia"
                )

                st.write(
                    f"**Código:** {datos.iloc[0]}"
                )

                st.write(
                    f"**Patología:** {datos.iloc[1]}"
                )

                st.write(
                    f"**Descripción breve:** {datos.iloc[2]}"
                )

                st.write(
                    f"**Causas frecuentes:** {datos.iloc[3]}"
                )

                st.write(
                    f"**Síntomas / Señales clave:** {datos.iloc[4]}"
                )

                st.write(
                    f"**Objetivo del paquete:** {datos.iloc[5]}"
                )

                st.write(
                    f"**Notas para el asesor:** {datos.iloc[6]}"
                )


    # ============================================================
    # 2. INGRESAR CÓDIGO O NOMBRE
    # ============================================================

    elif (
        tipo_busqueda_patologia
        == "Ingresar código o nombre de la patologia"
    ):

        st.write(
            "Ingrese el código o nombre de la patologia:"
        )

        texto_ingresado = st.text_input(
            "Código o nombre:",
            key="buscar_patologia_general"
        )

        if texto_ingresado.strip():

            texto_buscado = normalizar_patologia(
                texto_ingresado
            )

            resultados = []

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]

                if pd.isna(codigo):
                    codigo = ""

                if pd.isna(nombre):
                    nombre = ""

                codigo = str(codigo).strip()
                nombre = str(nombre).strip()

                codigo_normalizado = normalizar_patologia(
                    codigo
                )

                nombre_normalizado = normalizar_patologia(
                    nombre
                )

                coincidencia = False

                # ------------------------------------------------
                # CÓDIGO
                # ------------------------------------------------

                if texto_buscado in codigo_normalizado:
                    coincidencia = True

                # ------------------------------------------------
                # NOMBRE
                # ------------------------------------------------

                if texto_buscado in nombre_normalizado:
                    coincidencia = True

                # ------------------------------------------------
                # SIMILITUD
                # ------------------------------------------------

                if not coincidencia:

                    similitud = fuzz.ratio(
                        texto_buscado,
                        nombre_normalizado
                    )

                    if similitud >= 65:
                        coincidencia = True

                # ------------------------------------------------
                # SIMILITUD POR PALABRAS
                # ------------------------------------------------

                if not coincidencia:

                    palabras_buscadas = (
                        texto_buscado.split()
                    )

                    palabras_nombre = (
                        nombre_normalizado.split()
                    )

                    palabras_encontradas = 0

                    for palabra_buscada in palabras_buscadas:

                        for palabra_nombre in palabras_nombre:

                            similitud_palabra = fuzz.ratio(
                                palabra_buscada,
                                palabra_nombre
                            )

                            if similitud_palabra >= 70:

                                palabras_encontradas += 1
                                break

                    if palabras_buscadas:

                        porcentaje = (
                            palabras_encontradas
                            /
                            len(palabras_buscadas)
                        ) * 100

                        if porcentaje >= 70:
                            coincidencia = True

                # ------------------------------------------------
                # GUARDAR
                # ------------------------------------------------

                if coincidencia:

                    resultados.append(
                        {
                            "Codigo": codigo,
                            "Patologia": nombre
                        }
                    )

            # ----------------------------------------------------
            # ELIMINAR DUPLICADOS
            # ----------------------------------------------------

            resultados_unicos = {}

            for resultado in resultados:

                clave = (
                    resultado["Codigo"]
                    + "|"
                    + normalizar_patologia(
                        resultado["Patologia"]
                    )
                )

                if clave not in resultados_unicos:

                    resultados_unicos[clave] = resultado

            resultados = list(
                resultados_unicos.values()
            )

            # ----------------------------------------------------
            # ORDENAR
            # ----------------------------------------------------

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_patologia(
                        x["Patologia"]
                    )
            )

            # ----------------------------------------------------
            # MOSTRAR
            # ----------------------------------------------------

            if not resultados:

                st.warning(
                    "No se encontraron patologias "
                    "relacionadas con la búsqueda."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"posibles coincidencias."
                )

                opciones = [
                    "Seleccione una patologia"
                ]

                for resultado in resultados:

                    opciones.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion = st.selectbox(
                    "Seleccione la patologia que desea consultar:",
                    opciones,
                    key="resultado_busqueda_patologia"
                )

                if (
                    seleccion
                    != "Seleccione una patologia"
                ):

                    codigo_seleccionado = (
                        seleccion
                        .split(" — ")[0]
                        .strip()
                    )

                    mostrar_ficha_patologia(
                        codigo_seleccionado
                    )

# ============================================================
# ============================================================
# 3. PATOLOGÍA → CAUSA Y SÍNTOMA
# ============================================================

if (
    opcion_consulta == "Patologias"
    and st.session_state.get("tipo_busqueda_patologia")
    == "Patologia → causa y síntoma"
):

    st.subheader("Patología → causa y síntoma")

    # ========================================================
    # MENÚ INTERNO
    # ========================================================

    tipo_busqueda_causa_sintoma = st.selectbox(
        "¿Qué desea buscar?",
        [
            "Seleccione una opción",
            "Patología",
            "Causa",
            "Síntoma"
        ],
        key="tipo_busqueda_causa_sintoma"
    )

    # ========================================================
    # FUNCIÓN DE COMPARACIÓN DE TEXTO
    # ========================================================

    def coincide_texto_patologia(
        texto_buscado,
        texto_base,
        umbral=70
    ):

        buscado = normalizar_patologia(
            texto_buscado
        )

        base = normalizar_patologia(
            texto_base
        )

        if not buscado or not base:
            return False

        if buscado in base:
            return True

        palabras_buscadas = buscado.split()
        palabras_base = base.split()

        if not palabras_buscadas:
            return False

        palabras_encontradas = 0

        for palabra_buscada in palabras_buscadas:

            mejor_puntaje = 0

            for palabra_base in palabras_base:

                puntaje = fuzz.ratio(
                    palabra_buscada,
                    palabra_base
                )

                if puntaje > mejor_puntaje:
                    mejor_puntaje = puntaje

            if mejor_puntaje >= umbral:
                palabras_encontradas += 1

        porcentaje = (
            palabras_encontradas
            /
            len(palabras_buscadas)
        ) * 100

        return porcentaje >= umbral

    # ========================================================
    # FUNCIÓN — MOSTRAR FICHA COMPLETA
    # ========================================================

    def mostrar_ficha_patologia(
        codigo_seleccionado
    ):

        if Patologias is None:

            st.error(
                "La matriz de patologías no está disponible."
            )

            return

        if Patologias.empty:

            st.warning(
                "La hoja de patologías está vacía."
            )

            return

        ficha = Patologias[
            Patologias.iloc[:, 0]
            .astype(str)
            .str.strip()
            ==
            str(
                codigo_seleccionado
            ).strip()
        ]

        if ficha.empty:

            st.warning(
                "No fue posible encontrar "
                "la ficha de la patología."
            )

            return

        datos = ficha.iloc[0]

        st.divider()

        st.subheader(
            "Ficha completa de la patología"
        )

        st.write(
            f"**Código:** {datos.iloc[0]}"
        )

        st.write(
            f"**Patología:** {datos.iloc[1]}"
        )

        st.write(
            f"**Descripción breve:** {datos.iloc[2]}"
        )

        st.write(
            f"**Causas frecuentes:** {datos.iloc[3]}"
        )

        st.write(
            f"**Síntomas / Señales clave:** {datos.iloc[4]}"
        )

        st.write(
            f"**Objetivo del paquete:** {datos.iloc[5]}"
        )

        st.write(
            f"**Notas para el asesor:** {datos.iloc[6]}"
        )

    # ========================================================
    # 3.1 BUSCAR POR PATOLOGÍA
    # ========================================================

    if (
        tipo_busqueda_causa_sintoma
        == "Patología"
    ):

        texto_buscado = st.text_input(
            "Ingrese el nombre de la patología:",
            key="texto_busqueda_patologia_causa"
        )

        if texto_buscado.strip():

            resultados = []

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]

                if pd.isna(codigo):
                    continue

                if pd.isna(nombre):
                    continue

                if coincide_texto_patologia(
                    texto_buscado,
                    nombre,
                    umbral=70
                ):

                    resultados.append(
                        {
                            "Codigo":
                                str(
                                    codigo
                                ).strip(),

                            "Patologia":
                                str(
                                    nombre
                                ).strip()
                        }
                    )

            # ------------------------------------------------
            # ELIMINAR DUPLICADOS
            # ------------------------------------------------

            resultados_unicos = {}

            for resultado in resultados:

                clave = resultado[
                    "Codigo"
                ]

                if clave not in resultados_unicos:

                    resultados_unicos[
                        clave
                    ] = resultado

            resultados = list(
                resultados_unicos.values()
            )

            # ------------------------------------------------
            # ORDENAR
            # ------------------------------------------------

            resultados = sorted(
                resultados,
                key=lambda x:
                    normalizar_patologia(
                        x["Patologia"]
                    )
            )

            # ------------------------------------------------
            # MOSTRAR
            # ------------------------------------------------

            if not resultados:

                st.warning(
                    "No se encontraron patologías "
                    "relacionadas con la búsqueda."
                )

            else:

                st.success(
                    f"Se encontraron "
                    f"{len(resultados)} "
                    f"posibles coincidencias."
                )

                opciones = [
                    "Seleccione una patología"
                ]

                for resultado in resultados:

                    opciones.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion = st.selectbox(
                    "Seleccione la patología:",
                    opciones,
                    key="seleccion_patologia_causa"
                )

                if (
                    seleccion
                    !=
                    "Seleccione una patología"
                ):

                    codigo = (
                        seleccion
                        .split(" — ")[0]
                        .strip()
                    )

                    mostrar_ficha_patologia(
                        codigo
                    )
# ============================================================
# 3.2 BUSCAR POR CAUSA
# ============================================================

    elif (
        tipo_busqueda_causa_sintoma
        == "Causa"
    ):

        st.subheader(
            "Búsqueda por causas"
        )

# ========================================================
# CONFIGURACIÓN
# ========================================================

        UMBRAL_DIRECTO_CAUSA = 82.0
        UMBRAL_SEMANTICO_CAUSA = 65.0
        MAX_RESULTADOS_SEMANTICOS_CAUSA = 5

# ========================================================
# OBTENER MODELO BIOMÉDICO
# ========================================================

        modelo_biomedico_causa = globals().get(
            "modelo_biomedico",
            None
        )

        if modelo_biomedico_causa is None:

            st.error(
                "No está disponible el modelo biomédico "
                "para la búsqueda de causas."
            )

        else:

# ====================================================
# CONSTRUIR BASE DE CAUSAS DESDE LA MATRIZ
# ====================================================

            base_causas_3_2 = []

            for _, fila in Patologias.iterrows():

                codigo = fila.iloc[0]
                nombre = fila.iloc[1]
                causas = fila.iloc[3]

                if pd.isna(codigo):
                    continue

                if pd.isna(nombre):
                    continue

                if pd.isna(causas):
                    continue

                for elemento in str(causas).split(";"):

                    causa = elemento.strip()

                    if not causa:
                        continue

                    base_causas_3_2.append(
                        {
                            "Patologia_ID":
                                str(codigo).strip(),

                            "Patologia":
                                str(nombre).strip(),

                            "Causa":
                                causa
                        }
                    )

            df_causas_3_2 = pd.DataFrame(
                base_causas_3_2
            )

            if df_causas_3_2.empty:

                st.warning(
                    "No existen causas disponibles "
                    "en la matriz de patologías."
                )

            else:

# =================================================
# FUNCIÓN DE LIMPIEZA
# =================================================

                def limpiar_causa_3_2(
                    texto
                ):

                    if pd.isna(texto):
                        return ""

                    texto = str(
                        texto
                    )

                    texto = texto.lower()

                    texto = unidecode(
                        texto
                    )

                    texto = " ".join(
                        texto.split()
                    )

                    return texto.strip()

# =================================================
# PREPARAR TEXTO
# =================================================

                df_causas_3_2[
                    "Busqueda_3_2"
                ] = (
                    df_causas_3_2[
                        "Causa"
                    ]
                    .apply(
                        limpiar_causa_3_2
                    )
                )

                lista_causas_3_2 = (
                    df_causas_3_2[
                        "Busqueda_3_2"
                    ]
                    .drop_duplicates()
                    .tolist()
                )

# =================================================
# EMBEDDINGS DINÁMICOS DE CAUSAS
# =================================================
#
# NO se reutilizan embeddings de síntomas.
# Se generan para la columna de causas de la matriz.
# Se guardan en session_state y se regeneran solo
# cuando cambia el contenido de la base.

                firma_causas_3_2 = tuple(
                    df_causas_3_2[
                        "Busqueda_3_2"
                    ].tolist()
                )

                if (
                    "firma_embeddings_causas_3_2"
                    not in
                    st.session_state
                    or
                    st.session_state[
                        "firma_embeddings_causas_3_2"
                    ]
                    !=
                    firma_causas_3_2
                ):

                    textos_causas_3_2 = (
                        df_causas_3_2[
                            "Busqueda_3_2"
                        ]
                        .tolist()
                    )

                    st.session_state[
                        "embeddings_causas_3_2"
                    ] = (
                        modelo_biomedico_causa.encode(
                            textos_causas_3_2,
                            normalize_embeddings=True
                        )
                    )

                    st.session_state[
                        "firma_embeddings_causas_3_2"
                    ] = firma_causas_3_2

                embeddings_causas_3_2 = np.asarray(
                    st.session_state[
                        "embeddings_causas_3_2"
                    ],
                    dtype=np.float32
                )

# =================================================
# BÚSQUEDA DIRECTA
# =================================================

                def buscar_directa_causa_3_2(
                    consulta
                ):

                    consulta_limpia = (
                        limpiar_causa_3_2(
                            consulta
                        )
                    )

                    resultados = []

                    for _, fila in (
                        df_causas_3_2.iterrows()
                    ):

                        causa_limpia = (
                            fila[
                                "Busqueda_3_2"
                            ]
                        )

                        if (
                            consulta_limpia
                            ==
                            causa_limpia
                            or
                            consulta_limpia
                            in
                            causa_limpia
                            or
                            causa_limpia
                            in
                            consulta_limpia
                        ):

                            resultados.append(
                                {
                                    "Causa_consultada":
                                        consulta,

                                    "Causa_encontrada":
                                        fila[
                                            "Causa"
                                        ],

                                    "Patologia_ID":
                                        fila[
                                            "Patologia_ID"
                                        ],

                                    "Patologia":
                                        fila[
                                            "Patologia"
                                        ],

                                    "Tipo":
                                        "Exacta",

                                    "Puntaje":
                                        100.0
                                }
                            )

                    if resultados:
                        return resultados

# --------------------------------------------
# COINCIDENCIA APROXIMADA
# --------------------------------------------

                    coincidencias = process.extract(
                        consulta_limpia,
                        lista_causas_3_2,
                        scorer=fuzz.WRatio,
                        limit=5
                    )

                    for coincidencia in coincidencias:

                        causa_encontrada = coincidencia[0]
                        puntaje = float(
                            coincidencia[1]
                        )

                        if (
                            puntaje
                            <
                            UMBRAL_DIRECTO_CAUSA
                        ):
                            continue

                        filas = df_causas_3_2[
                            df_causas_3_2[
                                "Busqueda_3_2"
                            ]
                            ==
                            causa_encontrada
                        ]

                        for _, fila in (
                            filas.iterrows()
                        ):

                            resultados.append(
                                {
                                    "Causa_consultada":
                                        consulta,

                                    "Causa_encontrada":
                                        fila[
                                            "Causa"
                                        ],

                                    "Patologia_ID":
                                        fila[
                                            "Patologia_ID"
                                        ],

                                    "Patologia":
                                        fila[
                                            "Patologia"
                                        ],

                                    "Tipo":
                                        "Directa aproximada",

                                    "Puntaje":
                                        round(
                                            puntaje,
                                            2
                                        )
                                }
                            )

                    return resultados

# =================================================
# BÚSQUEDA SEMÁNTICA
# =================================================

                def buscar_semantica_causa_3_2(
                    consulta
                ):

                    try:

                        embedding_consulta = (
                            modelo_biomedico_causa.encode(
                                [consulta],
                                normalize_embeddings=True
                            )
                        )

                        embedding_consulta = np.asarray(
                            embedding_consulta,
                            dtype=np.float32
                        )[0]

                    except Exception as error:

                        st.error(
                            "No fue posible generar el embedding "
                            "de la causa consultada."
                        )

                        return []

                    norma_consulta = np.linalg.norm(
                        embedding_consulta
                    )

                    if norma_consulta == 0:
                        return []

                    normas_base = np.linalg.norm(
                        embeddings_causas_3_2,
                        axis=1
                    )

                    denominadores = (
                        normas_base
                        *
                        norma_consulta
                    )

                    similitudes = np.zeros(
                        len(
                            embeddings_causas_3_2
                        ),
                        dtype=np.float32
                    )

                    mascara = (
                        denominadores
                        >
                        0
                    )

                    similitudes[
                        mascara
                    ] = (
                        embeddings_causas_3_2[
                            mascara
                        ]
                        @
                        embedding_consulta
                    ) / denominadores[
                        mascara
                    ]

                    indices = np.argsort(
                        similitudes
                    )[::-1][
                        :MAX_RESULTADOS_SEMANTICOS_CAUSA
                    ]

                    resultados = []

                    for indice in indices:

                        puntaje = (
                            float(
                                similitudes[
                                    indice
                                ]
                            )
                            *
                            100
                        )

                        if (
                            puntaje
                            <
                            UMBRAL_SEMANTICO_CAUSA
                        ):
                            continue

                        fila = (
                            df_causas_3_2.iloc[
                                indice
                            ]
                        )

                        resultados.append(
                            {
                                "Causa_consultada":
                                    consulta,

                                "Causa_encontrada":
                                    fila[
                                        "Causa"
                                    ],

                                "Patologia_ID":
                                    fila[
                                        "Patologia_ID"
                                    ],

                                "Patologia":
                                    fila[
                                        "Patologia"
                                    ],

                                "Tipo":
                                    "Semántica",

                                "Puntaje":
                                    round(
                                        puntaje,
                                        2
                                    )
                            }
                        )

                    return resultados

# =================================================
# BÚSQUEDA HÍBRIDA
# =================================================

                def buscar_causa_hibrida_3_2(
                    consulta
                ):

                    resultados_directos = (
                        buscar_directa_causa_3_2(
                            consulta
                        )
                    )

                    if resultados_directos:
                        return resultados_directos

                    return (
                        buscar_semantica_causa_3_2(
                            consulta
                        )
                    )

# =================================================
# ELIMINAR DUPLICADOS
# =================================================

                def eliminar_duplicados_causa_3_2(
                    resultados
                ):

                    mejores = {}

                    for resultado in resultados:

                        clave = (
                            limpiar_causa_3_2(
                                resultado[
                                    "Causa_consultada"
                                ]
                            ),

                            str(
                                resultado[
                                    "Patologia_ID"
                                ]
                            ).strip(),

                            limpiar_causa_3_2(
                                resultado[
                                    "Causa_encontrada"
                                ]
                            )
                        )

                        if clave not in mejores:

                            mejores[
                                clave
                            ] = resultado

                        elif (
                            resultado[
                                "Puntaje"
                            ]
                            >
                            mejores[
                                clave
                            ][
                                "Puntaje"
                            ]
                        ):

                            mejores[
                                clave
                            ] = resultado

                    return list(
                        mejores.values()
                    )

# =================================================
# INGRESAR UNA O VARIAS CAUSAS
# =================================================

                texto_buscado = st.text_input(
                    "Ingrese una o varias causas:",
                    placeholder=(
                        "Ejemplo: infección, alteración hormonal"
                    ),
                    key=(
                        "texto_busqueda_causa_patologia"
                    )
                )

                st.caption(
                    "Puede ingresar varias causas "
                    "separadas por coma."
                )

# =================================================
# PROCESAR CONSULTA
# =================================================

                if texto_buscado.strip():

                    causas_consultadas = []

                    for elemento in (
                        texto_buscado.split(",")
                    ):

                        causa = (
                            elemento.strip()
                        )

                        if not causa:
                            continue

                        causa_limpia = (
                            limpiar_causa_3_2(
                                causa
                            )
                        )

                        if not causa_limpia:
                            continue

                        if not any(
                            limpiar_causa_3_2(
                                existente
                            )
                            ==
                            causa_limpia
                            for existente
                            in causas_consultadas
                        ):

                            causas_consultadas.append(
                                causa
                            )

                    if not causas_consultadas:

                        st.warning(
                            "No se ingresaron causas válidas."
                        )

                    else:

                        resultados_por_causa = {}

                        for causa in causas_consultadas:

                            resultados = (
                                buscar_causa_hibrida_3_2(
                                    causa
                                )
                            )

                            resultados = (
                                eliminar_duplicados_causa_3_2(
                                    resultados
                                )
                            )

                            resultados_por_causa[
                                causa
                            ] = resultados

# =========================================
# AGRUPAR POR PATOLOGÍA
# =========================================

                        patologias_causa = {}

                        for (
                            causa,
                            resultados
                        ) in (
                            resultados_por_causa.items()
                        ):

                            for resultado in resultados:

                                pid = str(
                                    resultado[
                                        "Patologia_ID"
                                    ]
                                ).strip()

                                if (
                                    pid
                                    not in
                                    patologias_causa
                                ):

                                    patologias_causa[
                                        pid
                                    ] = {

                                        "Patologia_ID":
                                            pid,

                                        "Patologia":
                                            str(
                                                resultado[
                                                    "Patologia"
                                                ]
                                            ).strip(),

                                        "Por_causa":
                                            {}
                                    }

                                if (
                                    causa
                                    not in
                                    patologias_causa[
                                        pid
                                    ][
                                        "Por_causa"
                                    ]
                                ):

                                    patologias_causa[
                                        pid
                                    ][
                                        "Por_causa"
                                    ][
                                        causa
                                    ] = []

                                patologias_causa[
                                    pid
                                ][
                                    "Por_causa"
                                ][
                                    causa
                                ].append(
                                    resultado
                                )

# =========================================
# EVALUAR PATOLOGÍAS
# =========================================

                        resultados_finales_causa = []

                        total_causas = len(
                            causas_consultadas
                        )

                        for (
                            pid,
                            datos
                        ) in (
                            patologias_causa.items()
                        ):

                            coincidencias_validas = []

                            for causa in causas_consultadas:

                                candidatos = (
                                    datos[
                                        "Por_causa"
                                    ].get(
                                        causa,
                                        []
                                    )
                                )

                                if not candidatos:
                                    continue

                                candidatos = sorted(
                                    candidatos,
                                    key=lambda x:
                                        x[
                                            "Puntaje"
                                        ],
                                    reverse=True
                                )

                                mejor = candidatos[0]

# ---------------------------------
# EVIDENCIA VÁLIDA
# ---------------------------------

                                if (
                                    mejor[
                                        "Tipo"
                                    ]
                                    ==
                                    "Exacta"
                                ):

                                    es_valida = True

                                elif (
                                    mejor[
                                        "Tipo"
                                    ]
                                    ==
                                    "Directa aproximada"
                                    and
                                    mejor[
                                        "Puntaje"
                                    ]
                                    >=
                                    UMBRAL_DIRECTO_CAUSA
                                ):

                                    es_valida = True

                                elif (
                                    mejor[
                                        "Tipo"
                                    ]
                                    ==
                                    "Semántica"
                                    and
                                    mejor[
                                        "Puntaje"
                                    ]
                                    >=
                                    UMBRAL_SEMANTICO_CAUSA
                                ):

                                    es_valida = True

                                else:

                                    es_valida = False

                                if es_valida:

                                    coincidencias_validas.append(
                                        mejor
                                    )

                            sintomas_respaldo = len(
                                coincidencias_validas
                            )

                            if (
                                sintomas_respaldo
                                == 0
                            ):
                                continue

                            cobertura = (
                                sintomas_respaldo
                                /
                                total_causas
                            ) * 100

                            puntajes = [
                                x[
                                    "Puntaje"
                                ]
                                for x
                                in coincidencias_validas
                            ]

                            promedio = (
                                sum(
                                    puntajes
                                )
                                /
                                len(
                                    puntajes
                                )
                            )

# ---------------------------------
# NIVEL
# ---------------------------------

                            if (
                                sintomas_respaldo
                                ==
                                total_causas
                                and
                                total_causas
                                >=
                                2
                            ):

                                nivel = (
                                    "EVIDENCIA ACUMULADA"
                                )

                            elif (
                                sintomas_respaldo
                                >=
                                2
                            ):

                                nivel = (
                                    "EVIDENCIA MULTIPLE"
                                )

                            elif (
                                promedio
                                >=
                                70
                            ):

                                nivel = (
                                    "COINCIDENCIA FUERTE"
                                )

                            else:

                                nivel = (
                                    "CANDIDATA - "
                                    "REQUIERE CONFIRMACION"
                                )

                            resultados_finales_causa.append(
                                {
                                    "Codigo":
                                        pid,

                                    "Patologia":
                                        datos[
                                            "Patologia"
                                        ],

                                    "Causas_respaldo":
                                        sintomas_respaldo,

                                    "Cobertura":
                                        round(
                                            cobertura,
                                            2
                                        ),

                                    "Promedio":
                                        round(
                                            promedio,
                                            2
                                        ),

                                    "Nivel":
                                        nivel,

                                    "Coincidencias":
                                        coincidencias_validas
                                }
                            )

# =========================================
# ORDENAR
# =========================================

                        resultados_finales_causa.sort(
                            key=lambda x: (
                                x[
                                    "Causas_respaldo"
                                ],

                                x[
                                    "Cobertura"
                                ],

                                x[
                                    "Promedio"
                                ]
                            ),
                            reverse=True
                        )

# =========================================
# MOSTRAR
# =========================================

                        if not resultados_finales_causa:

                            st.warning(
                                "No se encontraron patologías "
                                "con evidencia suficiente para "
                                "las causas ingresadas."
                            )

                        else:

                            st.success(
                                f"Se encontraron "
                                f"{len(resultados_finales_causa)} "
                                f"patologías relacionadas."
                            )

                            opciones = [
                                "Seleccione una patología"
                            ]

                            for resultado in (
                                resultados_finales_causa
                            ):

                                opciones.append(
                                    f"{resultado['Codigo']} — "
                                    f"{resultado['Patologia']} | "
                                    f"{resultado['Nivel']} | "
                                    f"{resultado['Cobertura']:.0f}%"
                                )

                            seleccion = st.selectbox(
                                "Seleccione la patología:",
                                opciones,
                                key=(
                                    "seleccion_patologia_causa_busqueda"
                                )
                            )

                            if (
                                seleccion
                                !=
                                "Seleccione una patología"
                            ):

                                indice = (
                                    opciones.index(
                                        seleccion
                                    )
                                    - 1
                                )

                                resultado = (
                                    resultados_finales_causa[
                                        indice
                                    ]
                                )

# =================================
# EVIDENCIA
# =================================

                                st.write(
                                    "**Evidencia encontrada:**"
                                )

                                col1, col2, col3 = (
                                    st.columns(3)
                                )

                                with col1:

                                    st.metric(
                                        "Causas de respaldo",
                                        (
                                            f"{resultado['Causas_respaldo']}"
                                            f"/{total_causas}"
                                        )
                                    )

                                with col2:

                                    st.metric(
                                        "Cobertura",
                                        (
                                            f"{resultado['Cobertura']:.2f}%"
                                        )
                                    )

                                with col3:

                                    st.metric(
                                        "Promedio",
                                        (
                                            f"{resultado['Promedio']:.2f}%"
                                        )
                                    )

                                st.info(
                                    f"Nivel de evidencia: "
                                    f"**{resultado['Nivel']}**"
                                )

                                st.write(
                                    "**Causas que respaldan "
                                    "la coincidencia:**"
                                )

                                for coincidencia in (
                                    resultado[
                                        "Coincidencias"
                                    ]
                                ):

                                    st.write(
                                        "• "
                                        f"**{coincidencia['Causa_consultada']}**"
                                        " → "
                                        f"{coincidencia['Causa_encontrada']}"
                                        " | "
                                        f"{coincidencia['Tipo']}"
                                        " | "
                                        f"{coincidencia['Puntaje']:.2f}%"
                                    )

# =================================
# FICHA
# =================================

                                mostrar_ficha_patologia(
                                    resultado[
                                        "Codigo"
                                    ]
                                )
# ============================================================
# 3.3 BUSCAR POR SÍNTOMA
# ============================================================

    elif (
        tipo_busqueda_causa_sintoma
        == "Síntoma"
    ):

        st.subheader(
            "Búsqueda por síntomas"
        )

# ========================================================
# CONFIGURACIÓN HÍBRIDA RECUPERADA DE COLAB
# ========================================================

        UMBRAL_DIRECTO = 82.0
        UMBRAL_SEMANTICO = 65.0
        MAX_RESULTADOS_SEMANTICOS = 5

# ========================================================
# IMPORTACIÓN LOCAL DE RAPIDFUZZ
# ========================================================

        from rapidfuzz import process, fuzz

# ========================================================
# OBTENER INFRAESTRUCTURA SEMÁNTICA
# ========================================================

        base_semantica_local = globals().get(
            "base_semantica",
            None
        )

        embeddings_sintomas_local = globals().get(
            "embeddings_sintomas",
            None
        )

        modelo_biomedico_local = globals().get(
            "modelo_biomedico",
            None
        )

# ========================================================
# VALIDAR INFRAESTRUCTURA
# ========================================================

        if base_semantica_local is None:

            st.error(
                "No está disponible la base semántica "
                "de síntomas."
            )

        elif embeddings_sintomas_local is None:

            st.error(
                "No están disponibles los embeddings "
                "de síntomas."
            )

        elif modelo_biomedico_local is None:

            st.error(
                "No está disponible el modelo biomédico."
            )

        else:

            df_3f = (
                base_semantica_local
                .copy()
            )

            columnas_necesarias = [
                "Sintoma",
                "Patologia_ID",
                "Patologia"
            ]

            columnas_faltantes = [
                columna
                for columna in columnas_necesarias
                if columna not in df_3f.columns
            ]

            if columnas_faltantes:

                st.error(
                    "Faltan columnas en la base semántica: "
                    +
                    ", ".join(
                        columnas_faltantes
                    )
                )

            else:

                # ====================================================
                # LIMPIEZA DE TEXTO
                # ====================================================

                def limpiar_texto_3f(
                    texto
                ):

                    if pd.isna(texto):

                        return ""

                    texto = str(
                        texto
                    )

                    texto = texto.lower()

                    texto = unidecode(
                        texto
                    )

                    texto = " ".join(
                        texto.split()
                    )

                    return texto.strip()

                # ====================================================
                # PREPARAR BASE DE BÚSQUEDA DIRECTA
                # ====================================================

                df_busqueda_3f = (
                    df_3f
                    .copy()
                )

                df_busqueda_3f[
                    "Busqueda_limpia"
                ] = (
                    df_busqueda_3f[
                        "Sintoma"
                    ]
                    .apply(
                        limpiar_texto_3f
                    )
                )

                lista_sintomas_directos = (
                    df_busqueda_3f[
                        "Busqueda_limpia"
                    ]
                    .drop_duplicates()
                    .tolist()
                )

                # ====================================================
                # PREPARAR EMBEDDINGS
                # ====================================================

                try:

                    embeddings_3f = np.asarray(
                        embeddings_sintomas_local,
                        dtype=np.float32
                    )

                except Exception:

                    st.error(
                        "No fue posible preparar "
                        "los embeddings de síntomas."
                    )

                    st.stop()

                if (
                    embeddings_3f.ndim
                    !=
                    2
                ):

                    st.error(
                        "Los embeddings no tienen "
                        "el formato esperado."
                    )

                elif (
                    len(
                        embeddings_3f
                    )
                    !=
                    len(
                        df_busqueda_3f
                    )
                ):

                    st.error(
                        "La cantidad de embeddings "
                        "no coincide con la cantidad "
                        "de registros de la base semántica."
                    )

                else:

                    # =================================================
                    # BÚSQUEDA DIRECTA + APROXIMADA
                    # =================================================

                    def buscar_directo_3f(
                        consulta
                    ):

                        consulta_limpia = (
                            limpiar_texto_3f(
                                consulta
                            )
                        )

                        resultados = []

                        # ---------------------------------------------
                        # 1. COINCIDENCIA EXACTA O CONTENIDA
                        # ---------------------------------------------

                        for _, fila in (
                            df_busqueda_3f.iterrows()
                        ):

                            sintoma_limpio = (
                                fila[
                                    "Busqueda_limpia"
                                ]
                            )

                            if (
                                consulta_limpia
                                ==
                                sintoma_limpio
                                or
                                consulta_limpia
                                in
                                sintoma_limpio
                                or
                                sintoma_limpio
                                in
                                consulta_limpia
                            ):

                                resultados.append(
                                    {
                                        "Sintoma_consultado":
                                            consulta,

                                        "Sintoma_encontrado":
                                            fila[
                                                "Sintoma"
                                            ],

                                        "Patologia_ID":
                                            fila[
                                                "Patologia_ID"
                                            ],

                                        "Patologia":
                                            fila[
                                                "Patologia"
                                            ],

                                        "Tipo":
                                            "Directa",

                                        "Puntaje":
                                            100.0
                                    }
                                )

                        if resultados:

                            return resultados

                        # ---------------------------------------------
                        # 2. COINCIDENCIA DIRECTA APROXIMADA
                        # ---------------------------------------------

                        coincidencias = process.extract(
                            consulta_limpia,
                            lista_sintomas_directos,
                            scorer=fuzz.WRatio,
                            limit=3
                        )

                        for coincidencia in (
                            coincidencias
                        ):

                            sintoma_encontrado = (
                                coincidencia[0]
                            )

                            puntaje = float(
                                coincidencia[1]
                            )

                            if (
                                puntaje
                                <
                                UMBRAL_DIRECTO
                            ):

                                continue

                            filas = (
                                df_busqueda_3f[
                                    df_busqueda_3f[
                                        "Busqueda_limpia"
                                    ]
                                    ==
                                    sintoma_encontrado
                                ]
                            )

                            for _, fila in (
                                filas.iterrows()
                            ):

                                resultados.append(
                                    {
                                        "Sintoma_consultado":
                                            consulta,

                                        "Sintoma_encontrado":
                                            fila[
                                                "Sintoma"
                                            ],

                                        "Patologia_ID":
                                            fila[
                                                "Patologia_ID"
                                            ],

                                        "Patologia":
                                            fila[
                                                "Patologia"
                                            ],

                                        "Tipo":
                                            "Directa aproximada",

                                        "Puntaje":
                                            round(
                                                puntaje,
                                                2
                                            )
                                    }
                                )

                        return resultados

                    # =================================================
                    # BÚSQUEDA SEMÁNTICA CON EMBEDDINGS
                    # =================================================

                    def buscar_semantica_3f(
                        consulta
                    ):

                        try:

                            embedding_consulta = (
                                modelo_biomedico_local.encode(
                                    [consulta],
                                    normalize_embeddings=True
                                )
                            )

                            embedding_consulta = np.asarray(
                                embedding_consulta,
                                dtype=np.float32
                            )

                        except Exception:

                            return []

                        if (
                            embedding_consulta.ndim
                            !=
                            2
                        ):

                            return []

                        vector_consulta = (
                            embedding_consulta[0]
                        )

                        norma_consulta = np.linalg.norm(
                            vector_consulta
                        )

                        if (
                            norma_consulta
                            ==
                            0
                        ):

                            return []

                        normas_base = np.linalg.norm(
                            embeddings_3f,
                            axis=1
                        )

                        denominadores = (
                            normas_base
                            *
                            norma_consulta
                        )

                        similitudes = np.zeros(
                            len(
                                embeddings_3f
                            ),
                            dtype=np.float32
                        )

                        mascara = (
                            denominadores
                            >
                            0
                        )

                        similitudes[
                            mascara
                        ] = (
                            embeddings_3f[
                                mascara
                            ]
                            @
                            vector_consulta
                        ) / denominadores[
                            mascara
                        ]

                        indices = np.argsort(
                            similitudes
                        )[::-1][
                            :MAX_RESULTADOS_SEMANTICOS
                        ]

                        resultados = []

                        for indice in indices:

                            puntaje = (
                                float(
                                    similitudes[
                                        indice
                                    ]
                                )
                                *
                                100
                            )

                            if (
                                puntaje
                                <
                                UMBRAL_SEMANTICO
                            ):

                                continue

                            fila = (
                                df_busqueda_3f.iloc[
                                    indice
                                ]
                            )

                            resultados.append(
                                {
                                    "Sintoma_consultado":
                                        consulta,

                                    "Sintoma_encontrado":
                                        fila[
                                            "Sintoma"
                                        ],

                                    "Patologia_ID":
                                        fila[
                                            "Patologia_ID"
                                        ],

                                    "Patologia":
                                        fila[
                                            "Patologia"
                                        ],

                                    "Tipo":
                                        "Semantica",

                                    "Puntaje":
                                        round(
                                            puntaje,
                                            2
                                        )
                                }
                            )

                        return resultados

                    # =================================================
                    # BÚSQUEDA HÍBRIDA
                    # =================================================

                    def buscar_sintoma_3f(
                        consulta
                    ):

                        resultados_directos = (
                            buscar_directo_3f(
                                consulta
                            )
                        )

                        # Primero directa/aproximada.
                        if resultados_directos:

                            return resultados_directos

                        # Solo si no hubo evidencia textual,
                        # consultar embeddings.
                        return buscar_semantica_3f(
                            consulta
                        )

                    # =================================================
                    # ELIMINAR DUPLICADOS CONSERVANDO EL MEJOR
                    # =================================================

                    def eliminar_duplicados_3f(
                        resultados
                    ):

                        mejores = {}

                        for resultado in resultados:

                            clave = (
                                limpiar_texto_3f(
                                    resultado[
                                        "Sintoma_consultado"
                                    ]
                                ),

                                str(
                                    resultado[
                                        "Patologia_ID"
                                    ]
                                ).strip(),

                                limpiar_texto_3f(
                                    resultado[
                                        "Sintoma_encontrado"
                                    ]
                                )
                            )

                            if (
                                clave
                                not in
                                mejores
                            ):

                                mejores[
                                    clave
                                ] = resultado

                                continue

                            if (
                                resultado[
                                    "Puntaje"
                                ]
                                >
                                mejores[
                                    clave
                                ][
                                    "Puntaje"
                                ]
                            ):

                                mejores[
                                    clave
                                ] = resultado

                        return list(
                            mejores.values()
                        )

                    # =================================================
                    # INGRESAR SÍNTOMAS
                    # =================================================

                    texto_buscado = st.text_input(
                        "Ingrese uno o varios síntomas o señales:",
                        placeholder=(
                            "Ejemplo: dificultad para orinar, "
                            "mal olor"
                        ),
                        key=(
                            "texto_busqueda_sintoma_patologia"
                        )
                    )

                    st.caption(
                        "Puede ingresar uno o varios síntomas. "
                        "Sepárelos por coma."
                    )

                    if texto_buscado.strip():

                        sintomas = []

                        for elemento in (
                            texto_buscado.split(",")
                        ):

                            sintoma = (
                                elemento.strip()
                            )

                            if not sintoma:

                                continue

                            sintoma_limpio = (
                                limpiar_texto_3f(
                                    sintoma
                                )
                            )

                            if not sintoma_limpio:

                                continue

                            if not any(
                                limpiar_texto_3f(
                                    existente
                                )
                                ==
                                sintoma_limpio
                                for existente
                                in sintomas
                            ):

                                sintomas.append(
                                    sintoma
                                )

                        if not sintomas:

                            st.warning(
                                "No se ingresaron síntomas "
                                "válidos."
                            )

                        else:

                            # =================================================
                            # BUSCAR CADA SÍNTOMA INDEPENDIENTEMENTE
                            # =================================================

                            resultados_por_sintoma = {}

                            for sintoma in sintomas:

                                resultados = (
                                    buscar_sintoma_3f(
                                        sintoma
                                    )
                                )

                                resultados = (
                                    eliminar_duplicados_3f(
                                        resultados
                                    )
                                )

                                resultados_por_sintoma[
                                    sintoma
                                ] = resultados

                            # =================================================
                            # INTEGRAR POR PATOLOGÍA
                            # =================================================

                            patologias = {}

                            for (
                                sintoma,
                                resultados
                            ) in (
                                resultados_por_sintoma.items()
                            ):

                                for resultado in resultados:

                                    pid = str(
                                        resultado[
                                            "Patologia_ID"
                                        ]
                                    ).strip()

                                    if (
                                        pid
                                        not in
                                        patologias
                                    ):

                                        patologias[
                                            pid
                                        ] = {
                                            "Patologia_ID":
                                                pid,

                                            "Patologia":
                                                str(
                                                    resultado[
                                                        "Patologia"
                                                    ]
                                                ).strip(),

                                            "Por_sintoma":
                                                {}
                                        }

                                    patologias[
                                        pid
                                    ][
                                        "Por_sintoma"
                                    ].setdefault(
                                        sintoma,
                                        []
                                    ).append(
                                        resultado
                                    )

                            # =================================================
                            # EVALUAR PATOLOGÍAS POR CONJUNTO DE SÍNTOMAS
                            # =================================================

                            # La matriz define dinámicamente qué síntomas
                            # pertenecen a cada patología. No se utilizan
                            # categorías rígidas de enfermedades.

                            resultados_finales = []

                            total_sintomas = len(
                                sintomas
                            )

                            for (
                                pid,
                                datos
                            ) in (
                                patologias.items()
                            ):

                                coincidencias_validas = []

                                for sintoma in sintomas:

                                    candidatos = (
                                        datos[
                                            "Por_sintoma"
                                        ].get(
                                            sintoma,
                                            []
                                        )
                                    )

                                    if not candidatos:

                                        continue

                                    # ---------------------------------------------
                                    # CONSERVAR LA MEJOR COINCIDENCIA DEL SÍNTOMA
                                    # DENTRO DE ESTA PATOLOGÍA
                                    # ---------------------------------------------

                                    candidatos = sorted(
                                        candidatos,
                                        key=lambda x:
                                            x[
                                                "Puntaje"
                                            ],
                                        reverse=True
                                    )

                                    mejor = candidatos[0]

                                    if (
                                        mejor[
                                            "Tipo"
                                        ]
                                        in
                                        (
                                            "Directa",
                                            "Directa aproximada"
                                        )
                                    ):

                                        es_valida = (
                                            mejor[
                                                "Puntaje"
                                            ]
                                            >=
                                            UMBRAL_DIRECTO
                                        )

                                    else:

                                        es_valida = (
                                            mejor[
                                                "Puntaje"
                                            ]
                                            >=
                                            UMBRAL_SEMANTICO
                                        )

                                    if es_valida:

                                        coincidencias_validas.append(
                                            mejor
                                        )

                                # ---------------------------------------------
                                # SIN COINCIDENCIAS: NO ES CANDIDATA
                                # ---------------------------------------------

                                if not coincidencias_validas:

                                    continue

                                sintomas_respaldo = len(
                                    coincidencias_validas
                                )

                                cobertura = (
                                    sintomas_respaldo
                                    /
                                    total_sintomas
                                ) * 100

                                puntajes = [
                                    x[
                                        "Puntaje"
                                    ]
                                    for x
                                    in coincidencias_validas
                                ]

                                promedio = (
                                    sum(
                                        puntajes
                                    )
                                    /
                                    len(
                                        puntajes
                                    )
                                )

                                mejor_puntaje = max(
                                    puntajes
                                )

                                # ---------------------------------------------
                                # COHERENCIA DEL CONJUNTO
                                # ---------------------------------------------
                                # Con múltiples síntomas, una patología no debe
                                # entrar al resultado principal solo porque
                                # comparte un síntoma genérico como "dolor".
                                # Necesitamos evidencia de al menos dos síntomas
                                # de la consulta, salvo una coincidencia aislada
                                # excepcionalmente fuerte.

                                if total_sintomas >= 2:

                                    if (
                                        sintomas_respaldo
                                        >=
                                        2
                                        and
                                        promedio
                                        >=
                                        70.0
                                    ):

                                        es_candidata = True

                                    elif (
                                        sintomas_respaldo
                                        ==
                                        total_sintomas
                                        and
                                        promedio
                                        >=
                                        65.0
                                    ):

                                        es_candidata = True

                                    else:

                                        es_candidata = False

                                else:

                                    # Para un único síntoma sí permitimos
                                    # candidatos, pero solamente si la evidencia
                                    # individual es suficientemente fuerte.
                                    es_candidata = (
                                        mejor_puntaje
                                        >=
                                        70.0
                                    )

                                if not es_candidata:

                                    continue

                                # ---------------------------------------------
                                # NIVEL DE EVIDENCIA
                                # ---------------------------------------------

                                if (
                                    sintomas_respaldo
                                    ==
                                    total_sintomas
                                    and
                                    total_sintomas
                                    >=
                                    2
                                    and
                                    promedio
                                    >=
                                    75.0
                                ):

                                    nivel = (
                                        "EVIDENCIA ACUMULADA"
                                    )

                                elif (
                                    sintomas_respaldo
                                    >=
                                    2
                                    and
                                    promedio
                                    >=
                                    70.0
                                ):

                                    nivel = (
                                        "EVIDENCIA MÚLTIPLE COHERENTE"
                                    )

                                elif (
                                    sintomas_respaldo
                                    ==
                                    1
                                    and
                                    promedio
                                    >=
                                    85.0
                                ):

                                    nivel = (
                                        "COINCIDENCIA FUERTE AISLADA"
                                    )

                                else:

                                    nivel = (
                                        "CANDIDATA - REQUIERE CONFIRMACIÓN"
                                    )

                                resultados_finales.append(
                                    {
                                        "Patologia_ID":
                                            pid,

                                        "Patologia":
                                            datos[
                                                "Patologia"
                                            ],

                                        "Sintomas_respaldo":
                                            sintomas_respaldo,

                                        "Cobertura":
                                            round(
                                                cobertura,
                                                2
                                            ),

                                        "Promedio":
                                            round(
                                                promedio,
                                                2
                                            ),

                                        "Mejor_puntaje":
                                            round(
                                                mejor_puntaje,
                                                2
                                            ),

                                        "Nivel":
                                            nivel,

                                        "Coincidencias":
                                            coincidencias_validas
                                    }
                                )

                            # =================================================
                            # ORDENAR POR EVIDENCIA
                            # =================================================

                            resultados_finales.sort(
                                key=lambda x: (
                                    x[
                                        "Sintomas_respaldo"
                                    ],

                                    x[
                                        "Cobertura"
                                    ],

                                    x[
                                        "Promedio"
                                    ],

                                    x[
                                        "Mejor_puntaje"
                                    ]
                                ),
                                reverse=True
                            )

                            # =================================================
                            # MOSTRAR RESULTADOS
                            # =================================================

                            st.subheader(
                                "Resultados de la búsqueda"
                            )

                            if not resultados_finales:

                                st.warning(
                                    "No se encontró evidencia "
                                    "suficiente para los síntomas ingresados."
                                )

                                st.info(
                                    "Pruebe describiendo el síntoma "
                                    "con otras palabras o agregando "
                                    "otro síntoma relacionado."
                                )

                            else:

                                st.success(
                                    f"Se encontraron "
                                    f"{len(resultados_finales)} "
                                    f"patologías candidatas."
                                )

                                opciones = [
                                    "Seleccione una patología"
                                ]

                                for resultado in (
                                    resultados_finales
                                ):

                                    opciones.append(
                                        f"{resultado['Patologia_ID']} — "
                                        f"{resultado['Patologia']} | "
                                        f"{resultado['Nivel']} | "
                                        f"{resultado['Sintomas_respaldo']}/{total_sintomas} síntomas"
                                    )

                                seleccion = st.selectbox(
                                    "Seleccione la patología:",
                                    opciones,
                                    key=(
                                        "seleccion_patologia_sintoma_busqueda"
                                    )
                                )

                                if (
                                    seleccion
                                    !=
                                    "Seleccione una patología"
                                ):

                                    indice = (
                                        opciones.index(
                                            seleccion
                                        )
                                        -
                                        1
                                    )

                                    resultado = (
                                        resultados_finales[
                                            indice
                                        ]
                                    )

                                    st.write(
                                        "**Evidencia encontrada:**"
                                    )

                                    col1, col2, col3, col4 = (
                                        st.columns(4)
                                    )

                                    with col1:

                                        st.metric(
                                            "Síntomas de respaldo",
                                            (
                                                f"{resultado['Sintomas_respaldo']}"
                                                f"/{total_sintomas}"
                                            )
                                        )

                                    with col2:

                                        st.metric(
                                            "Cobertura",
                                            (
                                                f"{resultado['Cobertura']:.2f}%"
                                            )
                                        )

                                    with col3:

                                        st.metric(
                                            "Promedio",
                                            (
                                                f"{resultado['Promedio']:.2f}%"
                                            )
                                        )

                                    with col4:

                                        st.metric(
                                            "Mejor coincidencia",
                                            (
                                                f"{resultado['Mejor_puntaje']:.2f}%"
                                            )
                                        )

                                    st.info(
                                        f"Nivel de evidencia: "
                                        f"**{resultado['Nivel']}**"
                                    )

                                    st.write(
                                        "**Síntomas que respaldan la coincidencia:**"
                                    )

                                    for coincidencia in (
                                        resultado[
                                            "Coincidencias"
                                        ]
                                    ):

                                        st.write(
                                            "• "
                                            f"**{coincidencia['Sintoma_consultado']}**"
                                            " → "
                                            f"{coincidencia['Sintoma_encontrado']}"
                                            " | "
                                            f"{coincidencia['Tipo']}"
                                            " | "
                                            f"{coincidencia['Puntaje']:.2f}%"
                                        )

                                    sintomas_resueltos = [
                                        limpiar_texto_3f(
                                            x[
                                                "Sintoma_consultado"
                                            ]
                                        )
                                        for x
                                        in resultado[
                                            "Coincidencias"
                                        ]
                                    ]

                                    sintomas_sin_coincidencia = [
                                        sintoma
                                        for sintoma
                                        in sintomas
                                        if (
                                            limpiar_texto_3f(
                                                sintoma
                                            )
                                            not in
                                            sintomas_resueltos
                                        )
                                    ]

                                    if (
                                        sintomas_sin_coincidencia
                                    ):

                                        st.warning(
                                            "Síntomas sin coincidencia "
                                            "suficiente:"
                                        )

                                        for sintoma in (
                                            sintomas_sin_coincidencia
                                        ):

                                            st.write(
                                                f"• {sintoma}"
                                            )

                                    if total_sintomas == 1:

                                        st.warning(
                                            "Con un solo síntoma, el resultado "
                                            "debe considerarse orientativo y requiere "
                                            "confirmación con más información."
                                        )

                                    elif (
                                        resultado[
                                            "Sintomas_respaldo"
                                        ]
                                        <
                                        total_sintomas
                                    ):

                                        st.warning(
                                            "La patología seleccionada está respaldada "
                                            "solo por una parte de los síntomas ingresados. "
                                            "Agregar otro síntoma puede aumentar la precisión."
                                        )

                                    mostrar_ficha_patologia(
                                        resultado[
                                            "Patologia_ID"
                                        ]
                                    )

# =============================================
# NAVEGACIÓN
# =============================================

        st.divider()

        siguiente_accion_sintoma = st.selectbox(
            "¿Qué desea hacer ahora?",
            [
                "Seleccione una opción",
                "Realizar otra búsqueda",
                "Ir al menú principal"
            ],
            key=(
                "navegacion_sintoma_patologia"
            )
        )

        if (
            siguiente_accion_sintoma
            == "Realizar otra búsqueda"
        ):

            st.info(
                "Ingrese nuevamente uno o varios síntomas "
                "separados por coma."
            )

        elif (
            siguiente_accion_sintoma
            == "Ir al menú principal"
        ):

            st.session_state[
                "volver_menu_principal"
            ] = True

            st.rerun()
# ============================================================
# BLOQUE — RESTRICCIONES
# ============================================================

if opcion_consulta == "Restricciones":

    st.subheader("Consulta de restricciones")

    # ========================================================
    # MENÚ DE CONSULTA DE RESTRICCIONES
    # ========================================================

    tipo_consulta_restriccion = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Restricción → precaución / contraindicación",
            "Producto → motivo y alternativa"
        ],
        key="tipo_consulta_restriccion"
    )

    # ========================================================
    # CONSULTA 3 — RESTRICCIÓN → PRECAUCIÓN / CONTRAINDICACIÓN
    # ========================================================

    if (
        tipo_consulta_restriccion
        == "Restricción → precaución / contraindicación"
    ):

        st.write(
            "Ingrese la precaución o contraindicación que desea buscar:"
        )

        texto_busqueda_restriccion = st.text_input(
            "Buscar restricción:",
            key="busqueda_restriccion_precaucion"
        )

        if texto_busqueda_restriccion.strip():

            consulta = (
                texto_busqueda_restriccion
                .strip()
                .lower()
            )

            resultados_restricciones = []

            # =================================================
            # BUSCAR EN LA COLUMNA
            # PRECAUCIÓN / CONTRAINDICACIÓN
            # =================================================

            for _, fila in Restricciones.iterrows():

                codigo = fila.iloc[0]
                producto = fila.iloc[1]
                tipo = fila.iloc[2]
                restriccion = fila.iloc[3]

                if pd.isna(restriccion):
                    continue

                restriccion_texto = (
                    str(restriccion)
                    .strip()
                )

                if not restriccion_texto:
                    continue

                # ---------------------------------------------
                # COINCIDENCIA DIRECTA
                # ---------------------------------------------

                if consulta in restriccion_texto.lower():

                    resultados_restricciones.append(
                        {
                            "codigo": codigo,
                            "producto": producto,
                            "tipo": tipo,
                            "restriccion": restriccion_texto
                        }
                    )

            # =================================================
            # BÚSQUEDA TOLERANTE A ERRORES
            # =================================================

            if not resultados_restricciones:

                candidatos_restricciones = []

                for _, fila in Restricciones.iterrows():

                    codigo = fila.iloc[0]
                    producto = fila.iloc[1]
                    tipo = fila.iloc[2]
                    restriccion = fila.iloc[3]

                    if pd.isna(restriccion):
                        continue

                    restriccion_texto = (
                        str(restriccion)
                        .strip()
                    )

                    if not restriccion_texto:
                        continue

                    puntuacion = fuzz.partial_ratio(
                        consulta,
                        restriccion_texto.lower()
                    )

                    if puntuacion >= 60:

                        candidatos_restricciones.append(
                            (
                                codigo,
                                producto,
                                tipo,
                                restriccion_texto,
                                puntuacion
                            )
                        )

                candidatos_restricciones.sort(
                    key=lambda x: x[4],
                    reverse=True
                )

                for (
                    codigo,
                    producto,
                    tipo,
                    restriccion_texto,
                    puntuacion
                ) in candidatos_restricciones[:20]:

                    resultados_restricciones.append(
                        {
                            "codigo": codigo,
                            "producto": producto,
                            "tipo": tipo,
                            "restriccion": restriccion_texto
                        }
                    )

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not resultados_restricciones:

                st.warning(
                    "No se encontraron productos relacionados "
                    "con esa precaución o contraindicación."
                )

            else:

                # ---------------------------------------------
                # OBTENER PRODUCTOS ÚNICOS
                # ---------------------------------------------

                productos_encontrados = {}

                for resultado in resultados_restricciones:

                    producto = resultado["producto"]

                    if pd.isna(producto):
                        continue

                    producto = str(producto).strip()

                    if not producto:
                        continue

                    clave = producto.lower()

                    if clave not in productos_encontrados:

                        productos_encontrados[
                            clave
                        ] = producto

                productos_ordenados = sorted(
                    productos_encontrados.values(),
                    key=lambda x: x.lower()
                )

                # ---------------------------------------------
                # LISTADO DE PRODUCTOS
                # ---------------------------------------------

                st.write(
                    "Productos relacionados encontrados:"
                )

                opciones_productos = [
                    "Seleccione un producto"
                ]

                opciones_productos.extend(
                    productos_ordenados
                )

                producto_seleccionado = st.selectbox(
                    "Seleccione el producto que desea consultar:",
                    opciones_productos,
                    key="producto_resultado_restriccion"
                )

                # =================================================
                # MOSTRAR FICHA COMPLETA DEL PRODUCTO
                # =================================================

                if (
                    producto_seleccionado
                    != "Seleccione un producto"
                ):

                    producto_normalizado = (
                        producto_seleccionado
                        .lower()
                        .strip()
                    )

                    ficha = Restricciones[
                        Restricciones.iloc[:, 1]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        == producto_normalizado
                    ]

                    if ficha.empty:

                        st.warning(
                            "No se encontraron restricciones "
                            "para este producto."
                        )

                    else:

                        st.divider()

                        st.subheader(
                            "Ficha de restricciones del producto"
                        )

                        st.write(
                            f"**Producto:** "
                            f"{producto_seleccionado}"
                        )

                        # =========================================
                        # MOSTRAR TODAS LAS RESTRICCIONES DEL PRODUCTO
                        # =========================================

                        for _, datos in ficha.iterrows():

                            st.markdown("---")

                            st.write(
                                f"**Restricción ID:** "
                                f"{datos.iloc[0]}"
                            )

                            st.write(
                                f"**Tipo:** "
                                f"{datos.iloc[2]}"
                            )

                            st.write(
                                f"**Precaución / "
                                f"Contraindicación:** "
                                f"{datos.iloc[3]}"
                            )

                            st.write(
                                f"**Motivo:** "
                                f"{datos.iloc[4]}"
                            )
    
    # ========================================================
    # CONSULTA 4 — PRODUCTO → MOTIVO Y ALTERNATIVA
    # ========================================================

    if (
        tipo_consulta_restriccion
        == "Producto → motivo y alternativa"
    ):

        st.write(
            "Ingrese el nombre o código del producto:"
        )

        texto_busqueda_producto_4 = st.text_input(
            "Buscar producto:",
            key="busqueda_producto_motivo_alternativa"
        )

        if texto_busqueda_producto_4.strip():

            consulta = (
                texto_busqueda_producto_4
                .strip()
                .lower()
            )

            # =================================================
            # BUSCAR PRODUCTOS
            # =================================================

            productos_encontrados_4 = {}

            for _, fila in Restricciones.iterrows():

                codigo = fila.iloc[0]
                producto = fila.iloc[1]

                if pd.isna(codigo) or pd.isna(producto):
                    continue

                codigo = str(codigo).strip()
                producto = str(producto).strip()

                if not codigo or not producto:
                    continue

                codigo_normalizado = codigo.lower()
                producto_normalizado = producto.lower()

                # ---------------------------------------------
                # COINCIDENCIA DIRECTA
                # ---------------------------------------------

                if (
                    consulta in codigo_normalizado
                    or consulta in producto_normalizado
                ):

                    productos_encontrados_4[
                        producto_normalizado
                    ] = producto

            # =================================================
            # BÚSQUEDA TOLERANTE A ERRORES
            # =================================================

            if not productos_encontrados_4:

                candidatos_4 = {}

                for _, fila in Restricciones.iterrows():

                    codigo = fila.iloc[0]
                    producto = fila.iloc[1]

                    if pd.isna(codigo) or pd.isna(producto):
                        continue

                    codigo = str(codigo).strip()
                    producto = str(producto).strip()

                    if not codigo or not producto:
                        continue

                    puntuacion_producto = fuzz.partial_ratio(
                        consulta,
                        producto.lower()
                    )

                    puntuacion_codigo = fuzz.partial_ratio(
                        consulta,
                        codigo.lower()
                    )

                    puntuacion = max(
                        puntuacion_producto,
                        puntuacion_codigo
                    )

                    if puntuacion >= 60:

                        clave = producto.lower()

                        if (
                            clave not in candidatos_4
                            or puntuacion
                            > candidatos_4[clave][1]
                        ):

                            candidatos_4[clave] = (
                                producto,
                                puntuacion
                            )

                candidatos_ordenados_4 = sorted(
                    candidatos_4.values(),
                    key=lambda x: x[1],
                    reverse=True
                )

                for producto, puntuacion in (
                    candidatos_ordenados_4[:10]
                ):

                    productos_encontrados_4[
                        producto.lower()
                    ] = producto

            # =================================================
            # MOSTRAR PRODUCTOS ENCONTRADOS
            # =================================================

            if not productos_encontrados_4:

                st.warning(
                    "No se encontraron productos "
                    "relacionados con la búsqueda."
                )

            else:

                st.write(
                    "Seleccione el producto que desea consultar:"
                )

                opciones_productos_4 = [
                    "Seleccione un producto"
                ]

                opciones_productos_4.extend(
                    sorted(
                        productos_encontrados_4.values(),
                        key=lambda x: x.lower()
                    )
                )

                producto_seleccionado_4 = st.selectbox(
                    "Productos encontrados:",
                    opciones_productos_4,
                    key="resultado_producto_motivo_alternativa"
                )

                # =================================================
                # MOSTRAR MOTIVO Y ALTERNATIVAS
                # =================================================

                if (
                    producto_seleccionado_4
                    != "Seleccione un producto"
                ):

                    producto_normalizado_4 = (
                        producto_seleccionado_4
                        .lower()
                        .strip()
                    )

                    ficha_4 = Restricciones[
                        Restricciones.iloc[:, 1]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        == producto_normalizado_4
                    ]

                    if ficha_4.empty:

                        st.warning(
                            "No se encontraron restricciones "
                            "para este producto."
                        )

                    else:

                        st.divider()

                        st.subheader(
                            "Motivos y alternativas del producto"
                        )

                        st.write(
                            f"**Producto:** "
                            f"{producto_seleccionado_4}"
                        )

                        # =========================================
                        # MOSTRAR CADA RESTRICCIÓN
                        # =========================================

                        for _, datos in ficha_4.iterrows():

                            st.markdown("---")

                            st.write(
                                f"**Restricción ID:** "
                                f"{datos.iloc[0]}"
                            )

                            st.write(
                                f"**Tipo:** "
                                f"{datos.iloc[2]}"
                            )

                            st.write(
                                f"**Precaución / "
                                f"Contraindicación:** "
                                f"{datos.iloc[3]}"
                            )

                            st.write(
                                f"**Motivo:** "
                                f"{datos.iloc[4]}"
                            )

                            st.write(
                                f"**Alternativas seguras:** "
                                f"{datos.iloc[5]}"
                            )

                        # =========================================
                        # NAVEGACIÓN
                        # =========================================

                        st.divider()

                        siguiente_accion_4 = st.selectbox(
                            "¿Qué desea hacer ahora?",
                            [
                                "Seleccione una opción",
                                "Seleccionar otro producto",
                                "Realizar otra búsqueda",
                                "Ir al menú principal"
                            ],
                            key="navegacion_restricciones_4"
                        )

                        if (
                            siguiente_accion_4
                            == "Seleccionar otro producto"
                        ):

                            st.info(
                                "Puede seleccionar otro "
                                "producto de los resultados."
                            )

                        elif (
                            siguiente_accion_4
                            == "Realizar otra búsqueda"
                        ):

                            st.info(
                                "Ingrese un nuevo "
                                "producto o código."
                            )

                        elif (
                            siguiente_accion_4
                            == "Ir al menú principal"
                        ):

                            st.session_state[
                                "volver_menu_principal"
                            ] = True

                            st.rerun()
    # ========================================================
# ============================================================
# ============================================================
# MÓDULO — COMPLEMENTARIOS
# ============================================================

if opcion_consulta == "Complementarios":

    st.subheader("Consulta de productos complementarios")

    tipo_consulta_complementario = st.selectbox(
        "¿Qué desea consultar?",
        [
            "Seleccione una opción",
            "Ver todos los productos",
            "Ingresar nombre del producto"
        ],
        key="menu_consulta_complementarios"
    )

    # ========================================================
    # CONSULTA 1 — VER TODOS LOS PRODUCTOS
    # ========================================================

    if (
        tipo_consulta_complementario
        == "Ver todos los productos"
    ):

        st.write(
            "Seleccione el producto que desea consultar:"
        )

        productos_unicos = {}

        for _, fila in Complementarios.iterrows():

            producto = fila.iloc[0]

            if pd.isna(producto):
                continue

            producto = str(producto).strip()

            if not producto:
                continue

            clave = producto.lower().strip()

            if clave not in productos_unicos:

                productos_unicos[clave] = producto

        productos_ordenados = sorted(
            productos_unicos.values(),
            key=lambda x: x.lower().strip()
        )

        opciones_productos = [
            "Seleccione un producto"
        ]

        opciones_productos.extend(
            productos_ordenados
        )

        producto_seleccionado = st.selectbox(
            "Productos disponibles:",
            opciones_productos,
            key="producto_complementario_lista"
        )

        # ====================================================
        # MOSTRAR FICHA
        # ====================================================

        if (
            producto_seleccionado
            != "Seleccione un producto"
        ):

            producto_normalizado = (
                producto_seleccionado
                .lower()
                .strip()
            )

            ficha = Complementarios[
                Complementarios.iloc[:, 0]
                .astype(str)
                .str.lower()
                .str.strip()
                == producto_normalizado
            ]

            if ficha.empty:

                st.warning(
                    "No se encontró información "
                    "para este producto."
                )

            else:

                datos = ficha.iloc[0]

                st.divider()

                st.subheader(
                    "Ficha del producto complementario"
                )

                st.write(
                    f"**Producto:** "
                    f"{datos.iloc[0]}"
                )

                st.write(
                    f"**Categoría principal:** "
                    f"{datos.iloc[1]}"
                )

                st.write(
                    f"**Indicaciones / Escenarios:** "
                    f"{datos.iloc[2]}"
                )

                st.write(
                    f"**Modo de acción resumido:** "
                    f"{datos.iloc[3]}"
                )

                st.write(
                    f"**Combinaciones estratégicas:** "
                    f"{datos.iloc[4]}"
                )

                # ============================================
                # NAVEGACIÓN
                # ============================================

                st.divider()

                siguiente_accion = st.selectbox(
                    "¿Qué desea hacer ahora?",
                    [
                        "Seleccione una opción",
                        "Seleccionar otro producto",
                        "Realizar otra búsqueda",
                        "Volver al menú de Complementarios",
                        "Ir al menú principal"
                    ],
                    key="navegacion_complementarios_lista"
                )

                if (
                    siguiente_accion
                    == "Seleccionar otro producto"
                ):

                    st.info(
                        "Seleccione otro producto "
                        "del listado."
                    )

                elif (
                    siguiente_accion
                    == "Realizar otra búsqueda"
                ):

                    st.info(
                        "Seleccione la opción "
                        "'Ingresar nombre del producto'."
                    )

                elif (
                    siguiente_accion
                    == "Volver al menú de Complementarios"
                ):

                    st.info(
                        "Seleccione nuevamente "
                        "el tipo de consulta."
                    )

                elif (
                    siguiente_accion
                    == "Ir al menú principal"
                ):

                    st.session_state[
                        "volver_menu_principal"
                    ] = True

                    st.rerun()


    # ========================================================
    # CONSULTA 2 — INGRESAR NOMBRE DEL PRODUCTO
    # ========================================================

    if (
        tipo_consulta_complementario
        == "Ingresar nombre del producto"
    ):

        st.write(
            "Ingrese el nombre del producto:"
        )

        texto_busqueda = st.text_input(
            "Buscar producto:",
            key="busqueda_complementario_manual"
        )

        if texto_busqueda.strip():

            consulta = (
                texto_busqueda
                .strip()
                .lower()
            )

            productos_encontrados = {}

            # =================================================
            # 1. COINCIDENCIA DIRECTA
            # =================================================

            for _, fila in Complementarios.iterrows():

                producto = fila.iloc[0]

                if pd.isna(producto):
                    continue

                producto = str(producto).strip()

                if not producto:
                    continue

                producto_normalizado = (
                    producto.lower()
                )

                if consulta in producto_normalizado:

                    productos_encontrados[
                        producto_normalizado
                    ] = producto

            # =================================================
            # 2. BÚSQUEDA TOLERANTE A ERRORES
            # =================================================

            if not productos_encontrados:

                candidatos = {}

                for _, fila in Complementarios.iterrows():

                    producto = fila.iloc[0]

                    if pd.isna(producto):
                        continue

                    producto = str(producto).strip()

                    if not producto:
                        continue

                    puntuacion = fuzz.partial_ratio(
                        consulta,
                        producto.lower()
                    )

                    if puntuacion >= 60:

                        clave = producto.lower()

                        if (
                            clave not in candidatos
                            or puntuacion
                            > candidatos[clave][1]
                        ):

                            candidatos[clave] = (
                                producto,
                                puntuacion
                            )

                candidatos_ordenados = sorted(
                    candidatos.values(),
                    key=lambda x: x[1],
                    reverse=True
                )

                for producto, puntuacion in (
                    candidatos_ordenados[:10]
                ):

                    productos_encontrados[
                        producto.lower()
                    ] = producto

            # =================================================
            # MOSTRAR RESULTADOS
            # =================================================

            if not productos_encontrados:

                st.warning(
                    "No se encontraron productos "
                    "que coincidan con la búsqueda."
                )

            else:

                st.write(
                    "Seleccione el producto que desea consultar:"
                )

                opciones_resultados = [
                    "Seleccione un producto"
                ]

                opciones_resultados.extend(
                    sorted(
                        productos_encontrados.values(),
                        key=lambda x: x.lower()
                    )
                )

                producto_resultado = st.selectbox(
                    "Productos encontrados:",
                    opciones_resultados,
                    key="resultado_complementario_manual"
                )

                # =============================================
                # MOSTRAR FICHA
                # =============================================

                if (
                    producto_resultado
                    != "Seleccione un producto"
                ):

                    producto_normalizado = (
                        producto_resultado
                        .lower()
                        .strip()
                    )

                    ficha = Complementarios[
                        Complementarios.iloc[:, 0]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        == producto_normalizado
                    ]

                    if ficha.empty:

                        st.warning(
                            "No se encontró información "
                            "para este producto."
                        )

                    else:

                        datos = ficha.iloc[0]

                        st.divider()

                        st.subheader(
                            "Ficha del producto complementario"
                        )

                        st.write(
                            f"**Producto:** "
                            f"{datos.iloc[0]}"
                        )

                        st.write(
                            f"**Categoría principal:** "
                            f"{datos.iloc[1]}"
                        )

                        st.write(
                            f"**Indicaciones / Escenarios:** "
                            f"{datos.iloc[2]}"
                        )

                        st.write(
                            f"**Modo de acción resumido:** "
                            f"{datos.iloc[3]}"
                        )

                        st.write(
                            f"**Combinaciones estratégicas:** "
                            f"{datos.iloc[4]}"
                        )

                        # =====================================
                        # NAVEGACIÓN
                        # =====================================

                        st.divider()

                        siguiente_accion_2 = st.selectbox(
                            "¿Qué desea hacer ahora?",
                            [
                                "Seleccione una opción",
                                "Seleccionar otro producto",
                                "Realizar otra búsqueda",
                                "Volver al menú de Complementarios",
                                "Ir al menú principal"
                            ],
                            key="navegacion_complementarios_manual"
                        )

                        if (
                            siguiente_accion_2
                            == "Seleccionar otro producto"
                        ):

                            st.info(
                                "Puede seleccionar otro "
                                "producto de los resultados."
                            )

                        elif (
                            siguiente_accion_2
                            == "Realizar otra búsqueda"
                        ):

                            st.info(
                                "Ingrese un nuevo nombre "
                                "de producto."
                            )

                        elif (
                            siguiente_accion_2
                            == "Volver al menú de Complementarios"
                        ):

                            st.info(
                                "Seleccione nuevamente "
                                "el tipo de consulta."
                            )

                        elif (
                            siguiente_accion_2
                            == "Ir al menú principal"
                        ):

                            st.session_state[
                                "volver_menu_principal"
                            ] = True

                            st.rerun()
# ============================================================
# ============================================================
# ============================================================
# 6. SECCIÓN ASESORÍA
# ============================================================

elif opcion_principal == "ASESORÍA":

    st.header("ASESORÍA")

    st.subheader("ENTREVISTA")

    # ========================================================
    # 6.1.1 — ENTRADA Y SELECCIÓN DE PATOLOGÍA
    # ========================================================

    metodo_busqueda_asesoria = st.radio(
        "¿Cómo desea buscar la patología?",
        [
            "Por código",
            "Por nombre"
        ],
        key="metodo_busqueda_patologia_asesoria"
    )

    # ========================================================
    # BÚSQUEDA POR CÓDIGO
    # ========================================================

    if metodo_busqueda_asesoria == "Por código":

        codigo_buscado_asesoria = st.text_input(
            "Ingrese el código de la patología:",
            key="codigo_patologia_asesoria"
        )

        if codigo_buscado_asesoria.strip():

            codigo_buscado_asesoria = (
                codigo_buscado_asesoria
                .strip()
                .upper()
            )

            resultado_codigo_asesoria = Patologias[
                Patologias.iloc[:, 0]
                .astype(str)
                .str.strip()
                .str.upper()
                == codigo_buscado_asesoria
            ]

            if resultado_codigo_asesoria.empty:

                st.warning(
                    "No se encontró una patología "
                    "con ese código."
                )

            else:

                fila_patologia_asesoria = (
                    resultado_codigo_asesoria.iloc[0]
                )

                patologia_id_asesoria = str(
                    fila_patologia_asesoria.iloc[0]
                ).strip()

                patologia_nombre_asesoria = str(
                    fila_patologia_asesoria.iloc[1]
                ).strip()

                st.session_state[
                    "patologia_id_asesoria"
                ] = patologia_id_asesoria

                st.session_state[
                    "patologia_nombre_asesoria"
                ] = patologia_nombre_asesoria

                st.success(
                    f"Patología seleccionada: "
                    f"{patologia_nombre_asesoria}"
                )

                st.write(
                    f"**Código:** "
                    f"{patologia_id_asesoria}"
                )

    # ========================================================
    # BÚSQUEDA POR NOMBRE
    # ========================================================

    elif metodo_busqueda_asesoria == "Por nombre":

        nombre_buscado_asesoria = st.text_input(
            "Ingrese el nombre de la patología:",
            key="nombre_patologia_asesoria"
        )

        if nombre_buscado_asesoria.strip():

            texto_busqueda_asesoria = (
                unidecode(
                    nombre_buscado_asesoria
                )
                .lower()
                .strip()
            )

            resultados_patologia_asesoria = []

            # =================================================
            # BUSCAR DIRECTAMENTE EN DATAFRAME PATOLOGIAS
            # =================================================

            for _, fila_patologia_asesoria in (
                Patologias.iterrows()
            ):

                codigo_patologia_asesoria = str(
                    fila_patologia_asesoria.iloc[0]
                ).strip()

                nombre_patologia_asesoria = str(
                    fila_patologia_asesoria.iloc[1]
                ).strip()

                if not nombre_patologia_asesoria:
                    continue

                nombre_normalizado_asesoria = (
                    unidecode(
                        nombre_patologia_asesoria
                    )
                    .lower()
                    .strip()
                )

                # ---------------------------------------------
                # COINCIDENCIA EXACTA
                # ---------------------------------------------

                if (
                    texto_busqueda_asesoria
                    == nombre_normalizado_asesoria
                ):

                    puntaje_asesoria = 100

                # ---------------------------------------------
                # COINCIDENCIA PARCIAL
                # ---------------------------------------------

                elif (
                    texto_busqueda_asesoria
                    in nombre_normalizado_asesoria
                ):

                    puntaje_asesoria = 95

                # ---------------------------------------------
                # TOLERANCIA A ERRORES DE DIGITACIÓN
                # ---------------------------------------------

                else:

                    puntaje_asesoria = fuzz.WRatio(
                        texto_busqueda_asesoria,
                        nombre_normalizado_asesoria
                    )

                if puntaje_asesoria >= 60:

                    resultados_patologia_asesoria.append(
                        {
                            "Codigo":
                                codigo_patologia_asesoria,
                            "Patologia":
                                nombre_patologia_asesoria,
                            "Puntaje":
                                round(
                                    puntaje_asesoria,
                                    2
                                )
                        }
                    )

            # =================================================
            # ELIMINAR DUPLICADOS
            # =================================================

            resultados_unicos_asesoria = {}

            for resultado in (
                resultados_patologia_asesoria
            ):

                codigo = resultado["Codigo"]

                if codigo not in (
                    resultados_unicos_asesoria
                ):

                    resultados_unicos_asesoria[
                        codigo
                    ] = resultado

            resultados_patologia_asesoria = list(
                resultados_unicos_asesoria.values()
            )

            # =================================================
            # ORDENAR POR COINCIDENCIA
            # =================================================

            resultados_patologia_asesoria = sorted(
                resultados_patologia_asesoria,
                key=lambda x: x["Puntaje"],
                reverse=True
            )

            # =================================================
            # SIN RESULTADOS
            # =================================================

            if not resultados_patologia_asesoria:

                st.warning(
                    "No se encontraron patologías "
                    "que coincidan con la búsqueda."
                )

            else:

                opciones_patologia_asesoria = [
                    "Seleccione una patología"
                ]

                for resultado in (
                    resultados_patologia_asesoria
                ):

                    opciones_patologia_asesoria.append(
                        f"{resultado['Codigo']} — "
                        f"{resultado['Patologia']}"
                    )

                seleccion_patologia_asesoria = (
                    st.selectbox(
                        "Seleccione la patología correspondiente:",
                        opciones_patologia_asesoria,
                        key="seleccion_patologia_asesoria"
                    )
                )

                # =================================================
                # GUARDAR PATOLOGÍA SELECCIONADA
                # =================================================

                if (
                    seleccion_patologia_asesoria
                    != "Seleccione una patología"
                ):

                    codigo_seleccionado_asesoria = (
                        seleccion_patologia_asesoria
                        .split(" — ", 1)[0]
                        .strip()
                    )

                    nombre_seleccionado_asesoria = (
                        seleccion_patologia_asesoria
                        .split(" — ", 1)[1]
                        .strip()
                    )

                    st.session_state[
                        "patologia_id_asesoria"
                    ] = codigo_seleccionado_asesoria

                    st.session_state[
                        "patologia_nombre_asesoria"
                    ] = nombre_seleccionado_asesoria

                    st.success(
                        f"Patología seleccionada: "
                        f"{nombre_seleccionado_asesoria}"
                    )

                    st.write(
                        f"**Código:** "
                        f"{codigo_seleccionado_asesoria}"
                    )

       # ========================================================
    # 6.1.2 — CARGA DE PREGUNTAS DE LA ENTREVISTA
    # ========================================================

    if (
        "patologia_id_asesoria"
        in st.session_state
    ):

        patologia_id_actual = (
            st.session_state[
                "patologia_id_asesoria"
            ]
        )

        entrevista_actual = Entrevista[
            Entrevista["Patologia_ID"]
            .astype(str)
            .str.strip()
            ==
            str(
                patologia_id_actual
            ).strip()
        ].copy()

        if entrevista_actual.empty:

            st.warning(
                "No existen preguntas de entrevista "
                "para la patología seleccionada."
            )

            st.session_state[
                "entrevista_actual"
            ] = pd.DataFrame()

        else:

            entrevista_actual = (
                entrevista_actual
                .sort_values(
                    by="Orden"
                )
                .reset_index(
                    drop=True
                )
            )

            st.session_state[
                "entrevista_actual"
            ] = entrevista_actual

            # =================================================
            # VALIDACIÓN DE CARGA
            # =================================================

            st.success(
                "Preguntas cargadas correctamente."
            )

            st.write(
                f"**Patología:** "
                f"{patologia_id_actual}"
            )

            st.write(
                f"**Número de preguntas cargadas:** "
                f"{len(entrevista_actual)}"
            )
    # ========================================================
    # 6.1.3 — INICIO Y REGISTRO DE LA ENTREVISTA
    # ========================================================

    if (
        "entrevista_actual"
        in st.session_state
        and not st.session_state[
            "entrevista_actual"
        ].empty
    ):

        entrevista_actual = st.session_state[
            "entrevista_actual"
        ]

        st.divider()

        st.subheader(
            "INICIO DE LA ENTREVISTA"
        )

        st.write(
            f"**Patología:** "
            f"{st.session_state['patologia_nombre_asesoria']}"
        )

        st.write(
            f"**Preguntas:** "
            f"{len(entrevista_actual)}"
        )

        st.info(
            "Las preguntas no son obligatorias. "
            "Puede dejar una pregunta sin responder."
        )

        respuestas_entrevista = {}

        # ====================================================
        # MOSTRAR PREGUNTAS
        # ====================================================

        for indice, (_, fila) in enumerate(
            entrevista_actual.iterrows(),
            start=1
        ):

            flujo_id = str(
                fila["Flujo_ID"]
            ).strip()

            condicion_id = str(
                fila["Condicion_ID"]
            ).strip()

            pregunta = str(
                fila["Pregunta"]
            ).strip()

            tipo_control = str(
                fila["Tipo_Control"]
            ).strip()

            opciones_texto = str(
                fila["Opciones"]
            ).strip()

            observaciones = str(
                fila["Observaciones"]
            ).strip()

            st.markdown(
                f"### Pregunta {indice} de "
                f"{len(entrevista_actual)}"
            )

            st.write(
                pregunta
            )

            # =================================================
            # OBSERVACIONES
            # =================================================

            if (
                observaciones
                and observaciones.lower()
                != "nan"
            ):

                st.caption(
                    observaciones
                )

            # =================================================
            # PREPARAR OPCIONES
            # =================================================

            opciones = []

            if (
                opciones_texto
                and opciones_texto.lower()
                != "nan"
            ):

                opciones = [
                    opcion.strip()
                    for opcion
                    in opciones_texto.split(";")
                    if opcion.strip()
                ]

            # =================================================
            # LISTA — UNA SOLA RESPUESTA
            # =================================================

            if tipo_control == "Lista":

                respuesta = st.radio(
                    "Seleccione una opción:",
                    opciones,
                    index=None,
                    key=f"respuesta_entrevista_{flujo_id}"
                )

            # =================================================
            # SÍ / NO
            # =================================================

            elif tipo_control == "Sí/No":

                respuesta = st.radio(
                    "Seleccione una opción:",
                    [
                        "Sí",
                        "No"
                    ],
                    index=None,
                    key=f"respuesta_entrevista_{flujo_id}"
                )

            # =================================================
            # NÚMERO
            # =================================================

            elif tipo_control == "Número":

                respuesta = st.number_input(
                    "Ingrese la respuesta:",
                    value=None,
                    placeholder="Opcional",
                    key=f"respuesta_entrevista_{flujo_id}"
                )

            # =================================================
            # SELECCIÓN MÚLTIPLE
            # =================================================

            elif tipo_control == "Selección múltiple":

                if opciones:

                    respuesta = st.multiselect(
                        "Seleccione las opciones que correspondan:",
                        opciones,
                        key=f"respuesta_entrevista_{flujo_id}"
                    )

                else:

                    respuesta = []

            # =================================================
            # TEXTO
            # =================================================

            elif tipo_control == "Texto":

                respuesta = st.text_input(
                    "Respuesta:",
                    placeholder="Opcional",
                    key=f"respuesta_entrevista_{flujo_id}"
                )

            # =================================================
            # CONTROL NO DEFINIDO
            # =================================================

            else:

                respuesta = st.text_input(
                    "Respuesta:",
                    placeholder="Opcional",
                    key=f"respuesta_entrevista_{flujo_id}"
                )

            # =================================================
            # REGISTRAR RESPUESTA TEMPORAL
            # =================================================

            respuestas_entrevista[
                flujo_id
            ] = {
                "Flujo_ID": flujo_id,
                "Condicion_ID": condicion_id,
                "Pregunta": pregunta,
                "Tipo_Control": tipo_control,
                "Respuesta": respuesta
            }

            st.divider()

        # ====================================================
        # FINALIZAR ENTREVISTA
        # ====================================================

        if st.button(
            "Finalizar entrevista",
            key="finalizar_entrevista"
        ):

            st.session_state[
                "respuestas_entrevista"
            ] = respuestas_entrevista

            st.session_state[
                "entrevista_finalizada"
            ] = True

            st.success(
                "Entrevista finalizada correctamente."
            )
    # ========================================================
    # 6.1.4 — RESUMEN DE LA ENTREVISTA
    # ========================================================

    if (
        st.session_state.get(
            "entrevista_finalizada",
            False
        )
        and
        "respuestas_entrevista"
        in st.session_state
    ):

        respuestas_entrevista = (
            st.session_state[
                "respuestas_entrevista"
            ]
        )

        st.divider()

        st.subheader(
            "RESUMEN DE LA ENTREVISTA"
        )

        total_preguntas = len(
            respuestas_entrevista
        )

        preguntas_respondidas = 0
        preguntas_sin_respuesta = 0

        # ====================================================
        # MOSTRAR RESUMEN
        # ====================================================

        for flujo_id, datos_respuesta in (
            respuestas_entrevista.items()
        ):

            pregunta = datos_respuesta[
                "Pregunta"
            ]

            respuesta = datos_respuesta[
                "Respuesta"
            ]

            st.write(
                f"**Pregunta:** {pregunta}"
            )

            # ------------------------------------------------
            # RESPUESTA VACÍA
            # ------------------------------------------------

            if (
                respuesta is None
                or respuesta == ""
                or respuesta == []
            ):

                st.caption(
                    "Sin respuesta"
                )

                preguntas_sin_respuesta += 1

            # ------------------------------------------------
            # RESPUESTA REGISTRADA
            # ------------------------------------------------

            else:

                preguntas_respondidas += 1

                if isinstance(
                    respuesta,
                    list
                ):

                    respuesta_mostrada = (
                        ", ".join(
                            str(item)
                            for item
                            in respuesta
                        )
                    )

                else:

                    respuesta_mostrada = str(
                        respuesta
                    )

                st.write(
                    f"**Respuesta:** "
                    f"{respuesta_mostrada}"
                )

            st.divider()

        # ====================================================
        # CONTADORES
        # ====================================================

        st.write(
            f"**Preguntas respondidas:** "
            f"{preguntas_respondidas} de "
            f"{total_preguntas}"
        )

        st.write(
            f"**Preguntas sin respuesta:** "
            f"{preguntas_sin_respuesta} de "
            f"{total_preguntas}"
        )

        # ====================================================
        # CONFIRMAR ENTREVISTA
        # ====================================================

        if st.button(
            "Confirmar y continuar",
            key="confirmar_resumen_entrevista"
        ):

            st.session_state[
                "resumen_entrevista_confirmado"
            ] = True

            st.success(
                "Resumen confirmado. "
                "La entrevista está lista "
                "para iniciar la evaluación de reglas."
            )


# ============================================================
# 6.2 — EVALUACIÓN DE REGLAS
# ============================================================

    if (
        st.session_state.get(
            "resumen_entrevista_confirmado",
            False
        )
        and
        "respuestas_entrevista"
        in st.session_state
        and
        "patologia_id_asesoria"
        in st.session_state
    ):

        st.divider()

        st.subheader(
            "EVALUACIÓN DE REGLAS"
        )

        respuestas_entrevista = (
            st.session_state[
                "respuestas_entrevista"
            ]
        )

        patologia_id_actual = (
            st.session_state[
                "patologia_id_asesoria"
            ]
        )

        # ====================================================
        # CONSTRUIR MAPA CONDICIÓN → RESPUESTA
        # ====================================================

        respuestas_por_condicion = {}

        for flujo_id, datos_respuesta in (
            respuestas_entrevista.items()
        ):

            condicion_id = str(
                datos_respuesta.get(
                    "Condicion_ID",
                    ""
                )
            ).strip()

            respuesta = datos_respuesta.get(
                "Respuesta",
                None
            )

            if not condicion_id:
                continue

            respuestas_por_condicion[
                condicion_id
            ] = respuesta

        # ====================================================
        # REGLAS DE LA PATOLOGÍA
        # ====================================================

        reglas_actuales = Reglas_Paquetes[
            Reglas_Paquetes["Patologia_ID"]
            .astype(str)
            .str.strip()
            .str.upper()
            ==
            str(
                patologia_id_actual
            ).strip().upper()
        ].copy()

        # ====================================================
        # INFORMACIÓN DE VALIDACIÓN
        # ====================================================

        st.write(
            f"**Patología:** "
            f"{patologia_id_actual} — "
            f"{st.session_state.get(
                'patologia_nombre_asesoria',
                ''
            )}"
        )

        st.write(
            f"**Respuestas registradas:** "
            f"{sum(
                1
                for valor
                in respuestas_por_condicion.values()
                if valor is not None
                and valor != ""
                and valor != []
            )}"
        )

        st.write(
            f"**Reglas encontradas:** "
            f"{len(reglas_actuales)}"
        )

        # ====================================================
        # FUNCIONES AUXILIARES
        # ====================================================

        def normalizar_valor_regla(valor):

            return (
                unidecode(
                    str(valor)
                )
                .lower()
                .strip()
            )


        def obtener_respuesta_condicion(
            condicion_id
        ):

            if condicion_id not in (
                respuestas_por_condicion
            ):
                return None

            return respuestas_por_condicion[
                condicion_id
            ]


        def evaluar_condicion_simple(
            expresion
        ):

            expresion = (
                expresion
                .strip()
            )

            # -----------------------------------------------
            # QUITAR PARÉNTESIS EXTERNOS
            # -----------------------------------------------

            while (
                expresion.startswith("(")
                and
                expresion.endswith(")")
            ):

                contenido = expresion[1:-1].strip()

                nivel = 0
                parentesis_externos = True

                for posicion, caracter in (
                    enumerate(contenido)
                ):

                    if caracter == "(":
                        nivel += 1

                    elif caracter == ")":
                        nivel -= 1

                        if (
                            nivel == 0
                            and
                            posicion
                            != len(contenido) - 1
                        ):

                            parentesis_externos = False
                            break

                if parentesis_externos:

                    expresion = contenido

                else:

                    break

            # -----------------------------------------------
            # SEPARAR CÓDIGO Y VALOR
            # -----------------------------------------------

            partes = expresion.split(
                "=",
                1
            )

            if len(partes) != 2:

                return False

            condicion_id = (
                partes[0]
                .strip()
                .upper()
            )

            valor_esperado = (
                partes[1]
                .strip()
            )

            respuesta = (
                obtener_respuesta_condicion(
                    condicion_id
                )
            )

            # -----------------------------------------------
            # SIN RESPUESTA
            # -----------------------------------------------

            if (
                respuesta is None
                or respuesta == ""
                or respuesta == []
            ):

                return False

            # -----------------------------------------------
            # NORMALIZAR RESPUESTA
            # -----------------------------------------------

            if isinstance(
                respuesta,
                list
            ):

                respuestas = [
                    normalizar_valor_regla(
                        valor
                    )
                    for valor
                    in respuesta
                ]

            else:

                respuestas = [
                    normalizar_valor_regla(
                        respuesta
                    )
                ]

            esperado = normalizar_valor_regla(
                valor_esperado
            )

            # -----------------------------------------------
            # INCLUYE
            # -----------------------------------------------

            if esperado.startswith(
                "incluye "
            ):

                valor_buscado = (
                    esperado[8:]
                    .strip()
                )

                return any(
                    valor_buscado
                    in respuesta_actual
                    for respuesta_actual
                    in respuestas
                )

            # -----------------------------------------------
            # CONTIENE
            # -----------------------------------------------

            if esperado.startswith(
                "contiene "
            ):

                valor_buscado = (
                    esperado[9:]
                    .strip()
                )

                return any(
                    valor_buscado
                    in respuesta_actual
                    for respuesta_actual
                    in respuestas
                )

            # -----------------------------------------------
            # MAYOR QUE
            # -----------------------------------------------

            if esperado.startswith(">"):

                try:

                    limite = float(
                        esperado[1:]
                        .strip()
                    )

                    valor = float(
                        respuesta
                    )

                    return valor > limite

                except (
                    ValueError,
                    TypeError
                ):

                    return False

            # -----------------------------------------------
            # MENOR QUE
            # -----------------------------------------------

            if esperado.startswith("<"):

                try:

                    limite = float(
                        esperado[1:]
                        .strip()
                    )

                    valor = float(
                        respuesta
                    )

                    return valor < limite

                except (
                    ValueError,
                    TypeError
                ):

                    return False

            # -----------------------------------------------
            # IGUALDAD
            # -----------------------------------------------

            return any(
                respuesta_actual
                == esperado
                for respuesta_actual
                in respuestas
            )


        def evaluar_expresion(
            expresion
        ):

            expresion = (
                expresion
                .strip()
            )

            # =================================================
            # QUITAR PARÉNTESIS EXTERNOS
            # =================================================

            while (
                expresion.startswith("(")
                and
                expresion.endswith(")")
            ):

                nivel = 0
                cubre_todo = True

                for posicion, caracter in (
                    enumerate(expresion)
                ):

                    if caracter == "(":

                        nivel += 1

                    elif caracter == ")":

                        nivel -= 1

                        if (
                            nivel == 0
                            and
                            posicion
                            != len(expresion) - 1
                        ):

                            cubre_todo = False
                            break

                if cubre_todo:

                    expresion = (
                        expresion[1:-1]
                        .strip()
                    )

                else:

                    break

            # =================================================
            # BUSCAR OR AL NIVEL PRINCIPAL
            # =================================================

            partes_or = []

            nivel = 0
            inicio = 0
            posicion = 0

            while posicion < len(
                expresion
            ):

                caracter = (
                    expresion[posicion]
                )

                if caracter == "(":

                    nivel += 1

                elif caracter == ")":

                    nivel -= 1

                elif (
                    nivel == 0
                    and
                    expresion[
                        posicion:
                        posicion + 4
                    ].upper()
                    == " OR "
                ):

                    partes_or.append(
                        expresion[
                            inicio:
                            posicion
                        ].strip()
                    )

                    inicio = (
                        posicion + 4
                    )

                    posicion += 4

                    continue

                posicion += 1

            if partes_or:

                partes_or.append(
                    expresion[
                        inicio:
                    ].strip()
                )

                return any(
                    evaluar_expresion(
                        parte
                    )
                    for parte
                    in partes_or
                )

            # =================================================
            # BUSCAR AND AL NIVEL PRINCIPAL
            # =================================================

            partes_and = []

            nivel = 0
            inicio = 0
            posicion = 0

            while posicion < len(
                expresion
            ):

                caracter = (
                    expresion[posicion]
                )

                if caracter == "(":

                    nivel += 1

                elif caracter == ")":

                    nivel -= 1

                elif (
                    nivel == 0
                    and
                    expresion[
                        posicion:
                        posicion + 5
                    ].upper()
                    == " AND "
                ):

                    partes_and.append(
                        expresion[
                            inicio:
                            posicion
                        ].strip()
                    )

                    inicio = (
                        posicion + 5
                    )

                    posicion += 5

                    continue

                posicion += 1

            if partes_and:

                partes_and.append(
                    expresion[
                        inicio:
                    ].strip()
                )

                return all(
                    evaluar_expresion(
                        parte
                    )
                    for parte
                    in partes_and
                )

            # =================================================
            # CONDICIÓN SIMPLE
            # =================================================

            return evaluar_condicion_simple(
                expresion
            )

        # ====================================================
        # NO HAY REGLAS
        # ====================================================

        if reglas_actuales.empty:

            st.warning(
                "No existen reglas configuradas "
                "para la patología seleccionada."
            )

            st.session_state[
                "reglas_activadas"
            ] = pd.DataFrame()

        # ====================================================
        # EVALUAR REGLAS
        # ====================================================

        else:

            reglas_activadas = []

            for _, regla in (
                reglas_actuales.iterrows()
            ):

                condiciones_regla = str(
                    regla[
                        "Condiciones (lógica)"
                    ]
                ).strip()

                if (
                    not condiciones_regla
                    or
                    condiciones_regla.lower()
                    == "nan"
                ):

                    continue

                regla_cumplida = (
                    evaluar_expresion(
                        condiciones_regla
                    )
                )

                if regla_cumplida:

                    reglas_activadas.append(
                        regla
                    )

            # =================================================
            # ORDENAR POR PRIORIDAD
            # =================================================

            if reglas_activadas:

                reglas_activadas_df = (
                    pd.DataFrame(
                        reglas_activadas
                    )
                    .sort_values(
                        by="Prioridad (1=alta)",
                        ascending=True
                    )
                    .reset_index(
                        drop=True
                    )
                )

            else:

                reglas_activadas_df = (
                    pd.DataFrame(
                        columns=
                        reglas_actuales.columns
                    )
                )

            # =================================================
            # GUARDAR RESULTADO
            # =================================================

            st.session_state[
                "reglas_activadas"
            ] = reglas_activadas_df

            # =================================================
            # RESULTADO
            # =================================================

            st.success(
                f"Se evaluaron "
                f"{len(reglas_actuales)} "
                f"reglas."
            )

            if not reglas_activadas_df.empty:

                st.success(
                    f"Se activaron "
                    f"{len(reglas_activadas_df)} "
                    f"reglas."
                )

                st.write(
                    "**Reglas activadas:**"
                )

                for _, regla in (
                    reglas_activadas_df.iterrows()
                ):

                    st.write(
                        f"- **{regla['Regla_ID']}** — "
                        f"{regla['Segmento/Perfil']}"
                    )

            else:

                st.info(
                    "No se activaron reglas "
                    "con las respuestas registradas."
                )

    
# ============================================================
# 6.4 — DEPURACIÓN DE REGLAS, RESTRICCIONES Y PRODUCTOS
# ============================================================

    if (
        "reglas_activadas"
        in st.session_state
    ):

        reglas_activadas_df = (
            st.session_state[
                "reglas_activadas"
            ]
        )

        st.divider()

        st.subheader(
            "DEPURACIÓN DE REGLAS ACTIVADAS"
        )

        # ====================================================
        # VERIFICAR SI EXISTEN REGLAS
        # ====================================================

        if reglas_activadas_df.empty:

            st.info(
                "No existen reglas activadas "
                "para generar productos."
            )

            st.session_state[
                "productos_principales_temporal"
            ] = []

            st.session_state[
                "productos_coadyuvantes_temporal"
            ] = []

            st.session_state[
                "restricciones_productos_temporal"
            ] = []

        else:

            # =================================================
            # LISTAS TEMPORALES
            # =================================================

            productos_principales = []

            productos_coadyuvantes = []

            restricciones_productos = []

            # =================================================
            # RECORRER REGLAS ACTIVADAS
            # =================================================

            for _, regla in (
                reglas_activadas_df.iterrows()
            ):

                regla_id = str(
                    regla.get(
                        "Regla_ID",
                        ""
                    )
                ).strip()

                st.markdown(
                    f"### {regla_id}"
                )

                # =============================================
                # COLUMNA 6 — PRODUCTO PRINCIPAL
                # =============================================

                producto_principal = str(
                    regla.get(
                        "Producto principal",
                        ""
                    )
                ).strip()

                if (
                    producto_principal
                    and
                    producto_principal.lower()
                    != "nan"
                ):

                    st.write(
                        f"**Producto principal:** "
                        f"{producto_principal}"
                    )

                    productos_principales.append(
                        producto_principal
                    )

                # =============================================
                # COLUMNA 7 — COADYUVANTES
                # =============================================

                coadyuvantes_texto = str(
                    regla.get(
                        "Coadyuvantes sugeridos (1-3)",
                        ""
                    )
                ).strip()

                if (
                    coadyuvantes_texto
                    and
                    coadyuvantes_texto.lower()
                    != "nan"
                ):

                    st.write(
                        f"**Coadyuvantes sugeridos:** "
                        f"{coadyuvantes_texto}"
                    )

                    coadyuvantes = [
                        producto.strip()
                        for producto
                        in coadyuvantes_texto.split(";")
                        if producto.strip()
                    ]

                    productos_coadyuvantes.extend(
                        coadyuvantes
                    )

                else:

                    coadyuvantes = []

                # =============================================
                # COLUMNA 8 — MOTIVO TÉCNICO
                # =============================================

                motivo_tecnico = str(
                    regla.get(
                        "Motivo técnico (mecanismo)",
                        ""
                    )
                ).strip()

                if (
                    motivo_tecnico
                    and
                    motivo_tecnico.lower()
                    != "nan"
                ):

                    st.write(
                        f"**Motivo técnico:** "
                        f"{motivo_tecnico}"
                    )

                # =============================================
                # COLUMNA 9 — MENSAJE COMERCIAL
                # =============================================

                mensaje_comercial = str(
                    regla.get(
                        "Mensaje comercial (1 frase)",
                        ""
                    )
                ).strip()

                if (
                    mensaje_comercial
                    and
                    mensaje_comercial.lower()
                    != "nan"
                ):

                    st.write(
                        f"**Mensaje comercial:** "
                        f"{mensaje_comercial}"
                    )

                # =============================================
                # COLUMNA 10 — RESTRICCIÓN
                # =============================================

                restriccion_texto = str(
                    regla.get(
                        "No sugerir si (restricción)",
                        ""
                    )
                ).strip()

                if (
                    not restriccion_texto
                    or
                    restriccion_texto.lower()
                    == "nan"
                ):

                    st.write(
                        "**Restricción:** "
                        "No registrada."
                    )

                else:

                    st.write(
                        f"**Restricción configurada:** "
                        f"{restriccion_texto}"
                    )

                    # =========================================
                    # BUSCAR CÓDIGOS DE RESTRICCIÓN
                    # =========================================

                    codigos_restriccion = [
                        codigo.strip().upper()
                        for codigo
                        in restriccion_texto
                        .replace(",", ";")
                        .split(";")
                        if codigo.strip()
                    ]

                    for codigo_restriccion in (
                        codigos_restriccion
                    ):

                        # -------------------------------------
                        # BUSCAR EN HOJA RESTRICCIONES
                        # -------------------------------------

                        if (
                            "Restriccion_ID"
                            not in Restricciones.columns
                        ):

                            continue

                        coincidencia_restriccion = (
                            Restricciones[
                                Restricciones[
                                    "Restriccion_ID"
                                ]
                                .astype(str)
                                .str.strip()
                                .str.upper()
                                ==
                                codigo_restriccion
                            ]
                        )

                        if (
                            coincidencia_restriccion.empty
                        ):

                            st.warning(
                                f"No se encontró la "
                                f"restricción "
                                f"{codigo_restriccion} "
                                f"en la hoja Restricciones."
                            )

                            continue

                        datos_restriccion = (
                            coincidencia_restriccion
                            .iloc[0]
                        )

                        # =====================================
                        # ÚLTIMAS 3 COLUMNAS
                        # =====================================

                        precaucion = str(
                            datos_restriccion.iloc[-3]
                        ).strip()

                        motivo_restriccion = str(
                            datos_restriccion.iloc[-2]
                        ).strip()

                        alternativas_seguras = str(
                            datos_restriccion.iloc[-1]
                        ).strip()

                        st.warning(
                            f"**{codigo_restriccion}**"
                        )

                        st.write(
                            f"**Precaución / "
                            f"Contraindicación:** "
                            f"{precaucion}"
                        )

                        st.write(
                            f"**Motivo:** "
                            f"{motivo_restriccion}"
                        )

                        st.write(
                            f"**Alternativas seguras:** "
                            f"{alternativas_seguras}"
                        )

                        restricciones_productos.append(
                            {
                                "Regla_ID":
                                    regla_id,
                                "Restriccion_ID":
                                    codigo_restriccion,
                                "Precaucion":
                                    precaucion,
                                "Motivo":
                                    motivo_restriccion,
                                "Alternativas":
                                    alternativas_seguras
                            }
                        )

                st.divider()

            # =================================================
            # GUARDAR INFORMACIÓN TEMPORAL
            # =================================================

            st.session_state[
                "productos_principales_temporal"
            ] = productos_principales

            st.session_state[
                "productos_coadyuvantes_temporal"
            ] = productos_coadyuvantes

            st.session_state[
                "restricciones_productos_temporal"
            ] = restricciones_productos

            # =================================================
            # RESUMEN DE PRODUCTOS ANTES DE DEPURAR
            # =================================================

            st.subheader(
                "Productos identificados antes "
                "de eliminar duplicados"
            )

            st.write(
                f"**Productos principales encontrados:** "
                f"{len(productos_principales)}"
            )

            st.write(
                f"**Coadyuvantes encontrados:** "
                f"{len(productos_coadyuvantes)}"
            )

            # =================================================
            # NORMALIZAR PRODUCTOS PARA COMPARACIÓN
            # =================================================

            def normalizar_producto(
                producto
            ):

                return (
                    unidecode(
                        str(producto)
                    )
                    .lower()
                    .strip()
                )

            # =================================================
            # CONSOLIDAR PRODUCTOS ÚNICOS
            # =================================================

            productos_unicos = {}

            # -----------------------------------------------
            # PRINCIPALES
            # -----------------------------------------------

            for producto in (
                productos_principales
            ):

                clave = normalizar_producto(
                    producto
                )

                if clave not in productos_unicos:

                    productos_unicos[
                        clave
                    ] = {
                        "Producto":
                            producto,
                        "Tipo":
                            "Principal"
                    }

            # -----------------------------------------------
            # COADYUVANTES
            # -----------------------------------------------

            for producto in (
                productos_coadyuvantes
            ):

                clave = normalizar_producto(
                    producto
                )

                if clave not in productos_unicos:

                    productos_unicos[
                        clave
                    ] = {
                        "Producto":
                            producto,
                        "Tipo":
                            "Coadyuvante"
                    }

            productos_unicos_df = pd.DataFrame(
                list(
                    productos_unicos.values()
                )
            )

            st.session_state[
                "productos_unicos_df"
            ] = productos_unicos_df

            # =================================================
            # RESULTADO DE DEPURACIÓN
            # =================================================

            st.subheader(
                "Productos únicos recomendados"
            )

            if productos_unicos_df.empty:

                st.info(
                    "No se identificaron productos "
                    "a partir de las reglas activadas."
                )

            else:

                st.success(
                    f"Se identificaron "
                    f"{len(productos_unicos_df)} "
                    f"productos únicos."
                )

                st.dataframe(
                    productos_unicos_df,
                    use_container_width=True,
                    hide_index=True
                )

# ============================================================
# 6.3.1 — VISUALIZACIÓN DE PRODUCTOS Y MODO DE ACCIÓN
# ============================================================

if (
    "reglas_activadas" in st.session_state
    and not st.session_state[
        "reglas_activadas"
    ].empty
):

    reglas_activadas = st.session_state[
        "reglas_activadas"
    ]

    st.divider()

    st.subheader(
        "PRODUCTOS RECOMENDADOS"
    )

    # ========================================================
    # ESTRUCTURA TEMPORAL DE PRODUCTOS
    # ========================================================

    productos_principales_631 = {}
    productos_coadyuvantes_631 = {}

    # ========================================================
    # RECORRER LAS REGLAS ACTIVADAS
    # ========================================================

    for _, regla in reglas_activadas.iterrows():

        regla_id = str(
            regla["Regla_ID"]
        ).strip()

        motivo_tecnico = str(
            regla["Motivo técnico (mecanismo)"]
        ).strip()

        if (
            not motivo_tecnico
            or motivo_tecnico.lower() == "nan"
        ):

            motivo_tecnico = (
                "Información de acción no disponible."
            )

        # ====================================================
        # PRODUCTO PRINCIPAL
        # ====================================================

        producto_principal = str(
            regla["Producto principal"]
        ).strip()

        if (
            producto_principal
            and producto_principal.lower()
            != "nan"
        ):

            if producto_principal not in (
                productos_principales_631
            ):

                productos_principales_631[
                    producto_principal
                ] = {
                    "Producto":
                        producto_principal,
                    "Tipo":
                        "Principal",
                    "Reglas": [],
                    "Motivos": []
                }

            if regla_id not in (
                productos_principales_631[
                    producto_principal
                ]["Reglas"]
            ):

                productos_principales_631[
                    producto_principal
                ]["Reglas"].append(
                    regla_id
                )

            if motivo_tecnico not in (
                productos_principales_631[
                    producto_principal
                ]["Motivos"]
            ):

                productos_principales_631[
                    producto_principal
                ]["Motivos"].append(
                    motivo_tecnico
                )

        # ====================================================
        # COADYUVANTES
        # ====================================================

        texto_coadyuvantes = str(
            regla[
                "Coadyuvantes sugeridos (1-3)"
            ]
        ).strip()

        if (
            texto_coadyuvantes
            and texto_coadyuvantes.lower()
            != "nan"
        ):

            coadyuvantes = (
                texto_coadyuvantes.split(";")
            )

            for producto_coadyuvante in (
                coadyuvantes
            ):

                producto_coadyuvante = (
                    producto_coadyuvante.strip()
                )

                if (
                    not producto_coadyuvante
                    or producto_coadyuvante.lower()
                    == "nan"
                ):
                    continue

                # --------------------------------------------
                # SI YA ES PRINCIPAL, NO SE CREA DUPLICADO
                # --------------------------------------------

                if producto_coadyuvante in (
                    productos_principales_631
                ):

                    continue

                if producto_coadyuvante not in (
                    productos_coadyuvantes_631
                ):

                    productos_coadyuvantes_631[
                        producto_coadyuvante
                    ] = {
                        "Producto":
                            producto_coadyuvante,
                        "Tipo":
                            "Coadyuvante",
                        "Reglas": [],
                        "Motivos": []
                    }

                if regla_id not in (
                    productos_coadyuvantes_631[
                        producto_coadyuvante
                    ]["Reglas"]
                ):

                    productos_coadyuvantes_631[
                        producto_coadyuvante
                    ]["Reglas"].append(
                        regla_id
                    )

                if motivo_tecnico not in (
                    productos_coadyuvantes_631[
                        producto_coadyuvante
                    ]["Motivos"]
                ):

                    productos_coadyuvantes_631[
                        producto_coadyuvante
                    ]["Motivos"].append(
                        motivo_tecnico
                    )

    # ========================================================
    # GUARDAR ESTRUCTURA TEMPORAL
    # ========================================================

    st.session_state[
        "productos_principales_asesoria"
    ] = productos_principales_631

    st.session_state[
        "productos_coadyuvantes_asesoria"
    ] = productos_coadyuvantes_631

    # ========================================================
    # FUNCIÓN PARA BUSCAR PRODUCTO EN BASE_PRODUCTOS
    # ========================================================

    def buscar_producto_base_631(
        nombre_producto
    ):

        resultado = Base_Productos[
            Base_Productos["Producto"]
            .astype(str)
            .str.strip()
            .str.casefold()
            ==
            str(nombre_producto)
            .strip()
            .casefold()
        ]

        if resultado.empty:

            return None

        return resultado.iloc[0]

    # ========================================================
    # FUNCIÓN PARA MOSTRAR TARJETA
    # ========================================================

    def mostrar_tarjeta_producto_631(
        datos_producto,
        informacion_producto
    ):

        nombre_producto = (
            informacion_producto[
                "Producto"
            ]
        )

        tipo_producto = (
            informacion_producto[
                "Tipo"
            ]
        )

        reglas_producto = (
            informacion_producto[
                "Reglas"
            ]
        )

        motivos_producto = (
            informacion_producto[
                "Motivos"
            ]
        )

        # ====================================================
        # CONTENEDOR VISUAL
        # ====================================================

        with st.container(
            border=True
        ):

            st.markdown(
                f"### {nombre_producto}"
            )

            st.caption(
                f"{tipo_producto}"
            )

            # =================================================
            # IMAGEN
            # =================================================

            imagen = None

            if datos_producto is not None:

                if "Foto" in Base_Productos.columns:

                    valor_imagen = (
                        datos_producto["Foto"]
                    )

                    if (
                        pd.notna(valor_imagen)
                        and str(valor_imagen).strip()
                    ):

                        imagen = str(
                            valor_imagen
                        ).strip()

            if imagen:

                try:

                    st.image(
                        imagen,
                        width=180
                    )

                except Exception:

                    st.caption(
                        "Imagen no disponible"
                    )

            else:

                st.caption(
                    "Imagen no disponible"
                )

            # =================================================
            # PRECIO
            # =================================================

            precio_disponible = False
            precio_numerico = None

            if datos_producto is not None:

                if (
                    "Precio público"
                    in Base_Productos.columns
                ):

                    precio = (
                        datos_producto[
                            "Precio público"
                        ]
                    )

                    if (
                        pd.notna(precio)
                        and str(precio).strip()
                        and str(precio).strip()
                        .lower() != "nan"
                    ):

                        try:

                            texto_precio = (
                                str(precio)
                                .replace("$", "")
                                .replace(" ", "")
                                .strip()
                            )

                            if (
                                ","
                                in texto_precio
                                and "."
                                in texto_precio
                            ):

                                texto_precio = (
                                    texto_precio
                                    .replace(
                                        ".",
                                        ""
                                    )
                                    .replace(
                                        ",",
                                        "."
                                    )
                                )

                            elif "," in texto_precio:

                                texto_precio = (
                                    texto_precio
                                    .replace(
                                        ",",
                                        "."
                                    )
                                )

                            elif (
                                "."
                                in texto_precio
                                and len(
                                    texto_precio
                                    .split(".")[-1]
                                ) == 3
                            ):

                                texto_precio = (
                                    texto_precio
                                    .replace(
                                        ".",
                                        ""
                                    )
                                )

                            precio_numerico = float(
                                texto_precio
                            )

                            precio_disponible = True

                        except (
                            ValueError,
                            TypeError
                        ):

                            precio_disponible = False

            if precio_disponible:

                st.markdown(
                    f"**Precio: "
                    f"${precio_numerico:,.0f}**"
                )

            else:

                st.warning(
                    "Precio no disponible"
                )

            # =================================================
            # MODO DE ACCIÓN
            # =================================================

            st.markdown(
                "**Modo de acción / motivo técnico:**"
            )

            if motivos_producto:

                for motivo in motivos_producto:

                    st.write(
                        motivo
                    )

            else:

                st.caption(
                    "Información de acción "
                    "no disponible."
                )

            # =================================================
            # REGLAS QUE ORIGINARON EL PRODUCTO
            # =================================================

            st.caption(
                "Regla(s) activada(s): "
                + ", ".join(
                    reglas_producto
                )
            )

    # ========================================================
    # PRODUCTOS PRINCIPALES
    # ========================================================

    if productos_principales_631:

        st.markdown(
            "### Productos principales"
        )

        columnas = st.columns(3)

        for indice, (
            nombre_producto,
            informacion_producto
        ) in enumerate(
            productos_principales_631.items()
        ):

            with columnas[
                indice % 3
            ]:

                datos_producto = (
                    buscar_producto_base_631(
                        nombre_producto
                    )
                )

                mostrar_tarjeta_producto_631(
                    datos_producto,
                    informacion_producto
                )

    # ========================================================
    # COADYUVANTES
    # ========================================================

    if productos_coadyuvantes_631:

        st.divider()

        st.markdown(
            "### Coadyuvantes sugeridos"
        )

        columnas = st.columns(3)

        for indice, (
            nombre_producto,
            informacion_producto
        ) in enumerate(
            productos_coadyuvantes_631.items()
        ):

            with columnas[
                indice % 3
            ]:

                datos_producto = (
                    buscar_producto_base_631(
                        nombre_producto
                    )
                )

                mostrar_tarjeta_producto_631(
                    datos_producto,
                    informacion_producto
                )

# ============================================================
# 6.3.2 — SELECCIÓN DE PRODUCTOS Y COTIZACIÓN
# ============================================================

if (
    "productos_principales_asesoria"
    in st.session_state
    or
    "productos_coadyuvantes_asesoria"
    in st.session_state
):

    productos_principales = st.session_state.get(
        "productos_principales_asesoria",
        {}
    )

    productos_coadyuvantes = st.session_state.get(
        "productos_coadyuvantes_asesoria",
        {}
    )

    # ========================================================
    # UNIFICAR PRODUCTOS
    # ========================================================

    productos_cotizacion = {}

    for nombre, datos in (
        productos_principales.items()
    ):

        productos_cotizacion[nombre] = datos

    for nombre, datos in (
        productos_coadyuvantes.items()
    ):

        if nombre not in productos_cotizacion:

            productos_cotizacion[nombre] = datos

    if productos_cotizacion:

        st.divider()

        st.subheader(
            "COTIZACIÓN"
        )

        st.write(
            "Seleccione los productos que "
            "el cliente desea llevar."
        )

        st.info(
            "La selección es opcional. "
            "Los productos sin precio disponible "
            "no se incluyen en el total."
        )

        # ====================================================
        # INICIALIZAR SELECCIÓN TEMPORAL
        # ====================================================

        if (
            "productos_seleccionados_cotizacion"
            not in st.session_state
        ):

            st.session_state[
                "productos_seleccionados_cotizacion"
            ] = {}

        # ====================================================
        # MOSTRAR PRODUCTOS
        # ====================================================

        columnas = st.columns(3)

        for indice, (
            nombre_producto,
            informacion_producto
        ) in enumerate(
            productos_cotizacion.items()
        ):

            with columnas[
                indice % 3
            ]:

                datos_producto = (
                    buscar_producto_base_631(
                        nombre_producto
                    )
                )

                with st.container(
                    border=True
                ):

                    st.markdown(
                        f"### {nombre_producto}"
                    )

                    st.caption(
                        informacion_producto[
                            "Tipo"
                        ]
                    )

                    # ========================================
                    # IMAGEN
                    # ========================================

                    imagen = None

                    if datos_producto is not None:

                        if (
                            "Foto"
                            in Base_Productos.columns
                        ):

                            valor_imagen = (
                                datos_producto[
                                    "Foto"
                                ]
                            )

                            if (
                                pd.notna(
                                    valor_imagen
                                )
                                and
                                str(
                                    valor_imagen
                                ).strip()
                            ):

                                imagen = str(
                                    valor_imagen
                                ).strip()

                    if imagen:

                        try:

                            st.image(
                                imagen,
                                width=180
                            )

                        except Exception:

                            st.caption(
                                "Imagen no disponible"
                            )

                    else:

                        st.caption(
                            "Imagen no disponible"
                        )

                    # ========================================
                    # PRECIO
                    # ========================================

                    precio_disponible = False
                    precio_numerico = None

                    if datos_producto is not None:

                        if (
                            "Precio público"
                            in Base_Productos.columns
                        ):

                            precio = (
                                datos_producto[
                                    "Precio público"
                                ]
                            )

                            if (
                                pd.notna(precio)
                                and
                                str(
                                    precio
                                ).strip()
                                and
                                str(
                                    precio
                                ).strip().lower()
                                != "nan"
                            ):

                                try:

                                    texto_precio = (
                                        str(
                                            precio
                                        )
                                        .replace(
                                            "$",
                                            ""
                                        )
                                        .replace(
                                            " ",
                                            ""
                                        )
                                        .strip()
                                    )

                                    if (
                                        ","
                                        in texto_precio
                                        and
                                        "."
                                        in texto_precio
                                    ):

                                        texto_precio = (
                                            texto_precio
                                            .replace(
                                                ".",
                                                ""
                                            )
                                            .replace(
                                                ",",
                                                "."
                                            )
                                        )

                                    elif (
                                        ","
                                        in texto_precio
                                    ):

                                        texto_precio = (
                                            texto_precio
                                            .replace(
                                                ",",
                                                "."
                                            )
                                        )

                                    elif (
                                        "."
                                        in texto_precio
                                        and
                                        len(
                                            texto_precio
                                            .split(
                                                "."
                                            )[-1]
                                        ) == 3
                                    ):

                                        texto_precio = (
                                            texto_precio
                                            .replace(
                                                ".",
                                                ""
                                            )
                                        )

                                    precio_numerico = float(
                                        texto_precio
                                    )

                                    precio_disponible = True

                                except (
                                    ValueError,
                                    TypeError
                                ):

                                    precio_disponible = False

                    if precio_disponible:

                        st.markdown(
                            f"**Precio: "
                            f"${precio_numerico:,.0f}**"
                        )

                    else:

                        st.warning(
                            "Precio no disponible"
                        )

                    # ========================================
                    # SELECCIÓN
                    # ========================================

                    seleccionado = st.checkbox(
                        "Llevar",
                        key=(
                            f"llevar_producto_"
                            f"{nombre_producto}"
                        )
                    )

                    # ========================================
                    # GUARDAR SELECCIÓN TEMPORAL
                    # ========================================

                    if seleccionado:

                        st.session_state[
                            "productos_seleccionados_cotizacion"
                        ][nombre_producto] = {
                            "Producto":
                                nombre_producto,
                            "Tipo":
                                informacion_producto[
                                    "Tipo"
                                ],
                            "Precio":
                                precio_numerico
                                if precio_disponible
                                else None,
                            "Precio_disponible":
                                precio_disponible
                        }

                    else:

                        st.session_state[
                            "productos_seleccionados_cotizacion"
                        ].pop(
                            nombre_producto,
                            None
                        )

        # ====================================================
        # RESUMEN DE COTIZACIÓN
        # ====================================================

        seleccionados = (
            st.session_state.get(
                "productos_seleccionados_cotizacion",
                {}
            )
        )

        if seleccionados:

            st.divider()

            st.subheader(
                "PRODUCTOS SELECCIONADOS"
            )

            total_cotizacion = 0

            productos_con_precio = 0
            productos_sin_precio = 0

            for (
                nombre_producto,
                datos_producto
            ) in seleccionados.items():

                precio = datos_producto[
                    "Precio"
                ]

                if (
                    datos_producto[
                        "Precio_disponible"
                    ]
                    and precio is not None
                ):

                    st.write(
                        f"**{nombre_producto}** "
                        f"— ${precio:,.0f}"
                    )

                    total_cotizacion += precio

                    productos_con_precio += 1

                else:

                    st.write(
                        f"**{nombre_producto}** "
                        f"— Precio no disponible"
                    )

                    productos_sin_precio += 1

            # =================================================
            # TOTAL
            # =================================================

            st.divider()

            st.markdown(
                f"## TOTAL: "
                f"${total_cotizacion:,.0f}"
            )

            st.caption(
                f"{productos_con_precio} producto(s) "
                f"con precio incluido en el total."
            )

            if productos_sin_precio:

                st.warning(
                    f"{productos_sin_precio} producto(s) "
                    f"seleccionado(s) no tienen precio "
                    f"disponible y no fueron incluidos "
                    f"en el total."
                )

            # =================================================
            # GUARDAR COTIZACIÓN TEMPORAL
            # =================================================

            st.session_state[
                "cotizacion_final"
            ] = {
                "Productos":
                    seleccionados,
                "Total":
                    total_cotizacion
            }

        else:

            st.info(
                "Aún no se han seleccionado "
                "productos para llevar."
            )
            # ============================================================
# 6.4 — FINALIZAR ASESORÍA Y REGRESAR AL MENÚ PRINCIPAL
# ============================================================

if (
    "cotizacion_final" in st.session_state
):

    st.divider()

    st.success(
        "La asesoría y cotización han finalizado correctamente."
    )

    if st.button(
        "Finalizar asesoría",
        key="finalizar_asesoria"
    ):

        # ====================================================
        # LIMPIAR INFORMACIÓN TEMPORAL DE LA ASESORÍA
        # ====================================================

        variables_temporales_asesoria = [

            "patologia_id_asesoria",
            "patologia_nombre_asesoria",

            "entrevista_actual",

            "respuestas_entrevista",
            "entrevista_finalizada",
            "resumen_entrevista_confirmado",

            "reglas_activadas",

            "productos_principales_asesoria",
            "productos_coadyuvantes_asesoria",

            "productos_seleccionados_cotizacion",
            "cotizacion_final"
        ]

        for variable in (
            variables_temporales_asesoria
        ):

            st.session_state.pop(
                variable,
                None
            )

        # ====================================================
        # LIMPIAR CONTROLES PROPIOS DE LA ASESORÍA
        # ====================================================

        controles_asesoria = [

            "metodo_busqueda_patologia_asesoria",
            "codigo_patologia_asesoria",
            "nombre_patologia_asesoria",
            "seleccion_patologia_asesoria"
        ]

        for variable in controles_asesoria:

            st.session_state.pop(
                variable,
                None
            )

        # ====================================================
        # REGRESAR AL MENÚ PRINCIPAL
        # ====================================================

        st.session_state[
            "opcion_principal"
        ] = "Seleccione una opción"

        st.rerun()
    

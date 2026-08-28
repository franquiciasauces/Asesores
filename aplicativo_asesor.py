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

    opciones_evaluacion = [
        "Seleccione una opción",
        "Evaluación general",
        "Evaluaciones controladas",
        "Historial de evaluaciones"
    ]

    opcion_evaluacion = st.selectbox(
        "Seleccione una opción:",
        opciones_evaluacion,
        key="menu_evaluacion"
    )
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
    



# ============================================================
# GESTIONEJECUCIONEVALUACION
# PARTE 1 — ACCESO A EVALUACIONES CONTROLADAS
# ============================================================

if (
    opcion_principal == "EVALUACIÓN"
    and opcion_evaluacion == "Evaluaciones controladas"
):

    st.subheader(
        "Evaluaciones controladas"
    )

    # ========================================================
    # 1. ARCHIVO PERSISTENTE
    # ========================================================

    ARCHIVO_EVALUACIONES_CONTROLADAS = (
        BASE_DIR
        / "evaluacion"
        / "Respositorioevaluacionescontroladas.csv"
    )

    # ========================================================
    # 2. VERIFICAR ARCHIVO
    # ========================================================

    if not ARCHIVO_EVALUACIONES_CONTROLADAS.exists():

        st.info(
            "No existen evaluaciones controladas disponibles."
        )

    else:

        # ====================================================
        # 3. CARGAR REPOSITORIO
        # ====================================================

        try:

            df_evaluaciones_controladas = pd.read_csv(
                ARCHIVO_EVALUACIONES_CONTROLADAS,
                dtype=str,
                keep_default_na=False
            )

            df_evaluaciones_controladas = (
                df_evaluaciones_controladas.fillna("")
            )

        except Exception as error:

            st.error(
                "No fue posible cargar el repositorio "
                "de evaluaciones controladas."
            )

            st.code(
                str(error)
            )

            df_evaluaciones_controladas = (
                pd.DataFrame()
            )

        # ====================================================
        # 4. GUARDAR REPOSITORIO EN MEMORIA TEMPORAL
        # ====================================================

        if not df_evaluaciones_controladas.empty:

            st.session_state[
                "df_evaluaciones_controladas_ejecucion"
            ] = df_evaluaciones_controladas.copy()

            # =================================================
            # 5. VERIFICAR COLUMNAS
            # =================================================

            columnas_necesarias = [
                "Evaluacion_ID",
                "Modulo",
                "Nombre_Evaluacion",
                "Descripcion",
                "Estado"
            ]

            columnas_faltantes = [
                columna
                for columna in columnas_necesarias
                if columna not in df_evaluaciones_controladas.columns
            ]

            if columnas_faltantes:

                st.error(
                    "El repositorio de evaluaciones controladas "
                    "no tiene la estructura esperada."
                )

                st.write(
                    "Columnas faltantes:",
                    columnas_faltantes
                )

            else:

                # =================================================
                # 6. NORMALIZAR CAMPOS PRINCIPALES
                # =================================================

                for columna in columnas_necesarias:

                    df_evaluaciones_controladas[
                        columna
                    ] = (
                        df_evaluaciones_controladas[
                            columna
                        ]
                        .astype(str)
                        .str.strip()
                    )

                # =================================================
                # 7. SOLO EVALUACIONES ACTIVAS
                # =================================================

                df_evaluaciones_activas = (
                    df_evaluaciones_controladas[
                        df_evaluaciones_controladas[
                            "Estado"
                        ]
                        .str.upper()
                        == "ACTIVA"
                    ]
                    .copy()
                )

                if df_evaluaciones_activas.empty:

                    st.info(
                        "No existen evaluaciones controladas "
                        "activas disponibles."
                    )

                    st.session_state.pop(
                        "evaluacion_controlada_ejecucion_id",
                        None
                    )

                else:

                    # =================================================
                    # 8. SELECCIONAR MÓDULO
                    # =================================================

                    st.markdown(
                        "### 1. Seleccione el módulo"
                    )

                    modulos = sorted(
                        df_evaluaciones_activas[
                            "Modulo"
                        ]
                        .replace("", pd.NA)
                        .dropna()
                        .unique()
                        .tolist()
                    )

                    modulo_seleccionado = st.selectbox(
                        "Módulo:",
                        [
                            "Seleccione un módulo"
                        ] + modulos,
                        key="modulo_ejecucion_controlada"
                    )

                    # =================================================
                    # 9. FILTRAR POR MÓDULO
                    # =================================================

                    if (
                        modulo_seleccionado
                        != "Seleccione un módulo"
                    ):

                        df_modulo = (
                            df_evaluaciones_activas[
                                df_evaluaciones_activas[
                                    "Modulo"
                                ]
                                == modulo_seleccionado
                            ]
                            .copy()
                        )

                        if df_modulo.empty:

                            st.info(
                                "No existen evaluaciones controladas "
                                "activas para el módulo seleccionado."
                            )

                        else:

                            # =========================================
                            # 10. MOSTRAR RESUMEN
                            # =========================================

                            st.markdown(
                                "### 2. Evaluaciones disponibles"
                            )

                            columnas_mostrar = [
                                "Evaluacion_ID",
                                "Modulo",
                                "Nombre_Evaluacion",
                                "Descripcion"
                            ]

                            st.dataframe(
                                df_modulo[
                                    columnas_mostrar
                                ],
                                use_container_width=True,
                                hide_index=True
                            )

                            # =========================================
                            # 11. SELECCIONAR EVALUACIÓN
                            # =========================================

                            st.markdown(
                                "### 3. Seleccione la evaluación que realizará"
                            )

                            opciones_evaluaciones = (
                                df_modulo[
                                    "Evaluacion_ID"
                                ]
                                .astype(str)
                                .str.strip()
                                .tolist()
                            )

                            evaluacion_seleccionada = (
                                st.selectbox(
                                    "Evaluación:",
                                    [
                                        "Seleccione una evaluación"
                                    ] + opciones_evaluaciones,
                                    key=(
                                        "evaluacion_controlada_selector"
                                    )
                                )
                            )

                            # =========================================
                            # 12. GUARDAR SELECCIÓN TEMPORAL
                            # =========================================

                            if (
                                evaluacion_seleccionada
                                != "Seleccione una evaluación"
                            ):

                                st.session_state[
                                    "evaluacion_controlada_ejecucion_id"
                                ] = (
                                    evaluacion_seleccionada
                                )

                                st.success(
                                    "Evaluación seleccionada: "
                                    + evaluacion_seleccionada
                                )

                            else:

                                st.session_state.pop(
                                    "evaluacion_controlada_ejecucion_id",
                                    None
                                )


# ============================================================
# GESTIONEJECUCIONEVALUACION
# PARTE 2 — CARGA Y RESPUESTA DE LA EVALUACIÓN
# ============================================================
# ============================================================

# GESTIONEJECUCIONEVALUACION

# PARTE 2 — CARGA Y RESPUESTA DE LA EVALUACIÓN

# ============================================================

if (
opcion_principal == "EVALUACIÓN"
and opcion_evaluacion == "Evaluaciones controladas"
and st.session_state.get(
"evaluacion_controlada_ejecucion_id"
)
):


evaluacion_id_seleccionada = str(
    st.session_state[
        "evaluacion_controlada_ejecucion_id"
    ]
).strip()

ARCHIVO_EVALUACIONES_CONTROLADAS = (
    BASE_DIR
    / "evaluacion"
    / "Respositorioevaluacionescontroladas.csv"
)

if not ARCHIVO_EVALUACIONES_CONTROLADAS.exists():

    st.error(
        "No se encuentra el repositorio de evaluaciones controladas."
    )

else:

    try:

        df_repositorio_ejecucion = pd.read_csv(
            ARCHIVO_EVALUACIONES_CONTROLADAS,
            dtype=str,
            keep_default_na=False
        )

        df_repositorio_ejecucion = (
            df_repositorio_ejecucion.fillna("")
        )

    except Exception as error:

        st.error(
            "No fue posible cargar el repositorio "
            "de evaluaciones controladas."
        )

        st.code(str(error))

        df_repositorio_ejecucion = pd.DataFrame()

    if not df_repositorio_ejecucion.empty:

        # ====================================================
        # FILTRAR LA EVALUACIÓN SELECCIONADA
        # ====================================================

        df_evaluacion = (
            df_repositorio_ejecucion[
                df_repositorio_ejecucion[
                    "Evaluacion_ID"
                ]
                .astype(str)
                .str.strip()
                == evaluacion_id_seleccionada
            ]
            .copy()
        )

        if df_evaluacion.empty:

            st.warning(
                "No se encontró la evaluación seleccionada "
                "en el repositorio de evaluaciones controladas."
            )

        else:

            fila_evaluacion = df_evaluacion.iloc[0]

            estado_evaluacion = str(
                fila_evaluacion.get(
                    "Estado",
                    ""
                )
            ).strip().upper()

            if estado_evaluacion != "ACTIVA":

                st.warning(
                    "La evaluación seleccionada "
                    "no se encuentra activa."
                )

            else:

                # =================================================
                # INFORMACIÓN GENERAL DE LA EVALUACIÓN
                # =================================================

                nombre_evaluacion = str(
                    fila_evaluacion.get(
                        "Nombre_Evaluacion",
                        ""
                    )
                ).strip()

                modulo_evaluacion = str(
                    fila_evaluacion.get(
                        "Modulo",
                        ""
                    )
                ).strip()

                descripcion_evaluacion = str(
                    fila_evaluacion.get(
                        "Descripcion",
                        ""
                    )
                ).strip()

                st.subheader(
                    nombre_evaluacion
                )

                st.write(
                    f"Módulo: {modulo_evaluacion}"
                )

                if descripcion_evaluacion:

                    st.write(
                        descripcion_evaluacion
                    )

                st.divider()

                # =================================================
                # OBTENER CAMPOS CONCATENADOS
                # =================================================

                preguntas_raw = str(
                    fila_evaluacion.get(
                        "Pregunta",
                        ""
                    )
                ).strip()

                opcion_1_raw = str(
                    fila_evaluacion.get(
                        "Opcion_1",
                        ""
                    )
                ).strip()

                opcion_2_raw = str(
                    fila_evaluacion.get(
                        "Opcion_2",
                        ""
                    )
                ).strip()

                opcion_3_raw = str(
                    fila_evaluacion.get(
                        "Opcion_3",
                        ""
                    )
                ).strip()

                opcion_4_raw = str(
                    fila_evaluacion.get(
                        "Opcion_4",
                        ""
                    )
                ).strip()

                # =================================================
                # SEPARADOR REAL DEL REPOSITORIO
                # =================================================

                separador = (
                    "EVAL_CONTROLADA_"
                    + evaluacion_id_seleccionada.split(
                        "EVAL_CONTROLADA_"
                    )[-1]
                    + "_P"
                )

                # =================================================
                # RECONSTRUIR PREGUNTAS
                # =================================================

                bloques_preguntas = (
                    preguntas_raw.split(
                        separador
                    )
                )

                bloques_preguntas = [
                    bloque.strip()
                    for bloque in bloques_preguntas
                    if bloque.strip()
                ]

                preguntas_ejecucion = []

                for bloque in bloques_preguntas:

                    partes = bloque.split(
                        " | ",
                        2
                    )

                    if len(partes) == 3:

                        numero_pregunta = (
                            partes[0]
                            .strip()
                        )

                        nivel = (
                            partes[1]
                            .strip()
                        )

                        pregunta_texto = (
                            partes[2]
                            .strip()
                        )

                        pregunta_id = (
                            evaluacion_id_seleccionada
                            + "_P"
                            + numero_pregunta
                        )

                        preguntas_ejecucion.append(
                            {
                                "Pregunta_ID":
                                    pregunta_id,
                                "Nivel":
                                    nivel,
                                "Pregunta":
                                    pregunta_texto
                            }
                        )

                # =================================================
                # RECONSTRUIR OPCIONES
                # =================================================

                bloques_opcion_1 = [
                    bloque.strip()
                    for bloque in opcion_1_raw.split(
                        separador
                    )
                    if bloque.strip()
                ]

                bloques_opcion_2 = [
                    bloque.strip()
                    for bloque in opcion_2_raw.split(
                        separador
                    )
                    if bloque.strip()
                ]

                bloques_opcion_3 = [
                    bloque.strip()
                    for bloque in opcion_3_raw.split(
                        separador
                    )
                    if bloque.strip()
                ]

                bloques_opcion_4 = [
                    bloque.strip()
                    for bloque in opcion_4_raw.split(
                        separador
                    )
                    if bloque.strip()
                ]

                opciones_por_pregunta = []

                for indice in range(
                    len(preguntas_ejecucion)
                ):

                    opciones = []

                    bloques_opciones = [
                        bloques_opcion_1,
                        bloques_opcion_2,
                        bloques_opcion_3,
                        bloques_opcion_4
                    ]

                    for bloques in bloques_opciones:

                        if indice < len(bloques):

                            texto_opcion = (
                                bloques[indice]
                                .strip()
                            )

                            partes_opcion = (
                                texto_opcion.split(
                                    " | ",
                                    1
                                )
                            )

                            if len(partes_opcion) == 2:

                                opcion = (
                                    partes_opcion[1]
                                    .strip()
                                )

                            else:

                                opcion = (
                                    texto_opcion
                                )

                            opciones.append(
                                opcion
                            )

                        else:

                            opciones.append("")

                    opciones_por_pregunta.append(
                        opciones
                    )

                # =================================================
                # VERIFICAR QUE SE HAYAN RECONSTRUIDO PREGUNTAS
                # =================================================

                if not preguntas_ejecucion:

                    st.warning(
                        "La evaluación seleccionada no contiene "
                        "preguntas que puedan ser presentadas."
                    )

                else:

                    # =================================================
                    # MEMORIA TEMPORAL DE RESPUESTAS
                    # =================================================

                    clave_respuestas = (
                        "respuestas_temporales_"
                        + evaluacion_id_seleccionada
                    )

                    if (
                        clave_respuestas
                        not in st.session_state
                    ):

                        st.session_state[
                            clave_respuestas
                        ] = {}

                    respuestas_temporales = (
                        st.session_state[
                            clave_respuestas
                        ]
                    )

                    # =================================================
                    # MOSTRAR PREGUNTAS Y OPCIONES
                    # =================================================

                    for indice, pregunta in enumerate(
                        preguntas_ejecucion
                    ):

                        pregunta_id = str(
                            pregunta[
                                "Pregunta_ID"
                            ]
                        ).strip()

                        nivel = str(
                            pregunta[
                                "Nivel"
                            ]
                        ).strip()

                        enunciado = str(
                            pregunta[
                                "Pregunta"
                            ]
                        ).strip()

                        opciones = (
                            opciones_por_pregunta[
                                indice
                            ]
                        )

                        opciones = [
                            opcion
                            for opcion in opciones
                            if str(opcion).strip() != ""
                        ]

                        st.markdown(
                            f"### Pregunta "
                            f"{indice + 1} de "
                            f"{len(preguntas_ejecucion)}"
                        )

                        if nivel:

                            st.caption(
                                f"Nivel: {nivel}"
                            )

                        st.write(
                            enunciado
                        )

                        respuesta_guardada = (
                            respuestas_temporales.get(
                                pregunta_id
                            )
                        )

                        if (
                            respuesta_guardada
                            in opciones
                        ):

                            indice_inicial = (
                                opciones.index(
                                    respuesta_guardada
                                )
                            )

                        else:

                            indice_inicial = None

                        respuesta_seleccionada = st.radio(
                            "Seleccione la respuesta que considera correcta:",
                            opciones,
                            index=indice_inicial,
                            key=(
                                "respuesta_controlada_"
                                + evaluacion_id_seleccionada
                                + "_"
                                + pregunta_id
                            )
                        )

                        if (
                            respuesta_seleccionada
                            is not None
                        ):

                            respuestas_temporales[
                                pregunta_id
                            ] = (
                                respuesta_seleccionada
                            )

                        st.divider()

                    # =================================================
                    # ACTUALIZAR MEMORIA TEMPORAL
                    # =================================================

                    st.session_state[
                        clave_respuestas
                    ] = respuestas_temporales

                    # =================================================
                    # INFORMACIÓN TEMPORAL
                    # =================================================

                    preguntas_respondidas = len(
                        respuestas_temporales
                    )

                    total_preguntas = len(
                        preguntas_ejecucion
                    )

                    st.info(
                        f"Preguntas respondidas: "
                        f"{preguntas_respondidas} de "
                        f"{total_preguntas}"
                    )


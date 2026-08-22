# ============================================================
# FITOASISTE
# APLICATIVO DE EVALUACIÓN
# NORMALIZACIÓN DE PRODUCTO - COMPONENTE - ACCIÓN GENERAL
# ============================================================

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


# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="FITOASISTE - Evaluación",
    page_icon="📝",
    layout="wide"
)


# ============================================================
# 2. RUTAS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

ARCHIVO_MATRIZ = (
    BASE_DIR /
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)

ARCHIVO_NORMALIZADO = (
    BASE_DIR /
    "DATAFRAME_PRODUCTO_COMPONENTE_ACCION.csv"
)


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

st.title("📝 FITOASISTE — EVALUACIÓN")

st.write(
    f"Administrador: **{USUARIO}**"
)

# ============================================================


# ============================================================
# 5.1 DIAGNÓSTICO Y LECTURA DE LA MATRIZ FUENTE
# ============================================================

st.markdown("### 5.1 Diagnóstico de la matriz fuente")

try:
    if not ARCHIVO_MATRIZ.exists():
        st.error(
            "❌ 5.1 ERROR: No se encontró el archivo de la matriz."
        )
    else:
        libro = pd.ExcelFile(ARCHIVO_MATRIZ)

        st.success(
            f"✅ 5.1 OK: Archivo de matriz encontrado. "
            f"Hojas disponibles: {len(libro.sheet_names)}"
        )

        st.write("**Hojas encontradas en la matriz:**")
        st.write(libro.sheet_names)

        hoja_fuente = st.selectbox(
            "Seleccione la hoja de la matriz que contiene la información:",
            libro.sheet_names,
            key="hoja_matriz_normalizacion"
        )

        df_fuente = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name=hoja_fuente
        )

        df_fuente = df_fuente.dropna(
            axis=1,
            how="all"
        )

        st.success(
            f"✅ Hoja cargada correctamente: **{hoja_fuente}**"
        )

        st.info(
            f"Registros encontrados: **{len(df_fuente)}** | "
            f"Columnas encontradas: **{len(df_fuente.columns)}**"
        )

        st.write("### Columnas REALES encontradas en la hoja")

        columnas_reales = pd.DataFrame({
            "N.º": range(1, len(df_fuente.columns) + 1),
            "Nombre real de la columna": [
                str(col)
                for col in df_fuente.columns
            ]
        })

        st.dataframe(
            columnas_reales,
            use_container_width=True,
            hide_index=True
        )

        st.write("### Primeros registros de la matriz original")

        st.dataframe(
            df_fuente.head(10),
            use_container_width=True,
            hide_index=True
        )

        st.success(
            "🟢 5.1 TERMINADO: La matriz fue leída correctamente. "
            "Todavía NO se ha realizado ninguna normalización."
        )

except Exception as e:
    st.error(
        f"🔴 5.1 ERROR al leer la matriz: "
        f"{type(e).__name__}: {e}"
    )

# ============================================================
# 5.2 SEPARAR ACCIONES GENERALES
# ============================================================

st.markdown("### 5.2 Separación de acciones generales")

try:
    import re

    requeridas_52 = [
        "Producto",
        "Acciones generales"
    ]

    faltantes_52 = [
        columna
        for columna in requeridas_52
        if columna not in df_fuente.columns
    ]

    if faltantes_52:
        st.error(
            "❌ 5.2 ERROR: Faltan columnas: "
            + ", ".join(faltantes_52)
        )
        st.stop()

    df_acciones_52 = df_fuente[
        requeridas_52
    ].copy()

    df_acciones_52 = df_acciones_52.fillna("")

    def separar_acciones_52(texto):
        texto = str(texto).strip()

        if not texto:
            return []

        partes = re.split(
            r"\s*;\s*|\s*,\s*|(?<=\.)\s+(?=[A-ZÁÉÍÓÚÑ])",
            texto
        )

        return [
            parte.strip(" ;,.")
            for parte in partes
            if parte.strip(" ;,.")
        ]

    df_acciones_52[
        "Acción general"
    ] = df_acciones_52[
        "Acciones generales"
    ].apply(separar_acciones_52)

    df_acciones_52 = (
        df_acciones_52
        .explode("Acción general")
        .reset_index(drop=True)
    )

    df_acciones_52["Producto"] = (
        df_acciones_52["Producto"]
        .astype(str)
        .str.strip()
    )

    df_acciones_52["Acción general"] = (
        df_acciones_52["Acción general"]
        .astype(str)
        .str.strip()
    )

    df_acciones_52 = df_acciones_52[
        (df_acciones_52["Producto"] != "")
        &
        (df_acciones_52["Acción general"] != "")
    ].copy()

    df_acciones_52 = df_acciones_52[
        [
            "Producto",
            "Acción general"
        ]
    ]

    st.session_state[
        "df_acciones_52"
    ] = df_acciones_52.copy()

    st.success(
        f"🟢 5.2 TERMINADO: "
        f"{len(df_fuente)} registros originales → "
        f"{len(df_acciones_52)} relaciones Producto–Acción."
    )

    st.dataframe(
        df_acciones_52,
        use_container_width=True,
        hide_index=True
    )

except Exception as e:
    st.error(
        f"🔴 5.2 ERROR: "
        f"{type(e).__name__}: {e}"
    )


# ============================================================
# ============================================================
# 5.3 CLASIFICACIÓN DE CADA ACCIÓN GENERAL
# ============================================================

st.markdown("### 5.3 Clasificación de acciones generales")

try:
    import re
    import unicodedata

    # --------------------------------------------------------
    # 1. RECUPERAR RESULTADO DE 5.2
    # --------------------------------------------------------

    df_acciones_52 = st.session_state.get(
        "df_acciones_52"
    )

    if (
        df_acciones_52 is None
        or not isinstance(df_acciones_52, pd.DataFrame)
        or df_acciones_52.empty
    ):
        st.error(
            "❌ 5.3 ERROR: No existe el resultado de 5.2."
        )
        st.stop()

    # --------------------------------------------------------
    # 2. VALIDAR COLUMNAS DE 5.2 Y MATRIZ
    # --------------------------------------------------------

    requeridas_53 = [
        "Producto",
        "Acción general"
    ]

    faltantes_53 = [
        columna
        for columna in requeridas_53
        if columna not in df_acciones_52.columns
    ]

    if faltantes_53:
        st.error(
            "❌ 5.3 ERROR: Faltan columnas de 5.2: "
            + ", ".join(faltantes_53)
        )
        st.stop()

    if "Componentes" not in df_fuente.columns:
        st.error(
            "❌ 5.3 ERROR: La matriz no contiene "
            "la columna real 'Componentes'."
        )
        st.stop()

    # --------------------------------------------------------
    # 3. FUNCIONES DE NORMALIZACIÓN
    # --------------------------------------------------------

    def normalizar_53(texto):
        texto = str(texto).strip().lower()

        texto = unicodedata.normalize(
            "NFD",
            texto
        )

        texto = "".join(
            caracter
            for caracter in texto
            if unicodedata.category(caracter) != "Mn"
        )

        texto = re.sub(
            r"[^a-z0-9]+",
            " ",
            texto
        )

        return re.sub(
            r"\s+",
            " ",
            texto
        ).strip()

    def contiene_termino_53(texto, termino):
        texto_n = normalizar_53(texto)
        termino_n = normalizar_53(termino)

        if not termino_n:
            return False

        patron = (
            r"(?<![a-z0-9])"
            + re.escape(termino_n)
            + r"(?![a-z0-9])"
        )

        return re.search(
            patron,
            texto_n
        ) is not None

    # --------------------------------------------------------
    # 4. CREAR COMPONENTES PROPIOS DE CADA PRODUCTO
    # --------------------------------------------------------

    componentes_producto_53 = {}

    for _, fila in df_fuente[
        [
            "Producto",
            "Componentes"
        ]
    ].fillna("").iterrows():

        producto = str(
            fila["Producto"]
        ).strip()

        componentes = str(
            fila["Componentes"]
        ).strip()

        if not producto:
            continue

        partes = re.split(
            r"\s*;\s*|\s*,\s*",
            componentes
        )

        componentes_limpios = []

        for componente in partes:
            componente = componente.strip()

            if componente:
                componentes_limpios.append(
                    componente
                )

        componentes_producto_53[
            normalizar_53(producto)
        ] = componentes_limpios

    # --------------------------------------------------------
    # 5. COMPONENTES DE TODA LA MATRIZ
    # --------------------------------------------------------

    componentes_globales_53 = set()

    for lista_componentes in (
        componentes_producto_53.values()
    ):
        for componente in lista_componentes:
            componente_n = normalizar_53(
                componente
            )

            if componente_n:
                componentes_globales_53.add(
                    componente_n
                )

    # --------------------------------------------------------
    # 6. PRODUCTOS DE TODA LA MATRIZ
    # --------------------------------------------------------

    productos_globales_53 = []

    for producto in df_fuente[
        "Producto"
    ].dropna().astype(str):

        producto = producto.strip()

        if producto:
            productos_globales_53.append(
                producto
            )

    # --------------------------------------------------------
    # 7. REGLAS EXPLÍCITAS DE CLASIFICACIÓN
    # --------------------------------------------------------

    patrones_eliminar_53 = [
        "frase comercial",
        "frase de venta",
        "frase venta"
    ]

    patrones_combinacion_53 = [
        "combinaciones",
        "combinacion",
        "se puede combinar",
        "puede combinarse",
        "combinar con"
    ]

    patrones_restriccion_53 = [
        "contraindicacion",
        "contraindicaciones",
        "restriccion",
        "restricciones",
        "no usar",
        "no recomendado",
        "no recomendada",
        "precaucion",
        "precauciones"
    ]

    patrones_uso_53 = [
        "posologia",
        "dosis",
        "modo de uso",
        "forma de uso",
        "uso externo",
        "uso topico",
        "aplicacion"
    ]

    # --------------------------------------------------------
    # 8. CLASIFICAR
    # --------------------------------------------------------

    def clasificar_accion_53(fila):

        producto = str(
            fila["Producto"]
        ).strip()

        accion = str(
            fila["Acción general"]
        ).strip()

        accion_n = normalizar_53(
            accion
        )

        producto_n = normalizar_53(
            producto
        )

        # --------------------------------------------
        # FRASE COMERCIAL
        # --------------------------------------------

        for patron in patrones_eliminar_53:
            if contiene_termino_53(
                accion_n,
                patron
            ):
                return "FRASE COMERCIAL"

        # --------------------------------------------
        # COMBINACIONES
        # --------------------------------------------

        for patron in patrones_combinacion_53:
            if contiene_termino_53(
                accion_n,
                patron
            ):
                return "RECOMENDACIÓN DE COMBINACIÓN CON OTRO PRODUCTO"


        # --------------------------------------------
        # RESTRICCIONES
        # --------------------------------------------

        for patron in patrones_restriccion_53:
            if contiene_termino_53(
                accion_n,
                patron
            ):
                return "RESTRICCIÓN"

        
        # --------------------------------------------
        # USO / POSOLOGÍA
        # --------------------------------------------

        for patron in patrones_uso_53:
            if contiene_termino_53(
                accion_n,
                patron
            ):
                return "USO / POSOLOGÍA / PRECAUCIÓN"

        # --------------------------------------------
        # COMPONENTES PROPIOS DEL PRODUCTO
        # --------------------------------------------

        componentes_propios = (
            componentes_producto_53.get(
                producto_n,
                []
            )
        )

        componente_propio_encontrado = False

        for componente in componentes_propios:
            if contiene_termino_53(
                accion_n,
                componente
            ):
                componente_propio_encontrado = True
                break

        if componente_propio_encontrado:
            return "ACCIÓN DE COMPONENTE"

     

        # --------------------------------------------
        # OTRO COMPONENTE
        #
        # Si aparece un componente que NO pertenece
        # al producto actual, es una posible
        # recomendación/complemento.
        # --------------------------------------------

        for componente_global in componentes_globales_53:

            if contiene_termino_53(
                accion_n,
                componente_global
            ):
                return "RECOMENDACIÓN DE COMBINACIÓN CON OTRO PRODUCTO"

        # --------------------------------------------
        # OTRO PRODUCTO
        # --------------------------------------------

        for otro_producto in productos_globales_53:

            otro_n = normalizar_53(
                otro_producto
            )

            if not otro_n:
                continue

            if otro_n == producto_n:
                continue

            if contiene_termino_53(
                accion_n,
                otro_producto
            ):
                return "RECOMENDACIÓN DE COMBINACIÓN CON OTRO PRODUCTO"

       
        # --------------------------------------------
        # ACCIÓN GENERAL
        # --------------------------------------------

        return "ACCIÓN GENERAL"

    # --------------------------------------------------------
    # 9. EJECUTAR CLASIFICACIÓN
    # --------------------------------------------------------

    df_clasificacion_53 = df_acciones_52.copy()

    df_clasificacion_53[
        "Clasificación"
    ] = df_clasificacion_53.apply(
        clasificar_accion_53,
        axis=1
    )

    # --------------------------------------------------------
    # 10. GUARDAR RESULTADO
    # --------------------------------------------------------

    st.session_state[
        "df_clasificacion_53"
    ] = df_clasificacion_53.copy()

    # --------------------------------------------------------
    # 11. MOSTRAR RESULTADO
    # --------------------------------------------------------

    st.success(
        f"🟢 5.3 TERMINADO: "
        f"{len(df_clasificacion_53)} relaciones clasificadas."
    )

    st.dataframe(
        df_clasificacion_53,
        use_container_width=True,
        hide_index=True
    )

except Exception as e:

    st.error(
        f"🔴 5.3 ERROR: "
        f"{type(e).__name__}: {e}"
    )

# ============================================================
# 5.4 APRENDIZAJE DE CLASIFICACIÓN DE ACCIONES
# ============================================================

st.markdown("### 5.4 Aprendizaje de clasificación")

try:

    import re
    import unicodedata
    import random

    # ========================================================
    # 1. RECUPERAR RESULTADO COMPLETO DE 5.3
    # ========================================================

    df_clasificacion_53 = st.session_state.get(
        "df_clasificacion_53"
    )

    if (
        df_clasificacion_53 is None
        or not isinstance(df_clasificacion_53, pd.DataFrame)
        or df_clasificacion_53.empty
    ):
        st.error(
            "❌ 5.4 ERROR: No existe el resultado completo de 5.3."
        )
        st.stop()

    # ========================================================
    # 2. VALIDAR COLUMNAS
    # ========================================================

    requeridas_54 = [
        "Producto",
        "Acción general",
        "Clasificación"
    ]

    faltantes_54 = [
        columna
        for columna in requeridas_54
        if columna not in df_clasificacion_53.columns
    ]

    if faltantes_54:
        st.error(
            "❌ 5.4 ERROR: Faltan columnas de 5.3: "
            + ", ".join(faltantes_54)
        )
        st.stop()

    # ========================================================
    # 3. CATEGORÍAS DE APRENDIZAJE
    # ========================================================

    categorias_aprendizaje_54 = [
        "ACCIÓN GENERAL",
        "COMPONENTE",
        "ACCIÓN DE COMPONENTE",
        "NO APLICA",
        "RECOMENDACIÓN DE COMBINACIÓN CON OTRO PRODUCTO",
        "FRASE COMERCIAL",
        "RESTRICCIÓN",
        "POSOLOGÍA"
    ]

    # ========================================================
    # 4. NORMALIZACIÓN
    # ========================================================

    def normalizar_54(texto):

        texto = str(texto).strip().lower()

        texto = unicodedata.normalize(
            "NFD",
            texto
        )

        texto = "".join(
            caracter
            for caracter in texto
            if unicodedata.category(caracter) != "Mn"
        )

        texto = re.sub(
            r"[^a-z0-9]+",
            " ",
            texto
        )

        return re.sub(
            r"\s+",
            " ",
            texto
        ).strip()

    # ========================================================
    # 5. CREAR CLAVE ÚNICA DE APRENDIZAJE
    #
    # IMPORTANTE:
    # Esta clave NO modifica el dataframe de 5.3.
    # Solo permite saber qué registros ya fueron enseñados.
    # ========================================================

    def clave_aprendizaje_54(fila):

        producto = normalizar_54(
            fila.get("Producto", "")
        )

        accion = normalizar_54(
            fila.get("Acción general", "")
        )

        return producto + "||" + accion

    # ========================================================
    # 6. RECUPERAR BASE DE APRENDIZAJE
    #
    # ESTA BASE ES INDEPENDIENTE DEL RESULTADO DE 5.3.
    # ========================================================

    if "aprendizaje_54" not in st.session_state:

        st.session_state["aprendizaje_54"] = pd.DataFrame(
            columns=[
                "Producto",
                "Acción general",
                "Clasificación 5.3",
                "Validación"
            ]
        )

    df_aprendizaje_54 = st.session_state[
        "aprendizaje_54"
    ].copy()

    # ========================================================
    # 7. DETERMINAR REGISTROS YA APRENDIDOS
    # ========================================================

    claves_aprendidas_54 = set()

    if not df_aprendizaje_54.empty:

        for _, fila_aprendida in df_aprendizaje_54.iterrows():

            clave = (
                normalizar_54(
                    fila_aprendida.get(
                        "Producto",
                        ""
                    )
                )
                + "||"
                + normalizar_54(
                    fila_aprendida.get(
                        "Acción general",
                        ""
                    )
                )
            )

            if clave != "||":
                claves_aprendidas_54.add(
                    clave
                )

    # ========================================================
    # 8. CREAR COPIA DEL RESULTADO DE 5.3
    #
    # NUNCA SE MODIFICA df_clasificacion_53
    # ========================================================

    df_trabajo_54 = df_clasificacion_53.copy()

    df_trabajo_54[
        "_clave_54"
    ] = df_trabajo_54.apply(
        clave_aprendizaje_54,
        axis=1
    )

    # ========================================================
    # 9. EXCLUIR LO QUE YA FUE ENSEÑADO
    # ========================================================

    df_pendientes_54 = df_trabajo_54[
        ~df_trabajo_54[
            "_clave_54"
        ].isin(
            claves_aprendidas_54
        )
    ].copy()

    # ========================================================
    # 10. CONFIGURACIÓN DEL LOTE
    # ========================================================

    TAMANO_LOTE_54 = 10

    if "lote_54" not in st.session_state:
        st.session_state["lote_54"] = 0

    # ========================================================
    # 11. INFORMACIÓN DEL APRENDIZAJE
    # ========================================================

    total_53 = len(
        df_clasificacion_53
    )

    total_aprendidas = len(
        df_aprendizaje_54
    )

    total_pendientes = len(
        df_pendientes_54
    )

    st.info(
        f"5.3 contiene **{total_53:,} funciones**. "
        f"Ya se han utilizado **{total_aprendidas:,}** "
        f"para aprendizaje. "
        f"Quedan **{total_pendientes:,}** disponibles."
    )

    # ========================================================
    # 12. SI NO HAY MÁS REGISTROS
    # ========================================================

    if df_pendientes_54.empty:

        st.success(
            "🟢 No hay más funciones pendientes "
            "para este proceso de aprendizaje."
        )

        st.stop()

    # ========================================================
    # 13. GENERAR LOTE
    #
    # El lote se conserva en session_state.
    # No cambia cada vez que Streamlit hace rerun.
    # ========================================================

    if (
        "lote_registros_54"
        not in st.session_state
    ):

        # ----------------------------------------------------
        # Intentar distribuir el lote entre las clasificaciones
        # automáticas existentes.
        # ----------------------------------------------------

        grupos_54 = []

        clasificaciones_existentes = (
            df_pendientes_54[
                "Clasificación"
            ]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        for clasificacion in (
            clasificaciones_existentes
        ):

            grupo = df_pendientes_54[
                df_pendientes_54[
                    "Clasificación"
                ].astype(str)
                == str(clasificacion)
            ]

            if not grupo.empty:
                grupos_54.append(
                    grupo
                )

        seleccionados_54 = []

        # ----------------------------------------------------
        # Tomar primero una muestra de cada grupo.
        # Esto evita que el primer lote tenga únicamente
        # ACCIÓN GENERAL.
        # ----------------------------------------------------

        for grupo in grupos_54:

            if len(
                seleccionados_54
            ) >= TAMANO_LOTE_54:
                break

            fila = grupo.sample(
                n=1,
                random_state=(
                    1000
                    + st.session_state[
                        "lote_54"
                    ]
                )
            )

            seleccionados_54.append(
                fila
            )

        # ----------------------------------------------------
        # Completar hasta 10 registros.
        # ----------------------------------------------------

        if seleccionados_54:

            df_seleccionados_54 = pd.concat(
                seleccionados_54,
                ignore_index=True
            )

        else:

            df_seleccionados_54 = pd.DataFrame()

        claves_seleccionadas_54 = set(
            df_seleccionados_54[
                "_clave_54"
            ].tolist()
        )

        restantes_54 = df_pendientes_54[
            ~df_pendientes_54[
                "_clave_54"
            ].isin(
                claves_seleccionadas_54
            )
        ]

        faltan_54 = (
            TAMANO_LOTE_54
            - len(
                df_seleccionados_54
            )
        )

        if (
            faltan_54 > 0
            and not restantes_54.empty
        ):

            faltantes_lote_54 = restantes_54.sample(
                n=min(
                    faltan_54,
                    len(restantes_54)
                ),
                random_state=(
                    2000
                    + st.session_state[
                        "lote_54"
                    ]
                )
            )

            df_seleccionados_54 = pd.concat(
                [
                    df_seleccionados_54,
                    faltantes_lote_54
                ],
                ignore_index=True
            )

        st.session_state[
            "lote_registros_54"
        ] = df_seleccionados_54.copy()

    # ========================================================
    # 14. RECUPERAR LOTE ACTUAL
    # ========================================================

    lote_actual_54 = st.session_state[
        "lote_registros_54"
    ].copy()

    if lote_actual_54.empty:

        st.warning(
            "No se pudo generar un lote de aprendizaje."
        )

        st.stop()

    # ========================================================
    # 15. MOSTRAR LOTE
    # ========================================================

    st.subheader(
        "Lote de aprendizaje"
    )

    st.write(
        f"Se muestran **{len(lote_actual_54)} "
        f"funciones** para validar."
    )

    st.caption(
        "La validación realizada aquí solamente "
        "alimenta el aprendizaje. No reemplaza "
        "ni modifica el resultado completo de 5.3."
    )

    # ========================================================
    # 16. FORMULARIO
    #
    # IMPORTANTE:
    # Los selectores están dentro de un formulario.
    # Por eso seleccionar una categoría NO provoca
    # rerun ni mueve el scroll.
    # ========================================================

    with st.form(
        "formulario_aprendizaje_54",
        clear_on_submit=False
    ):

        validaciones_54 = []

        for posicion, (
            indice,
            fila
        ) in enumerate(
            lote_actual_54.iterrows()
        ):

            producto = str(
                fila["Producto"]
            ).strip()

            accion = str(
                fila["Acción general"]
            ).strip()

            clasificacion_53 = str(
                fila["Clasificación"]
            ).strip()

            st.markdown(
                f"**{posicion + 1}. {accion}**"
            )

            st.caption(
                f"Producto: {producto}"
            )

            st.caption(
                f"Clasificación automática 5.3: "
                f"**{clasificacion_53}**"
            )

            clasificacion_manual = st.selectbox(
                "Clasificación correcta:",
                [
                    "Seleccione..."
                ]
                + categorias_aprendizaje_54,
                key=(
                    f"clasificacion_54_{indice}"
                )
            )

            col1, col2 = st.columns(2)

            with col1:

                restriccion = st.checkbox(
                    "Restricciones",
                    key=(
                        f"restriccion_54_{indice}"
                    )
                )

            with col2:

                posologia = st.checkbox(
                    "Posología",
                    key=(
                        f"posologia_54_{indice}"
                    )
                )

            validaciones_54.append(
                {
                    "indice": indice,
                    "Producto": producto,
                    "Acción general": accion,
                    "Clasificación 5.3":
                        clasificacion_53,
                    "Validación":
                        clasificacion_manual,
                   
                }
            )

            st.divider()

        guardar_54 = st.form_submit_button(
            "💾 Guardar aprendizaje",
            use_container_width=True
        )

    # ========================================================
    # 17. GUARDAR APRENDIZAJE
    # ========================================================

    if guardar_54:

        validaciones_validas_54 = [
            registro
            for registro in validaciones_54
            if registro[
                "Validación"
            ] != "Seleccione..."
        ]

        if not validaciones_validas_54:

            st.warning(
                "⚠️ No se seleccionó ninguna "
                "clasificación para guardar."
            )

        else:

            df_nuevo_aprendizaje_54 = pd.DataFrame(
                validaciones_validas_54
            )

            # ------------------------------------------------
            # Eliminar la columna interna.
            # ------------------------------------------------

            columnas_aprendizaje_54 = [
                "Producto",
                "Acción general",
                "Clasificación 5.3",
                "Validación",
               
            ]

            df_nuevo_aprendizaje_54 = (
                df_nuevo_aprendizaje_54[
                    columnas_aprendizaje_54
                ].copy()
            )

            # ------------------------------------------------
            # Agregar al aprendizaje existente.
            # ------------------------------------------------

            df_aprendizaje_54 = pd.concat(
                [
                    df_aprendizaje_54,
                    df_nuevo_aprendizaje_54
                ],
                ignore_index=True
            )

            # ------------------------------------------------
            # Eliminar duplicados.
            # ------------------------------------------------

            df_aprendizaje_54["_clave"] = (
                df_aprendizaje_54[
                    "Producto"
                ].map(
                    normalizar_54
                )
                + "||"
                + df_aprendizaje_54[
                    "Acción general"
                ].map(
                    normalizar_54
                )
            )

            df_aprendizaje_54 = (
                df_aprendizaje_54
                .drop_duplicates(
                    subset=["_clave"],
                    keep="last"
                )
                .drop(
                    columns=["_clave"]
                )
                .reset_index(
                    drop=True
                )
            )

            # ------------------------------------------------
            # GUARDAR SOLAMENTE LA BASE DE APRENDIZAJE
            #
            # NO SE MODIFICA 5.3.
            # ------------------------------------------------

            st.session_state[
                "aprendizaje_54"
            ] = df_aprendizaje_54.copy()

            # ------------------------------------------------
            # El lote actual deja de estar pendiente.
            # ------------------------------------------------

            st.session_state.pop(
                "lote_registros_54",
                None
            )

            st.session_state[
                "lote_54"
            ] += 1

            st.success(
                f"🟢 Se guardaron "
                f"{len(validaciones_validas_54)} "
                f"clasificaciones como aprendizaje."
            )

            st.info(
                "El resultado completo de 5.3 "
                "NO fue modificado."
            )

            st.rerun()

    # ========================================================
    # 18. GENERAR OTRO LOTE
    # ========================================================

    st.divider()

    col_a, col_b = st.columns(2)

    with col_a:

        if st.button(
            "➕ Generar otro lote",
            key="otro_lote_54",
            use_container_width=True
        ):

            st.session_state.pop(
                "lote_registros_54",
                None
            )

            st.rerun()

    with col_b:

        st.metric(
            "Aprendizajes registrados",
            len(
                st.session_state[
                    "aprendizaje_54"
                ]
            )
        )

    # ========================================================
    # 19. MOSTRAR BASE DE APRENDIZAJE
    #
    # SOLO PARA CONSULTA.
    # NO ES EL DATAFRAME FINAL.
    # ========================================================

    if not df_aprendizaje_54.empty:

        with st.expander(
            "Ver aprendizaje acumulado"
        ):

            st.dataframe(
                df_aprendizaje_54,
                use_container_width=True,
                hide_index=True
            )

except Exception as e:

    st.error(
        f"🔴 5.4 ERROR: "
        f"{type(e).__name__}: {e}"
    )
# ============================================================
# 5.5 PERSISTENCIA Y ANÁLISIS DEL APRENDIZAJE DE 5.4
# ============================================================

st.markdown("### 5.5 Persistencia y análisis del aprendizaje")

try:

    import base64
    import json
    import urllib.request
    import urllib.error

    # ========================================================
    # 1. CONFIGURACIÓN DE PERSISTENCIA
    # ========================================================

    GITHUB_USUARIO_55 = "franquiciasauces"
    GITHUB_REPOSITORIO_55 = "Asesores"
    GITHUB_RAMA_55 = "main"

    ARCHIVO_APRENDIZAJE_55 = (
        "APRENDIZAJE_54.csv"
    )

    URL_GITHUB_55 = (
        "https://api.github.com/repos/"
        f"{GITHUB_USUARIO_55}/"
        f"{GITHUB_REPOSITORIO_55}/contents/"
        f"{ARCHIVO_APRENDIZAJE_55}"
    )

    # ========================================================
    # 2. VALIDAR TOKEN
    # ========================================================

    if not GITHUB_TOKEN:
        st.error(
            "❌ 5.5 ERROR: No existe GITHUB_TOKEN."
        )
        st.stop()

    # ========================================================
    # 3. COLUMNAS OBLIGATORIAS DEL APRENDIZAJE
    #
    # Estas son las columnas reales que debe conservar 5.5.
    # ========================================================

    columnas_aprendizaje_55 = [
        "Producto",
        "Acción general",
        "Clasificación 5.3",
        "Validación"
    ]

    # ========================================================
    # 4. CATEGORÍAS VÁLIDAS DE VALIDACIÓN
    # ========================================================

    categorias_validacion_55 = [
        "ACCIÓN GENERAL",
        "COMPONENTE",
        "ACCIÓN DE COMPONENTE",
        "NO APLICA",
        "RECOMENDACIÓN DE COMBINACIÓN CON OTRO PRODUCTO",
        "FRASE COMERCIAL",
        "RESTRICCIÓN",
        "POSOLOGÍA"
    ]

    # ========================================================
    # 5. FUNCIÓN DE NORMALIZACIÓN
    # ========================================================

    def normalizar_55(texto):

        import unicodedata
        import re

        texto = str(texto).strip().lower()

        texto = unicodedata.normalize(
            "NFD",
            texto
        )

        texto = "".join(
            caracter
            for caracter in texto
            if unicodedata.category(caracter) != "Mn"
        )

        texto = re.sub(
            r"[^a-z0-9]+",
            " ",
            texto
        )

        return re.sub(
            r"\s+",
            " ",
            texto
        ).strip()

    # ========================================================
    # 6. CLAVE ÚNICA DE LA RELACIÓN
    #
    # IMPORTANTE:
    # Esta clave NO modifica 5.3.
    #
    # Sirve únicamente para identificar una relación
    # Producto + Acción general dentro del aprendizaje.
    # ========================================================

    def clave_relacion_55(fila):

        producto = normalizar_55(
            fila.get(
                "Producto",
                ""
            )
        )

        accion = normalizar_55(
            fila.get(
                "Acción general",
                ""
            )
        )

        return (
            producto
            + "||"
            + accion
        )

    # ========================================================
    # 7. LEER APRENDIZAJE TEMPORAL DE 5.4
    # ========================================================

    aprendizaje_temporal_55 = (
        st.session_state.get(
            "aprendizaje_54"
        )
    )

    if (
        aprendizaje_temporal_55 is None
        or not isinstance(
            aprendizaje_temporal_55,
            pd.DataFrame
        )
    ):

        aprendizaje_temporal_55 = pd.DataFrame(
            columns=columnas_aprendizaje_55
        )

    else:

        aprendizaje_temporal_55 = (
            aprendizaje_temporal_55.copy()
        )

    # ========================================================
    # 8. VALIDAR Y NORMALIZAR COLUMNAS
    # ========================================================

    faltantes_temporal_55 = [
        columna
        for columna in columnas_aprendizaje_55
        if columna not in aprendizaje_temporal_55.columns
    ]

    if faltantes_temporal_55:

        st.error(
            "❌ 5.5 ERROR: El aprendizaje de 5.4 "
            "no contiene las columnas obligatorias: "
            + ", ".join(
                faltantes_temporal_55
            )
        )

        st.stop()

    aprendizaje_temporal_55 = (
        aprendizaje_temporal_55[
            columnas_aprendizaje_55
        ].copy()
    )

    aprendizaje_temporal_55 = (
        aprendizaje_temporal_55.fillna("")
    )

    # ========================================================
    # 9. VALIDAR CATEGORÍAS DE VALIDACIÓN
    # ========================================================

    aprendizaje_temporal_55 = (
        aprendizaje_temporal_55[
            aprendizaje_temporal_55[
                "Validación"
            ].astype(str).isin(
                categorias_validacion_55
            )
        ].copy()
    )

    # ========================================================
    # 10. ANALIZAR COHERENCIA 5.3 VS VALIDACIÓN
    #
    # IMPORTANTE:
    #
    # Este análisis NO cambia la clasificación de 5.3.
    # Solamente registra si la validación humana confirmó
    # o corrigió la propuesta inicial.
    # ========================================================

    def analizar_coherencia_55(fila):

        clasificacion_53 = (
            str(
                fila[
                    "Clasificación 5.3"
                ]
            )
            .strip()
            .upper()
        )

        validacion = (
            str(
                fila[
                    "Validación"
                ]
            )
            .strip()
            .upper()
        )

        if (
            not clasificacion_53
            or not validacion
        ):
            return "SIN DATOS"

        if (
            clasificacion_53
            == validacion
        ):
            return "CORRECTA"

        return "CORREGIDA"

    aprendizaje_temporal_55[
        "Coherencia 5.3"
    ] = aprendizaje_temporal_55.apply(
        analizar_coherencia_55,
        axis=1
    )

    # ========================================================
    # 11. REGISTRAR LA CORRECCIÓN
    #
    # No reemplazamos la clasificación 5.3.
    # Conservamos ambas:
    #
    # Clasificación 5.3 = propuesta inicial
    # Validación = decisión humana
    # ========================================================

    aprendizaje_temporal_55[
        "Corrección 5.3"
    ] = aprendizaje_temporal_55.apply(
        lambda fila:
        "SIN CORRECCIÓN"
        if fila["Coherencia 5.3"] == "CORRECTA"
        else (
            str(
                fila["Clasificación 5.3"]
            ).strip()
            + " → "
            + str(
                fila["Validación"]
            ).strip()
        ),
        axis=1
    )

    # ========================================================
    # 12. LEER APRENDIZAJE PERSISTENTE DE GITHUB
    # ========================================================

    aprendizaje_persistente_55 = pd.DataFrame(
        columns=columnas_aprendizaje_55
    )

    sha_github_55 = None

    solicitud_get_55 = urllib.request.Request(
        URL_GITHUB_55,
        method="GET"
    )

    solicitud_get_55.add_header(
        "Authorization",
        f"Bearer {GITHUB_TOKEN}"
    )

    solicitud_get_55.add_header(
        "Accept",
        "application/vnd.github+json"
    )

    try:

        with urllib.request.urlopen(
            solicitud_get_55,
            timeout=30
        ) as respuesta_55:

            contenido_respuesta_55 = (
                json.loads(
                    respuesta_55.read().decode(
                        "utf-8"
                    )
                )
            )

        sha_github_55 = (
            contenido_respuesta_55.get(
                "sha"
            )
        )

        contenido_base64_55 = (
            contenido_respuesta_55.get(
                "content",
                ""
            )
        )

        if contenido_base64_55:

            contenido_bytes_55 = (
                base64.b64decode(
                    contenido_base64_55
                )
            )

            texto_csv_55 = (
                contenido_bytes_55.decode(
                    "utf-8-sig"
                )
            )

            if texto_csv_55.strip():

                aprendizaje_persistente_55 = (
                    pd.read_csv(
                        __import__(
                            "io"
                        ).StringIO(
                            texto_csv_55
                        ),
                        dtype=str
                    )
                )

    except urllib.error.HTTPError as error_55:

        if error_55.code == 404:

            st.info(
                "ℹ️ 5.5: Todavía no existe "
                "un archivo de aprendizaje persistente. "
                "Se creará con las validaciones actuales."
            )

        else:

            st.error(
                "❌ 5.5 ERROR al consultar GitHub: "
                f"HTTP {error_55.code}"
            )

            st.stop()

    except Exception as error_lectura_55:

        st.error(
            "❌ 5.5 ERROR al leer el aprendizaje "
            f"persistente: "
            f"{type(error_lectura_55).__name__}: "
            f"{error_lectura_55}"
        )

        st.stop()

    # ========================================================
    # 13. ASEGURAR COLUMNAS DEL ARCHIVO PERSISTENTE
    # ========================================================

    for columna in columnas_aprendizaje_55:

        if columna not in aprendizaje_persistente_55.columns:

            aprendizaje_persistente_55[
                columna
            ] = ""

    aprendizaje_persistente_55 = (
        aprendizaje_persistente_55[
            columnas_aprendizaje_55
        ].copy()
    )

    aprendizaje_persistente_55 = (
        aprendizaje_persistente_55.fillna("")
    )

    # ========================================================
    # 14. UNIR APRENDIZAJE HISTÓRICO + NUEVO
    #
    # El aprendizaje anterior NO se reemplaza.
    # ========================================================

    aprendizaje_consolidado_55 = pd.concat(
        [
            aprendizaje_persistente_55,
            aprendizaje_temporal_55[
                columnas_aprendizaje_55
            ]
        ],
        ignore_index=True
    )

    # ========================================================
    # 15. ELIMINAR DUPLICADOS
    #
    # Si una relación ya existe en el aprendizaje histórico,
    # la validación más reciente prevalece.
    #
    # Esto NO altera 5.3.
    # ========================================================

    aprendizaje_consolidado_55[
        "_clave_55"
    ] = aprendizaje_consolidado_55.apply(
        clave_relacion_55,
        axis=1
    )

    aprendizaje_consolidado_55 = (
        aprendizaje_consolidado_55[
            aprendizaje_consolidado_55[
                "_clave_55"
            ] != "||"
        ]
        .drop_duplicates(
            subset=["_clave_55"],
            keep="last"
        )
        .drop(
            columns=["_clave_55"]
        )
        .reset_index(
            drop=True
        )
    )

    # ========================================================
    # 16. VOLVER A CALCULAR EL ANÁLISIS SOBRE EL
    # APRENDIZAJE CONSOLIDADO
    # ========================================================

    aprendizaje_consolidado_55[
        "Coherencia 5.3"
    ] = aprendizaje_consolidado_55.apply(
        analizar_coherencia_55,
        axis=1
    )

    aprendizaje_consolidado_55[
        "Corrección 5.3"
    ] = aprendizaje_consolidado_55.apply(
        lambda fila:
        "SIN CORRECCIÓN"
        if fila["Coherencia 5.3"] == "CORRECTA"
        else (
            str(
                fila["Clasificación 5.3"]
            ).strip()
            + " → "
            + str(
                fila["Validación"]
            ).strip()
        ),
        axis=1
    )

    # ========================================================
    # 17. MOSTRAR RESUMEN DEL APRENDIZAJE
    # ========================================================

    total_historico_55 = len(
        aprendizaje_persistente_55
    )

    total_nuevo_55 = len(
        aprendizaje_temporal_55
    )

    total_consolidado_55 = len(
        aprendizaje_consolidado_55
    )

    total_correctas_55 = len(
        aprendizaje_consolidado_55[
            aprendizaje_consolidado_55[
                "Coherencia 5.3"
            ] == "CORRECTA"
        ]
    )

    total_corregidas_55 = len(
        aprendizaje_consolidado_55[
            aprendizaje_consolidado_55[
                "Coherencia 5.3"
            ] == "CORREGIDA"
        ]
    )

    st.info(
        f"Aprendizaje histórico: **{total_historico_55:,}**  \n"
        f"Validaciones nuevas de 5.4: **{total_nuevo_55:,}**  \n"
        f"Aprendizaje consolidado: **{total_consolidado_55:,}**"
    )

    col_55_a, col_55_b = st.columns(2)

    with col_55_a:

        st.metric(
            "5.3 correctas",
            total_correctas_55
        )

    with col_55_b:

        st.metric(
            "5.3 corregidas",
            total_corregidas_55
        )

    # ========================================================
    # 18. MOSTRAR CORRECCIONES
    # ========================================================

    correcciones_55 = (
        aprendizaje_consolidado_55[
            aprendizaje_consolidado_55[
                "Coherencia 5.3"
            ] == "CORREGIDA"
        ][
            [
                "Producto",
                "Acción general",
                "Clasificación 5.3",
                "Validación",
                "Corrección 5.3"
            ]
        ].copy()
    )

    if not correcciones_55.empty:

        with st.expander(
            "Ver correcciones de 5.3"
        ):

            st.dataframe(
                correcciones_55,
                use_container_width=True,
                hide_index=True
            )

    # ========================================================
    # 19. GUARDAR APRENDIZAJE PERSISTENTE EN GITHUB
    # ========================================================

    if st.button(
        "💾 Consolidar y guardar aprendizaje permanente",
        key="guardar_aprendizaje_55",
        use_container_width=True
    ):

        # ----------------------------------------------------
        # Convertir a CSV UTF-8
        # ----------------------------------------------------

        csv_55 = (
            aprendizaje_consolidado_55[
                [
                    "Producto",
                    "Acción general",
                    "Clasificación 5.3",
                    "Validación",
                    "Coherencia 5.3",
                    "Corrección 5.3"
                ]
            ]
            .to_csv(
                index=False
            )
        )

        contenido_base64_55 = (
            base64.b64encode(
                csv_55.encode(
                    "utf-8"
                )
            )
            .decode(
                "utf-8"
            )
        )

        # ----------------------------------------------------
        # Crear mensaje de commit
        # ----------------------------------------------------

        mensaje_commit_55 = (
            "FITOASISTE: actualizar aprendizaje 5.4"
        )

        datos_github_55 = {
            "message": mensaje_commit_55,
            "content": contenido_base64_55,
            "branch": GITHUB_RAMA_55
        }

        # ----------------------------------------------------
        # Si el archivo ya existe, GitHub exige SHA
        # ----------------------------------------------------

        if sha_github_55:

            datos_github_55[
                "sha"
            ] = sha_github_55

        cuerpo_github_55 = json.dumps(
            datos_github_55
        ).encode(
            "utf-8"
        )

        solicitud_put_55 = urllib.request.Request(
            URL_GITHUB_55,
            data=cuerpo_github_55,
            method="PUT"
        )

        solicitud_put_55.add_header(
            "Authorization",
            f"Bearer {GITHUB_TOKEN}"
        )

        solicitud_put_55.add_header(
            "Accept",
            "application/vnd.github+json"
        )

        solicitud_put_55.add_header(
            "Content-Type",
            "application/json"
        )

        try:

            with urllib.request.urlopen(
                solicitud_put_55,
                timeout=30
            ) as respuesta_put_55:

                resultado_put_55 = json.loads(
                    respuesta_put_55.read().decode(
                        "utf-8"
                    )
                )

            st.success(
                "🟢 5.5 TERMINADO: "
                "El aprendizaje fue guardado "
                "permanentemente en GitHub."
            )

            st.info(
                f"Archivo persistente: "
                f"**{ARCHIVO_APRENDIZAJE_55}**  \n"
                f"Registros consolidados: "
                f"**{total_consolidado_55:,}**"
            )

            # ------------------------------------------------
            # Actualizar la sesión con el aprendizaje
            # consolidado, pero SIN modificar 5.3.
            # ------------------------------------------------

            st.session_state[
                "aprendizaje_54"
            ] = aprendizaje_consolidado_55[
                columnas_aprendizaje_55
            ].copy()

        except urllib.error.HTTPError as error_put_55:

            detalle_55 = ""

            try:

                detalle_55 = (
                    error_put_55.read()
                    .decode(
                        "utf-8"
                    )
                )

            except Exception:
                detalle_55 = ""

            st.error(
                "❌ 5.5 ERROR al guardar en GitHub: "
                f"HTTP {error_put_55.code}"
            )

            if detalle_55:

                st.code(
                    detalle_55
                )

        except Exception as error_guardado_55:

            st.error(
                "❌ 5.5 ERROR al guardar el aprendizaje: "
                f"{type(error_guardado_55).__name__}: "
                f"{error_guardado_55}"
            )

    # ========================================================
    # 20. MOSTRAR APRENDIZAJE CONSOLIDADO
    #
    # SOLO CONSULTA.
    #
    # NO ES EL RESULTADO DE 5.3.
    # NO MODIFICA 5.3.
    # ========================================================

    with st.expander(
        "📚 Ver aprendizaje persistente consolidado"
    ):

        st.dataframe(
            aprendizaje_consolidado_55,
            use_container_width=True,
            hide_index=True
        )

    # ========================================================
    # 21. MENSAJE DE SEGURIDAD DEL FLUJO
    # ========================================================

    st.success(
        "🔒 5.5 funciona de forma independiente: "
        "el aprendizaje se obtiene de las validaciones humanas "
        "de 5.4, se conserva persistentemente y NO modifica "
        "ni retroalimenta directamente 5.3."
    )

except Exception as e:

    st.error(
        f"🔴 5.5 ERROR: "
        f"{type(e).__name__}: {e}"
    )
# ============================================================
# 5.6 — CONSOLIDACIÓN DE RELACIONES Y CLASIFICACIÓN FINAL
# ============================================================
# INSUMOS:
#   1. st.session_state["df_acciones_52"]
#      Producto + Acción general
#
#   2. APRENDIZAJE_54.csv
#      Resultado de las validaciones de 5.4 / aprendizaje de 5.5
#
# SALIDA:
#   MATRIZ_RELACIONES_56.csv
#
# IMPORTANTE:
#   5.6 NO vuelve a aprender.
#   5.6 utiliza la clasificación final aprendida en 5.5.
# ============================================================

import streamlit as st
import pandas as pd
import re
import unicodedata
from pathlib import Path


# ------------------------------------------------------------
# 5.6.1 — RUTAS
# ------------------------------------------------------------

BASE_DIR_56 = Path(__file__).resolve().parent

ARCHIVO_APRENDIZAJE_54 = BASE_DIR_56 / "APRENDIZAJE_54.csv"
ARCHIVO_MATRIZ_56 = BASE_DIR_56 / "MATRIZ_RELACIONES_56.csv"


# ------------------------------------------------------------
# 5.6.2 — NORMALIZACIÓN DE TEXTO
# ------------------------------------------------------------

def normalizar_texto_56(valor):
    """
    Normaliza texto únicamente para comparar.
    No modifica el texto original que se guarda.
    """

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    # Corrige casos de codificación UTF-8 interpretados como Latin-1
    # Ejemplo: AcciÃ³n -> Acción
    try:
        if "Ã" in texto or "Â" in texto or "â" in texto:
            texto_corregido = texto.encode(
                "latin1"
            ).decode(
                "utf-8"
            )
            texto = texto_corregido
    except Exception:
        pass

    texto = unicodedata.normalize(
        "NFKC",
        texto
    )

    texto = texto.replace("\n", " ")
    texto = texto.replace("\r", " ")

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip().casefold()


def limpiar_texto_56(valor):
    """
    Limpieza para conservar el valor legible.
    """

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    try:
        if "Ã" in texto or "Â" in texto or "â" in texto:
            texto = texto.encode(
                "latin1"
            ).decode(
                "utf-8"
            )
    except Exception:
        pass

    texto = unicodedata.normalize(
        "NFKC",
        texto
    )

    texto = texto.replace("\n", " ")
    texto = texto.replace("\r", " ")

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


# ------------------------------------------------------------
# 5.6.3 — NORMALIZACIÓN DE NOMBRES DE COLUMNAS
# ------------------------------------------------------------

def normalizar_nombre_columna_56(nombre):
    nombre = limpiar_texto_56(nombre)

    nombre = nombre.replace(
        "Acciónn general",
        "Acción general"
    )

    nombre = nombre.replace(
        "Clasificaciónn 5.3",
        "Clasificación 5.3"
    )

    nombre = nombre.replace(
        "Validaciónn",
        "Validación"
    )

    nombre = nombre.replace(
        "Correcciónn 5.3",
        "Corrección 5.3"
    )

    return nombre.strip()


def normalizar_columnas_56(df):

    df = df.copy()

    df.columns = [
        normalizar_nombre_columna_56(col)
        for col in df.columns
    ]

    return df


# ------------------------------------------------------------
# 5.6.4 — LEER APRENDIZAJE 5.5
# ------------------------------------------------------------

def cargar_aprendizaje_54_56():

    if not ARCHIVO_APRENDIZAJE_54.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo: "
            f"{ARCHIVO_APRENDIZAJE_54.name}"
        )

    ultimo_error = None

    # Se prueban las codificaciones habituales porque el archivo
    # actualmente puede contener texto con problemas de codificación.
    for encoding in [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin1"
    ]:

        try:

            df = pd.read_csv(
                ARCHIVO_APRENDIZAJE_54,
                encoding=encoding
            )

            df = normalizar_columnas_56(df)

            columnas_requeridas = {
                "Producto",
                "Acción general",
                "Clasificación 5.3",
                "Validación",
                "Coherencia 5.3",
                "Corrección 5.3"
            }

            faltantes = (
                columnas_requeridas
                - set(df.columns)
            )

            if not faltantes:
                return df

        except Exception as error:
            ultimo_error = error

    raise ValueError(
        "No fue posible leer APRENDIZAJE_54.csv "
        "con una estructura válida. "
        f"Último error: {ultimo_error}"
    )


# ------------------------------------------------------------
# 5.6.5 — OBTENER CLASIFICACIÓN FINAL
# ------------------------------------------------------------

def obtener_clasificacion_final_56(fila):

    validacion = limpiar_texto_56(
        fila.get("Validación", "")
    )

    correccion = limpiar_texto_56(
        fila.get("Corrección 5.3", "")
    )

    clasificacion_53 = limpiar_texto_56(
        fila.get("Clasificación 5.3", "")
    )

    # --------------------------------------------------------
    # NO APLICA
    # --------------------------------------------------------

    if normalizar_texto_56(validacion) == "no aplica":
        return "NO APLICA"

    # --------------------------------------------------------
    # La VALIDACIÓN humana es la clasificación final.
    #
    # Esto es especialmente importante cuando hubo corrección.
    # --------------------------------------------------------

    if validacion:
        clasificacion_final = validacion

    elif correccion:
        # Respaldo únicamente si por alguna razón la validación
        # está vacía.
        if "→" in correccion:
            clasificacion_final = correccion.split(
                "→"
            )[-1].strip()
        else:
            clasificacion_final = correccion

    else:
        clasificacion_final = clasificacion_53

    # --------------------------------------------------------
    # HOMOLOGACIÓN DEFINITIVA
    #
    # ELIMINAR ya no es una categoría válida.
    # Se conserva FRASE COMERCIAL.
    # --------------------------------------------------------

    if normalizar_texto_56(
        clasificacion_final
    ) == "eliminar":

        clasificacion_final = "FRASE COMERCIAL"

    return limpiar_texto_56(
        clasificacion_final
    )


# ------------------------------------------------------------
# 5.6.6 — CARGAR RELACIONES DE 5.2
# ------------------------------------------------------------

def obtener_acciones_52_56():

    if "df_acciones_52" not in st.session_state:

        raise ValueError(
            "No existe df_acciones_52 en st.session_state. "
            "Debe ejecutarse 5.2 antes de ejecutar 5.6."
        )

    df = st.session_state[
        "df_acciones_52"
    ].copy()

    df = normalizar_columnas_56(df)

    columnas_requeridas = {
        "Producto",
        "Acción general"
    }

    faltantes = (
        columnas_requeridas
        - set(df.columns)
    )

    if faltantes:

        raise ValueError(
            "df_acciones_52 no contiene las columnas "
            f"requeridas: {sorted(faltantes)}"
        )

    df = df[
        [
            "Producto",
            "Acción general"
        ]
    ].copy()

    df["Producto"] = df["Producto"].apply(
        limpiar_texto_56
    )

    df["Acción general"] = df[
        "Acción general"
    ].apply(
        limpiar_texto_56
    )

    # Elimina filas vacías
    df = df[
        (
            df["Producto"] != ""
        )
        &
        (
            df["Acción general"] != ""
        )
    ].copy()

    # Elimina duplicados de la estructura original
    df = df.drop_duplicates(
        subset=[
            "Producto",
            "Acción general"
        ]
    )

    return df.reset_index(
        drop=True
    )


# ------------------------------------------------------------
# 5.6.7 — CONSTRUIR CLAVES DE RELACIÓN
# ------------------------------------------------------------

def agregar_claves_56(df):

    df = df.copy()

    df["_producto_key_56"] = df[
        "Producto"
    ].apply(
        normalizar_texto_56
    )

    df["_accion_key_56"] = df[
        "Acción general"
    ].apply(
        normalizar_texto_56
    )

    return df


# ------------------------------------------------------------
# 5.6.8 — CONSOLIDAR 5.2 + APRENDIZAJE 5.5
# ------------------------------------------------------------

def construir_matriz_56():

    # --------------------------------------------
    # ORIGEN 5.2
    # --------------------------------------------

    df_acciones = obtener_acciones_52_56()

    # --------------------------------------------
    # APRENDIZAJE 5.5
    # --------------------------------------------

    df_aprendizaje = cargar_aprendizaje_54_56()

    # --------------------------------------------
    # CLASIFICACIÓN FINAL
    # --------------------------------------------

    df_aprendizaje[
        "Clasificación final"
    ] = df_aprendizaje.apply(
        obtener_clasificacion_final_56,
        axis=1
    )

    # --------------------------------------------
    # Limpiar datos
    # --------------------------------------------

    df_aprendizaje["Producto"] = (
        df_aprendizaje["Producto"]
        .apply(limpiar_texto_56)
    )

    df_aprendizaje["Acción general"] = (
        df_aprendizaje["Acción general"]
        .apply(limpiar_texto_56)
    )

    # --------------------------------------------
    # Crear claves
    # --------------------------------------------

    df_acciones = agregar_claves_56(
        df_acciones
    )

    df_aprendizaje = agregar_claves_56(
        df_aprendizaje
    )

    # --------------------------------------------
    # Solo necesitamos una clasificación final
    # por cada relación Producto + Acción.
    #
    # Si existen varias validaciones para la misma
    # relación, se conserva la última clasificación
    # disponible en el archivo de aprendizaje.
    # --------------------------------------------

    df_aprendizaje = df_aprendizaje[
        [
            "Producto",
            "Acción general",
            "Clasificación final",
            "_producto_key_56",
            "_accion_key_56"
        ]
    ].copy()

    df_aprendizaje = df_aprendizaje.drop_duplicates(
        subset=[
            "_producto_key_56",
            "_accion_key_56"
        ],
        keep="last"
    )

    # --------------------------------------------
    # CRUCE:
    #
    # 5.2 = estructura Producto + Acción
    # 5.5 = clasificación final aprendida
    # --------------------------------------------

    df_matriz = pd.merge(
        df_acciones,
        df_aprendizaje[
            [
                "_producto_key_56",
                "_accion_key_56",
                "Clasificación final"
            ]
        ],
        on=[
            "_producto_key_56",
            "_accion_key_56"
        ],
        how="left"
    )

    # --------------------------------------------
    # Marcar las relaciones sin aprendizaje.
    # NO inventamos una clasificación.
    # --------------------------------------------

    df_matriz[
        "Estado aprendizaje"
    ] = df_matriz[
        "Clasificación final"
    ].apply(
        lambda x:
        "APRENDIDA"
        if limpiar_texto_56(x)
        else "SIN APRENDIZAJE"
    )

    # --------------------------------------------
    # Orden definitivo
    # --------------------------------------------

    df_matriz = df_matriz[
        [
            "Producto",
            "Acción general",
            "Clasificación final",
            "Estado aprendizaje"
        ]
    ].copy()

    # --------------------------------------------
    # Limpieza final
    # --------------------------------------------

    df_matriz[
        "Clasificación final"
    ] = df_matriz[
        "Clasificación final"
    ].apply(
        limpiar_texto_56
    )

    # Homologación final de seguridad
    df_matriz[
        "Clasificación final"
    ] = df_matriz[
        "Clasificación final"
    ].replace(
        {
            "ELIMINAR": "FRASE COMERCIAL"
        }
    )

    df_matriz = df_matriz.drop_duplicates(
        subset=[
            "Producto",
            "Acción general"
        ]
    )

    return df_matriz.reset_index(
        drop=True
    )


# ------------------------------------------------------------
# 5.6.9 — EJECUTAR 5.6
# ------------------------------------------------------------

st.markdown(
    "### 5.6 — Consolidación de relaciones y clasificación final"
)

if st.button(
    "Generar matriz persistente 5.6",
    key="btn_generar_matriz_56"
):

    try:

        with st.spinner(
            "Construyendo matriz 5.6..."
        ):

            df_matriz_56 = construir_matriz_56()

            # ----------------------------------------
            # Guardar archivo persistente
            # ----------------------------------------

            df_matriz_56.to_csv(
                ARCHIVO_MATRIZ_56,
                index=False,
                encoding="utf-8-sig"
            )

            # ----------------------------------------
            # También queda disponible durante
            # la sesión actual
            # ----------------------------------------

            st.session_state[
                "df_matriz_56"
            ] = df_matriz_56.copy()

        st.success(
            "Matriz 5.6 generada correctamente."
        )

        st.write(
            f"Registros generados: "
            f"**{len(df_matriz_56):,}**"
        )

        st.write(
            f"Archivo persistente: "
            f"**{ARCHIVO_MATRIZ_56.name}**"
        )

        st.dataframe(
            df_matriz_56,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            label="Descargar matriz 5.6",
            data=df_matriz_56.to_csv(
                index=False,
                encoding="utf-8-sig"
            ),
            file_name="MATRIZ_RELACIONES_56.csv",
            mime="text/csv",
            key="download_matriz_56"
        )

    except Exception as e:

        st.error(
            "No fue posible generar la matriz 5.6."
        )

        st.exception(e)

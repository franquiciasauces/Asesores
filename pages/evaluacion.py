# =============================================GE===============
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
# ============================================================
# ============================================================
# FITOASISTE
# 5.6 — APLICACIÓN DEL APRENDIZAJE DE 5.5
# ============================================================
#
# 5.2:
#   df_acciones_52
#
# 5.5:
#   APRENDIZAJE_54.csv
#
# 5.6:
#   - aprende de las validaciones de 5.5
#   - genera embeddings EXCLUSIVOS para este aprendizaje
#   - aplica similitud semántica a TODAS las relaciones de 5.2
#   - no utiliza embeddings de síntomas
#   - filtra las relaciones que no quedan como ACCIÓN GENERAL
#   - genera archivo persistente
#
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata

from pathlib import Path

from sentence_transformers import SentenceTransformer
from sklearn.neighbors import KNeighborsClassifier
from sklearn.preprocessing import normalize


# ============================================================
# 5.6.1 — RUTAS
# ============================================================

BASE_DIR_56 = Path(__file__).resolve().parent

# evaluacion.py está en /pages
# Los archivos persistentes están en la raíz del repositorio.

RAIZ_REPOSITORIO_56 = BASE_DIR_56.parent

ARCHIVO_APRENDIZAJE_54 = (
    RAIZ_REPOSITORIO_56 /
    "APRENDIZAJE_54.csv"
)

ARCHIVO_EMBEDDINGS_56 = (
    RAIZ_REPOSITORIO_56 /
    "EMBEDDINGS_APRENDIZAJE_54.npy"
)

ARCHIVO_MEMORIA_56 = (
    RAIZ_REPOSITORIO_56 /
    "MEMORIA_APRENDIZAJE_54.csv"
)

ARCHIVO_SALIDA_56 = (
    RAIZ_REPOSITORIO_56 /
    "MATRIZ_56.csv"
)


# ============================================================
# 5.6.2 — MODELO SEMÁNTICO EXCLUSIVO
# ============================================================
#
# IMPORTANTE:
#
# Este modelo NO utiliza:
#
# embeddings_sintomas.npy
#
# Los embeddings de este módulo son independientes.
# ============================================================

MODELO_APRENDIZAJE_56 = (
    "SINAI/ALIA-MrBERT-es-biomedical-embeddings"
)


# ============================================================
# 5.6.3 — NORMALIZACIÓN
# ============================================================

def limpiar_texto_56(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    try:

        if (
            "Ã" in texto
            or "Â" in texto
            or "â" in texto
        ):
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

    texto = texto.replace(
        "\n",
        " "
    )

    texto = texto.replace(
        "\r",
        " "
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


def clave_56(valor):

    return limpiar_texto_56(
        valor
    ).casefold()


# ============================================================
# 5.6.4 — NORMALIZAR COLUMNAS
# ============================================================

def normalizar_columnas_56(df):

    df = df.copy()

    columnas = []

    for columna in df.columns:

        nombre = limpiar_texto_56(
            columna
        )

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

        columnas.append(
            nombre
        )

    df.columns = columnas

    return df


# ============================================================
# 5.6.5 — CARGAR APRENDIZAJE 5.5
# ============================================================

def cargar_aprendizaje_54_56():

    if not ARCHIVO_APRENDIZAJE_54.exists():

        raise FileNotFoundError(
            "No se encontró APRENDIZAJE_54.csv en: "
            f"{ARCHIVO_APRENDIZAJE_54}"
        )

    errores = []

    for encoding in (
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin1"
    ):

        try:

            df = pd.read_csv(
                ARCHIVO_APRENDIZAJE_54,
                encoding=encoding
            )

            df = normalizar_columnas_56(
                df
            )

            requeridas = {
                "Producto",
                "Acción general",
                "Clasificación 5.3",
                "Validación",
                "Coherencia 5.3",
                "Corrección 5.3"
            }

            faltantes = (
                requeridas
                - set(df.columns)
            )

            if not faltantes:
                return df

        except Exception as error:

            errores.append(
                str(error)
            )

    raise ValueError(
        "APRENDIZAJE_54.csv no tiene la "
        "estructura esperada. "
        f"Último error: {errores[-1:]}"
    )


# ============================================================
# 5.6.6 — OBTENER CLASIFICACIÓN FINAL
# ============================================================

def clasificacion_final_56(fila):

    validacion = limpiar_texto_56(
        fila.get(
            "Validación",
            ""
        )
    )

    correccion = limpiar_texto_56(
        fila.get(
            "Corrección 5.3",
            ""
        )
    )

    clasificacion_53 = limpiar_texto_56(
        fila.get(
            "Clasificación 5.3",
            ""
        )
    )

    # --------------------------------------------------------
    # La validación humana es la decisión final.
    # --------------------------------------------------------

    if validacion:

        resultado = validacion

    elif correccion:

        if "→" in correccion:

            resultado = (
                correccion
                .split("→")[-1]
                .strip()
            )

        else:

            resultado = correccion

    else:

        resultado = clasificacion_53

    # --------------------------------------------------------
    # HOMOLOGACIÓN
    #
    # ELIMINAR ya no existe como clasificación.
    # --------------------------------------------------------

    if clave_56(resultado) == "eliminar":

        resultado = "FRASE COMERCIAL"

    return limpiar_texto_56(
        resultado
    )


# ============================================================
# 5.6.7 — PREPARAR EJEMPLOS DE APRENDIZAJE
# ============================================================

def preparar_ejemplos_56(
    df
):

    df = df.copy()

    df[
        "Clasificación final"
    ] = df.apply(
        clasificacion_final_56,
        axis=1
    )

    df["Producto"] = (
        df["Producto"]
        .apply(limpiar_texto_56)
    )

    df["Acción general"] = (
        df["Acción general"]
        .apply(limpiar_texto_56)
    )

    df["Clasificación 5.3"] = (
        df["Clasificación 5.3"]
        .apply(limpiar_texto_56)
    )

    # --------------------------------------------------------
    # Solo usamos validaciones que realmente tienen
    # una decisión final.
    # --------------------------------------------------------

    df = df[
        df["Clasificación final"] != ""
    ].copy()

    return df.reset_index(
        drop=True
    )


# ============================================================
# 5.6.8 — TEXTO PARA EMBEDDING
# ============================================================
#
# No usamos solamente la acción.
#
# Combinamos:
#
#   acción
#   clasificación anterior
#   clasificación final
#
# Esto permite que el aprendizaje represente el tipo
# de relación que fue validado.
# ============================================================

def texto_aprendizaje_56(
    fila
):

    accion = limpiar_texto_56(
        fila["Acción general"]
    )

    clasificacion = limpiar_texto_56(
        fila["Clasificación 5.3"]
    )

    final = limpiar_texto_56(
        fila["Clasificación final"]
    )

    return (
        f"Acción: {accion}. "
        f"Clasificación inicial: {clasificacion}. "
        f"Clasificación validada: {final}."
    )


# ============================================================
# 5.6.9 — CARGAR MODELO
# ============================================================

@st.cache_resource(
    show_spinner=False
)
def cargar_modelo_aprendizaje_56():

    return SentenceTransformer(
        MODELO_APRENDIZAJE_56
    )


# ============================================================
# 5.6.10 — GENERAR EMBEDDINGS DEL APRENDIZAJE
# ============================================================

def generar_embeddings_aprendizaje_56(
    df
):

    modelo = (
        cargar_modelo_aprendizaje_56()
    )

    textos = [
        texto_aprendizaje_56(
            fila
        )
        for _, fila in df.iterrows()
    ]

    embeddings = modelo.encode(
        textos,
        convert_to_numpy=True,
        normalize_embeddings=True,
        show_progress_bar=False
    )

    embeddings = np.asarray(
        embeddings,
        dtype=np.float32
    )

    np.save(
        ARCHIVO_EMBEDDINGS_56,
        embeddings
    )

    return embeddings


# ============================================================
# 5.6.11 — CONSTRUIR MEMORIA PERSISTENTE
# ============================================================

def construir_memoria_56(
    df_aprendizaje,
    embeddings
):

    memoria = df_aprendizaje[
        [
            "Producto",
            "Acción general",
            "Clasificación 5.3",
            "Validación",
            "Coherencia 5.3",
            "Corrección 5.3",
            "Clasificación final"
        ]
    ].copy()

    memoria[
        "Embedding_ID"
    ] = np.arange(
        len(memoria)
    )

    memoria.to_csv(
        ARCHIVO_MEMORIA_56,
        index=False,
        encoding="utf-8-sig"
    )

    return memoria


# ============================================================
# 5.6.12 — OBTENER RELACIONES DE 5.2
# ============================================================

def obtener_relaciones_52_56():

    if (
        "df_acciones_52"
        not in st.session_state
    ):

        raise ValueError(
            "No existe df_acciones_52 en "
            "st.session_state. "
            "Debe ejecutarse 5.2 antes de 5.6."
        )

    df = st.session_state[
        "df_acciones_52"
    ].copy()

    if not isinstance(
        df,
        pd.DataFrame
    ):

        raise TypeError(
            "df_acciones_52 no es un DataFrame."
        )

    df = normalizar_columnas_56(
        df
    )

    requeridas = {
        "Producto",
        "Acción general"
    }

    faltantes = (
        requeridas
        - set(df.columns)
    )

    if faltantes:

        raise ValueError(
            "df_acciones_52 no contiene: "
            f"{sorted(faltantes)}"
        )

    # --------------------------------------------------------
    # NO reducimos las columnas.
    # Se conserva TODO lo que produjo 5.2.
    # --------------------------------------------------------

    for columna in [
        "Producto",
        "Acción general"
    ]:

        df[columna] = (
            df[columna]
            .apply(
                limpiar_texto_56
            )
        )

    df = df[
        (
            df["Producto"] != ""
        )
        &
        (
            df["Acción general"] != ""
        )
    ].copy()

    return df.reset_index(
        drop=True
    )


# ============================================================
# 5.6.13 — CLASIFICAR LAS 685 RELACIONES
# ============================================================
#
# ESTA ES LA PARTE FUNDAMENTAL.
#
# No hacemos merge exacto con APRENDIZAJE_54.
#
# El aprendizaje se utiliza para clasificar relaciones
# nuevas mediante similitud semántica.
# ============================================================

def clasificar_relaciones_56(
    df_52,
    df_aprendizaje,
    embeddings_aprendizaje
):

    modelo = (
        cargar_modelo_aprendizaje_56()
    )

    # --------------------------------------------------------
    # Texto de cada relación nueva.
    # --------------------------------------------------------

    textos_nuevos = []

    for _, fila in df_52.iterrows():

        accion = limpiar_texto_56(
            fila["Acción general"]
        )

        # Si 5.2 tiene una clasificación inicial,
        # la incorporamos.
        clasificacion_inicial = ""

        if (
            "Clasificación 5.3"
            in df_52.columns
        ):

            clasificacion_inicial = (
                limpiar_texto_56(
                    fila[
                        "Clasificación 5.3"
                    ]
                )
            )

        textos_nuevos.append(
            (
                f"Acción: {accion}. "
                f"Clasificación inicial: "
                f"{clasificacion_inicial}."
            )
        )

    embeddings_nuevos = modelo.encode(
        textos_nuevos,
        convert_to_numpy=True,
        normalize_embeddings=True,
        show_progress_bar=False
    )

    embeddings_nuevos = np.asarray(
        embeddings_nuevos,
        dtype=np.float32
    )

    # --------------------------------------------------------
    # KNN semántico.
    #
    # Usamos los ejemplos validados como memoria.
    # --------------------------------------------------------

    n_ejemplos = len(
        df_aprendizaje
    )

    if n_ejemplos < 1:

        raise ValueError(
            "No existen ejemplos de aprendizaje "
            "válidos en APRENDIZAJE_54.csv."
        )

    # --------------------------------------------------------
    # Número de vecinos.
    #
    # Con pocos ejemplos no debemos utilizar un K excesivo.
    # --------------------------------------------------------

    k = min(
        5,
        n_ejemplos
    )

    modelo_knn = KNeighborsClassifier(
        n_neighbors=k,
        weights="distance",
        metric="cosine"
    )

    etiquetas = (
        df_aprendizaje[
            "Clasificación final"
        ]
        .astype(str)
        .values
    )

    modelo_knn.fit(
        embeddings_aprendizaje,
        etiquetas
    )

    distancias, indices = (
        modelo_knn.kneighbors(
            embeddings_nuevos,
            n_neighbors=k
        )
    )

    resultados = []

    # --------------------------------------------------------
    # Para cada relación nueva:
    # obtener evidencia de los vecinos aprendidos.
    # --------------------------------------------------------

    for i in range(
        len(df_52)
    ):

        vecinos = indices[i]

        dist = distancias[i]

        votos = {}

        for posicion, distancia in zip(
            vecinos,
            dist
        ):

            etiqueta = limpiar_texto_56(
                etiquetas[posicion]
            )

            peso = 1.0 / (
                float(distancia)
                + 0.0001
            )

            votos[etiqueta] = (
                votos.get(
                    etiqueta,
                    0.0
                )
                + peso
            )

        if votos:

            clasificacion = max(
                votos,
                key=votos.get
            )

            peso_total = sum(
                votos.values()
            )

            confianza = (
                votos[
                    clasificacion
                ]
                / peso_total
                if peso_total > 0
                else 0.0
            )

        else:

            clasificacion = ""
            confianza = 0.0

        # ----------------------------------------------------
        # Homologación
        # ----------------------------------------------------

        if (
            clave_56(
                clasificacion
            )
            ==
            "eliminar"
        ):

            clasificacion = (
                "FRASE COMERCIAL"
            )

        resultados.append(
            (
                limpiar_texto_56(
                    clasificacion
                ),
                float(confianza)
            )
        )

    return resultados


# ============================================================
# 5.6.14 — CONSTRUIR RESULTADO FINAL
# ============================================================

def construir_56():

    # --------------------------------------------------------
    # 1. Relaciones completas de 5.2
    # --------------------------------------------------------

    df_52 = (
        obtener_relaciones_52_56()
    )

    total_52 = len(
        df_52
    )

    # --------------------------------------------------------
    # 2. Aprendizaje de 5.5
    # --------------------------------------------------------

    df_aprendizaje_original = (
        cargar_aprendizaje_54_56()
    )

    df_aprendizaje = (
        preparar_ejemplos_56(
            df_aprendizaje_original
        )
    )

    if len(
        df_aprendizaje
    ) == 0:

        raise ValueError(
            "APRENDIZAJE_54.csv no contiene "
            "ejemplos válidos de aprendizaje."
        )

    # --------------------------------------------------------
    # 3. Embeddings exclusivos de 5.6
    # --------------------------------------------------------

    embeddings = (
        generar_embeddings_aprendizaje_56(
            df_aprendizaje
        )
    )

    # --------------------------------------------------------
    # 4. Guardar memoria legible
    # --------------------------------------------------------

    construir_memoria_56(
        df_aprendizaje,
        embeddings
    )

    # --------------------------------------------------------
    # 5. Aplicar aprendizaje a TODAS las relaciones
    # --------------------------------------------------------

    resultados = (
        clasificar_relaciones_56(
            df_52,
            df_aprendizaje,
            embeddings
        )
    )

    df_resultado = df_52.copy()

    df_resultado[
        "Clasificación final 5.6"
    ] = [
        resultado[0]
        for resultado in resultados
    ]

    df_resultado[
        "Confianza 5.6"
    ] = [
        resultado[1]
        for resultado in resultados
    ]

    # --------------------------------------------------------
    # 6. Determinar cuáles son Acción General
    # --------------------------------------------------------

    df_resultado[
        "Es Acción general"
    ] = (
        df_resultado[
            "Clasificación final 5.6"
        ]
        .apply(
            lambda x:
            clave_56(x)
            ==
            "acción general"
        )
    )

    # --------------------------------------------------------
    # 7. FILTRO
    #
    # Solo pasan las relaciones que el aprendizaje
    # determina como ACCIÓN GENERAL.
    # --------------------------------------------------------

    df_final = df_resultado[
        df_resultado[
            "Es Acción general"
        ]
    ].copy()

    # --------------------------------------------------------
    # 8. Eliminar columnas técnicas
    # --------------------------------------------------------

    df_final = df_final.drop(
        columns=[
            "Es Acción general"
        ],
        errors="ignore"
    )

    # --------------------------------------------------------
    # 9. Eliminar duplicados de relación
    # --------------------------------------------------------

    df_final = df_final.drop_duplicates(
        subset=[
            "Producto",
            "Acción general"
        ]
    ).reset_index(
        drop=True
    )

    total_final = len(
        df_final
    )

    total_filtradas = (
        total_52
        - total_final
    )

    # --------------------------------------------------------
    # 10. Control matemático
    # --------------------------------------------------------

    if (
        total_final
        +
        total_filtradas
        !=
        total_52
    ):

        raise ValueError(
            "Error de integridad en 5.6: "
            "las relaciones no cuadran."
        )

    resumen = {
        "total_5_2": total_52,
        "total_aprendizaje": len(
            df_aprendizaje
        ),
        "total_clasificadas": len(
            df_resultado
        ),
        "total_conservadas": total_final,
        "total_filtradas": total_filtradas
    }

    return (
        df_final,
        df_resultado,
        resumen
    )


# ============================================================
# 5.6.15 — INTERFAZ
# ============================================================

st.markdown(
    "### 5.6 — Aplicación del aprendizaje de 5.5"
)

st.caption(
    "El aprendizaje validado se aplica semánticamente "
    "a todas las relaciones generadas por 5.2."
)


if st.button(
    "Ejecutar 5.6",
    key="ejecutar_56"
):

    try:

        with st.spinner(
            "Aplicando aprendizaje a las relaciones de 5.2..."
        ):

            (
                df_final_56,
                df_clasificado_56,
                resumen_56
            ) = construir_56()

            # ------------------------------------------------
            # Guardar archivo persistente
            # ------------------------------------------------

            df_final_56.to_csv(
                ARCHIVO_SALIDA_56,
                index=False,
                encoding="utf-8-sig"
            )

            # ------------------------------------------------
            # Guardar en sesión
            # ------------------------------------------------

            st.session_state[
                "df_matriz_56"
            ] = df_final_56.copy()

            st.session_state[
                "df_clasificado_56"
            ] = df_clasificado_56.copy()

        # ====================================================
        # RESULTADOS
        # ====================================================

        st.success(
            "5.6 terminó correctamente."
        )

        st.info(
            f"De {resumen_56['total_5_2']:,} "
            f"relaciones recibidas de 5.2, "
            f"se clasificaron "
            f"{resumen_56['total_clasificadas']:,} "
            f"mediante el aprendizaje de 5.5."
        )

        st.info(
            f"Se conservaron "
            f"{resumen_56['total_conservadas']:,} "
            f"relaciones que corresponden a "
            f"ACCIÓN GENERAL y se filtraron "
            f"{resumen_56['total_filtradas']:,}."
        )

        # ----------------------------------------------------
        # CONTROL
        # ----------------------------------------------------

        if (
            resumen_56["total_conservadas"]
            +
            resumen_56["total_filtradas"]
            ==
            resumen_56["total_5_2"]
        ):

            st.success(
                "Control correcto: "
                "conservadas + filtradas = relaciones de 5.2."
            )

        else:

            st.error(
                "ERROR: los conteos no coinciden."
            )

        # ----------------------------------------------------
        # Información
        # ----------------------------------------------------

        st.write(
            "Ejemplos utilizados para aprendizaje: "
            f"**{resumen_56['total_aprendizaje']:,}**"
        )

        st.write(
            "Archivo persistente generado: "
            f"**{ARCHIVO_SALIDA_56.name}**"
        )

        st.write(
            "Embeddings exclusivos del aprendizaje: "
            f"**{ARCHIVO_EMBEDDINGS_56.name}**"
        )

        # ----------------------------------------------------
        # Vista previa
        # ----------------------------------------------------

        st.dataframe(
            df_final_56,
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # Descargar
        # ----------------------------------------------------

        st.download_button(
            label="Descargar resultado 5.6",
            data=df_final_56.to_csv(
                index=False,
                encoding="utf-8-sig"
            ),
            file_name="MATRIZ_56.csv",
            mime="text/csv",
            key="descargar_56"
        )

    except Exception as error:

        st.error(
            "No fue posible ejecutar 5.6."
        )

        st.exception(
            error
        )

# ============================================================
# 5.7 — PERSISTENCIA DE PRODUCTO - ACCIONES GENERALES
# ============================================================
# Toma directamente el resultado generado por 5.6:
#
#     st.session_state["df_matriz_56"]
#
# y lo guarda persistentemente en GitHub como:
#
#     producto_accionesgenerales.csv
#
# NO clasifica.
# NO aprende.
# NO modifica las relaciones.
# NO elimina información.
# ============================================================


# ============================================================
# 1. CONFIGURACIÓN DE PERSISTENCIA
# ============================================================

GITHUB_USUARIO_57 = "franquiciasauces"
GITHUB_REPOSITORIO_57 = "Asesores"
GITHUB_RAMA_57 = "main"

ARCHIVO_PERSISTENTE_57 = (
    "producto_accionesgenerales.csv"
)

URL_GITHUB_57 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_57}/"
    f"{GITHUB_REPOSITORIO_57}/contents/"
    f"{ARCHIVO_PERSISTENTE_57}"
)


# ============================================================
# 2. VALIDAR TOKEN
# ============================================================

if not GITHUB_TOKEN:

    st.error(
        "❌ 5.7 ERROR: No existe GITHUB_TOKEN."
    )

    st.stop()


# ============================================================
# 3. FUNCIÓN DE PERSISTENCIA
# ============================================================

def guardar_producto_accionesgenerales_57(
    df_57
):

    # --------------------------------------------------------
    # Validar DataFrame
    # --------------------------------------------------------

    if df_57 is None:

        st.error(
            "❌ 5.7 ERROR: No se recibió el resultado de 5.6."
        )

        return False

    if df_57.empty:

        st.error(
            "❌ 5.7 ERROR: El resultado de 5.6 está vacío."
        )

        return False

    # --------------------------------------------------------
    # Convertir DataFrame a CSV
    # --------------------------------------------------------

    try:

        contenido_csv_57 = df_57.to_csv(
            index=False,
            encoding="utf-8-sig"
        )

    except Exception as error:

        st.error(
            "❌ 5.7 ERROR al convertir el resultado a CSV."
        )

        st.exception(error)

        return False

    # --------------------------------------------------------
    # Codificar contenido
    # --------------------------------------------------------

    contenido_base64_57 = base64.b64encode(
        contenido_csv_57.encode("utf-8-sig")
    ).decode("utf-8")

    # --------------------------------------------------------
    # Headers GitHub
    # --------------------------------------------------------

    headers_57 = {
        "Authorization": (
            f"token {GITHUB_TOKEN}"
        ),
        "Accept": (
            "application/vnd.github.v3+json"
        ),
        "Content-Type": (
            "application/json"
        )
    }

    # --------------------------------------------------------
    # Verificar si el archivo ya existe
    # --------------------------------------------------------

    sha_57 = None

    try:

        solicitud_get_57 = urllib.request.Request(
            URL_GITHUB_57,
            headers=headers_57,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud_get_57,
            timeout=30
        ) as respuesta_57:

            datos_existentes_57 = json.loads(
                respuesta_57.read().decode("utf-8")
            )

            sha_57 = datos_existentes_57.get(
                "sha"
            )

    except urllib.error.HTTPError as error:

        if error.code != 404:

            st.error(
                "❌ 5.7 ERROR consultando "
                "el archivo en GitHub."
            )

            st.exception(error)

            return False

        # 404 significa que todavía no existe.
        sha_57 = None

    except Exception as error:

        st.error(
            "❌ 5.7 ERROR conectando con GitHub."
        )

        st.exception(error)

        return False

    # --------------------------------------------------------
    # Preparar datos para GitHub
    # --------------------------------------------------------

    datos_github_57 = {
        "message": (
            "Persistencia 5.7 - "
            "producto_accionesgenerales.csv"
        ),
        "content": contenido_base64_57,
        "branch": GITHUB_RAMA_57
    }

    # --------------------------------------------------------
    # Si ya existe, enviar SHA para actualizarlo
    # --------------------------------------------------------

    if sha_57:

        datos_github_57["sha"] = sha_57

    # --------------------------------------------------------
    # Guardar / actualizar
    # --------------------------------------------------------

    try:

        cuerpo_57 = json.dumps(
            datos_github_57
        ).encode("utf-8")

        solicitud_put_57 = urllib.request.Request(
            URL_GITHUB_57,
            data=cuerpo_57,
            headers=headers_57,
            method="PUT"
        )

        with urllib.request.urlopen(
            solicitud_put_57,
            timeout=30
        ) as respuesta_put_57:

            resultado_github_57 = json.loads(
                respuesta_put_57.read().decode(
                    "utf-8"
                )
            )

        # ----------------------------------------------------
        # Confirmación
        # ----------------------------------------------------

        if resultado_github_57:

            st.success(
                "✅ 5.7: "
                "producto_accionesgenerales.csv "
                "quedó guardado persistentemente."
            )

            st.info(
                f"Relaciones persistidas: "
                f"**{len(df_57):,}**"
            )

            return True

    except urllib.error.HTTPError as error:

        try:

            detalle_error_57 = (
                error.read()
                .decode("utf-8")
            )

        except Exception:

            detalle_error_57 = str(error)

        st.error(
            "❌ 5.7 ERROR guardando en GitHub."
        )

        st.code(
            detalle_error_57
        )

        return False

    except Exception as error:

        st.error(
            "❌ 5.7 ERROR inesperado."
        )

        st.exception(error)

        return False

    return False


# ============================================================
# 4. INTERFAZ 5.7
# ============================================================

st.markdown(
    "### 5.7 — Persistencia de producto_accionesgenerales"
)


if (
    "df_matriz_56"
    not in st.session_state
):

    st.warning(
        "Primero ejecute 5.6. "
        "El resultado todavía no está disponible."
    )

else:

    df_resultado_57 = (
        st.session_state[
            "df_matriz_56"
        ].copy()
    )

    st.success(
        "Resultado de 5.6 disponible."
    )

    st.write(
        f"Relaciones listas para persistencia: "
        f"**{len(df_resultado_57):,}**"
    )

    if st.button(
        "💾 Guardar producto_accionesgenerales.csv",
        key="guardar_producto_accionesgenerales_57"
    ):

        guardar_producto_accionesgenerales_57(
            df_resultado_57
        )

        # ----------------------------------------------------
        # Descargar el mismo archivo persistido
        # ----------------------------------------------------

        st.download_button(
            label=(
                "⬇️ Descargar "
                "producto_accionesgenerales.csv"
            ),
            data=df_resultado_57.to_csv(
                index=False,
                encoding="utf-8-sig"
            ),
            file_name=(
                "producto_accionesgenerales.csv"
            ),
            mime="text/csv",
            key="descargar_producto_accionesgenerales_57"
        )

# ============================================================
# 5.8 — COMPONENTES Y ACCIONES
# ============================================================
# FUENTE:
# MATRIZ_PRODUCTO_PATOLOGIAS-PAQUETES
#
# SALIDA:
# COMPONENTES_Y_ACCIONES.csv
#
# COLUMNAS:
# Producto | Componente | Acciones
#
# IMPORTANTE:
# - NO modifica 5.7
# - NO muestra GITHUB_TOKEN
# - Usa el mismo GITHUB_TOKEN existente en la aplicación
# - La persistencia es automática
# ============================================================


# ============================================================
# 1. IMPORTACIONES ESPECÍFICAS DE 5.8
# ============================================================

import re
import base64
import urllib.request
import urllib.error
import json


# ============================================================
# 2. CONFIGURACIÓN DE 5.8
# ============================================================

GITHUB_USUARIO_58 = "franquiciasauces"

GITHUB_REPOSITORIO_58 = "Asesores"

GITHUB_RAMA_58 = "main"

ARCHIVO_SALIDA_58 = "COMPONENTES_Y_ACCIONES.csv"

URL_GITHUB_58 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_58}/"
    f"{GITHUB_REPOSITORIO_58}/contents/"
    f"{ARCHIVO_SALIDA_58}"
)


# ============================================================
# 3. TOKEN
# ============================================================
# IMPORTANTE:
# NO imprimir, mostrar ni escribir el token en pantalla.
#
# Si GITHUB_TOKEN ya fue definido al inicio de evaluacion.py,
# se reutiliza.
# ============================================================

if "GITHUB_TOKEN" not in globals():

    try:

        GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]

    except Exception:

        GITHUB_TOKEN = ""


# ============================================================
# 4. NORMALIZAR TEXTO
# ============================================================

def normalizar_texto_58(texto):

    if texto is None:
        return ""

    try:

        if pd.isna(texto):
            return ""

    except Exception:

        pass

    texto = str(texto).strip()

    texto = unidecode(
        texto
    ).lower()

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


# ============================================================
# 5. SEPARAR COMPONENTES
# ============================================================

def separar_componentes_58(valor):

    if valor is None:
        return []

    try:

        if pd.isna(valor):
            return []

    except Exception:

        pass

    texto = str(
        valor
    ).strip()

    if not texto:
        return []

    partes = re.split(
        r"\s*;\s*|\s*\|\s*",
        texto
    )

    resultado = []

    for parte in partes:

        parte = parte.strip()

        if not parte:
            continue

        if normalizar_texto_58(
            parte
        ) in {
            "nan",
            "none",
            "-"
        }:
            continue

        if parte not in resultado:

            resultado.append(
                parte
            )

    return resultado


# ============================================================
# 6. MARCADORES QUE NO SON ACCIONES
# ============================================================

MARCADORES_CORTE_58 = [

    "combinaciones:",

    "combinaciones",

    "se recomienda:",

    "se recomienda",

    "recomendado:",

    "recomendado",

    "recomendación:",

    "recomendación",

    "recomendaciones:",

    "recomendaciones",

    "frase de venta:",

    "frase de venta",

    "frase comercial:",

    "frase comercial",

]


# ============================================================
# 7. CORTAR CONTENIDO POSTERIOR
# ============================================================

def cortar_contenido_58(texto):

    if not texto:
        return ""

    original = str(
        texto
    )

    normalizado = normalizar_texto_58(
        original
    )

    posiciones = []

    for marcador in MARCADORES_CORTE_58:

        marcador_n = normalizar_texto_58(
            marcador
        )

        posicion = normalizado.find(
            marcador_n
        )

        if posicion >= 0:

            posiciones.append(
                posicion
            )

    if posiciones:

        limite = min(
            posiciones
        )

        original = original[
            :limite
        ]

    return original.strip()


# ============================================================
# 8. ELIMINAR "MODO DE ACCIÓN"
# ============================================================

def limpiar_modo_accion_58(texto):

    if not texto:
        return ""

    resultado = str(
        texto
    ).strip()

    patrones = [

        r"^\s*MODO\s+DE\s+ACCI[ÓO]N\s*:\s*",

        r"^\s*MODO\s+DE\s+ACCI[ÓO]N\s*-\s*",

        r"^\s*MODO\s+DE\s+ACCI[ÓO]N\s*=\s*",

    ]

    cambio = True

    while cambio:

        cambio = False

        for patron in patrones:

            nuevo = re.sub(
                patron,
                "",
                resultado,
                flags=re.IGNORECASE
            )

            if nuevo != resultado:

                resultado = nuevo.strip()

                cambio = True

    return resultado


# ============================================================
# 9. SEPARAR ACCIONES
# ============================================================

def separar_acciones_58(texto):

    if not texto:
        return []

    texto = cortar_contenido_58(
        texto
    )

    texto = limpiar_modo_accion_58(
        texto
    )

    if not texto:
        return []

    partes = re.split(
        r"\s*;\s*",
        texto
    )

    acciones = []

    for parte in partes:

        accion = parte.strip()

        if not accion:
            continue

        accion = limpiar_modo_accion_58(
            accion
        )

        normalizada = normalizar_texto_58(
            accion
        )

        if not normalizada:
            continue

        # ----------------------------------------------------
        # No permitir que estas secciones entren como acción.
        # ----------------------------------------------------

        prohibidas = [

            "se recomienda",

            "recomendado",

            "recomendacion",

            "recomendaciones",

            "frase de venta",

            "frase comercial",

        ]

        if any(
            normalizada.startswith(x)
            for x in prohibidas
        ):

            continue

        accion = re.sub(
            r"\s+",
            " ",
            accion
        ).strip()

        if accion and accion not in acciones:

            acciones.append(
                accion
            )

    return acciones


# ============================================================
# 10. DETECTAR COMPONENTES EXPLÍCITOS
# ============================================================

def componentes_mencionados_58(
    accion,
    componentes
):

    if not accion:
        return []

    accion_normalizada = normalizar_texto_58(
        accion
    )

    encontrados = []

    componentes_ordenados = sorted(
        componentes,
        key=lambda x: len(
            normalizar_texto_58(x)
        ),
        reverse=True
    )

    for componente in componentes_ordenados:

        componente_n = normalizar_texto_58(
            componente
        )

        if not componente_n:
            continue

        patron = (
            r"(?<!\w)"
            +
            re.escape(
                componente_n
            )
            +
            r"(?!\w)"
        )

        if re.search(
            patron,
            accion_normalizada
        ):

            if componente not in encontrados:

                encontrados.append(
                    componente
                )

    return encontrados


# ============================================================
# 11. LIMPIAR REFERENCIAS AL COMPONENTE
# ============================================================

def limpiar_referencias_componentes_58(
    accion,
    componentes_identificados
):

    if not accion:
        return ""

    resultado = str(
        accion
    ).strip()

    if not componentes_identificados:

        return resultado

    # --------------------------------------------------------
    # Primero eliminar paréntesis cuyo contenido corresponde
    # únicamente a componentes identificados.
    # --------------------------------------------------------

    def limpiar_parentesis(match):

        contenido = match.group(
            1
        ).strip()

        componentes_dentro = (
            componentes_mencionados_58(
                contenido,
                componentes_identificados
            )
        )

        if not componentes_dentro:

            return match.group(
                0
            )

        restante = normalizar_texto_58(
            contenido
        )

        for componente in componentes_identificados:

            componente_n = (
                normalizar_texto_58(
                    componente
                )
            )

            restante = restante.replace(
                componente_n,
                ""
            )

        restante = re.sub(
            r"[\+\&,;/]+",
            "",
            restante
        )

        restante = restante.replace(
            " y ",
            ""
        )

        restante = re.sub(
            r"\s+",
            "",
            restante
        )

        if not restante:

            return ""

        return match.group(
            0
        )

    resultado = re.sub(
        r"\(([^()]*)\)",
        limpiar_parentesis,
        resultado
    )

    # --------------------------------------------------------
    # Después eliminar referencia directa restante.
    # --------------------------------------------------------

    for componente in sorted(
        componentes_identificados,
        key=lambda x: len(
            normalizar_texto_58(x)
        ),
        reverse=True
    ):

        componente_n = normalizar_texto_58(
            componente
        )

        if not componente_n:
            continue

        patron = (
            r"(?<!\w)"
            +
            re.escape(
                componente_n
            )
            +
            r"(?!\w)"
        )

        resultado = re.sub(
            patron,
            "",
            resultado,
            flags=re.IGNORECASE
        )

    # --------------------------------------------------------
    # Limpieza final.
    # --------------------------------------------------------

    resultado = re.sub(
        r"\(\s*\)",
        "",
        resultado
    )

    resultado = re.sub(
        r"\s+",
        " ",
        resultado
    )

    resultado = re.sub(
        r"\s+([,.;:])",
        r"\1",
        resultado
    )

    resultado = resultado.strip(
        " ;,-"
    )

    return resultado.strip()


# ============================================================
# 12. GENERAR RELACIONES
# ============================================================

def generar_relaciones_58(
    df_origen
):

    relaciones = []

    for _, fila in df_origen.iterrows():

        # ====================================================
        # PRODUCTO
        # ====================================================

        producto = str(
            fila.get(
                "Producto",
                ""
            )
        ).strip()

        if not producto:
            continue

        # ====================================================
        # COMPONENTES
        # ====================================================

        componentes = (
            separar_componentes_58(
                fila.get(
                    "Componentes",
                    ""
                )
            )
        )

        if not componentes:
            continue

        # ====================================================
        # ACCIONES
        # ====================================================

        acciones = (
            separar_acciones_58(
                fila.get(
                    "Acciones generales",
                    ""
                )
            )
        )

        if not acciones:
            continue

        # ====================================================
        # PRODUCTO CON UN SOLO COMPONENTE
        # ====================================================

        if len(componentes) == 1:

            componente = componentes[0]

            for accion in acciones:

                accion_limpia = (
                    limpiar_referencias_componentes_58(
                        accion,
                        [componente]
                    )
                )

                if not accion_limpia:
                    continue

                relaciones.append({

                    "Producto":
                        producto,

                    "Componente":
                        componente,

                    "Acciones":
                        accion_limpia

                })

        # ====================================================
        # PRODUCTO CON VARIOS COMPONENTES
        # ====================================================

        else:

            for accion in acciones:

                componentes_identificados = (
                    componentes_mencionados_58(
                        accion,
                        componentes
                    )
                )

                # --------------------------------------------
                # REGLA FUNDAMENTAL:
                # si no aparece explícitamente el componente,
                # NO se crea relación.
                # --------------------------------------------

                if not componentes_identificados:

                    continue

                accion_limpia = (
                    limpiar_referencias_componentes_58(
                        accion,
                        componentes_identificados
                    )
                )

                if not accion_limpia:
                    continue

                for componente in (
                    componentes_identificados
                ):

                    relaciones.append({

                        "Producto":
                            producto,

                        "Componente":
                            componente,

                        "Acciones":
                            accion_limpia

                    })

    # ========================================================
    # DATAFRAME FINAL
    # ========================================================

    df_resultado = pd.DataFrame(
        relaciones,
        columns=[
            "Producto",
            "Componente",
            "Acciones"
        ]
    )

    if not df_resultado.empty:

        df_resultado = (
            df_resultado
            .drop_duplicates(
                subset=[
                    "Producto",
                    "Componente",
                    "Acciones"
                ]
            )
            .reset_index(
                drop=True
            )
        )

    return df_resultado


# ============================================================
# 13. CARGAR MATRIZ
# ============================================================

def cargar_matriz_58():

    nombres = [

        "MATRIZ_PRODUCTO_PATOLOGIAS-PAQUETES.xlsx",

        "MATRIZ_PRODUCTO_PATOLOGIAS-PAQUETES.xls",

        "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx",

        "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xls",

    ]

    for nombre in nombres:

        try:

            df = pd.read_excel(
                nombre
            )

            columnas = set(
                df.columns
            )

            if {
                "Producto",
                "Componentes",
                "Acciones generales"
            }.issubset(columnas):

                return df

        except Exception:

            continue

    return None


# ============================================================
# 14. PERSISTENCIA
# ============================================================

def guardar_github_58(
    df_resultado
):

    # --------------------------------------------------------
    # NO mostrar jamás GITHUB_TOKEN.
    # --------------------------------------------------------

    if not GITHUB_TOKEN:

        st.error(
            "❌ 5.8: GITHUB_TOKEN no está configurado."
        )

        return False

    try:

        csv_texto = (
            df_resultado.to_csv(
                index=False,
                encoding="utf-8-sig"
            )
        )

        contenido = (
            base64.b64encode(
                csv_texto.encode(
                    "utf-8-sig"
                )
            ).decode(
                "utf-8"
            )
        )

        headers = {

            "Authorization":
                f"token {GITHUB_TOKEN}",

            "Accept":
                "application/vnd.github+json",

            "Content-Type":
                "application/json",

            "User-Agent":
                "FITOASISTE"

        }

        sha = None

        # ====================================================
        # CONSULTAR ARCHIVO
        # ====================================================

        try:

            request_get = (
                urllib.request.Request(
                    URL_GITHUB_58,
                    headers=headers,
                    method="GET"
                )
            )

            with urllib.request.urlopen(
                request_get,
                timeout=30
            ) as response:

                datos = json.loads(
                    response.read().decode(
                        "utf-8"
                    )
                )

                sha = datos.get(
                    "sha"
                )

        except urllib.error.HTTPError as error:

            if error.code != 404:

                st.error(
                    "❌ 5.8: No fue posible "
                    "consultar el archivo en GitHub."
                )

                return False

        # ====================================================
        # CREAR / ACTUALIZAR
        # ====================================================

        payload = {

            "message":
                "5.8 - Actualizar "
                "COMPONENTES_Y_ACCIONES.csv",

            "content":
                contenido,

            "branch":
                GITHUB_RAMA_58

        }

        if sha:

            payload[
                "sha"
            ] = sha

        body = json.dumps(
            payload
        ).encode(
            "utf-8"
        )

        request_put = (
            urllib.request.Request(
                URL_GITHUB_58,
                data=body,
                headers=headers,
                method="PUT"
            )
        )

        with urllib.request.urlopen(
            request_put,
            timeout=30
        ) as response:

            resultado = json.loads(
                response.read().decode(
                    "utf-8"
                )
            )

        if resultado.get(
            "content"
        ):

            return True

        return False

    except urllib.error.HTTPError as error:

        st.error(
            "❌ 5.8: GitHub rechazó "
            "la actualización."
        )

        # IMPORTANTE:
        # NO mostramos el contenido del error
        # porque podría contener información sensible.

        return False

    except Exception:

        st.error(
            "❌ 5.8: Error inesperado "
            "durante la persistencia."
        )

        return False


# ============================================================
# 15. INTERFAZ
# ============================================================

st.markdown(
    "### 5.8 — Componentes y Acciones"
)

st.write(
    "Genera las relaciones "
    "Producto + Componente + Acción."
)


# ============================================================
# 16. GENERAR
# ============================================================

if st.button(
    "🔎 Generar COMPONENTES_Y_ACCIONES.csv",
    key="generar_componentes_acciones_58"
):

    # ========================================================
    # CARGAR
    # ========================================================

    df_matriz_58 = (
        cargar_matriz_58()
    )

    if df_matriz_58 is None:

        st.error(
            "❌ 5.8: No se encontró "
            "MATRIZ_PRODUCTO_PATOLOGIAS-PAQUETES "
            "con las columnas requeridas:"
        )

        st.code(
            "Producto\n"
            "Componentes\n"
            "Acciones generales"
        )

        st.stop()

    # ========================================================
    # GENERAR
    # ========================================================

    df_resultado_58 = (
        generar_relaciones_58(
            df_matriz_58
        )
    )

    # ========================================================
    # SESIÓN
    # ========================================================

    st.session_state[
        "df_componentes_acciones_58"
    ] = df_resultado_58.copy()

    # ========================================================
    # TOTALIZACIÓN
    # ========================================================

    total_relaciones_58 = len(
        df_resultado_58
    )

    if not df_resultado_58.empty:

        total_productos_58 = (
            df_resultado_58[
                "Producto"
            ].nunique()
        )

        total_componentes_58 = (
            df_resultado_58[
                "Componente"
            ].nunique()
        )

    else:

        total_productos_58 = 0

        total_componentes_58 = 0

    # ========================================================
    # RESULTADOS
    # ========================================================

    st.success(
        "✅ 5.8 terminó correctamente."
    )

    st.info(
        f"Relaciones generadas: "
        f"**{total_relaciones_58:,}**"
    )

    st.info(
        f"Productos involucrados: "
        f"**{total_productos_58:,}**"
    )

    st.info(
        f"Componentes involucrados: "
        f"**{total_componentes_58:,}**"
    )

    # ========================================================
    # PERSISTENCIA AUTOMÁTICA
    # ========================================================

    with st.spinner(
        "Guardando persistentemente..."
    ):

        guardado_58 = (
            guardar_github_58(
                df_resultado_58
            )
        )

    if guardado_58:

        st.success(
            "✅ Persistencia correcta."
        )

        st.write(
            "Archivo persistente: "
            "**COMPONENTES_Y_ACCIONES.csv**"
        )

    else:

        st.error(
            "❌ Se generó el resultado, "
            "pero no fue posible guardarlo "
            "persistentemente."
        )

    # ========================================================
    # VISTA PREVIA
    # ========================================================

    st.markdown(
        "#### Vista previa"
    )

    st.dataframe(
        df_resultado_58,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # DESCARGA
    # ========================================================

    csv_58 = (
        df_resultado_58.to_csv(
            index=False,
            encoding="utf-8-sig"
        )
    )

    st.download_button(
        label=(
            "⬇️ Descargar "
            "COMPONENTES_Y_ACCIONES.csv"
        ),
        data=csv_58,
        file_name=(
            "COMPONENTES_Y_ACCIONES.csv"
        ),
        mime="text/csv",
        key="descargar_componentes_acciones_58"
    )


# ============================================================
# 6.1 - PRODUCTO / ACCIÓN GENERAL
# PARTE 1 - CARGA Y CONTROL DE FUENTES
# ============================================================

# ------------------------------------------------------------
# ARCHIVOS
# ------------------------------------------------------------

ARCHIVO_FUENTE_61 = "producto_accionesgenerales.csv"
ARCHIVO_BANCO_61 = "BANCO_PREGUNTAS_GENERALES.csv"


# ------------------------------------------------------------
# FUNCIONES
# ------------------------------------------------------------

def normalizar_61(valor):

    if pd.isna(valor):
        return ""

    return (
        str(valor)
        .strip()
        .lower()
        .replace("\n", " ")
    )


def cargar_fuente_61():

    try:

        df = pd.read_csv(
            ARCHIVO_FUENTE_61,
            encoding="utf-8-sig"
        )

    except UnicodeDecodeError:

        df = pd.read_csv(
            ARCHIVO_FUENTE_61,
            encoding="latin-1"
        )

    columnas = [
        "Producto",
        "Acción general",
        "Clasificación final 5.6",
        "Confianza 5.6"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "6.1 ERROR: faltan columnas en "
            f"{ARCHIVO_FUENTE_61}: "
            f"{', '.join(faltantes)}"
        )

        return None

    df = df[columnas].copy()

    df["Producto"] = (
        df["Producto"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Acción general"] = (
        df["Acción general"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df = df[
        (df["Producto"] != "")
        &
        (df["Acción general"] != "")
    ].copy()

    df["Fuente_ID"] = [
        f"PTAG-F{i:06d}"
        for i in range(1, len(df) + 1)
    ]

    df["_clave"] = (
        df["Producto"]
        .map(normalizar_61)
        + "||"
        + df["Acción general"]
        .map(normalizar_61)
    )

    df = (
        df
        .drop_duplicates(
            subset="_clave"
        )
        .reset_index(drop=True)
    )

    return df


def cargar_banco_61():

    try:

        df = pd.read_csv(
            ARCHIVO_BANCO_61,
            encoding="utf-8-sig"
        )

    except FileNotFoundError:

        return pd.DataFrame()

    except UnicodeDecodeError:

        df = pd.read_csv(
            ARCHIVO_BANCO_61,
            encoding="latin-1"
        )

    return df


# ------------------------------------------------------------
# IDENTIFICAR RELACIONES YA UTILIZADAS
# ------------------------------------------------------------

def obtener_fuentes_usadas_61(df_banco):

    usadas = set()

    if df_banco.empty:
        return usadas

    if "Fuente_ID" not in df_banco.columns:
        return usadas

    for valor in df_banco["Fuente_ID"].fillna(""):

        texto = str(valor).strip()

        if not texto:
            continue

        for fuente in texto.split(";"):

            fuente = fuente.strip()

            if fuente:
                usadas.add(fuente)

    return usadas


# ------------------------------------------------------------
# IDENTIFICAR PREGUNTAS EXISTENTES
# ------------------------------------------------------------

def obtener_preguntas_existentes_61(df_banco):

    preguntas = set()

    if df_banco.empty:
        return preguntas

    if "Pregunta" not in df_banco.columns:
        return preguntas

    for pregunta in df_banco["Pregunta"].fillna(""):

        clave = normalizar_61(pregunta)

        if clave:
            preguntas.add(clave)

    return preguntas


# ============================================================
# INTERFAZ 6.1
# ============================================================

st.markdown(
    "## 6.1 Producto - Acción General"
)

st.write(
    "Control de relaciones disponibles para "
    "la generación de preguntas."
)


if st.button(
    "🔎 CARGAR Y VALIDAR FUENTES 6.1",
    key="cargar_fuentes_61"
):

    df_fuente_61 = cargar_fuente_61()

    if df_fuente_61 is None:
        st.stop()

    df_banco_61 = cargar_banco_61()

    fuentes_usadas_61 = (
        obtener_fuentes_usadas_61(
            df_banco_61
        )
    )

    preguntas_existentes_61 = (
        obtener_preguntas_existentes_61(
            df_banco_61
        )
    )

    # --------------------------------------------------------
    # RELACIONES DISPONIBLES
    # --------------------------------------------------------

    df_disponible_61 = (
        df_fuente_61[
            ~df_fuente_61["Fuente_ID"].isin(
                fuentes_usadas_61
            )
        ]
        .copy()
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # GUARDAR EN SESIÓN
    # --------------------------------------------------------

    st.session_state[
        "df_fuente_61"
    ] = df_fuente_61.copy()

    st.session_state[
        "df_banco_61"
    ] = df_banco_61.copy()

    st.session_state[
        "df_disponible_61"
    ] = df_disponible_61.copy()

    st.session_state[
        "fuentes_usadas_61"
    ] = fuentes_usadas_61

    st.session_state[
        "preguntas_existentes_61"
    ] = preguntas_existentes_61


# ============================================================
# MOSTRAR CONTROL
# ============================================================

if (
    "df_fuente_61"
    in st.session_state
):

    df_fuente_61 = (
        st.session_state[
            "df_fuente_61"
        ]
    )

    df_banco_61 = (
        st.session_state[
            "df_banco_61"
        ]
    )

    df_disponible_61 = (
        st.session_state[
            "df_disponible_61"
        ]
    )

    fuentes_usadas_61 = (
        st.session_state[
            "fuentes_usadas_61"
        ]
    )

    preguntas_existentes_61 = (
        st.session_state[
            "preguntas_existentes_61"
        ]
    )


    # --------------------------------------------------------
    # RESUMEN
    # --------------------------------------------------------

    st.success(
        "6.1 cargó correctamente las fuentes."
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Relaciones fuente",
            len(df_fuente_61)
        )

    with col2:

        st.metric(
            "Relaciones ya utilizadas",
            len(fuentes_usadas_61)
        )

    with col3:

        st.metric(
            "Relaciones disponibles",
            len(df_disponible_61)
        )


    st.info(
        "Preguntas existentes en el banco: "
        f"{len(preguntas_existentes_61):,}"
    )


    # --------------------------------------------------------
    # CONTROL DE CONSISTENCIA
    # --------------------------------------------------------

    total_fuente = len(
        df_fuente_61
    )

    total_usadas = len(
        fuentes_usadas_61
        &
        set(
            df_fuente_61[
                "Fuente_ID"
            ]
        )
    )

    total_disponibles = len(
        df_disponible_61
    )


    if (
        total_usadas
        +
        total_disponibles
        ==
        total_fuente
    ):

        st.success(
            "Control correcto: "
            "utilizadas + disponibles = "
            "relaciones fuente."
        )

    else:

        st.error(
            "6.1 ERROR: los conteos "
            "de relaciones no coinciden."
        )


    # --------------------------------------------------------
    # VISTA PREVIA
    # --------------------------------------------------------

    st.markdown(
        "### Relaciones disponibles"
    )

    st.dataframe(
        df_disponible_61[
            [
                "Fuente_ID",
                "Producto",
                "Acción general"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
# ============================================================
# 6.1 - PARTE 2
# GENERADOR PRODUCTO - ACCIÓN GENERAL
# ============================================================

# ============================================================
# 1. FUNCIONES DE CONTROL
# ============================================================

def clave_relacion_61(producto, accion):

    return (
        normalizar_61(producto)
        + "||"
        + normalizar_61(accion)
    )


def obtener_relaciones_consumidas_61():

    consumidas = set()

    if "df_banco_61" in st.session_state:

        df_banco = st.session_state[
            "df_banco_61"
        ]

        if (
            not df_banco.empty
            and "Fuente_ID" in df_banco.columns
        ):

            for valor in df_banco[
                "Fuente_ID"
            ].fillna(""):

                for fuente in str(
                    valor
                ).split(";"):

                    fuente = fuente.strip()

                    if fuente:
                        consumidas.add(fuente)

    # Relaciones consumidas durante esta sesión

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_61",
            set()
        )
    )

    return consumidas


def obtener_claves_preguntas_61():

    claves = set()

    if "df_banco_61" not in st.session_state:
        return claves

    df_banco = st.session_state[
        "df_banco_61"
    ]

    if df_banco.empty:
        return claves

    # --------------------------------------------------------
    # Se utiliza también Fuente_ID para evitar reutilización.
    # --------------------------------------------------------

    if "Fuente_ID" in df_banco.columns:

        for valor in df_banco[
            "Fuente_ID"
        ].fillna(""):

            fuentes = sorted(
                [
                    x.strip()
                    for x in str(valor).split(";")
                    if x.strip()
                ]
            )

            if fuentes:
                claves.add(
                    "FUENTES::"
                    + "||".join(fuentes)
                )

    return claves


def siguiente_id_61():

    mayor = 0

    df_banco = st.session_state.get(
        "df_banco_61",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco[
            "Pregunta_ID"
        ].fillna(""):

            coincidencia = re.match(
                r"PTAG-(\d+)",
                str(valor).strip()
            )

            if coincidencia:

                mayor = max(
                    mayor,
                    int(
                        coincidencia.group(1)
                    )
                )

    preguntas_actuales = st.session_state.get(
        "preguntas_generadas_61",
        []
    )

    for pregunta in preguntas_actuales:

        coincidencia = re.match(
            r"PTAG-(\d+)",
            str(
                pregunta.get(
                    "Pregunta_ID",
                    ""
                )
            )
        )

        if coincidencia:

            mayor = max(
                mayor,
                int(
                    coincidencia.group(1)
                )
            )

    return f"PTAG-{mayor + 1:06d}"


# ============================================================
# 2. GENERAR NIVEL 1
# ============================================================

def generar_nivel_1_61(df_disponible, consumidas):

    if len(df_disponible) < 4:
        return None

    candidatos = df_disponible[
        ~df_disponible[
            "Fuente_ID"
        ].isin(consumidas)
    ].copy()

    if len(candidatos) < 4:
        return None

    # --------------------------------------------------------
    # Buscar una relación verdadera
    # --------------------------------------------------------

    verdaderas = candidatos.sample(
        frac=1,
        random_state=None
    )

    for _, verdadera in verdaderas.iterrows():

        producto = verdadera[
            "Producto"
        ]

        clave_producto = normalizar_61(
            producto
        )

        # ----------------------------------------------------
        # Las falsas deben corresponder a otros productos.
        # ----------------------------------------------------

        falsas = candidatos[
            candidatos["Producto"].map(
                normalizar_61
            ) != clave_producto
        ].copy()

        if len(falsas) < 3:
            continue

        falsas = falsas.sample(
            n=3,
            random_state=None
        )

        opciones = pd.concat(
            [
                pd.DataFrame([verdadera]),
                falsas
            ],
            ignore_index=True
        )

        opciones = opciones.sample(
            frac=1,
            random_state=None
        ).reset_index(drop=True)

        fuentes = list(
            opciones["Fuente_ID"]
        )

        if len(set(fuentes)) != 4:
            continue

        return {
            "Producto": producto,
            "Opciones": opciones,
            "Correctas": [
                int(
                    opciones.index[
                        opciones["Fuente_ID"]
                        == verdadera["Fuente_ID"]
                    ][0]
                ) + 1
            ]
        }

    return None


# ============================================================
# 3. GENERAR NIVEL 2
# ============================================================

def generar_nivel_2_61(df_disponible, consumidas):

    candidatos = df_disponible[
        ~df_disponible[
            "Fuente_ID"
        ].isin(consumidas)
    ].copy()

    if len(candidatos) < 4:
        return None

    # --------------------------------------------------------
    # Buscar productos que tengan por lo menos 2 acciones.
    # --------------------------------------------------------

    grupos = (
        candidatos
        .groupby("Producto", sort=False)
        .filter(
            lambda grupo: len(grupo) >= 2
        )
    )

    if grupos.empty:
        return None

    productos = (
        grupos["Producto"]
        .drop_duplicates()
        .tolist()
    )

    np.random.shuffle(
        productos
    )

    for producto in productos:

        verdaderas = grupos[
            grupos["Producto"] == producto
        ].copy()

        if len(verdaderas) < 2:
            continue

        verdaderas = verdaderas.sample(
            n=2,
            random_state=None
        )

        claves_verdaderas = set(
            verdaderas["Fuente_ID"]
        )

        # ----------------------------------------------------
        # Las falsas deben provenir de otros productos.
        # ----------------------------------------------------

        falsas = candidatos[
            ~candidatos["Fuente_ID"].isin(
                claves_verdaderas
            )
        ].copy()

        falsas = falsas[
            falsas["Producto"].map(
                normalizar_61
            ) != normalizar_61(producto)
        ]

        if len(falsas) < 2:
            continue

        falsas = falsas.sample(
            n=2,
            random_state=None
        )

        opciones = pd.concat(
            [
                verdaderas.assign(
                    _correcta=True
                ),
                falsas.assign(
                    _correcta=False
                )
            ],
            ignore_index=True
        )

        opciones = opciones.sample(
            frac=1,
            random_state=None
        ).reset_index(drop=True)

        fuentes = list(
            opciones["Fuente_ID"]
        )

        if len(set(fuentes)) != 4:
            continue

        correctas = [
            i + 1
            for i, valor in enumerate(
                opciones["_correcta"]
            )
            if valor
        ]

        return {
            "Producto": producto,
            "Opciones": opciones,
            "Correctas": correctas
        }

    return None


# ============================================================
# 4. CONSTRUIR PREGUNTA
# ============================================================

def construir_pregunta_61(
    resultado,
    nivel
):

    opciones = resultado[
        "Opciones"
    ]

    producto = resultado[
        "Producto"
    ]

    if nivel == "Nivel 1":

        texto = (
            "¿Cuál de las siguientes "
            "acciones generales corresponde "
            f"al producto {producto}?"
        )

    else:

        texto = (
            "¿Cuáles de las siguientes "
            "acciones generales corresponden "
            f"al producto {producto}? "
            "Seleccione las dos opciones correctas."
        )

    pregunta_id = siguiente_id_61()

    fuentes = ";".join(
        opciones[
            "Fuente_ID"
        ].tolist()
    )

    return {

        "Pregunta_ID":
            pregunta_id,

        "Modulo":
            "Producto",

        "Tema":
            "Acción General",

        "Nivel":
            nivel,

        "Tipo_Relacion":
            "Producto-Acción General",

        "Pregunta":
            texto,

        "Respuesta_1":
            opciones.iloc[0][
                "Acción general"
            ],

        "Respuesta_2":
            opciones.iloc[1][
                "Acción general"
            ],

        "Respuesta_3":
            opciones.iloc[2][
                "Acción general"
            ],

        "Respuesta_4":
            opciones.iloc[3][
                "Acción general"
            ],

        "Respuesta_Correcta":
            ";".join(
                str(x)
                for x in resultado[
                    "Correctas"
                ]
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
            fuentes
    }


# ============================================================
# 5. GENERADOR GENERAL
# ============================================================

def generar_preguntas_61(
    cantidad,
    nivel
):

    df_disponible = st.session_state.get(
        "df_disponible_61",
        pd.DataFrame()
    )

    if df_disponible.empty:
        return []

    consumidas = obtener_relaciones_consumidas_61()

    preguntas = []

    while len(preguntas) < cantidad:

        if nivel == "Nivel 1":

            resultado = generar_nivel_1_61(
                df_disponible,
                consumidas
            )

        else:

            resultado = generar_nivel_2_61(
                df_disponible,
                consumidas
            )

        if resultado is None:
            break

        pregunta = construir_pregunta_61(
            resultado,
            nivel
        )

        preguntas.append(
            pregunta
        )

        # ----------------------------------------------------
        # Consumir las 4 relaciones inmediatamente.
        # ----------------------------------------------------

        fuentes = (
            resultado["Opciones"]
            ["Fuente_ID"]
            .tolist()
        )

        consumidas.update(
            fuentes
        )

    return preguntas


# ============================================================
# 6. INTERFAZ DE GENERACIÓN
# ============================================================

if (
    "df_disponible_61"
    in st.session_state
):

    st.markdown(
        "### Generación de preguntas"
    )

    cantidad_61 = st.number_input(
        "¿Cuántas preguntas desea generar?",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_61"
    )

    nivel_61 = st.selectbox(
        "Nivel de evaluación",
        [
            "Nivel 1",
            "Nivel 2"
        ],
        key="nivel_generar_61"
    )

    if st.button(
        "GENERAR PREGUNTAS",
        key="generar_preguntas_61"
    ):

        nuevas = generar_preguntas_61(
            cantidad_61,
            nivel_61
        )

        if not nuevas:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar nuevas "
                "preguntas con las condiciones "
                "establecidas."
            )

        else:

            # ------------------------------------------------
            # Guardar preguntas generadas
            # ------------------------------------------------

            st.session_state[
                "preguntas_generadas_61"
            ] = nuevas

            # ------------------------------------------------
            # Marcar relaciones consumidas
            # ------------------------------------------------

            consumidas = (
                st.session_state.get(
                    "fuentes_consumidas_61",
                    set()
                )
            )

            for pregunta in nuevas:

                for fuente in str(
                    pregunta[
                        "Fuente_ID"
                    ]
                ).split(";"):

                    fuente = fuente.strip()

                    if fuente:
                        consumidas.add(
                            fuente
                        )

            st.session_state[
                "fuentes_consumidas_61"
            ] = consumidas

            st.success(
                f"Se generaron "
                f"{len(nuevas)} preguntas."
            )

            st.info(
                "Las relaciones utilizadas en "
                "estas preguntas quedan consumidas "
                "y no volverán a utilizarse."
            )


# ============================================================
# 7. MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_61 = st.session_state.get(
    "preguntas_generadas_61",
    []
)

if preguntas_61:

    st.markdown(
        "### Preguntas generadas"
    )

    for pregunta in preguntas_61:

        st.markdown(
            f"**{pregunta['Pregunta_ID']} — "
            f"{pregunta['Nivel']}**"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"1. {pregunta['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta['Respuesta_4']}"
        )

        st.caption(
            "Respuesta correcta: "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta['Fuente_ID']}"
        )

        st.divider()
# ============================================================
# 6.1 - PARTE 3A
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS
# ============================================================

if preguntas_61:

    st.markdown(
        "## Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(preguntas_61):

        st.markdown(
            f"### {pregunta['Pregunta_ID']}"
        )

        st.write(
            f"**Nivel:** {pregunta['Nivel']}"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"**1.** {pregunta['Respuesta_1']}"
        )

        st.write(
            f"**2.** {pregunta['Respuesta_2']}"
        )

        st.write(
            f"**3.** {pregunta['Respuesta_3']}"
        )

        st.write(
            f"**4.** {pregunta['Respuesta_4']}"
        )

        st.write(
            "**Respuesta correcta:** "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            f"Fuente utilizada: {pregunta['Fuente_ID']}"
        )

        # ----------------------------------------------------
        # ESTADO ACTUAL
        # ----------------------------------------------------

        estado_actual = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado_actual}"
        )

        # ----------------------------------------------------
        # OBSERVACIÓN
        # ----------------------------------------------------

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_61_{i}"
        )

        # ----------------------------------------------------
        # BOTONES
        # ----------------------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "✅ APROBAR",
                key=f"aprobar_61_{i}"
            ):

                preguntas_61[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_61[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_61"
                ] = preguntas_61

                st.success(
                    f"{pregunta['Pregunta_ID']} "
                    "fue aprobada."
                )

                st.rerun()

        with col2:

            if st.button(
                "❌ RECHAZAR",
                key=f"rechazar_61_{i}"
            ):

                preguntas_61[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_61[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_61"
                ] = preguntas_61

                st.warning(
                    f"{pregunta['Pregunta_ID']} "
                    "fue rechazada."
                )

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN
# ============================================================

if preguntas_61:

    aprobadas_61 = sum(
        1
        for p in preguntas_61
        if p.get("Estado") == "APROBADA"
    )

    rechazadas_61 = sum(
        1
        for p in preguntas_61
        if p.get("Estado") == "RECHAZADA"
    )

    pendientes_61 = sum(
        1
        for p in preguntas_61
        if p.get("Estado", "PENDIENTE")
        == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_61
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_61
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_61
        )

    if pendientes_61 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "La sincronización con el banco "
            "se habilitará en la siguiente parte."
        )

# ============================================================
# 6.1 PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_61 = "franquiciasauces"
GITHUB_REPOSITORIO_61 = "Asesores"
GITHUB_RAMA_61 = "main"

GITHUB_ARCHIVO_61 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_61 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_61}/"
    f"{GITHUB_REPOSITORIO_61}/contents/"
    f"{GITHUB_ARCHIVO_61}"
)


def sincronizar_banco_61():

    preguntas = st.session_state.get(
        "preguntas_generadas_61",
        []
    )

    if not preguntas:
        st.warning(
            "No hay preguntas para sincronizar."
        )
        return

    if any(
        p.get("Estado", "PENDIENTE") == "PENDIENTE"
        for p in preguntas
    ):
        st.error(
            "Todavía hay preguntas pendientes de revisión."
        )
        return

    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    try:

        # ----------------------------------------------------
        # LEER ARCHIVO EXISTENTE DE GITHUB
        # ----------------------------------------------------

        solicitud = urllib.request.Request(
            URL_GITHUB_61,
            headers=headers,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            datos = json.loads(
                respuesta.read().decode("utf-8")
            )

        sha = datos["sha"]

        contenido = base64.b64decode(
            datos["content"].replace("\n", "")
        )

        df_banco = pd.read_excel(
            contenido
        )

        total_antes = len(df_banco)

        # ----------------------------------------------------
        # CONVERTIR PREGUNTAS VALIDADAS
        # ----------------------------------------------------

        df_nuevas = pd.DataFrame(
            preguntas
        )

        columnas = [
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

        df_nuevas = df_nuevas[
            columnas
        ].copy()

        # ----------------------------------------------------
        # EVITAR DUPLICAR PREGUNTAS YA EXISTENTES
        # ----------------------------------------------------

        if "Pregunta_ID" in df_banco.columns:

            existentes = set(
                df_banco[
                    "Pregunta_ID"
                ]
                .astype(str)
                .str.strip()
            )

            df_nuevas = df_nuevas[
                ~df_nuevas[
                    "Pregunta_ID"
                ]
                .astype(str)
                .str.strip()
                .isin(existentes)
            ]

        total_nuevas = len(df_nuevas)

        if total_nuevas == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                f"Preguntas existentes: "
                f"**{total_antes:,}**"
            )

            return

        # ----------------------------------------------------
        # AGREGAR AL BANCO
        # ----------------------------------------------------

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

        # ----------------------------------------------------
        # CREAR EXCEL EN MEMORIA
        # ----------------------------------------------------

        import io

        memoria = io.BytesIO()

        with pd.ExcelWriter(
            memoria,
            engine="openpyxl"
        ) as writer:

            df_final.to_excel(
                writer,
                index=False,
                sheet_name="Banco"
            )

        contenido_nuevo = base64.b64encode(
            memoria.getvalue()
        ).decode("utf-8")

        # ----------------------------------------------------
        # ACTUALIZAR GITHUB
        # ----------------------------------------------------

        datos_actualizacion = {
            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_61,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_61,
            data=cuerpo,
            headers={
                **headers,
                "Content-Type":
                    "application/json"
            },
            method="PUT"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            respuesta.read()

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        total_despues = len(df_final)

        st.success(
            "✅ Banco de preguntas actualizado "
            "correctamente en GitHub."
        )

        st.info(
            f"Preguntas existentes antes: "
            f"**{total_antes:,}**"
        )

        st.info(
            f"Preguntas incorporadas: "
            f"**{total_nuevas:,}**"
        )

        st.info(
            f"Preguntas totales después: "
            f"**{total_despues:,}**"
        )

        st.dataframe(
            df_nuevas,
            use_container_width=True,
            hide_index=True
        )

    except Exception as error:

        st.error(
            "No fue posible actualizar "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(error)


# ============================================================
# BOTÓN DE SINCRONIZACIÓN
# ============================================================

if preguntas_61:

    pendientes_61 = sum(
        1
        for p in preguntas_61
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_61 == 0:

        if st.button(
            "🔄 SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_61"
        ):

            sincronizar_banco_61()


# ============================================================
# 6.2 - PRODUCTO / COMPONENTE / ACCIÓN
# PARTE 1 - CARGA Y CONTROL DE FUENTES
# ============================================================

# ------------------------------------------------------------
# ARCHIVOS
# ------------------------------------------------------------

ARCHIVO_FUENTE_62 = "COMPONENTES_Y_ACCIONES.csv"
ARCHIVO_BANCO_62 = "BANCO_PREGUNTAS_GENERALES.csv"


# ------------------------------------------------------------
# FUNCIONES
# ------------------------------------------------------------

def normalizar_62(valor):

    if pd.isna(valor):
        return ""

    return (
        str(valor)
        .strip()
        .lower()
        .replace("\n", " ")
    )


def cargar_fuente_62():

    try:

        df = pd.read_csv(
            ARCHIVO_FUENTE_62,
            encoding="utf-8-sig"
        )

    except UnicodeDecodeError:

        df = pd.read_csv(
            ARCHIVO_FUENTE_62,
            encoding="latin-1"
        )

    columnas = [
        "Producto",
        "Componente",
        "Acciones"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "6.2 ERROR: faltan columnas en "
            f"{ARCHIVO_FUENTE_62}: "
            f"{', '.join(faltantes)}"
        )

        return None

    df = df[columnas].copy()

    for columna in columnas:

        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df = df[
        (df["Producto"] != "")
        &
        (df["Componente"] != "")
        &
        (df["Acciones"] != "")
    ].copy()

    df["Fuente_ID"] = [
        f"PTCA-F{i:06d}"
        for i in range(1, len(df) + 1)
    ]

    df["_clave"] = (
        df["Producto"]
        .map(normalizar_62)
        + "||"
        + df["Componente"]
        .map(normalizar_62)
        + "||"
        + df["Acciones"]
        .map(normalizar_62)
    )

    df = (
        df
        .drop_duplicates(
            subset="_clave"
        )
        .reset_index(drop=True)
    )

    return df


def cargar_banco_62():

    try:

        df = pd.read_csv(
            ARCHIVO_BANCO_62,
            encoding="utf-8-sig"
        )

    except FileNotFoundError:

        return pd.DataFrame()

    except UnicodeDecodeError:

        df = pd.read_csv(
            ARCHIVO_BANCO_62,
            encoding="latin-1"
        )

    return df


# ------------------------------------------------------------
# IDENTIFICAR RELACIONES YA UTILIZADAS
# ------------------------------------------------------------

def obtener_fuentes_usadas_62(df_banco):

    usadas = set()

    if df_banco.empty:
        return usadas

    if "Fuente_ID" not in df_banco.columns:
        return usadas

    for valor in df_banco["Fuente_ID"].fillna(""):

        texto = str(valor).strip()

        if not texto:
            continue

        for fuente in texto.split(";"):

            fuente = fuente.strip()

            if fuente:
                usadas.add(fuente)

    return usadas


# ------------------------------------------------------------
# IDENTIFICAR PREGUNTAS EXISTENTES
# ------------------------------------------------------------

def obtener_preguntas_existentes_62(df_banco):

    preguntas = set()

    if df_banco.empty:
        return preguntas

    if "Pregunta" not in df_banco.columns:
        return preguntas

    for pregunta in df_banco["Pregunta"].fillna(""):

        clave = normalizar_62(pregunta)

        if clave:
            preguntas.add(clave)

    return preguntas


# ============================================================
# INTERFAZ 6.2
# ============================================================

st.markdown(
    "## 6.2 Producto - Componente - Acción"
)

st.write(
    "Control de relaciones disponibles para "
    "la generación de preguntas."
)


if st.button(
    "🔎 CARGAR Y VALIDAR FUENTES 6.2",
    key="cargar_fuentes_62"
):

    df_fuente_62 = cargar_fuente_62()

    if df_fuente_62 is None:
        st.stop()

    df_banco_62 = cargar_banco_62()

    fuentes_usadas_62 = (
        obtener_fuentes_usadas_62(
            df_banco_62
        )
    )

    preguntas_existentes_62 = (
        obtener_preguntas_existentes_62(
            df_banco_62
        )
    )

    # --------------------------------------------------------
    # RELACIONES DISPONIBLES
    # --------------------------------------------------------

    df_disponible_62 = (
        df_fuente_62[
            ~df_fuente_62["Fuente_ID"].isin(
                fuentes_usadas_62
            )
        ]
        .copy()
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # GUARDAR EN SESIÓN
    # --------------------------------------------------------

    st.session_state[
        "df_fuente_62"
    ] = df_fuente_62.copy()

    st.session_state[
        "df_banco_62"
    ] = df_banco_62.copy()

    st.session_state[
        "df_disponible_62"
    ] = df_disponible_62.copy()

    st.session_state[
        "fuentes_usadas_62"
    ] = fuentes_usadas_62

    st.session_state[
        "preguntas_existentes_62"
    ] = preguntas_existentes_62


# ============================================================
# MOSTRAR CONTROL
# ============================================================

if (
    "df_fuente_62"
    in st.session_state
):

    df_fuente_62 = (
        st.session_state[
            "df_fuente_62"
        ]
    )

    df_banco_62 = (
        st.session_state[
            "df_banco_62"
        ]
    )

    df_disponible_62 = (
        st.session_state[
            "df_disponible_62"
        ]
    )

    fuentes_usadas_62 = (
        st.session_state[
            "fuentes_usadas_62"
        ]
    )

    preguntas_existentes_62 = (
        st.session_state[
            "preguntas_existentes_62"
        ]
    )


    # --------------------------------------------------------
    # RESUMEN
    # --------------------------------------------------------

    st.success(
        "6.2 cargó correctamente las fuentes."
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Relaciones fuente",
            len(df_fuente_62)
        )

    with col2:

        st.metric(
            "Relaciones ya utilizadas",
            len(fuentes_usadas_62)
        )

    with col3:

        st.metric(
            "Relaciones disponibles",
            len(df_disponible_62)
        )


    st.info(
        "Preguntas existentes en el banco: "
        f"{len(preguntas_existentes_62):,}"
    )


    # --------------------------------------------------------
    # CONTROL DE CONSISTENCIA
    # --------------------------------------------------------

    total_fuente = len(
        df_fuente_62
    )

    total_usadas = len(
        fuentes_usadas_62
        &
        set(
            df_fuente_62[
                "Fuente_ID"
            ]
        )
    )

    total_disponibles = len(
        df_disponible_62
    )


    if (
        total_usadas
        +
        total_disponibles
        ==
        total_fuente
    ):

        st.success(
            "Control correcto: "
            "utilizadas + disponibles = "
            "relaciones fuente."
        )

    else:

        st.error(
            "6.2 ERROR: los conteos "
            "de relaciones no coinciden."
        )


    # --------------------------------------------------------
    # VISTA PREVIA
    # --------------------------------------------------------

    st.markdown(
        "### Relaciones disponibles"
    )

    st.dataframe(
        df_disponible_62[
            [
                "Fuente_ID",
                "Producto",
                "Componente",
                "Acciones"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
# ============================================================
# 6.2 - PARTE 2
# GENERADOR PRODUCTO - COMPONENTE - ACCIÓN
# ============================================================

# ============================================================
# 1. FUNCIONES DE CONTROL
# ============================================================

def clave_relacion_62(producto, componente, accion):

    return (
        normalizar_62(producto)
        + "||"
        + normalizar_62(componente)
        + "||"
        + normalizar_62(accion)
    )


def obtener_relaciones_consumidas_62():

    consumidas = set()

    if "df_banco_62" in st.session_state:

        df_banco = st.session_state[
            "df_banco_62"
        ]

        if (
            not df_banco.empty
            and "Fuente_ID" in df_banco.columns
        ):

            for valor in df_banco[
                "Fuente_ID"
            ].fillna(""):

                for fuente in str(
                    valor
                ).split(";"):

                    fuente = fuente.strip()

                    if fuente:
                        consumidas.add(fuente)

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_62",
            set()
        )
    )

    return consumidas


def obtener_claves_preguntas_62():

    claves = set()

    if "df_banco_62" not in st.session_state:
        return claves

    df_banco = st.session_state[
        "df_banco_62"
    ]

    if df_banco.empty:
        return claves

    if "Fuente_ID" in df_banco.columns:

        for valor in df_banco[
            "Fuente_ID"
        ].fillna(""):

            fuentes = sorted(
                [
                    x.strip()
                    for x in str(
                        valor
                    ).split(";")
                    if x.strip()
                ]
            )

            if fuentes:

                claves.add(
                    "FUENTES::"
                    + "||".join(fuentes)
                )

    return claves


def siguiente_id_62():

    mayor = 0

    df_banco = st.session_state.get(
        "df_banco_62",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco[
            "Pregunta_ID"
        ].fillna(""):

            coincidencia = re.match(
                r"PTCA-(\d+)",
                str(valor).strip()
            )

            if coincidencia:

                mayor = max(
                    mayor,
                    int(
                        coincidencia.group(1)
                    )
                )

    preguntas_actuales = st.session_state.get(
        "preguntas_generadas_62",
        []
    )

    for pregunta in preguntas_actuales:

        coincidencia = re.match(
            r"PTCA-(\d+)",
            str(
                pregunta.get(
                    "Pregunta_ID",
                    ""
                )
            )
        )

        if coincidencia:

            mayor = max(
                mayor,
                int(
                    coincidencia.group(1)
                )
            )

    return f"PTCA-{mayor + 1:06d}"


# ============================================================
# 2. GENERAR NIVEL 1
# ============================================================

def generar_nivel_1_62(
    df_disponible,
    consumidas
):

    if len(df_disponible) < 4:
        return None

    candidatos = df_disponible[
        ~df_disponible[
            "Fuente_ID"
        ].isin(consumidas)
    ].copy()

    if len(candidatos) < 4:
        return None

    verdaderas = candidatos.sample(
        frac=1,
        random_state=None
    )

    for _, verdadera in verdaderas.iterrows():

        producto = verdadera[
            "Producto"
        ]

        componente = verdadera[
            "Componente"
        ]

        clave_producto = normalizar_62(
            producto
        )

        clave_componente = normalizar_62(
            componente
        )

        # ----------------------------------------------------
        # Las falsas deben corresponder a otros componentes.
        # ----------------------------------------------------

        falsas = candidatos[
            (
                candidatos["Producto"].map(
                    normalizar_62
                ) != clave_producto
            )
            |
            (
                candidatos["Componente"].map(
                    normalizar_62
                ) != clave_componente
            )
        ].copy()

        if len(falsas) < 3:
            continue

        falsas = falsas.sample(
            n=3,
            random_state=None
        )

        opciones = pd.concat(
            [
                pd.DataFrame([verdadera]),
                falsas
            ],
            ignore_index=True
        )

        opciones = opciones.sample(
            frac=1,
            random_state=None
        ).reset_index(drop=True)

        fuentes = list(
            opciones["Fuente_ID"]
        )

        if len(set(fuentes)) != 4:
            continue

        return {
            "Producto": producto,
            "Componente": componente,
            "Opciones": opciones,
            "Correctas": [
                int(
                    opciones.index[
                        opciones["Fuente_ID"]
                        == verdadera[
                            "Fuente_ID"
                        ]
                    ][0]
                ) + 1
            ]
        }

    return None


# ============================================================
# 3. GENERAR NIVEL 2
# ============================================================

def generar_nivel_2_62(
    df_disponible,
    consumidas
):

    candidatos = df_disponible[
        ~df_disponible[
            "Fuente_ID"
        ].isin(consumidas)
    ].copy()

    if len(candidatos) < 4:
        return None

    # --------------------------------------------------------
    # Buscar producto + componente con mínimo 2 acciones.
    # --------------------------------------------------------

    grupos = (
        candidatos
        .groupby(
            [
                "Producto",
                "Componente"
            ],
            sort=False
        )
        .filter(
            lambda grupo: len(grupo) >= 2
        )
    )

    if grupos.empty:
        return None

    combinaciones = (
        grupos[
            [
                "Producto",
                "Componente"
            ]
        ]
        .drop_duplicates()
        .values.tolist()
    )

    np.random.shuffle(
        combinaciones
    )

    for producto, componente in combinaciones:

        verdaderas = grupos[
            (
                grupos["Producto"] == producto
            )
            &
            (
                grupos["Componente"]
                == componente
            )
        ].copy()

        if len(verdaderas) < 2:
            continue

        verdaderas = verdaderas.sample(
            n=2,
            random_state=None
        )

        claves_verdaderas = set(
            verdaderas["Fuente_ID"]
        )

        # ----------------------------------------------------
        # Las falsas deben ser de otros componentes.
        # ----------------------------------------------------

        falsas = candidatos[
            ~candidatos[
                "Fuente_ID"
            ].isin(claves_verdaderas)
        ].copy()

        falsas = falsas[
            (
                falsas["Producto"].map(
                    normalizar_62
                )
                !=
                normalizar_62(producto)
            )
            |
            (
                falsas["Componente"].map(
                    normalizar_62
                )
                !=
                normalizar_62(componente)
            )
        ]

        if len(falsas) < 2:
            continue

        falsas = falsas.sample(
            n=2,
            random_state=None
        )

        opciones = pd.concat(
            [
                verdaderas.assign(
                    _correcta=True
                ),
                falsas.assign(
                    _correcta=False
                )
            ],
            ignore_index=True
        )

        opciones = opciones.sample(
            frac=1,
            random_state=None
        ).reset_index(drop=True)

        fuentes = list(
            opciones["Fuente_ID"]
        )

        if len(set(fuentes)) != 4:
            continue

        correctas = [
            i + 1
            for i, valor in enumerate(
                opciones["_correcta"]
            )
            if valor
        ]

        return {
            "Producto": producto,
            "Componente": componente,
            "Opciones": opciones,
            "Correctas": correctas
        }

    return None


# ============================================================
# 4. CONSTRUIR PREGUNTA
# ============================================================

def construir_pregunta_62(
    resultado,
    nivel
):

    opciones = resultado[
        "Opciones"
    ]

    producto = resultado[
        "Producto"
    ]

    componente = resultado[
        "Componente"
    ]

    if nivel == "Nivel 1":

        texto = (
            f"En el producto {producto}, "
            f"¿cuál de las siguientes acciones "
            f"corresponde específicamente al "
            f"componente {componente}?"
        )

    else:

        texto = (
            f"En el producto {producto}, "
            f"¿cuáles de las siguientes acciones "
            f"corresponden específicamente al "
            f"componente {componente}? "
            "Seleccione las dos opciones correctas."
        )

    pregunta_id = siguiente_id_62()

    fuentes = ";".join(
        opciones[
            "Fuente_ID"
        ].tolist()
    )

    return {

        "Pregunta_ID":
            pregunta_id,

        "Modulo":
            "Producto",

        "Tema":
            "Componente - Acción",

        "Nivel":
            nivel,

        "Tipo_Relacion":
            "Producto-Componente-Acción",

        "Pregunta":
            texto,

        "Respuesta_1":
            opciones.iloc[0][
                "Acciones"
            ],

        "Respuesta_2":
            opciones.iloc[1][
                "Acciones"
            ],

        "Respuesta_3":
            opciones.iloc[2][
                "Acciones"
            ],

        "Respuesta_4":
            opciones.iloc[3][
                "Acciones"
            ],

        "Respuesta_Correcta":
            ";".join(
                str(x)
                for x in resultado[
                    "Correctas"
                ]
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
            fuentes
    }


# ============================================================
# 5. GENERADOR GENERAL
# ============================================================

def generar_preguntas_62(
    cantidad,
    nivel
):

    df_disponible = st.session_state.get(
        "df_disponible_62",
        pd.DataFrame()
    )

    if df_disponible.empty:
        return []

    consumidas = (
        obtener_relaciones_consumidas_62()
    )

    preguntas = []

    while len(preguntas) < cantidad:

        if nivel == "Nivel 1":

            resultado = generar_nivel_1_62(
                df_disponible,
                consumidas
            )

        else:

            resultado = generar_nivel_2_62(
                df_disponible,
                consumidas
            )

        if resultado is None:
            break

        pregunta = construir_pregunta_62(
            resultado,
            nivel
        )

        preguntas.append(
            pregunta
        )

        # ----------------------------------------------------
        # Consumir las 4 relaciones utilizadas.
        # ----------------------------------------------------

        fuentes = (
            resultado[
                "Opciones"
            ][
                "Fuente_ID"
            ].tolist()
        )

        consumidas.update(
            fuentes
        )

    return preguntas


# ============================================================
# 6. INTERFAZ DE GENERACIÓN
# ============================================================

if (
    "df_disponible_62"
    in st.session_state
):

    st.markdown(
        "### Generación de preguntas"
    )

    cantidad_62 = st.number_input(
        "¿Cuántas preguntas desea generar?",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_62"
    )

    nivel_62 = st.selectbox(
        "Nivel de evaluación",
        [
            "Nivel 1",
            "Nivel 2"
        ],
        key="nivel_generar_62"
    )

    if st.button(
        "GENERAR PREGUNTAS",
        key="generar_preguntas_62"
    ):

        nuevas = generar_preguntas_62(
            cantidad_62,
            nivel_62
        )

        if not nuevas:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar nuevas "
                "preguntas con las condiciones "
                "establecidas."
            )

        else:

            st.session_state[
                "preguntas_generadas_62"
            ] = nuevas

            consumidas = (
                st.session_state.get(
                    "fuentes_consumidas_62",
                    set()
                )
            )

            for pregunta in nuevas:

                for fuente in str(
                    pregunta[
                        "Fuente_ID"
                    ]
                ).split(";"):

                    fuente = fuente.strip()

                    if fuente:
                        consumidas.add(
                            fuente
                        )

            st.session_state[
                "fuentes_consumidas_62"
            ] = consumidas

            st.success(
                f"Se generaron "
                f"{len(nuevas)} preguntas."
            )

            st.info(
                "Las relaciones utilizadas "
                "quedan consumidas y no se "
                "volverán a utilizar."
            )


# ============================================================
# 7. MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_62 = st.session_state.get(
    "preguntas_generadas_62",
    []
)

if preguntas_62:

    st.markdown(
        "### Preguntas generadas"
    )

    for pregunta in preguntas_62:

        st.markdown(
            f"**{pregunta['Pregunta_ID']} — "
            f"{pregunta['Nivel']}**"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"1. {pregunta['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta['Respuesta_4']}"
        )

        st.caption(
            "Respuesta correcta: "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta['Fuente_ID']}"
        )

        st.divider()
# ============================================================
# 6.2 - PARTE 3A
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS
# ============================================================

if preguntas_62:

    st.markdown(
        "## Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(preguntas_62):

        st.markdown(
            f"### {pregunta['Pregunta_ID']}"
        )

        st.write(
            f"**Nivel:** {pregunta['Nivel']}"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"**1.** {pregunta['Respuesta_1']}"
        )

        st.write(
            f"**2.** {pregunta['Respuesta_2']}"
        )

        st.write(
            f"**3.** {pregunta['Respuesta_3']}"
        )

        st.write(
            f"**4.** {pregunta['Respuesta_4']}"
        )

        st.write(
            "**Respuesta correcta:** "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            f"Fuente utilizada: {pregunta['Fuente_ID']}"
        )

        estado_actual = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado_actual}"
        )

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_62_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "✅ APROBAR",
                key=f"aprobar_62_{i}"
            ):

                preguntas_62[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_62[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_62"
                ] = preguntas_62

                st.success(
                    f"{pregunta['Pregunta_ID']} "
                    "fue aprobada."
                )

                st.rerun()

        with col2:

            if st.button(
                "❌ RECHAZAR",
                key=f"rechazar_62_{i}"
            ):

                preguntas_62[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_62[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_62"
                ] = preguntas_62

                st.warning(
                    f"{pregunta['Pregunta_ID']} "
                    "fue rechazada."
                )

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN
# ============================================================

if preguntas_62:

    aprobadas_62 = sum(
        1
        for p in preguntas_62
        if p.get("Estado") == "APROBADA"
    )

    rechazadas_62 = sum(
        1
        for p in preguntas_62
        if p.get("Estado") == "RECHAZADA"
    )

    pendientes_62 = sum(
        1
        for p in preguntas_62
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_62
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_62
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_62
        )

    if pendientes_62 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "La sincronización con el banco "
            "se habilitará a continuación."
        )

# ============================================================
# 6.2 - PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_62 = "franquiciasauces"
GITHUB_REPOSITORIO_62 = "Asesores"
GITHUB_RAMA_62 = "main"

GITHUB_ARCHIVO_62 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_62 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_62}/"
    f"{GITHUB_REPOSITORIO_62}/contents/"
    f"{GITHUB_ARCHIVO_62}"
)


def sincronizar_banco_62():

    import io

    preguntas = st.session_state.get(
        "preguntas_generadas_62",
        []
    )

    if not preguntas:

        st.warning(
            "No hay preguntas para sincronizar."
        )

        return

    # --------------------------------------------------------
    # TODAS LAS PREGUNTAS DEBEN ESTAR REVISADAS
    # --------------------------------------------------------

    if any(
        p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
        for p in preguntas
    ):

        st.error(
            "Todavía hay preguntas pendientes "
            "de revisión."
        )

        return

    headers = {
        "Authorization":
            f"Bearer {GITHUB_TOKEN}",

        "Accept":
            "application/vnd.github+json"
    }

    try:

        # ----------------------------------------------------
        # LEER ARCHIVO EXISTENTE DESDE GITHUB
        # ----------------------------------------------------

        solicitud = urllib.request.Request(
            URL_GITHUB_62,
            headers=headers,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            datos = json.loads(
                respuesta.read().decode(
                    "utf-8"
                )
            )

        sha = datos["sha"]

        contenido = base64.b64decode(
            datos["content"].replace(
                "\n",
                ""
            )
        )

        # ----------------------------------------------------
        # LEER EXCEL EN MEMORIA
        # ----------------------------------------------------

        df_banco = pd.read_excel(
            io.BytesIO(contenido),
            engine="openpyxl"
        )

        total_antes = len(
            df_banco
        )

        # ----------------------------------------------------
        # CONVERTIR PREGUNTAS GENERADAS
        # ----------------------------------------------------

        df_nuevas = pd.DataFrame(
            preguntas
        )

        columnas = [
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

        faltantes = [
            columna
            for columna in columnas
            if columna not in df_nuevas.columns
        ]

        if faltantes:

            st.error(
                "6.2 ERROR: faltan columnas "
                "en las preguntas generadas: "
                + ", ".join(faltantes)
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

        # ----------------------------------------------------
        # NO DUPLICAR PREGUNTAS YA EXISTENTES
        # ----------------------------------------------------

        if "Pregunta_ID" in df_banco.columns:

            existentes = set(
                df_banco[
                    "Pregunta_ID"
                ]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df_nuevas = df_nuevas[
                ~df_nuevas[
                    "Pregunta_ID"
                ]
                .fillna("")
                .astype(str)
                .str.strip()
                .isin(existentes)
            ].copy()

        total_nuevas = len(
            df_nuevas
        )

        # ----------------------------------------------------
        # NO HAY PREGUNTAS NUEVAS
        # ----------------------------------------------------

        if total_nuevas == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                f"Preguntas existentes: "
                f"**{total_antes:,}**"
            )

            return

        # ----------------------------------------------------
        # AGREGAR AL BANCO
        # ----------------------------------------------------

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

        # ----------------------------------------------------
        # CREAR NUEVO EXCEL EN MEMORIA
        # ----------------------------------------------------

        memoria = io.BytesIO()

        with pd.ExcelWriter(
            memoria,
            engine="openpyxl"
        ) as writer:

            df_final.to_excel(
                writer,
                index=False,
                sheet_name="Banco"
            )

        contenido_nuevo = base64.b64encode(
            memoria.getvalue()
        ).decode("utf-8")

        # ----------------------------------------------------
        # ACTUALIZAR ARCHIVO EN GITHUB
        # ----------------------------------------------------

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_62,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_62,
            data=cuerpo,
            headers={
                **headers,
                "Content-Type":
                    "application/json"
            },
            method="PUT"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            respuesta.read()

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        total_despues = len(
            df_final
        )

        st.success(
            "✅ BANCO_PREGUNTAS_GENERALES.xlsx "
            "fue actualizado correctamente en GitHub."
        )

        st.info(
            f"Preguntas existentes antes: "
            f"**{total_antes:,}**"
        )

        st.info(
            f"Preguntas incorporadas: "
            f"**{total_nuevas:,}**"
        )

        st.info(
            f"Preguntas totales después: "
            f"**{total_despues:,}**"
        )

        st.dataframe(
            df_nuevas,
            use_container_width=True,
            hide_index=True
        )

    except Exception as error:

        st.error(
            "No fue posible actualizar "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(error)


# ============================================================
# BOTÓN DE SINCRONIZACIÓN
# ============================================================

if preguntas_62:

    pendientes_62 = sum(
        1
        for p in preguntas_62
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_62 == 0:

        if st.button(
            "🔄 SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_62"
        ):

            sincronizar_banco_62()
# ============================================================
# 6.3 - PRODUCTO / CATEGORÍA PRINCIPAL
# PARTE 1 - CARGA Y CONTROL DE FUENTE
# ============================================================

# ------------------------------------------------------------
# ARCHIVOS
# ------------------------------------------------------------

ARCHIVO_FUENTE_63 = (
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)

ARCHIVO_BANCO_63 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)


# ------------------------------------------------------------
# FUNCIONES
# ------------------------------------------------------------

def normalizar_63(valor):

    if pd.isna(valor):
        return ""

    return (
        str(valor)
        .strip()
        .lower()
        .replace("\n", " ")
    )


def cargar_fuente_63():

    try:

        df = pd.read_excel(
            ARCHIVO_FUENTE_63,
            engine="openpyxl"
        )

    except Exception as error:

        st.error(
            "6.3 ERROR al cargar "
            f"{ARCHIVO_FUENTE_63}: "
            f"{error}"
        )

        return None

    columnas = [
        "Producto",
        "Categoría principal"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "6.3 ERROR: faltan columnas en "
            f"{ARCHIVO_FUENTE_63}: "
            f"{', '.join(faltantes)}"
        )

        return None

    # --------------------------------------------------------
    # CONSERVAR SOLAMENTE LAS COLUMNAS NECESARIAS
    # --------------------------------------------------------

    df = df[
        columnas
    ].copy()

    # --------------------------------------------------------
    # LIMPIEZA
    # --------------------------------------------------------

    df["Producto"] = (
        df["Producto"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Categoría principal"] = (
        df["Categoría principal"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # ELIMINAR REGISTROS INCOMPLETOS
    # --------------------------------------------------------

    df = df[
        (df["Producto"] != "")
        &
        (df["Categoría principal"] != "")
    ].copy()

    # --------------------------------------------------------
    # IDENTIFICADOR DE FUENTE
    # --------------------------------------------------------

    df["Fuente_ID"] = [
        f"PTCP-F{i:06d}"
        for i in range(
            1,
            len(df) + 1
        )
    ]

    # --------------------------------------------------------
    # CLAVE ÚNICA
    # --------------------------------------------------------

    df["_clave"] = (
        df["Producto"]
        .map(normalizar_63)
        + "||"
        + df["Categoría principal"]
        .map(normalizar_63)
    )

    # --------------------------------------------------------
    # ELIMINAR DUPLICADOS
    # --------------------------------------------------------

    df = (
        df
        .drop_duplicates(
            subset="_clave"
        )
        .reset_index(
            drop=True
        )
    )

    return df


def cargar_banco_63():

    try:

        df = pd.read_excel(
            ARCHIVO_BANCO_63,
            engine="openpyxl"
        )

    except FileNotFoundError:

        return pd.DataFrame()

    except Exception:

        return pd.DataFrame()

    return df


# ============================================================
# IDENTIFICAR RELACIONES YA UTILIZADAS
# ============================================================

def obtener_fuentes_usadas_63(
    df_banco
):

    usadas = set()

    if df_banco.empty:

        return usadas

    if "Fuente_ID" not in df_banco.columns:

        return usadas

    for valor in df_banco[
        "Fuente_ID"
    ].fillna(""):

        for fuente in str(
            valor
        ).split(";"):

            fuente = fuente.strip()

            if fuente:

                usadas.add(
                    fuente
                )

    return usadas


# ============================================================
# IDENTIFICAR PREGUNTAS EXISTENTES
# ============================================================

def obtener_preguntas_existentes_63(
    df_banco
):

    preguntas = set()

    if df_banco.empty:

        return preguntas

    if "Pregunta" not in df_banco.columns:

        return preguntas

    for pregunta in df_banco[
        "Pregunta"
    ].fillna(""):

        clave = normalizar_63(
            pregunta
        )

        if clave:

            preguntas.add(
                clave
            )

    return preguntas


# ============================================================
# INTERFAZ 6.3
# ============================================================

st.markdown(
    "## 6.3 Producto - Categoría principal"
)

st.write(
    "Control de relaciones disponibles para "
    "la generación de preguntas Nivel 1."
)


if st.button(
    "🔎 CARGAR Y VALIDAR FUENTE 6.3",
    key="cargar_fuentes_63"
):

    df_fuente_63 = (
        cargar_fuente_63()
    )

    if df_fuente_63 is None:

        st.stop()

    df_banco_63 = (
        cargar_banco_63()
    )

    fuentes_usadas_63 = (
        obtener_fuentes_usadas_63(
            df_banco_63
        )
    )

    preguntas_existentes_63 = (
        obtener_preguntas_existentes_63(
            df_banco_63
        )
    )

    # --------------------------------------------------------
    # RELACIONES DISPONIBLES
    # --------------------------------------------------------

    df_disponible_63 = (
        df_fuente_63[
            ~df_fuente_63[
                "Fuente_ID"
            ].isin(
                fuentes_usadas_63
            )
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )

    # --------------------------------------------------------
    # GUARDAR EN SESIÓN
    # --------------------------------------------------------

    st.session_state[
        "df_fuente_63"
    ] = df_fuente_63.copy()

    st.session_state[
        "df_banco_63"
    ] = df_banco_63.copy()

    st.session_state[
        "df_disponible_63"
    ] = df_disponible_63.copy()

    st.session_state[
        "fuentes_usadas_63"
    ] = fuentes_usadas_63

    st.session_state[
        "preguntas_existentes_63"
    ] = preguntas_existentes_63


# ============================================================
# MOSTRAR CONTROL
# ============================================================

if (
    "df_fuente_63"
    in st.session_state
):

    df_fuente_63 = (
        st.session_state[
            "df_fuente_63"
        ]
    )

    df_banco_63 = (
        st.session_state[
            "df_banco_63"
        ]
    )

    df_disponible_63 = (
        st.session_state[
            "df_disponible_63"
        ]
    )

    fuentes_usadas_63 = (
        st.session_state[
            "fuentes_usadas_63"
        ]
    )

    preguntas_existentes_63 = (
        st.session_state[
            "preguntas_existentes_63"
        ]
    )

    # --------------------------------------------------------
    # RESUMEN
    # --------------------------------------------------------

    st.success(
        "6.3 cargó correctamente la fuente."
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Relaciones fuente",
            len(df_fuente_63)
        )

    with col2:

        st.metric(
            "Relaciones ya utilizadas",
            len(fuentes_usadas_63)
        )

    with col3:

        st.metric(
            "Relaciones disponibles",
            len(df_disponible_63)
        )

    st.info(
        "Preguntas existentes en el banco: "
        f"{len(preguntas_existentes_63):,}"
    )

    # --------------------------------------------------------
    # CONTROL DE CONSISTENCIA
    # --------------------------------------------------------

    total_fuente = len(
        df_fuente_63
    )

    total_usadas = len(
        fuentes_usadas_63
        &
        set(
            df_fuente_63[
                "Fuente_ID"
            ]
        )
    )

    total_disponibles = len(
        df_disponible_63
    )

    if (
        total_usadas
        +
        total_disponibles
        ==
        total_fuente
    ):

        st.success(
            "Control correcto: "
            "utilizadas + disponibles = "
            "relaciones fuente."
        )

    else:

        st.error(
            "6.3 ERROR: los conteos "
            "de relaciones no coinciden."
        )

    # --------------------------------------------------------
    # VISTA PREVIA
    # --------------------------------------------------------

    st.markdown(
        "### Relaciones disponibles"
    )

    st.dataframe(
        df_disponible_63[
            [
                "Fuente_ID",
                "Producto",
                "Categoría principal"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
# ============================================================
# 6.3 - PARTE 2
# GENERADOR PRODUCTO - CATEGORÍA PRINCIPAL
# SOLO NIVEL 1
# ============================================================

def obtener_relaciones_consumidas_63():

    consumidas = set()

    df_banco = st.session_state.get(
        "df_banco_63",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Fuente_ID" in df_banco.columns
    ):

        for valor in df_banco["Fuente_ID"].fillna(""):

            for fuente in str(valor).split(";"):

                fuente = fuente.strip()

                if fuente:
                    consumidas.add(fuente)

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_63",
            set()
        )
    )

    return consumidas


def siguiente_id_63():

    mayor = 0

    df_banco = st.session_state.get(
        "df_banco_63",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco["Pregunta_ID"].fillna(""):

            coincidencia = re.match(
                r"PTCP-(\d+)",
                str(valor).strip()
            )

            if coincidencia:

                mayor = max(
                    mayor,
                    int(coincidencia.group(1))
                )

    for pregunta in st.session_state.get(
        "preguntas_generadas_63",
        []
    ):

        coincidencia = re.match(
            r"PTCP-(\d+)",
            str(
                pregunta.get(
                    "Pregunta_ID",
                    ""
                )
            )
        )

        if coincidencia:

            mayor = max(
                mayor,
                int(coincidencia.group(1))
            )

    return f"PTCP-{mayor + 1:06d}"


def generar_nivel_1_63(
    df_disponible,
    consumidas
):

    candidatos = df_disponible[
        ~df_disponible["Fuente_ID"].isin(
            consumidas
        )
    ].copy()

    if len(candidatos) < 4:

        return None

    candidatos = candidatos.sample(
        frac=1
    ).reset_index(drop=True)

    for _, verdadera in candidatos.iterrows():

        producto = verdadera["Producto"]

        categoria_correcta = (
            verdadera[
                "Categoría principal"
            ]
        )

        clave_correcta = normalizar_63(
            categoria_correcta
        )

        # ----------------------------------------------------
        # Buscar 3 categorías diferentes de la correcta
        # ----------------------------------------------------

        falsas = candidatos[
            candidatos[
                "Categoría principal"
            ].map(
                normalizar_63
            ) != clave_correcta
        ].copy()

        if falsas.empty:

            continue

        falsas = falsas.sample(
            frac=1
        ).reset_index(drop=True)

        seleccionadas = []

        categorias = {
            clave_correcta
        }

        for _, falsa in falsas.iterrows():

            categoria = falsa[
                "Categoría principal"
            ]

            clave = normalizar_63(
                categoria
            )

            if not clave:
                continue

            if clave in categorias:
                continue

            seleccionadas.append(
                falsa
            )

            categorias.add(
                clave
            )

            if len(seleccionadas) == 3:
                break

        if len(seleccionadas) != 3:

            continue

        opciones = pd.concat(
            [
                pd.DataFrame([verdadera]),
                pd.DataFrame(seleccionadas)
            ],
            ignore_index=True
        )

        opciones = opciones.sample(
            frac=1
        ).reset_index(drop=True)

        # ----------------------------------------------------
        # Verificar 4 categorías diferentes
        # ----------------------------------------------------

        categorias_finales = [
            normalizar_63(x)
            for x in opciones[
                "Categoría principal"
            ]
        ]

        if len(
            set(categorias_finales)
        ) != 4:

            continue

        # ----------------------------------------------------
        # Verificar 4 fuentes diferentes
        # ----------------------------------------------------

        fuentes = list(
            opciones["Fuente_ID"]
        )

        if len(set(fuentes)) != 4:

            continue

        correcta = (
            opciones.index[
                opciones["Fuente_ID"]
                ==
                verdadera["Fuente_ID"]
            ][0]
            + 1
        )

        return {
            "Producto": producto,
            "Opciones": opciones,
            "Correcta": correcta
        }

    return None


def construir_pregunta_63(
    resultado
):

    opciones = resultado["Opciones"]

    pregunta_id = siguiente_id_63()

    fuentes = ";".join(
        opciones["Fuente_ID"].tolist()
    )

    texto = (
        "¿A qué categoría principal "
        "pertenece el producto "
        f"{resultado['Producto']}?"
    )

    return {

        "Pregunta_ID":
            pregunta_id,

        "Modulo":
            "Producto",

        "Tema":
            "Categoría principal",

        "Nivel":
            "Nivel 1",

        "Tipo_Relacion":
            "Producto-Categoría principal",

        "Pregunta":
            texto,

        "Respuesta_1":
            opciones.iloc[0][
                "Categoría principal"
            ],

        "Respuesta_2":
            opciones.iloc[1][
                "Categoría principal"
            ],

        "Respuesta_3":
            opciones.iloc[2][
                "Categoría principal"
            ],

        "Respuesta_4":
            opciones.iloc[3][
                "Categoría principal"
            ],

        "Respuesta_Correcta":
            str(
                resultado["Correcta"]
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
            fuentes
    }


def generar_preguntas_63(
    cantidad
):

    df_disponible = st.session_state.get(
        "df_disponible_63",
        pd.DataFrame()
    )

    if df_disponible.empty:

        return []

    consumidas = (
        obtener_relaciones_consumidas_63()
    )

    preguntas = []

    while len(preguntas) < cantidad:

        resultado = generar_nivel_1_63(
            df_disponible,
            consumidas
        )

        if resultado is None:

            break

        pregunta = construir_pregunta_63(
            resultado
        )

        preguntas.append(
            pregunta
        )

        fuentes = (
            resultado["Opciones"]
            ["Fuente_ID"]
            .tolist()
        )

        consumidas.update(
            fuentes
        )

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR
# ============================================================

if "df_disponible_63" in st.session_state:

    st.markdown(
        "### Generador Producto - Categoría principal"
    )

    cantidad_63 = st.number_input(
        "¿Cuántas preguntas desea generar?",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_63"
    )

    st.info(
        "Nivel 1: una categoría correcta "
        "y tres categorías falsas."
    )

    if st.button(
        "GENERAR PREGUNTAS 6.3",
        key="generar_preguntas_63"
    ):

        nuevas_63 = generar_preguntas_63(
            cantidad_63
        )

        if not nuevas_63:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar preguntas "
                "con cuatro categorías diferentes."
            )

        else:

            st.session_state[
                "preguntas_generadas_63"
            ] = nuevas_63

            consumidas_63 = (
                st.session_state.get(
                    "fuentes_consumidas_63",
                    set()
                )
            )

            for pregunta in nuevas_63:

                for fuente in str(
                    pregunta["Fuente_ID"]
                ).split(";"):

                    fuente = fuente.strip()

                    if fuente:
                        consumidas_63.add(
                            fuente
                        )

            st.session_state[
                "fuentes_consumidas_63"
            ] = consumidas_63

            st.success(
                f"Se generaron "
                f"{len(nuevas_63)} preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_63 = st.session_state.get(
    "preguntas_generadas_63",
    []
)

if preguntas_63:

    st.markdown(
        "### Preguntas generadas"
    )

    for pregunta in preguntas_63:

        st.markdown(
            f"**{pregunta['Pregunta_ID']} — "
            f"{pregunta['Nivel']}**"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"1. {pregunta['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta['Respuesta_4']}"
        )

        st.caption(
            "Respuesta correcta: "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta['Fuente_ID']}"
        )

        st.divider()
# ============================================================
# 6.3 - PARTE 3
# VALIDACIÓN DE PREGUNTAS
# ============================================================

preguntas_63 = st.session_state.get(
    "preguntas_generadas_63",
    []
)


# ============================================================
# VALIDACIÓN INDIVIDUAL
# ============================================================

if preguntas_63:

    st.markdown(
        "## Validación de preguntas 6.3"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(
        preguntas_63
    ):

        st.markdown(
            f"### {pregunta['Pregunta_ID']}"
        )

        st.write(
            "**Nivel:** "
            f"{pregunta['Nivel']}"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"**1.** {pregunta['Respuesta_1']}"
        )

        st.write(
            f"**2.** {pregunta['Respuesta_2']}"
        )

        st.write(
            f"**3.** {pregunta['Respuesta_3']}"
        )

        st.write(
            f"**4.** {pregunta['Respuesta_4']}"
        )

        st.write(
            "**Respuesta correcta:** "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente utilizada: "
            f"{pregunta['Fuente_ID']}"
        )

        # ----------------------------------------------------
        # ESTADO ACTUAL
        # ----------------------------------------------------

        estado_actual = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            "**Estado actual:** "
            f"{estado_actual}"
        )

        # ----------------------------------------------------
        # OBSERVACIÓN
        # ----------------------------------------------------

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_63_{i}"
        )

        # ----------------------------------------------------
        # BOTONES
        # ----------------------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "✅ APROBAR",
                key=f"aprobar_63_{i}"
            ):

                preguntas_63[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_63[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_63"
                ] = preguntas_63

                st.success(
                    f"{pregunta['Pregunta_ID']} "
                    "fue aprobada."
                )

                st.rerun()

        with col2:

            if st.button(
                "❌ RECHAZAR",
                key=f"rechazar_63_{i}"
            ):

                preguntas_63[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_63[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_63"
                ] = preguntas_63

                st.warning(
                    f"{pregunta['Pregunta_ID']} "
                    "fue rechazada."
                )

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN
# ============================================================

if preguntas_63:

    aprobadas_63 = sum(
        1
        for p in preguntas_63
        if p.get("Estado")
        == "APROBADA"
    )

    rechazadas_63 = sum(
        1
        for p in preguntas_63
        if p.get("Estado")
        == "RECHAZADA"
    )

    pendientes_63 = sum(
        1
        for p in preguntas_63
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_63
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_63
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_63
        )

    if pendientes_63 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas."
        )

        st.info(
            "La sincronización con el Banco de "
            "Preguntas se habilitará en la Parte 4."
        )
# ============================================================
# 6.3 - PARTE 4
# SINCRONIZAR CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_63 = "franquiciasauces"
GITHUB_REPOSITORIO_63 = "Asesores"
GITHUB_RAMA_63 = "main"

GITHUB_ARCHIVO_63 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_63 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_63}/"
    f"{GITHUB_REPOSITORIO_63}/contents/"
    f"{GITHUB_ARCHIVO_63}"
)


def sincronizar_banco_63():

    preguntas = st.session_state.get(
        "preguntas_generadas_63",
        []
    )

    if not preguntas:

        st.warning(
            "No hay preguntas para sincronizar."
        )

        return

    # --------------------------------------------------------
    # TODAS DEBEN ESTAR REVISADAS
    # --------------------------------------------------------

    if any(
        p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
        for p in preguntas
    ):

        st.error(
            "Todavía hay preguntas pendientes "
            "de revisión."
        )

        return

    headers = {
        "Authorization":
            f"Bearer {GITHUB_TOKEN}",

        "Accept":
            "application/vnd.github+json"
    }

    try:

        # ----------------------------------------------------
        # LEER BANCO EXISTENTE DESDE GITHUB
        # ----------------------------------------------------

        solicitud = urllib.request.Request(
            URL_GITHUB_63,
            headers=headers,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            datos = json.loads(
                respuesta.read().decode(
                    "utf-8"
                )
            )

        sha = datos["sha"]

        contenido = base64.b64decode(
            datos["content"].replace(
                "\n",
                ""
            )
        )

        # ----------------------------------------------------
        # LEER EXCEL DIRECTAMENTE DESDE LOS BYTES
        # ----------------------------------------------------

        df_banco = pd.read_excel(
            contenido
        )

        total_antes = len(
            df_banco
        )

        # ----------------------------------------------------
        # CONVERTIR PREGUNTAS VALIDADAS
        # ----------------------------------------------------

        df_nuevas = pd.DataFrame(
            preguntas
        )

        columnas = [
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

        df_nuevas = df_nuevas[
            columnas
        ].copy()

        # ----------------------------------------------------
        # EVITAR DUPLICADOS POR Pregunta_ID
        # ----------------------------------------------------

        if "Pregunta_ID" in df_banco.columns:

            existentes = set(
                df_banco[
                    "Pregunta_ID"
                ]
                .astype(str)
                .str.strip()
            )

            df_nuevas = df_nuevas[
                ~df_nuevas[
                    "Pregunta_ID"
                ]
                .astype(str)
                .str.strip()
                .isin(existentes)
            ].copy()

        total_nuevas = len(
            df_nuevas
        )

        # ----------------------------------------------------
        # NO HAY NADA NUEVO
        # ----------------------------------------------------

        if total_nuevas == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                "Preguntas existentes: "
                f"**{total_antes:,}**"
            )

            return

        # ----------------------------------------------------
        # AGREGAR PREGUNTAS AL BANCO
        # ----------------------------------------------------

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

        # ----------------------------------------------------
        # PREPARAR EXCEL PARA GITHUB
        # ----------------------------------------------------

        import io

        memoria = io.BytesIO()

        with pd.ExcelWriter(
            memoria,
            engine="openpyxl"
        ) as writer:

            df_final.to_excel(
                writer,
                index=False,
                sheet_name="Banco"
            )

        contenido_nuevo = base64.b64encode(
            memoria.getvalue()
        ).decode(
            "utf-8"
        )

        # ----------------------------------------------------
        # ACTUALIZAR GITHUB
        # ----------------------------------------------------

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_63,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode(
            "utf-8"
        )

        solicitud = urllib.request.Request(
            URL_GITHUB_63,
            data=cuerpo,
            headers={
                **headers,
                "Content-Type":
                    "application/json"
            },
            method="PUT"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            respuesta.read()

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        total_despues = len(
            df_final
        )

        st.success(
            "✅ Banco de preguntas actualizado "
            "correctamente en GitHub."
        )

        st.info(
            "Preguntas existentes antes: "
            f"**{total_antes:,}**"
        )

        st.info(
            "Preguntas incorporadas: "
            f"**{total_nuevas:,}**"
        )

        st.info(
            "Preguntas totales después: "
            f"**{total_despues:,}**"
        )

        st.dataframe(
            df_nuevas,
            use_container_width=True,
            hide_index=True
        )

    except Exception as error:

        st.error(
            "No fue posible actualizar "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(
            error
        )


# ============================================================
# BOTÓN DE SINCRONIZACIÓN
# ============================================================

if preguntas_63:

    pendientes_63 = sum(
        1
        for p in preguntas_63
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_63 == 0:

        if st.button(
            "🔄 SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_63"
        ):

            sincronizar_banco_63()


# ============================================================
# 6.4 - PRODUCTO / CATEGORÍA PRINCIPAL + COMPLEMENTARIA
# PARTE 1 - CARGA Y CONTROL DE FUENTE
# ============================================================

# ------------------------------------------------------------
# ARCHIVOS
# ------------------------------------------------------------

ARCHIVO_FUENTE_64 = (
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)

ARCHIVO_BANCO_64 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)


# ------------------------------------------------------------
# FUNCIONES
# ------------------------------------------------------------

def normalizar_64(valor):

    if pd.isna(valor):
        return ""

    return (
        str(valor)
        .strip()
        .lower()
        .replace("\n", " ")
    )


def cargar_fuente_64():

    try:

        df = pd.read_excel(
            ARCHIVO_FUENTE_64,
            engine="openpyxl"
        )

    except Exception as error:

        st.error(
            "6.4 ERROR al cargar "
            f"{ARCHIVO_FUENTE_64}: "
            f"{error}"
        )

        return None

    columnas = [
        "Producto",
        "Categoría principal",
        "Categorías complementarias"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "6.4 ERROR: faltan columnas en "
            f"{ARCHIVO_FUENTE_64}: "
            f"{', '.join(faltantes)}"
        )

        return None

    # --------------------------------------------------------
    # CONSERVAR SOLAMENTE LAS COLUMNAS NECESARIAS
    # --------------------------------------------------------

    df = df[
        columnas
    ].copy()

    # --------------------------------------------------------
    # LIMPIEZA
    # --------------------------------------------------------

    df["Producto"] = (
        df["Producto"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Categoría principal"] = (
        df["Categoría principal"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Categorías complementarias"] = (
        df["Categorías complementarias"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # ELIMINAR REGISTROS INCOMPLETOS
    # --------------------------------------------------------

    df = df[
        (df["Producto"] != "")
        &
        (df["Categoría principal"] != "")
        &
        (df["Categorías complementarias"] != "")
    ].copy()

    # --------------------------------------------------------
    # IDENTIFICADOR DE FUENTE
    # --------------------------------------------------------

    df["Fuente_ID"] = [
        f"PTCC-F{i:06d}"
        for i in range(
            1,
            len(df) + 1
        )
    ]

    # --------------------------------------------------------
    # CLAVE ÚNICA
    # --------------------------------------------------------

    df["_clave"] = (
        df["Producto"]
        .map(normalizar_64)
        + "||"
        + df["Categoría principal"]
        .map(normalizar_64)
        + "||"
        + df["Categorías complementarias"]
        .map(normalizar_64)
    )

    # --------------------------------------------------------
    # ELIMINAR DUPLICADOS
    # --------------------------------------------------------

    df = (
        df
        .drop_duplicates(
            subset="_clave"
        )
        .reset_index(
            drop=True
        )
    )

    return df


def cargar_banco_64():

    try:

        df = pd.read_excel(
            ARCHIVO_BANCO_64,
            engine="openpyxl"
        )

    except FileNotFoundError:

        return pd.DataFrame()

    except Exception:

        return pd.DataFrame()

    return df


# ============================================================
# IDENTIFICAR RELACIONES YA UTILIZADAS
# ============================================================

def obtener_fuentes_usadas_64(
    df_banco
):

    usadas = set()

    if df_banco.empty:

        return usadas

    if "Fuente_ID" not in df_banco.columns:

        return usadas

    for valor in df_banco[
        "Fuente_ID"
    ].fillna(""):

        for fuente in str(
            valor
        ).split(";"):

            fuente = fuente.strip()

            if fuente:

                usadas.add(
                    fuente
                )

    return usadas


# ============================================================
# IDENTIFICAR PREGUNTAS EXISTENTES
# ============================================================

def obtener_preguntas_existentes_64(
    df_banco
):

    preguntas = set()

    if df_banco.empty:

        return preguntas

    if "Pregunta" not in df_banco.columns:

        return preguntas

    for pregunta in df_banco[
        "Pregunta"
    ].fillna(""):

        clave = normalizar_64(
            pregunta
        )

        if clave:

            preguntas.add(
                clave
            )

    return preguntas


# ============================================================
# INTERFAZ 6.4
# ============================================================

st.markdown(
    "## 6.4 Producto - Categoría principal + complementaria"
)

st.write(
    "Control de relaciones disponibles para "
    "la generación de preguntas Nivel 2."
)


if st.button(
    "🔎 CARGAR Y VALIDAR FUENTE 6.4",
    key="cargar_fuentes_64"
):

    df_fuente_64 = (
        cargar_fuente_64()
    )

    if df_fuente_64 is None:

        st.stop()

    df_banco_64 = (
        cargar_banco_64()
    )

    fuentes_usadas_64 = (
        obtener_fuentes_usadas_64(
            df_banco_64
        )
    )

    preguntas_existentes_64 = (
        obtener_preguntas_existentes_64(
            df_banco_64
        )
    )

    # --------------------------------------------------------
    # RELACIONES DISPONIBLES
    # --------------------------------------------------------

    df_disponible_64 = (
        df_fuente_64[
            ~df_fuente_64[
                "Fuente_ID"
            ].isin(
                fuentes_usadas_64
            )
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )

    # --------------------------------------------------------
    # GUARDAR EN SESIÓN
    # --------------------------------------------------------

    st.session_state[
        "df_fuente_64"
    ] = df_fuente_64.copy()

    st.session_state[
        "df_banco_64"
    ] = df_banco_64.copy()

    st.session_state[
        "df_disponible_64"
    ] = df_disponible_64.copy()

    st.session_state[
        "fuentes_usadas_64"
    ] = fuentes_usadas_64

    st.session_state[
        "preguntas_existentes_64"
    ] = preguntas_existentes_64


# ============================================================
# MOSTRAR CONTROL
# ============================================================

if (
    "df_fuente_64"
    in st.session_state
):

    df_fuente_64 = (
        st.session_state[
            "df_fuente_64"
        ]
    )

    df_banco_64 = (
        st.session_state[
            "df_banco_64"
        ]
    )

    df_disponible_64 = (
        st.session_state[
            "df_disponible_64"
        ]
    )

    fuentes_usadas_64 = (
        st.session_state[
            "fuentes_usadas_64"
        ]
    )

    preguntas_existentes_64 = (
        st.session_state[
            "preguntas_existentes_64"
        ]
    )

    # --------------------------------------------------------
    # RESUMEN
    # --------------------------------------------------------

    st.success(
        "6.4 cargó correctamente la fuente."
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Relaciones fuente",
            len(df_fuente_64)
        )

    with col2:

        st.metric(
            "Relaciones ya utilizadas",
            len(fuentes_usadas_64)
        )

    with col3:

        st.metric(
            "Relaciones disponibles",
            len(df_disponible_64)
        )

    st.info(
        "Preguntas existentes en el banco: "
        f"{len(preguntas_existentes_64):,}"
    )

    # --------------------------------------------------------
    # CONTROL DE CONSISTENCIA
    # --------------------------------------------------------

    total_fuente = len(
        df_fuente_64
    )

    total_usadas = len(
        fuentes_usadas_64
        &
        set(
            df_fuente_64[
                "Fuente_ID"
            ]
        )
    )

    total_disponibles = len(
        df_disponible_64
    )

    if (
        total_usadas
        +
        total_disponibles
        ==
        total_fuente
    ):

        st.success(
            "Control correcto: "
            "utilizadas + disponibles = "
            "relaciones fuente."
        )

    else:

        st.error(
            "6.4 ERROR: los conteos "
            "de relaciones no coinciden."
        )

    # --------------------------------------------------------
    # VISTA PREVIA
    # --------------------------------------------------------

    st.markdown(
        "### Relaciones disponibles"
    )

    st.dataframe(
        df_disponible_64[
            [
                "Fuente_ID",
                "Producto",
                "Categoría principal",
                "Categorías complementarias"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

# ============================================================
# 6.4 - PARTE 2
# GENERADOR PRODUCTO - CATEGORÍA PRINCIPAL + COMPLEMENTARIA
# NIVEL 2
# ============================================================

def obtener_relaciones_consumidas_64():

    consumidas = set()

    df_banco = st.session_state.get(
        "df_banco_64",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Fuente_ID" in df_banco.columns
    ):

        for valor in df_banco["Fuente_ID"].fillna(""):

            for fuente in str(valor).split(";"):

                fuente = fuente.strip()

                if fuente:
                    consumidas.add(fuente)

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_64",
            set()
        )
    )

    return consumidas


def siguiente_id_64():

    mayor = 0

    df_banco = st.session_state.get(
        "df_banco_64",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco["Pregunta_ID"].fillna(""):

            coincidencia = re.match(
                r"PTCC-(\d+)",
                str(valor).strip()
            )

            if coincidencia:
                mayor = max(
                    mayor,
                    int(coincidencia.group(1))
                )

    return f"PTCC-{mayor + 1:06d}"


def generar_nivel_2_64(
    df_disponible,
    consumidas
):

    candidatos = df_disponible[
        ~df_disponible["Fuente_ID"].isin(
            consumidas
        )
    ].copy()

    if len(candidatos) < 2:
        return None

    for _, verdadera in candidatos.iterrows():

        principal = str(
            verdadera["Categoría principal"]
        ).strip()

        complementaria = str(
            verdadera["Categorías complementarias"]
        ).strip()

        if not principal or not complementaria:
            continue

        falsas = candidatos[
            candidatos["Fuente_ID"]
            != verdadera["Fuente_ID"]
        ].copy()

        for _, falsa in falsas.iterrows():

            falsa_principal = str(
                falsa["Categoría principal"]
            ).strip()

            falsa_complementaria = str(
                falsa["Categorías complementarias"]
            ).strip()

            if not falsa_principal or not falsa_complementaria:
                continue

            if falsa_principal.lower() == principal.lower():
                continue

            if (
                falsa_complementaria.lower()
                == complementaria.lower()
            ):
                continue

            opciones = [
                principal,
                complementaria,
                falsa_principal,
                falsa_complementaria
            ]

            if len({
                x.lower()
                for x in opciones
            }) != 4:
                continue

            return {
                "Producto":
                    verdadera["Producto"],

                "Correctas": [
                    principal,
                    complementaria
                ],

                "Falsas": [
                    falsa_principal,
                    falsa_complementaria
                ],

                "Fuente_ID":
                    verdadera["Fuente_ID"]
                    + ";"
                    + falsa["Fuente_ID"]
            }

    return None


def construir_pregunta_64(resultado):

    opciones = (
        resultado["Correctas"]
        + resultado["Falsas"]
    )

    np.random.shuffle(opciones)

    correctas = []

    for i, opcion in enumerate(opciones):

        if opcion in resultado["Correctas"]:
            correctas.append(i + 1)

    return {

        "Pregunta_ID":
            siguiente_id_64(),

        "Modulo":
            "Producto",

        "Tema":
            "Categoría principal y complementaria",

        "Nivel":
            "Nivel 2",

        "Tipo_Relacion":
            "Producto-Categoría principal-Categoría complementaria",

        "Pregunta":
            "¿Cuáles de las siguientes categorías "
            "corresponden al producto "
            f"{resultado['Producto']}? "
            "Seleccione las dos opciones correctas.",

        "Respuesta_1":
            opciones[0],

        "Respuesta_2":
            opciones[1],

        "Respuesta_3":
            opciones[2],

        "Respuesta_4":
            opciones[3],

        "Respuesta_Correcta":
            ";".join(
                str(x)
                for x in sorted(correctas)
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
            resultado["Fuente_ID"]
    }


def generar_preguntas_64(cantidad):

    df_disponible = st.session_state.get(
        "df_disponible_64",
        pd.DataFrame()
    )

    if df_disponible.empty:
        return []

    consumidas = (
        obtener_relaciones_consumidas_64()
    )

    preguntas = []

    while len(preguntas) < cantidad:

        resultado = generar_nivel_2_64(
            df_disponible,
            consumidas
        )

        if resultado is None:
            break

        pregunta = construir_pregunta_64(
            resultado
        )

        preguntas.append(pregunta)

        consumidas.update(
            resultado["Fuente_ID"].split(";")
        )

    return preguntas


# ============================================================
# INTERFAZ
# ============================================================

if "df_disponible_64" in st.session_state:

    st.markdown(
        "### Generador Producto - Categorías"
    )

    cantidad_64 = st.number_input(
        "¿Cuántas preguntas desea generar?",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_64"
    )

    st.info(
        "Nivel 2: una categoría principal correcta, "
        "una categoría complementaria correcta "
        "y dos categorías falsas."
    )

    if st.button(
        "GENERAR PREGUNTAS 6.4",
        key="generar_preguntas_64"
    ):

        nuevas_64 = generar_preguntas_64(
            cantidad_64
        )

        if not nuevas_64:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar preguntas."
            )

        else:

            st.session_state[
                "preguntas_generadas_64"
            ] = nuevas_64

            consumidas_64 = st.session_state.get(
                "fuentes_consumidas_64",
                set()
            )

            for pregunta in nuevas_64:

                for fuente in str(
                    pregunta["Fuente_ID"]
                ).split(";"):

                    fuente = fuente.strip()

                    if fuente:
                        consumidas_64.add(fuente)

            st.session_state[
                "fuentes_consumidas_64"
            ] = consumidas_64

            st.success(
                f"Se generaron "
                f"{len(nuevas_64)} preguntas."
            )


preguntas_64 = st.session_state.get(
    "preguntas_generadas_64",
    []
)

if preguntas_64:

    st.markdown(
        "### Preguntas generadas"
    )

    for pregunta in preguntas_64:

        st.markdown(
            f"**{pregunta['Pregunta_ID']} — "
            f"{pregunta['Nivel']}**"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"1. {pregunta['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta['Respuesta_4']}"
        )

        st.caption(
            "Respuestas correctas: "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta['Fuente_ID']}"
        )

        st.divider()

# ============================================================
# 6.4 - PARTE 3
# VALIDACIÓN
# ============================================================

preguntas_64 = st.session_state.get(
    "preguntas_generadas_64",
    []
)

if preguntas_64:

    st.markdown("## Validación 6.4")

    for i, p in enumerate(preguntas_64):

        st.markdown(
            f"### {p['Pregunta_ID']}"
        )

        st.write(p["Pregunta"])

        st.write(f"1. {p['Respuesta_1']}")
        st.write(f"2. {p['Respuesta_2']}")
        st.write(f"3. {p['Respuesta_3']}")
        st.write(f"4. {p['Respuesta_4']}")

        st.caption(
            f"Correctas: {p['Respuesta_Correcta']}"
        )

        obs = st.text_input(
            "Observación",
            value=p.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"obs64_{i}"
        )

        c1, c2 = st.columns(2)

        with c1:
            if st.button(
                "APROBAR",
                key=f"ok64_{i}"
            ):

                p["Estado"] = "APROBADA"
                p["Observacion_Administrador"] = obs

                st.session_state[
                    "preguntas_generadas_64"
                ] = preguntas_64

                st.rerun()

        with c2:
            if st.button(
                "RECHAZAR",
                key=f"no64_{i}"
            ):

                p["Estado"] = "RECHAZADA"
                p["Observacion_Administrador"] = obs

                st.session_state[
                    "preguntas_generadas_64"
                ] = preguntas_64

                st.rerun()

        st.divider()

    aprobadas = sum(
        p.get("Estado") == "APROBADA"
        for p in preguntas_64
    )

    rechazadas = sum(
        p.get("Estado") == "RECHAZADA"
        for p in preguntas_64
    )

    pendientes = len(preguntas_64) - aprobadas - rechazadas

    st.write(
        f"**Aprobadas:** {aprobadas} | "
        f"**Rechazadas:** {rechazadas} | "
        f"**Pendientes:** {pendientes}"
    )
# ============================================================
# 6.4 PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_64 = "franquiciasauces"
GITHUB_REPOSITORIO_64 = "Asesores"
GITHUB_RAMA_64 = "main"

GITHUB_ARCHIVO_64 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_64 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_64}/"
    f"{GITHUB_REPOSITORIO_64}/contents/"
    f"{GITHUB_ARCHIVO_64}"
)


def sincronizar_banco_64():

    preguntas = st.session_state.get(
        "preguntas_generadas_64",
        []
    )

    if not preguntas:

        st.warning(
            "No hay preguntas para sincronizar."
        )

        return

    if any(
        p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
        for p in preguntas
    ):

        st.error(
            "Todavía hay preguntas pendientes de revisión."
        )

        return

    headers = {
        "Authorization":
            f"Bearer {GITHUB_TOKEN}",

        "Accept":
            "application/vnd.github+json"
    }

    try:

        # ----------------------------------------------------
        # LEER BANCO EXISTENTE
        # ----------------------------------------------------

        solicitud = urllib.request.Request(
            URL_GITHUB_64,
            headers=headers,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            datos = json.loads(
                respuesta.read().decode("utf-8")
            )

        sha = datos["sha"]

        contenido = base64.b64decode(
            datos["content"].replace("\n", "")
        )

        df_banco = pd.read_excel(
            contenido
        )

        total_antes = len(
            df_banco
        )

        # ----------------------------------------------------
        # PREGUNTAS VALIDADAS
        # ----------------------------------------------------

        df_nuevas = pd.DataFrame(
            preguntas
        )

        columnas = [
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

        faltantes = [
            columna
            for columna in columnas
            if columna not in df_nuevas.columns
        ]

        if faltantes:

            st.error(
                "6.4 ERROR: faltan columnas "
                "en las preguntas generadas: "
                f"{', '.join(faltantes)}"
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

        # ----------------------------------------------------
        # EVITAR DUPLICADOS
        # ----------------------------------------------------

        if "Pregunta_ID" in df_banco.columns:

            existentes = set(
                df_banco[
                    "Pregunta_ID"
                ]
                .astype(str)
                .str.strip()
            )

            df_nuevas = df_nuevas[
                ~df_nuevas[
                    "Pregunta_ID"
                ]
                .astype(str)
                .str.strip()
                .isin(
                    existentes
                )
            ]

        total_nuevas = len(
            df_nuevas
        )

        if total_nuevas == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                f"Preguntas existentes: "
                f"**{total_antes:,}**"
            )

            return

        # ----------------------------------------------------
        # AGREGAR AL BANCO
        # ----------------------------------------------------

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

        # ----------------------------------------------------
        # CREAR EXCEL EN MEMORIA
        # ----------------------------------------------------

        import io

        memoria = io.BytesIO()

        with pd.ExcelWriter(
            memoria,
            engine="openpyxl"
        ) as writer:

            df_final.to_excel(
                writer,
                index=False,
                sheet_name="Banco"
            )

        contenido_nuevo = base64.b64encode(
            memoria.getvalue()
        ).decode("utf-8")

        # ----------------------------------------------------
        # ACTUALIZAR GITHUB
        # ----------------------------------------------------

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_64,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_64,
            data=cuerpo,
            headers={
                **headers,
                "Content-Type":
                    "application/json"
            },
            method="PUT"
        )

        with urllib.request.urlopen(
            solicitud,
            timeout=30
        ) as respuesta:

            respuesta.read()

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        total_despues = len(
            df_final
        )

        st.success(
            "Banco de preguntas actualizado "
            "correctamente en GitHub."
        )

        st.info(
            f"Preguntas existentes antes: "
            f"**{total_antes:,}**"
        )

        st.info(
            f"Preguntas incorporadas: "
            f"**{total_nuevas:,}**"
        )

        st.info(
            f"Preguntas totales después: "
            f"**{total_despues:,}**"
        )

        st.dataframe(
            df_nuevas,
            use_container_width=True,
            hide_index=True
        )

    except Exception as error:

        st.error(
            "No fue posible actualizar "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(error)


# ============================================================
# BOTÓN DE SINCRONIZACIÓN
# ============================================================

if preguntas_64:

    pendientes_64 = sum(
        1
        for p in preguntas_64
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_64 == 0:

        if st.button(
            "SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_64"
        ):

            sincronizar_banco_64()
# ============================================================
# 7.1 - PATOLOGÍA
# PARTE 1 - CARGA Y CONTROL DE FUENTE
# ============================================================

ARCHIVO_FUENTE_71 = (
    "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
)

HOJA_FUENTE_71 = "Patologias"

ARCHIVO_BANCO_71 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)


# ============================================================
# FUNCIONES
# ============================================================

def normalizar_71(valor):

    if pd.isna(valor):
        return ""

    return (
        str(valor)
        .strip()
        .lower()
        .replace("\n", " ")
    )


def cargar_fuente_71():

    try:

        df = pd.read_excel(
            ARCHIVO_FUENTE_71,
            sheet_name=HOJA_FUENTE_71,
            engine="openpyxl"
        )

    except Exception as error:

        st.error(
            "7.1 ERROR al cargar la hoja "
            f"{HOJA_FUENTE_71}: {error}"
        )

        return None

    columnas = [
        "Patologia_ID",
        "Patología",
        "Descripción breve (para cliente)",
        "Causas frecuentes (resumen)",
        "Síntomas/Señales clave (checklist)",
        "Objetivo del paquete",
        "Notas (para asesor)"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "7.1 ERROR: faltan columnas en "
            f"la hoja {HOJA_FUENTE_71}: "
            f"{', '.join(faltantes)}"
        )

        return None

    df = df[columnas].copy()

    for columna in columnas:

        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df = df[
        (df["Patologia_ID"] != "")
        &
        (df["Patología"] != "")
    ].copy()

    df["Fuente_ID"] = [
        f"PTG-F{i:06d}"
        for i in range(1, len(df) + 1)
    ]

    df["_clave"] = (
        df["Patologia_ID"]
        .map(normalizar_71)
    )

    df = (
        df
        .drop_duplicates(
            subset="_clave"
        )
        .reset_index(drop=True)
    )

    return df


def cargar_banco_71():

    try:

        df = pd.read_excel(
            ARCHIVO_BANCO_71,
            engine="openpyxl"
        )

    except FileNotFoundError:

        return pd.DataFrame()

    except Exception:

        return pd.DataFrame()

    return df


# ============================================================
# RELACIONES YA UTILIZADAS
# ============================================================

def obtener_fuentes_usadas_71(
    df_banco
):

    usadas = set()

    if df_banco.empty:

        return usadas

    if "Fuente_ID" not in df_banco.columns:

        return usadas

    for valor in df_banco[
        "Fuente_ID"
    ].fillna(""):

        for fuente in str(
            valor
        ).split(";"):

            fuente = fuente.strip()

            if fuente:

                usadas.add(
                    fuente
                )

    return usadas


# ============================================================
# PREGUNTAS EXISTENTES
# ============================================================

def obtener_preguntas_existentes_71(
    df_banco
):

    preguntas = set()

    if df_banco.empty:

        return preguntas

    if "Pregunta" not in df_banco.columns:

        return preguntas

    for pregunta in df_banco[
        "Pregunta"
    ].fillna(""):

        clave = normalizar_71(
            pregunta
        )

        if clave:

            preguntas.add(
                clave
            )

    return preguntas


# ============================================================
# INTERFAZ 7.1
# ============================================================

st.markdown(
    "## 7.1 Patología"
)

st.write(
    "Carga y control de la hoja "
    "Patologias para los generadores "
    "de definición, causas y síntomas."
)


if st.button(
    "CARGAR Y VALIDAR FUENTE 7.1",
    key="cargar_fuentes_71"
):

    df_fuente_71 = (
        cargar_fuente_71()
    )

    if df_fuente_71 is None:

        st.stop()

    df_banco_71 = (
        cargar_banco_71()
    )

    fuentes_usadas_71 = (
        obtener_fuentes_usadas_71(
            df_banco_71
        )
    )

    preguntas_existentes_71 = (
        obtener_preguntas_existentes_71(
            df_banco_71
        )
    )

    df_disponible_71 = (
        df_fuente_71[
            ~df_fuente_71[
                "Fuente_ID"
            ].isin(
                fuentes_usadas_71
            )
        ]
        .copy()
        .reset_index(drop=True)
    )

    st.session_state[
        "df_fuente_71"
    ] = df_fuente_71.copy()

    st.session_state[
        "df_banco_71"
    ] = df_banco_71.copy()

    st.session_state[
        "df_disponible_71"
    ] = df_disponible_71.copy()

    st.session_state[
        "fuentes_usadas_71"
    ] = fuentes_usadas_71

    st.session_state[
        "preguntas_existentes_71"
    ] = preguntas_existentes_71


# ============================================================
# MOSTRAR CONTROL
# ============================================================

if "df_fuente_71" in st.session_state:

    df_fuente_71 = st.session_state[
        "df_fuente_71"
    ]

    df_banco_71 = st.session_state[
        "df_banco_71"
    ]

    df_disponible_71 = st.session_state[
        "df_disponible_71"
    ]

    fuentes_usadas_71 = st.session_state[
        "fuentes_usadas_71"
    ]

    preguntas_existentes_71 = st.session_state[
        "preguntas_existentes_71"
    ]

    st.success(
        "7.1 cargó correctamente la hoja "
        "Patologias."
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Patologías fuente",
            len(df_fuente_71)
        )

    with col2:

        st.metric(
            "Fuentes utilizadas",
            len(
                fuentes_usadas_71
                &
                set(
                    df_fuente_71[
                        "Fuente_ID"
                    ]
                )
            )
        )

    with col3:

        st.metric(
            "Patologías disponibles",
            len(df_disponible_71)
        )

    st.info(
        "Preguntas existentes en el banco: "
        f"{len(preguntas_existentes_71):,}"
    )

    st.markdown(
        "### Estructura cargada"
    )

    st.dataframe(
        df_disponible_71[
            [
                "Fuente_ID",
                "Patologia_ID",
                "Patología",
                "Descripción breve (para cliente)",
                "Causas frecuentes (resumen)",
                "Síntomas/Señales clave (checklist)"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

# ============================================================
# 7.2 - PARTE 2
# GENERADOR PATOLOGÍA - DEFINICIÓN
# NIVEL 1 Y NIVEL 2
# ============================================================

def normalizar_72(valor):

    if pd.isna(valor):
        return ""

    return " ".join(
        str(valor).strip().lower().split()
    )


def siguiente_id_72():

    mayor = 0

    for pregunta in st.session_state.get(
        "preguntas_generadas_72",
        []
    ):

        texto = str(
            pregunta.get(
                "Pregunta_ID",
                ""
            )
        )

        if texto.startswith("PTG-DS-"):

            try:

                numero = int(
                    texto.replace(
                        "PTG-DS-",
                        ""
                    )
                )

                mayor = max(
                    mayor,
                    numero
                )

            except ValueError:
                pass

    return (
        f"PTG-DS-{mayor + 1:06d}"
    )


def obtener_consumidas_72():

    consumidas = set()

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_72",
            set()
        )
    )

    return consumidas


def generar_nivel_72(
    df,
    consumidas,
    nivel
):

    candidatos = df[
        ~df["Fuente_ID"].isin(
            consumidas
        )
    ].copy()

    if len(candidatos) < 4:

        return None

    candidatos = candidatos.sample(
        frac=1
    ).reset_index(
        drop=True
    )

    for _, verdadera in candidatos.iterrows():

        fuente = str(
            verdadera["Fuente_ID"]
        ).strip()

        patologia = str(
            verdadera["Patología"]
        ).strip()

        definicion = str(
            verdadera[
                "Descripción breve (para cliente)"
            ]
        ).strip()

        if (
            not fuente
            or not patologia
            or not definicion
        ):

            continue

        clave_correcta = normalizar_72(
            definicion
        )

        falsas = candidatos[
            candidatos["Fuente_ID"] != fuente
        ].copy()

        falsas["Definicion"] = (
            falsas[
                "Descripción breve (para cliente)"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        falsas = falsas[
            falsas["Definicion"] != ""
        ]

        falsas = falsas[
            falsas["Definicion"].map(
                normalizar_72
            ) != clave_correcta
        ]

        falsas = falsas.drop_duplicates(
            subset=["Definicion"]
        )

        if len(falsas) < 3:

            continue

        falsas = falsas.sample(
            frac=1
        ).reset_index(
            drop=True
        )

        seleccionadas = []

        for _, falsa in falsas.iterrows():

            texto = falsa[
                "Definicion"
            ]

            clave = normalizar_72(
                texto
            )

            if clave in [
                normalizar_72(x)
                for x in seleccionadas
            ]:

                continue

            seleccionadas.append(
                texto
            )

            if len(seleccionadas) == 3:

                break

        if len(seleccionadas) != 3:

            continue

        opciones = [
            definicion,
            seleccionadas[0],
            seleccionadas[1],
            seleccionadas[2]
        ]

        opciones = pd.Series(
            opciones
        ).sample(
            frac=1
        ).tolist()

        correcta = (
            opciones.index(
                definicion
            ) + 1
        )

        return {
            "Patología": patologia,
            "Opciones": opciones,
            "Correcta": correcta,
            "Fuente_ID": fuente,
            "Nivel": nivel
        }

    return None


def construir_pregunta_72(
    resultado
):

    opciones = resultado[
        "Opciones"
    ]

    return {

        "Pregunta_ID":
            siguiente_id_72(),

        "Modulo":
            "Patología",

        "Tema":
            "Definición",

        "Nivel":
            resultado["Nivel"],

        "Tipo_Relacion":
            "Patología-Definición",

        "Pregunta":
            (
                "¿Cuál de las siguientes "
                "opciones describe correctamente "
                "la patología "
                f"{resultado['Patología']}?"
            ),

        "Respuesta_1":
            opciones[0],

        "Respuesta_2":
            opciones[1],

        "Respuesta_3":
            opciones[2],

        "Respuesta_4":
            opciones[3],

        "Respuesta_Correcta":
            str(
                resultado["Correcta"]
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
            resultado["Fuente_ID"]
    }


def generar_preguntas_72(
    cantidad,
    modo
):

    df = st.session_state.get(
        "df_disponible_71",
        pd.DataFrame()
    )

    if df.empty:

        return []

    consumidas = (
        obtener_consumidas_72()
    )

    preguntas = []

    if modo == "Nivel 1":

        niveles = [
            "Nivel 1"
        ]

    elif modo == "Nivel 2":

        niveles = [
            "Nivel 2"
        ]

    else:

        niveles = [
            "Nivel 1",
            "Nivel 2"
        ]

    while len(preguntas) < cantidad:

        generado = False

        for nivel in niveles:

            if len(preguntas) >= cantidad:

                break

            resultado = generar_nivel_72(
                df,
                consumidas,
                nivel
            )

            if resultado is None:

                continue

            pregunta = construir_pregunta_72(
                resultado
            )

            preguntas.append(
                pregunta
            )

            consumidas.add(
                resultado["Fuente_ID"]
            )

            generado = True

        if not generado:

            break

    st.session_state[
        "fuentes_consumidas_72"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR
# ============================================================

if "df_disponible_71" in st.session_state:

    st.markdown(
        "### 7.2 Parte 2 - Generador "
        "Patología - Definición"
    )

    modo_72 = st.selectbox(
        "Seleccione el nivel",
        [
            "Nivel 1",
            "Nivel 2",
            "Niveles 1 y 2"
        ],
        key="modo_generacion_72"
    )

    cantidad_72 = st.number_input(
        "Cantidad máxima de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_72"
    )

    if st.button(
        "GENERAR PREGUNTAS 7.2",
        key="generar_preguntas_72"
    ):

        nuevas_72 = generar_preguntas_72(
            cantidad_72,
            modo_72
        )

        st.session_state[
            "preguntas_generadas_72"
        ] = nuevas_72

        if nuevas_72:

            st.success(
                f"Se generaron "
                f"{len(nuevas_72)} preguntas."
            )

        else:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_72 = st.session_state.get(
    "preguntas_generadas_72",
    []
)

if preguntas_72:

    st.markdown(
        "### Preguntas generadas 7.2"
    )

    for pregunta in preguntas_72:

        st.markdown(
            f"**{pregunta['Pregunta_ID']} — "
            f"{pregunta['Nivel']}**"
        )

        st.write(
            pregunta["Pregunta"]
        )

        st.write(
            f"1. {pregunta['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta['Respuesta_4']}"
        )

        st.caption(
            "Respuesta correcta: "
            f"{pregunta['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta['Fuente_ID']}"
        )

        st.divider()




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

# ============================================================
# 7.2 - PARTE 3
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS PATOLOGÍA - DEFINICIÓN
# ============================================================

preguntas_72 = st.session_state.get(
    "preguntas_generadas_72",
    []
)

if preguntas_72:

    st.markdown(
        "## 7.2 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(
        preguntas_72
    ):

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
            f"Fuente utilizada: "
            f"{pregunta['Fuente_ID']}"
        )

        estado = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado}"
        )

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_72_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_72_{i}"
            ):

                preguntas_72[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_72[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_72"
                ] = preguntas_72

                st.rerun()

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_72_{i}"
            ):

                preguntas_72[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_72[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_72"
                ] = preguntas_72

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 7.2
# ============================================================

if preguntas_72:

    aprobadas_72 = sum(
        1
        for p in preguntas_72
        if p.get("Estado") == "APROBADA"
    )

    rechazadas_72 = sum(
        1
        for p in preguntas_72
        if p.get("Estado") == "RECHAZADA"
    )

    pendientes_72 = sum(
        1
        for p in preguntas_72
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 7.2"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_72
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_72
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_72
        )

    if pendientes_72 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "Las preguntas aprobadas quedan "
            "listas para la sincronización."
        )
# ============================================================

# ============================================================
# 7.2 PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_72 = "franquiciasauces"
GITHUB_REPOSITORIO_72 = "Asesores"
GITHUB_RAMA_72 = "main"

GITHUB_ARCHIVO_72 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_72 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_72}/"
    f"{GITHUB_REPOSITORIO_72}/contents/"
    f"{GITHUB_ARCHIVO_72}"
)


def sincronizar_banco_72():

    preguntas = st.session_state.get(
        "preguntas_generadas_72",
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

        solicitud = urllib.request.Request(
            URL_GITHUB_72,
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
                "7.2 ERROR: faltan columnas "
                "en las preguntas generadas: "
                f"{', '.join(faltantes)}"
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

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

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

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

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_72,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_72,
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

if preguntas_72:

    pendientes_72 = sum(
        1
        for p in preguntas_72
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_72 == 0:

        if st.button(
            "SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_72"
        ):

            sincronizar_banco_72()
# ============================================================
# 7.3 - PARTE 2
# GENERADOR PATOLOGÍA - CAUSAS
# NIVEL 1 Y NIVEL 2
# ============================================================

def normalizar_73(valor):

    if pd.isna(valor):
        return ""

    return " ".join(
        str(valor).strip().lower().split()
    )


def siguiente_id_73():

    mayor = 0

    for pregunta in st.session_state.get(
        "preguntas_generadas_73",
        []
    ):

        texto = str(
            pregunta.get(
                "Pregunta_ID",
                ""
            )
        )

        if texto.startswith("PTG-PC-"):

            try:

                numero = int(
                    texto.replace(
                        "PTG-PC-",
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
        f"PTG-PC-{mayor + 1:06d}"
    )


def obtener_consumidas_73():

    consumidas = set()

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_73",
            set()
        )
    )

    return consumidas


def generar_nivel_73(
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

        causas = str(
            verdadera[
                "Causas frecuentes (resumen)"
            ]
        ).strip()

        if (
            not fuente
            or not patologia
            or not causas
        ):

            continue

        clave_correcta = normalizar_73(
            causas
        )

        falsas = candidatos[
            candidatos["Fuente_ID"] != fuente
        ].copy()

        falsas["Causa"] = (
            falsas[
                "Causas frecuentes (resumen)"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        falsas = falsas[
            falsas["Causa"] != ""
        ]

        falsas = falsas[
            falsas["Causa"].map(
                normalizar_73
            ) != clave_correcta
        ]

        falsas = falsas.drop_duplicates(
            subset=["Causa"]
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
                "Causa"
            ]

            clave = normalizar_73(
                texto
            )

            if clave in [
                normalizar_73(x)
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
            causas,
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
                causas
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


def construir_pregunta_73(
    resultado
):

    opciones = resultado[
        "Opciones"
    ]

    return {

        "Pregunta_ID":
            siguiente_id_73(),

        "Modulo":
            "Patología",

        "Tema":
            "Causas",

        "Nivel":
            resultado["Nivel"],

        "Tipo_Relacion":
            "Patología-Causas",

        "Pregunta":
            (
                "¿Cuál de las siguientes "
                "opciones corresponde a las "
                "causas frecuentes de la patología "
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


def generar_preguntas_73(
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
        obtener_consumidas_73()
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

            resultado = generar_nivel_73(
                df,
                consumidas,
                nivel
            )

            if resultado is None:

                continue

            pregunta = construir_pregunta_73(
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
        "fuentes_consumidas_73"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR 7.3
# ============================================================

if "df_disponible_71" in st.session_state:

    st.markdown(
        "### 7.3 Parte 2 - Generador "
        "Patología - Causas"
    )

    modo_73 = st.selectbox(
        "Seleccione el nivel",
        [
            "Nivel 1",
            "Nivel 2",
            "Niveles 1 y 2"
        ],
        key="modo_generacion_73"
    )

    cantidad_73 = st.number_input(
        "Cantidad máxima de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_73"
    )

    if st.button(
        "GENERAR PREGUNTAS 7.3",
        key="generar_preguntas_73"
    ):

        nuevas_73 = generar_preguntas_73(
            cantidad_73,
            modo_73
        )

        st.session_state[
            "preguntas_generadas_73"
        ] = nuevas_73

        if nuevas_73:

            st.success(
                f"Se generaron "
                f"{len(nuevas_73)} preguntas."
            )

        else:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_73 = st.session_state.get(
    "preguntas_generadas_73",
    []
)

if preguntas_73:

    st.markdown(
        "### Preguntas generadas 7.3"
    )

    for pregunta in preguntas_73:

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
# 7.3 - PARTE 3
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS
# PATOLOGÍA - CAUSAS
# ============================================================

preguntas_73 = st.session_state.get(
    "preguntas_generadas_73",
    []
)

if preguntas_73:

    st.markdown(
        "## 7.3 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(
        preguntas_73
    ):

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
            f"Fuente utilizada: "
            f"{pregunta['Fuente_ID']}"
        )

        estado = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado}"
        )

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_73_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_73_{i}"
            ):

                preguntas_73[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_73[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_73"
                ] = preguntas_73

                st.rerun()

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_73_{i}"
            ):

                preguntas_73[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_73[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_73"
                ] = preguntas_73

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 7.3
# ============================================================

if preguntas_73:

    aprobadas_73 = sum(
        1
        for p in preguntas_73
        if p.get("Estado") == "APROBADA"
    )

    rechazadas_73 = sum(
        1
        for p in preguntas_73
        if p.get("Estado") == "RECHAZADA"
    )

    pendientes_73 = sum(
        1
        for p in preguntas_73
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 7.3"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_73
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_73
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_73
        )

    if pendientes_73 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "Las preguntas aprobadas quedan "
            "listas para la sincronización."
        )


# ============================================================
# 7.3 - PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_73 = "franquiciasauces"
GITHUB_REPOSITORIO_73 = "Asesores"
GITHUB_RAMA_73 = "main"

GITHUB_ARCHIVO_73 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_73 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_73}/"
    f"{GITHUB_REPOSITORIO_73}/contents/"
    f"{GITHUB_ARCHIVO_73}"
)


def sincronizar_banco_73():

    preguntas = st.session_state.get(
        "preguntas_generadas_73",
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

        solicitud = urllib.request.Request(
            URL_GITHUB_73,
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
                "7.3 ERROR: faltan columnas "
                "en las preguntas generadas: "
                f"{', '.join(faltantes)}"
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

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

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

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

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_73,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_73,
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

if preguntas_73:

    pendientes_73 = sum(
        1
        for p in preguntas_73
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_73 == 0:

        if st.button(
            "SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_73"
        ):

            sincronizar_banco_73()
# ============================================================
# 7.4 - PARTE 2
# GENERADOR PATOLOGÍA - SÍNTOMAS
# NIVEL 1 Y NIVEL 2
# ============================================================

def normalizar_74(valor):

    if pd.isna(valor):
        return ""

    return " ".join(
        str(valor).strip().lower().split()
    )


def siguiente_id_74():

    mayor = 0

    for pregunta in st.session_state.get(
        "preguntas_generadas_74",
        []
    ):

        texto = str(
            pregunta.get(
                "Pregunta_ID",
                ""
            )
        )

        if texto.startswith("PTG-ST-"):

            try:

                numero = int(
                    texto.replace(
                        "PTG-ST-",
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
        f"PTG-ST-{mayor + 1:06d}"
    )


def obtener_consumidas_74():

    consumidas = set()

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_74",
            set()
        )
    )

    return consumidas


def generar_nivel_74(
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

        sintomas = str(
            verdadera[
                "Síntomas/Señales clave (checklist)"
            ]
        ).strip()

        if (
            not fuente
            or not patologia
            or not sintomas
        ):

            continue

        clave_correcta = normalizar_74(
            sintomas
        )

        falsas = candidatos[
            candidatos["Fuente_ID"] != fuente
        ].copy()

        falsas["Sintomas"] = (
            falsas[
                "Síntomas/Señales clave (checklist)"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        falsas = falsas[
            falsas["Sintomas"] != ""
        ]

        falsas = falsas[
            falsas["Sintomas"].map(
                normalizar_74
            ) != clave_correcta
        ]

        falsas = falsas.drop_duplicates(
            subset=["Sintomas"]
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
                "Sintomas"
            ]

            clave = normalizar_74(
                texto
            )

            if clave in [
                normalizar_74(x)
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
            sintomas,
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
                sintomas
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


def construir_pregunta_74(
    resultado
):

    opciones = resultado[
        "Opciones"
    ]

    return {

        "Pregunta_ID":
            siguiente_id_74(),

        "Modulo":
            "Patología",

        "Tema":
            "Síntomas",

        "Nivel":
            resultado["Nivel"],

        "Tipo_Relacion":
            "Patología-Síntomas",

        "Pregunta":
            (
                "¿Cuáles de las siguientes "
                "opciones corresponden a los "
                "síntomas o señales clave de la "
                f"patología {resultado['Patología']}?"
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


def generar_preguntas_74(
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
        obtener_consumidas_74()
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

            resultado = generar_nivel_74(
                df,
                consumidas,
                nivel
            )

            if resultado is None:

                continue

            pregunta = construir_pregunta_74(
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
        "fuentes_consumidas_74"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR 7.4
# ============================================================

if "df_disponible_71" in st.session_state:

    st.markdown(
        "### 7.4 Parte 2 - Generador "
        "Patología - Síntomas"
    )

    modo_74 = st.selectbox(
        "Seleccione el nivel",
        [
            "Nivel 1",
            "Nivel 2",
            "Niveles 1 y 2"
        ],
        key="modo_generacion_74"
    )

    cantidad_74 = st.number_input(
        "Cantidad máxima de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_74"
    )

    if st.button(
        "GENERAR PREGUNTAS 7.4",
        key="generar_preguntas_74"
    ):

        nuevas_74 = generar_preguntas_74(
            cantidad_74,
            modo_74
        )

        st.session_state[
            "preguntas_generadas_74"
        ] = nuevas_74

        if nuevas_74:

            st.success(
                f"Se generaron "
                f"{len(nuevas_74)} preguntas."
            )

        else:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS 7.4
# ============================================================

preguntas_74 = st.session_state.get(
    "preguntas_generadas_74",
    []
)

if preguntas_74:

    st.markdown(
        "### Preguntas generadas 7.4"
    )

    for pregunta in preguntas_74:

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
# 7.4 - PARTE 3
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS
# PATOLOGÍA - SÍNTOMAS
# ============================================================

if "preguntas_generadas_74" in st.session_state:

    preguntas_74 = st.session_state[
        "preguntas_generadas_74"
    ]

else:

    preguntas_74 = []


if preguntas_74:

    st.markdown(
        "## 7.4 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(preguntas_74):

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
            "Fuente utilizada: "
            f"{pregunta['Fuente_ID']}"
        )

        st.write(
            "**Estado actual:** "
            f"{pregunta.get('Estado', 'PENDIENTE')}"
        )

        observacion_74 = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_validador_74_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_validador_74_{i}"
            ):

                preguntas_actualizadas_74 = (
                    st.session_state[
                        "preguntas_generadas_74"
                    ]
                )

                preguntas_actualizadas_74[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_actualizadas_74[i][
                    "Observacion_Administrador"
                ] = observacion_74

                st.session_state[
                    "preguntas_generadas_74"
                ] = preguntas_actualizadas_74

                st.rerun()

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_validador_74_{i}"
            ):

                preguntas_actualizadas_74 = (
                    st.session_state[
                        "preguntas_generadas_74"
                    ]
                )

                preguntas_actualizadas_74[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_actualizadas_74[i][
                    "Observacion_Administrador"
                ] = observacion_74

                st.session_state[
                    "preguntas_generadas_74"
                ] = preguntas_actualizadas_74

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 7.4
# ============================================================

if preguntas_74:

    aprobadas_74 = sum(
        1
        for pregunta in preguntas_74
        if pregunta.get("Estado") == "APROBADA"
    )

    rechazadas_74 = sum(
        1
        for pregunta in preguntas_74
        if pregunta.get("Estado") == "RECHAZADA"
    )

    pendientes_74 = sum(
        1
        for pregunta in preguntas_74
        if pregunta.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 7.4"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_74
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_74
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_74
        )

    if pendientes_74 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        if aprobadas_74 > 0:

            st.info(
                f"{aprobadas_74} preguntas aprobadas "
                "quedan disponibles para sincronización."
            )

        if rechazadas_74 > 0:

            st.info(
                f"{rechazadas_74} preguntas rechazadas "
                "no serán incorporadas al banco."
            )
# ============================================================
# 7.4 - PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_74 = "franquiciasauces"
GITHUB_REPOSITORIO_74 = "Asesores"
GITHUB_RAMA_74 = "main"

GITHUB_ARCHIVO_74 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_74 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_74}/"
    f"{GITHUB_REPOSITORIO_74}/contents/"
    f"{GITHUB_ARCHIVO_74}"
)


def sincronizar_banco_74():

    preguntas = st.session_state.get(
        "preguntas_generadas_74",
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

        solicitud = urllib.request.Request(
            URL_GITHUB_74,
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
                "7.4 ERROR: faltan columnas "
                "en las preguntas generadas: "
                f"{', '.join(faltantes)}"
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

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

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

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

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_74,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_74,
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
# BOTÓN DE SINCRONIZACIÓN 7.4
# ============================================================

if preguntas_74:

    pendientes_74 = sum(
        1
        for p in preguntas_74
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_74 == 0:

        if st.button(
            "SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_74"
        ):

            sincronizar_banco_74()
# ============================================================
# 7.5 - PARTE 1
# CARGA PATOLOGIAS Y REGLAS_PAQUETES
# CONSTRUYE DATAFRAME PATOLOGIA-PRODUCTO
# SOLO PRIORIDAD 1
# ============================================================

st.markdown(
    "## 7.5 Patología - Producto"
)

st.write(
    "Carga las hojas Patologias y Reglas_Paquetes "
    "y construye las relaciones de prioridad 1."
)


# ============================================================
# CARGAR HOJAS
# ============================================================

if st.button(
    "CARGAR DATOS 7.5",
    key="cargar_fuentes_75"
):

    df_patologias_75 = pd.read_excel(
        ARCHIVO_FUENTE_71,
        sheet_name="Patologias",
        engine="openpyxl"
    )

    df_reglas_75 = pd.read_excel(
        ARCHIVO_FUENTE_71,
        sheet_name="Reglas_Paquetes",
        engine="openpyxl"
    )


    # ========================================================
    # COLUMNAS NECESARIAS DE PATOLOGIAS
    # ========================================================

    columnas_patologias_75 = [
        "Patologia_ID",
        "Patología",
        "Descripción breve (para cliente)"
    ]

    faltantes_patologias_75 = [
        columna
        for columna in columnas_patologias_75
        if columna not in df_patologias_75.columns
    ]

    if faltantes_patologias_75:

        st.error(
            "7.5 ERROR: faltan columnas en "
            "Patologias: "
            + ", ".join(
                faltantes_patologias_75
            )
        )

        st.stop()


    # ========================================================
    # COLUMNAS NECESARIAS DE REGLAS_PAQUETES
    # ========================================================

    columnas_reglas_75 = [
        "Patologia_ID",
        "Prioridad (1=alta)",
        "Segmento/Perfil",
        "Producto principal",
        "Coadyuvantes sugeridos (1-3)"
    ]

    faltantes_reglas_75 = [
        columna
        for columna in columnas_reglas_75
        if columna not in df_reglas_75.columns
    ]

    if faltantes_reglas_75:

        st.error(
            "7.5 ERROR: faltan columnas en "
            "Reglas_Paquetes: "
            + ", ".join(
                faltantes_reglas_75
            )
        )

        st.stop()


    # ========================================================
    # CONSERVAR SOLO LAS COLUMNAS NECESARIAS
    # ========================================================

    df_patologias_75 = df_patologias_75[
        columnas_patologias_75
    ].copy()

    df_reglas_75 = df_reglas_75[
        columnas_reglas_75
    ].copy()


    # ========================================================
    # LIMPIAR CAMPOS
    # ========================================================

    for columna in columnas_patologias_75:

        df_patologias_75[columna] = (
            df_patologias_75[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    for columna in columnas_reglas_75:

        df_reglas_75[columna] = (
            df_reglas_75[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )


    # ========================================================
    # SOLO PRIORIDAD 1
    # ========================================================

    df_reglas_75 = df_reglas_75[
        pd.to_numeric(
            df_reglas_75[
                "Prioridad (1=alta)"
            ],
            errors="coerce"
        ) == 1
    ].copy()


    # ========================================================
    # UNIR POR PATOLOGIA_ID
    # ========================================================

    df_trabajo_75 = pd.merge(
        df_patologias_75,
        df_reglas_75,
        on="Patologia_ID",
        how="inner"
    )


    # ========================================================
    # ESTRUCTURA FINAL
    # ========================================================

    df_trabajo_75 = df_trabajo_75[
        [
            "Patologia_ID",
            "Patología",
            "Descripción breve (para cliente)",
            "Prioridad (1=alta)",
            "Segmento/Perfil",
            "Producto principal",
            "Coadyuvantes sugeridos (1-3)"
        ]
    ].drop_duplicates(
        ignore_index=True
    )


    # ========================================================
    # GUARDAR EN SESSION STATE
    # ========================================================

    st.session_state[
        "df_trabajo_75"
    ] = df_trabajo_75.copy()


    # ========================================================
    # RESULTADO
    # ========================================================

    st.success(
        "7.5 Parte 1 cargada correctamente."
    )

    st.info(
        f"Patologías cargadas: "
        f"{len(df_patologias_75):,}"
    )

    st.info(
        f"Reglas de prioridad 1: "
        f"{len(df_reglas_75):,}"
    )

    st.info(
        f"Relaciones Patología-Producto: "
        f"{len(df_trabajo_75):,}"
    )

    st.dataframe(
        df_trabajo_75,
        use_container_width=True,
        hide_index=True
    )
# ============================================================
# 7.5 - PARTE 2
# GENERADOR PATOLOGÍA - PRODUCTO
# NIVEL 1 Y NIVEL 2
# ============================================================

def normalizar_75(valor):

    if pd.isna(valor):
        return ""

    return " ".join(
        str(valor).strip().lower().split()
    )


def siguiente_id_75():

    mayor = 0

    for pregunta in st.session_state.get(
        "preguntas_generadas_75",
        []
    ):

        texto = str(
            pregunta.get(
                "Pregunta_ID",
                ""
            )
        )

        if texto.startswith("PTG-PP-"):

            try:

                numero = int(
                    texto.replace(
                        "PTG-PP-",
                        ""
                    )
                )

                mayor = max(
                    mayor,
                    numero
                )

            except ValueError:
                pass

    return f"PTG-PP-{mayor + 1:06d}"


def obtener_consumidas_75():

    consumidas = set()

    consumidas.update(
        st.session_state.get(
            "fuentes_consumidas_75",
            set()
        )
    )

    return consumidas


def generar_nivel_75(
    df,
    consumidas,
    nivel
):

    candidatos = df[
        ~df["Patologia_ID"].isin(
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

        patologia_id = str(
            verdadera["Patologia_ID"]
        ).strip()

        patologia = str(
            verdadera["Patología"]
        ).strip()

        producto = str(
            verdadera["Producto principal"]
        ).strip()

        if (
            not patologia_id
            or not patologia
            or not producto
        ):

            continue

        clave_producto = normalizar_75(
            producto
        )

        falsas = candidatos[
            candidatos["Patologia_ID"]
            != patologia_id
        ].copy()

        falsas["Producto"] = (
            falsas[
                "Producto principal"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        falsas = falsas[
            falsas["Producto"] != ""
        ]

        falsas = falsas[
            falsas["Producto"].map(
                normalizar_75
            ) != clave_producto
        ]

        falsas = falsas.drop_duplicates(
            subset=["Producto"]
        )

        if len(falsas) < 3:

            continue

        falsas = falsas.sample(
            frac=1
        ).reset_index(
            drop=True
        )

        seleccionadas = []

        productos_usados = {
            clave_producto
        }

        for _, falsa in falsas.iterrows():

            producto_falso = falsa[
                "Producto"
            ]

            clave_falso = normalizar_75(
                producto_falso
            )

            if (
                not clave_falso
                or clave_falso in productos_usados
            ):

                continue

            seleccionadas.append(
                producto_falso
            )

            productos_usados.add(
                clave_falso
            )

            if len(seleccionadas) == 3:

                break

        if len(seleccionadas) != 3:

            continue

        opciones = [
            producto,
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
                producto
            ) + 1
        )

        return {
            "Patologia_ID":
                patologia_id,

            "Patología":
                patologia,

            "Opciones":
                opciones,

            "Correcta":
                correcta,

            "Nivel":
                nivel
        }

    return None


def construir_pregunta_75(
    resultado
):

    opciones = resultado[
        "Opciones"
    ]

    return {

        "Pregunta_ID":
            siguiente_id_75(),

        "Modulo":
            "Patología",

        "Tema":
            "Producto",

        "Nivel":
            resultado["Nivel"],

        "Tipo_Relacion":
            "Patología-Producto",

        "Pregunta":
            (
                "¿Cuál de los siguientes "
                "productos está recomendado "
                "para la patología "
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
            resultado["Patologia_ID"]
    }


def generar_preguntas_75(
    cantidad,
    modo
):

    df = st.session_state.get(
        "df_trabajo_75",
        pd.DataFrame()
    )

    if df.empty:

        return []

    consumidas = (
        obtener_consumidas_75()
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

            resultado = generar_nivel_75(
                df,
                consumidas,
                nivel
            )

            if resultado is None:

                continue

            pregunta = construir_pregunta_75(
                resultado
            )

            preguntas.append(
                pregunta
            )

            consumidas.add(
                resultado["Patologia_ID"]
            )

            generado = True

        if not generado:

            break

    st.session_state[
        "fuentes_consumidas_75"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR 7.5
# ============================================================

if "df_trabajo_75" in st.session_state:

    st.markdown(
        "### 7.5 Parte 2 - Generador "
        "Patología - Producto"
    )

    modo_75 = st.selectbox(
        "Seleccione el nivel",
        [
            "Nivel 1",
            "Nivel 2",
            "Niveles 1 y 2"
        ],
        key="modo_generacion_75"
    )

    cantidad_75 = st.number_input(
        "Cantidad máxima de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_75"
    )

    if st.button(
        "GENERAR PREGUNTAS 7.5",
        key="generar_preguntas_75"
    ):

        nuevas_75 = generar_preguntas_75(
            cantidad_75,
            modo_75
        )

        st.session_state[
            "preguntas_generadas_75"
        ] = nuevas_75

        if nuevas_75:

            st.success(
                f"Se generaron "
                f"{len(nuevas_75)} preguntas."
            )

        else:

            st.warning(
                "No hay suficientes relaciones "
                "disponibles para generar preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_75 = st.session_state.get(
    "preguntas_generadas_75",
    []
)

if preguntas_75:

    st.markdown(
        "### Preguntas generadas 7.5"
    )

    for pregunta in preguntas_75:

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

# ============================================================
# 7.5 - PARTE 3
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS
# PATOLOGÍA - PRODUCTO
# ============================================================

preguntas_75 = st.session_state.get(
    "preguntas_generadas_75",
    []
)

if preguntas_75:

    st.markdown(
        "## 7.5 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(preguntas_75):

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
            f"Fuente utilizada: "
            f"{pregunta['Fuente_ID']}"
        )

        estado = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado}"
        )

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_75_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_75_{i}"
            ):

                preguntas_75[i]["Estado"] = (
                    "APROBADA"
                )

                preguntas_75[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_75"
                ] = preguntas_75

                st.rerun()

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_75_{i}"
            ):

                preguntas_75[i]["Estado"] = (
                    "RECHAZADA"
                )

                preguntas_75[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_75"
                ] = preguntas_75

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 7.5
# ============================================================

if preguntas_75:

    aprobadas_75 = sum(
        1
        for p in preguntas_75
        if p.get("Estado") == "APROBADA"
    )

    rechazadas_75 = sum(
        1
        for p in preguntas_75
        if p.get("Estado") == "RECHAZADA"
    )

    pendientes_75 = sum(
        1
        for p in preguntas_75
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 7.5"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_75
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_75
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_75
        )

    if pendientes_75 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "Las preguntas aprobadas quedan "
            "listas para la sincronización."
        )


# ============================================================
# 7.5 - PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_75 = "franquiciasauces"

GITHUB_REPOSITORIO_75 = "Asesores"

GITHUB_RAMA_75 = "main"

GITHUB_ARCHIVO_75 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_75 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_75}/"
    f"{GITHUB_REPOSITORIO_75}/contents/"
    f"{GITHUB_ARCHIVO_75}"
)


def sincronizar_banco_75():

    preguntas = st.session_state.get(
        "preguntas_generadas_75",
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

        solicitud = urllib.request.Request(
            URL_GITHUB_75,
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

        df_banco = pd.read_excel(
            contenido
        )

        total_antes = len(
            df_banco
        )

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
                "7.5 ERROR: faltan columnas "
                "en las preguntas generadas: "
                + ", ".join(faltantes)
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

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

        total_nuevas = len(
            df_nuevas
        )

        if total_nuevas == 0:

            st.info(
                "No hay preguntas nuevas "
                "para agregar."
            )

            st.info(
                f"Preguntas existentes: "
                f"**{total_antes:,}**"
            )

            return

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

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

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_75,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode(
            "utf-8"
        )

        solicitud = urllib.request.Request(
            URL_GITHUB_75,
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
# BOTÓN DE SINCRONIZACIÓN 7.5
# ============================================================

if preguntas_75:

    pendientes_75 = sum(
        1
        for p in preguntas_75
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_75 == 0:

        if st.button(
            "SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_75"
        ):

            sincronizar_banco_75()

# ============================================================
# 7.6 - PARTE 2
# GENERADOR PATOLOGÍA - DESCRIPCIÓN - PRODUCTO
# NIVEL 1 Y NIVEL 2
# ============================================================


def normalizar_76(valor):

    if pd.isna(valor):
        return ""

    return " ".join(
        str(valor).strip().lower().split()
    )


def siguiente_id_76():

    mayor = 0

    for pregunta in st.session_state.get(
        "preguntas_generadas_76",
        []
    ):

        texto = str(
            pregunta.get(
                "Pregunta_ID",
                ""
            )
        ).strip()

        if texto.startswith("PTG-PD-"):

            try:

                numero = int(
                    texto.replace(
                        "PTG-PD-",
                        ""
                    )
                )

                mayor = max(
                    mayor,
                    numero
                )

            except ValueError:
                pass

    return f"PTG-PD-{mayor + 1:06d}"


def obtener_consumidas_76():

    return set(
        st.session_state.get(
            "fuentes_consumidas_76",
            set()
        )
    )


def generar_nivel_76(
    df,
    consumidas,
    nivel
):

    candidatos = df.copy()

    # --------------------------------------------------------
    # SOLO PRIORIDAD 1
    # --------------------------------------------------------

    if "Prioridad (1=alta)" in candidatos.columns:

        candidatos = candidatos[
            pd.to_numeric(
                candidatos[
                    "Prioridad (1=alta)"
                ],
                errors="coerce"
            ) == 1
        ].copy()

    # --------------------------------------------------------
    # EXCLUIR PATOLOGÍAS YA UTILIZADAS
    # --------------------------------------------------------

    candidatos = candidatos[
        ~candidatos[
            "Patologia_ID"
        ].astype(str).str.strip().isin(
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

    # --------------------------------------------------------
    # BUSCAR RELACIÓN CORRECTA
    # --------------------------------------------------------

    for _, verdadera in candidatos.iterrows():

        patologia_id = str(
            verdadera[
                "Patologia_ID"
            ]
        ).strip()

        patologia = str(
            verdadera[
                "Patología"
            ]
        ).strip()

        descripcion = str(
            verdadera[
                "Descripción breve (para cliente)"
            ]
        ).strip()

        producto = str(
            verdadera[
                "Producto principal"
            ]
        ).strip()

        if (
            not patologia_id
            or not patologia
            or not descripcion
            or not producto
        ):

            continue

        clave_descripcion = normalizar_76(
            descripcion
        )

        clave_producto = normalizar_76(
            producto
        )

        # ----------------------------------------------------
        # BUSCAR PRODUCTOS FALSOS
        # ----------------------------------------------------

        falsas = candidatos[
            candidatos[
                "Patologia_ID"
            ].astype(str).str.strip()
            != patologia_id
        ].copy()

        falsas["Producto"] = (
            falsas[
                "Producto principal"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        falsas["Descripcion"] = (
            falsas[
                "Descripción breve (para cliente)"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        falsas = falsas[
            falsas["Producto"] != ""
        ].copy()

        falsas = falsas[
            falsas["Descripcion"] != ""
        ].copy()

        falsas = falsas[
            falsas["Producto"].map(
                normalizar_76
            ) != clave_producto
        ].copy()

        falsas = falsas[
            falsas["Descripcion"].map(
                normalizar_76
            ) != clave_descripcion
        ].copy()

        falsas = falsas.drop_duplicates(
            subset=["Producto"]
        )

        if len(falsas) < 3:

            continue

        falsas = falsas.sample(
            frac=1
        ).reset_index(
            drop=True
        )

        seleccionadas = []

        productos_usados = {
            clave_producto
        }

        for _, falsa in falsas.iterrows():

            producto_falso = str(
                falsa[
                    "Producto"
                ]
            ).strip()

            clave_falso = normalizar_76(
                producto_falso
            )

            if not clave_falso:

                continue

            if clave_falso in productos_usados:

                continue

            seleccionadas.append(
                producto_falso
            )

            productos_usados.add(
                clave_falso
            )

            if len(seleccionadas) == 3:

                break

        if len(seleccionadas) != 3:

            continue

        # ----------------------------------------------------
        # CONSTRUIR LAS 4 OPCIONES
        # ----------------------------------------------------

        opciones = [
            producto,
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
                producto
            ) + 1
        )

        return {

            "Patologia_ID":
                patologia_id,

            "Patología":
                patologia,

            "Descripción":
                descripcion,

            "Opciones":
                opciones,

            "Correcta":
                correcta,

            "Nivel":
                nivel
        }

    return None


def construir_pregunta_76(
    resultado
):

    opciones = resultado[
        "Opciones"
    ]

    descripcion = resultado[
        "Descripción"
    ]

    return {

        "Pregunta_ID":
            siguiente_id_76(),

        "Modulo":
            "Patología",

        "Tema":
            "Descripción-Producto",

        "Nivel":
            resultado[
                "Nivel"
            ],

        "Tipo_Relacion":
            "Patología-Descripción-Producto",

        "Pregunta":
            (
                "¿Cuál de los siguientes "
                "productos está recomendado "
                "para una patología cuya "
                "descripción es la siguiente: "
                f"{descripcion}?"
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
                resultado[
                    "Correcta"
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
            resultado[
                "Patologia_ID"
            ]
    }


def generar_preguntas_76(
    cantidad,
    modo
):

    df = st.session_state.get(
        "df_trabajo_75",
        pd.DataFrame()
    )

    if df.empty:

        return []

    consumidas = (
        obtener_consumidas_76()
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

            resultado = generar_nivel_76(
                df,
                consumidas,
                nivel
            )

            if resultado is None:

                continue

            pregunta = construir_pregunta_76(
                resultado
            )

            preguntas.append(
                pregunta
            )

            consumidas.add(
                str(
                    resultado[
                        "Patologia_ID"
                    ]
                ).strip()
            )

            generado = True

        if not generado:

            break

    st.session_state[
        "fuentes_consumidas_76"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR 7.6
# ============================================================

if "df_trabajo_75" in st.session_state:

    st.markdown(
        "### 7.6 Parte 2 - Generador "
        "Patología - Descripción - Producto"
    )

    modo_76 = st.selectbox(
        "Seleccione el nivel",
        [
            "Nivel 1",
            "Nivel 2",
            "Niveles 1 y 2"
        ],
        key="modo_generacion_patologia_descripcion_producto_76"
    )

    cantidad_76 = st.number_input(
        "Cantidad máxima de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_patologia_descripcion_producto_76"
    )

    if st.button(
        "GENERAR PREGUNTAS 7.6",
        key="generar_preguntas_patologia_descripcion_producto_76"
    ):

        nuevas_76 = generar_preguntas_76(
            cantidad_76,
            modo_76
        )

        st.session_state[
            "preguntas_generadas_76"
        ] = nuevas_76

        if nuevas_76:

            st.success(
                f"Se generaron "
                f"{len(nuevas_76)} preguntas."
            )

        else:

            st.warning(
                "No hay suficientes relaciones "
                "de prioridad 1 disponibles "
                "para generar preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS 7.6
# ============================================================

preguntas_76 = st.session_state.get(
    "preguntas_generadas_76",
    []
)

if preguntas_76:

    st.markdown(
        "### Preguntas generadas 7.6"
    )

    for pregunta in preguntas_76:

        st.markdown(
            f"**{pregunta['Pregunta_ID']} — "
            f"{pregunta['Nivel']}**"
        )

        st.write(
            pregunta[
                "Pregunta"
            ]
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
# 7.6 - PARTE 3
# VALIDACIÓN INDIVIDUAL DE PREGUNTAS
# PATOLOGÍA - PRODUCTO DESDE DESCRIPCIÓN
# ============================================================

preguntas_76 = st.session_state.get(
    "preguntas_generadas_76",
    []
)

if preguntas_76:

    st.markdown(
        "## 7.6 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(
        preguntas_76
    ):

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
            "Fuente utilizada: "
            f"{pregunta['Fuente_ID']}"
        )

        estado = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado}"
        )

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_76_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_76_{i}"
            ):

                preguntas_76[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_76[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_76"
                ] = preguntas_76

                st.rerun()

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_76_{i}"
            ):

                preguntas_76[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_76[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_76"
                ] = preguntas_76

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 7.6
# ============================================================

if preguntas_76:

    aprobadas_76 = sum(
        1
        for p in preguntas_76
        if p.get(
            "Estado"
        ) == "APROBADA"
    )

    rechazadas_76 = sum(
        1
        for p in preguntas_76
        if p.get(
            "Estado"
        ) == "RECHAZADA"
    )

    pendientes_76 = sum(
        1
        for p in preguntas_76
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 7.6"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_76
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_76
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_76
        )

    if pendientes_76 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "Las preguntas aprobadas quedan "
            "listas para la sincronización."
        )

# ============================================================
# 7.6 - PARTE 4
# SINCRONIZAR PREGUNTAS CON BANCO GENERAL
# ============================================================

GITHUB_USUARIO_76 = "franquiciasauces"
GITHUB_REPOSITORIO_76 = "Asesores"
GITHUB_RAMA_76 = "main"

GITHUB_ARCHIVO_76 = (
    "BANCO_PREGUNTAS_GENERALES.xlsx"
)

URL_GITHUB_76 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_76}/"
    f"{GITHUB_REPOSITORIO_76}/contents/"
    f"{GITHUB_ARCHIVO_76}"
)


def sincronizar_banco_76():

    preguntas = st.session_state.get(
        "preguntas_generadas_76",
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

        solicitud = urllib.request.Request(
            URL_GITHUB_76,
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

        # ====================================================
        # LEER BANCO EXISTENTE
        # ====================================================

        df_banco = pd.read_excel(
            contenido,
            engine="openpyxl"
        )

        total_antes = len(
            df_banco
        )

        # ====================================================
        # PREPARAR PREGUNTAS NUEVAS
        # ====================================================

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
                "7.6 ERROR: faltan columnas "
                "en las preguntas generadas: "
                f"{', '.join(faltantes)}"
            )

            return

        df_nuevas = df_nuevas[
            columnas
        ].copy()

        # ====================================================
        # EVITAR DUPLICADOS
        # ====================================================

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

        # ====================================================
        # UNIR BANCO CON PREGUNTAS NUEVAS
        # ====================================================

        df_final = pd.concat(
            [
                df_banco,
                df_nuevas
            ],
            ignore_index=True
        )

        # ====================================================
        # CREAR ARCHIVO EXCEL SIN IO
        # ====================================================

        ruta_excel_76 = (
            "/tmp/BANCO_PREGUNTAS_GENERALES_76.xlsx"
        )

        with pd.ExcelWriter(
            ruta_excel_76,
            engine="openpyxl"
        ) as writer:

            df_final.to_excel(
                writer,
                index=False,
                sheet_name="Banco"
            )

        with open(
            ruta_excel_76,
            "rb"
        ) as archivo:

            contenido_nuevo = base64.b64encode(
                archivo.read()
            ).decode("utf-8")

        # ====================================================
        # ACTUALIZAR GITHUB
        # ====================================================

        datos_actualizacion = {

            "message":
                "Actualizar BANCO_PREGUNTAS_GENERALES",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_76,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_76,
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

        # ====================================================
        # RESULTADO
        # ====================================================

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

if preguntas_76:

    pendientes_76 = sum(
        1
        for p in preguntas_76
        if p.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    if pendientes_76 == 0:

        if st.button(
            "SINCRONIZAR CON BANCO DE PREGUNTAS",
            key="sincronizar_banco_76"
        ):

            sincronizar_banco_76()


# ============================================================
# 7.7 - PARTE 1
# GENERADOR PATOLOGÍA - CONDICIÓN - PRODUCTO + COADYUVANTES
# NIVEL 2
# ============================================================

def siguiente_id_77():

    mayor = 0

    df_banco = st.session_state.get(
        "df_banco_71",
        pd.DataFrame()
    )

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco["Pregunta_ID"].fillna(""):

            texto = str(valor).strip()

            if texto.startswith("PTG-PC-"):

                try:

                    numero = int(
                        texto.replace(
                            "PTG-PC-",
                            ""
                        )
                    )

                    mayor = max(
                        mayor,
                        numero
                    )

                except ValueError:

                    pass

    return f"PTG-PC-{mayor + 1:06d}"


def generar_preguntas_77(cantidad):

    df = st.session_state.get(
        "df_trabajo_75",
        pd.DataFrame()
    )

    if df.empty:

        return []

    df = df[
        pd.to_numeric(
            df["Prioridad (1=alta)"],
            errors="coerce"
        ) == 1
    ].copy()

    df = df[
        (df["Patologia_ID"].fillna("").astype(str).str.strip() != "")
        &
        (df["Patología"].fillna("").astype(str).str.strip() != "")
        &
        (df["Segmento/Perfil"].fillna("").astype(str).str.strip() != "")
        &
        (df["Producto principal"].fillna("").astype(str).str.strip() != "")
        &
        (
            df["Coadyuvantes sugeridos (1-3)"]
            .fillna("")
            .astype(str)
            .str.strip()
            != ""
        )
    ].copy()

    if len(df) < 4:

        return []

    usadas = st.session_state.get(
        "fuentes_consumidas_77",
        set()
    )

    df = df[
        ~df["Patologia_ID"]
        .astype(str)
        .str.strip()
        .isin(usadas)
    ].copy()

    if len(df) < 4:

        return []

    df = df.sample(
        frac=1
    ).reset_index(
        drop=True
    )

    preguntas = []

    for _, verdadera in df.iterrows():

        if len(preguntas) >= cantidad:

            break

        patologia_id = str(
            verdadera["Patologia_ID"]
        ).strip()

        patologia = str(
            verdadera["Patología"]
        ).strip()

        segmento = str(
            verdadera["Segmento/Perfil"]
        ).strip()

        producto = str(
            verdadera["Producto principal"]
        ).strip()

        coadyuvantes = str(
            verdadera[
                "Coadyuvantes sugeridos (1-3)"
            ]
        ).strip()

        paquete_correcto = (
            producto
            + " | "
            + coadyuvantes
        )

        falsas = df[
            df["Patologia_ID"]
            .astype(str)
            .str.strip()
            != patologia_id
        ].copy()

        falsas = falsas.sample(
            frac=1
        ).reset_index(
            drop=True
        )

        paquetes_falsos = []

        for _, falsa in falsas.iterrows():

            producto_falso = str(
                falsa["Producto principal"]
            ).strip()

            coadyuvantes_falsos = str(
                falsa[
                    "Coadyuvantes sugeridos (1-3)"
                ]
            ).strip()

            if (
                not producto_falso
                or not coadyuvantes_falsos
            ):

                continue

            paquete_falso = (
                producto_falso
                + " | "
                + coadyuvantes_falsos
            )

            if (
                paquete_falso.lower()
                == paquete_correcto.lower()
            ):

                continue

            if any(
                paquete_falso.lower()
                == existente.lower()
                for existente in paquetes_falsos
            ):

                continue

            paquetes_falsos.append(
                paquete_falso
            )

            if len(paquetes_falsos) == 3:

                break

        if len(paquetes_falsos) != 3:

            continue

        opciones = [
            paquete_correcto,
            paquetes_falsos[0],
            paquetes_falsos[1],
            paquetes_falsos[2]
        ]

        opciones = pd.Series(
            opciones
        ).sample(
            frac=1
        ).tolist()

        correcta = (
            opciones.index(
                paquete_correcto
            ) + 1
        )

        preguntas.append({

            "Pregunta_ID":
                siguiente_id_77(),

            "Modulo":
                "Patología",

            "Tema":
                "Condición - Producto + Coadyuvantes",

            "Nivel":
                "Nivel 2",

            "Tipo_Relacion":
                "Patología-Condición-Producto-Coadyuvantes",

            "Pregunta":
                (
                    f"Para la patología {patologia}, "
                    f"si se presenta el perfil o condición "
                    f"{segmento}, ¿cuál sería el paquete "
                    "recomendado?"
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
                patologia_id
        })

        usadas.add(
            patologia_id
        )

    st.session_state[
        "fuentes_consumidas_77"
    ] = usadas

    return preguntas


# ============================================================
# INTERFAZ 7.7
# ============================================================

st.markdown(
    "### 7.7 Parte 1 - "
    "Patología + Condición + Producto + Coadyuvantes"
)

st.info(
    "Nivel 2: relaciona la patología y el "
    "segmento o condición con el paquete recomendado."
)

cantidad_77 = st.number_input(
    "Cantidad máxima de preguntas",
    min_value=1,
    max_value=500,
    value=10,
    step=1,
    key="cantidad_generar_77_parte1"
)

if st.button(
    "GENERAR PREGUNTAS 7.7",
    key="generar_preguntas_77_parte1"
):

    if "df_trabajo_75" not in st.session_state:

        st.error(
            "Primero debe cargar la Parte 1 de 7.5."
        )

    else:

        nuevas_77 = generar_preguntas_77(
            cantidad_77
        )

        st.session_state[
            "preguntas_generadas_77"
        ] = nuevas_77

        if nuevas_77:

            st.success(
                f"Se generaron "
                f"{len(nuevas_77)} preguntas."
            )

        else:

            st.warning(
                "No hay suficientes relaciones de "
                "prioridad 1 para generar preguntas."
            )


# ============================================================
# MOSTRAR PREGUNTAS
# ============================================================

preguntas_77 = st.session_state.get(
    "preguntas_generadas_77",
    []
)

if preguntas_77:

    st.markdown(
        "### Preguntas generadas 7.7"
    )

    for pregunta in preguntas_77:

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
# 7.7 - PARTE 3
# VALIDACIÓN INDIVIDUAL
# PATOLOGÍA - CONDICIÓN - PRODUCTO - COADYUVANTES
# ============================================================

preguntas_77 = st.session_state.get(
    "preguntas_generadas_77",
    []
)

if preguntas_77:

    st.markdown(
        "## 7.7 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta individualmente. "
        "Una pregunta rechazada no afecta las demás."
    )

    for i, pregunta in enumerate(preguntas_77):

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

        estado = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"**Estado actual:** {estado}"
        )

        observacion = st.text_input(
            "Observación del administrador",
            value=pregunta.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_validacion_77_{i}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_validacion_77_{i}"
            ):

                preguntas_77[i]["Estado"] = "APROBADA"

                preguntas_77[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_77"
                ] = preguntas_77

                st.rerun()

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_validacion_77_{i}"
            ):

                preguntas_77[i]["Estado"] = "RECHAZADA"

                preguntas_77[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_77"
                ] = preguntas_77

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 7.7
# ============================================================

if preguntas_77:

    aprobadas_77 = sum(
        1
        for pregunta in preguntas_77
        if pregunta.get("Estado") == "APROBADA"
    )

    rechazadas_77 = sum(
        1
        for pregunta in preguntas_77
        if pregunta.get("Estado") == "RECHAZADA"
    )

    pendientes_77 = sum(
        1
        for pregunta in preguntas_77
        if pregunta.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 7.7"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_77
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_77
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_77
        )

    if pendientes_77 == 0:

        st.success(
            "Todas las preguntas fueron "
            "revisadas individualmente."
        )

        st.info(
            "Las preguntas aprobadas quedan "
            "listas para sincronización."
        )
# ============================================================
# 7.7 - PARTE 4
# SINCRONIZAR PREGUNTAS APROBADAS CON BANCO GENERAL
# ============================================================

import io

GITHUB_USUARIO_77 = "franquiciasauces"
GITHUB_REPOSITORIO_77 = "Asesores"
GITHUB_RAMA_77 = "main"

GITHUB_ARCHIVO_77 = "BANCO_PREGUNTAS_GENERALES.xlsx"

URL_GITHUB_77 = (
    "https://api.github.com/repos/"
    f"{GITHUB_USUARIO_77}/"
    f"{GITHUB_REPOSITORIO_77}/contents/"
    f"{GITHUB_ARCHIVO_77}"
)


def sincronizar_banco_77():

    preguntas = st.session_state.get(
        "preguntas_generadas_77",
        []
    )

    if not preguntas:

        st.warning(
            "No hay preguntas 7.7 para sincronizar."
        )

        return

    pendientes = [
        pregunta
        for pregunta in preguntas
        if pregunta.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    ]

    if pendientes:

        st.error(
            f"Hay {len(pendientes)} preguntas pendientes "
            "de validación."
        )

        return

    aprobadas = [
        pregunta
        for pregunta in preguntas
        if pregunta.get(
            "Estado",
            ""
        ) == "APROBADA"
    ]

    if not aprobadas:

        st.warning(
            "No hay preguntas aprobadas para sincronizar."
        )

        return

    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    try:

        solicitud = urllib.request.Request(
            URL_GITHUB_77,
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
            io.BytesIO(contenido),
            engine="openpyxl"
        )

        total_antes = len(df_banco)

        df_aprobadas = pd.DataFrame(
            aprobadas
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
            if columna not in df_aprobadas.columns
        ]

        if faltantes:

            st.error(
                "7.7 ERROR: faltan columnas: "
                + ", ".join(faltantes)
            )

            return

        df_aprobadas = df_aprobadas[
            columnas
        ].copy()

        nuevas = 0
        actualizadas = 0

        if "Pregunta_ID" not in df_banco.columns:

            df_final = pd.concat(
                [
                    df_banco,
                    df_aprobadas
                ],
                ignore_index=True
            )

            nuevas = len(df_aprobadas)

        else:

            df_final = df_banco.copy()

            for _, pregunta in df_aprobadas.iterrows():

                pregunta_id = str(
                    pregunta["Pregunta_ID"]
                ).strip()

                coincidencias = (
                    df_final["Pregunta_ID"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    == pregunta_id
                )

                if coincidencias.any():

                    indice = df_final.index[
                        coincidencias
                    ][0]

                    for columna in columnas:

                        df_final.at[
                            indice,
                            columna
                        ] = pregunta[columna]

                    actualizadas += 1

                else:

                    df_final = pd.concat(
                        [
                            df_final,
                            pd.DataFrame(
                                [pregunta]
                            )
                        ],
                        ignore_index=True
                    )

                    nuevas += 1

        total_despues = len(df_final)

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

        datos_actualizacion = {
            "message":
                "Actualizar preguntas aprobadas 7.7",

            "content":
                contenido_nuevo,

            "branch":
                GITHUB_RAMA_77,

            "sha":
                sha
        }

        cuerpo = json.dumps(
            datos_actualizacion
        ).encode("utf-8")

        solicitud = urllib.request.Request(
            URL_GITHUB_77,
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

        st.success(
            "Banco de preguntas actualizado correctamente."
        )

        st.info(
            f"Preguntas en el banco antes: {total_antes:,}"
        )

        st.info(
            f"Preguntas aprobadas procesadas: "
            f"{len(aprobadas):,}"
        )

        st.info(
            f"Preguntas nuevas incorporadas: {nuevas:,}"
        )

        st.info(
            f"Preguntas existentes actualizadas: "
            f"{actualizadas:,}"
        )

        st.info(
            f"Preguntas en el banco después: "
            f"{total_despues:,}"
        )

        st.dataframe(
            df_aprobadas,
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
# BOTÓN DE SINCRONIZACIÓN 7.7
# ============================================================

preguntas_77 = st.session_state.get(
    "preguntas_generadas_77",
    []
)

if preguntas_77:

    pendientes_77 = sum(
        1
        for pregunta in preguntas_77
        if pregunta.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    aprobadas_77 = sum(
        1
        for pregunta in preguntas_77
        if pregunta.get(
            "Estado",
            ""
        ) == "APROBADA"
    )

    if pendientes_77 == 0:

        if aprobadas_77 > 0:

            if st.button(
                "SINCRONIZAR APROBADAS 7.7",
                key="sincronizar_aprobadas_77"
            ):

                sincronizar_banco_77()

        else:

            st.warning(
                "No hay preguntas aprobadas para sincronizar."
            )

# ============================================================
# 8.1 - RESTRICCIONES
# PARTE 1 - CARGA Y CONTROL DE FUENTE
# ============================================================

ARCHIVO_FUENTE_81 = "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
HOJA_FUENTE_81 = "Restricciones"
ARCHIVO_BANCO_81 = "BANCO_PREGUNTAS_GENERALES.xlsx"


def cargar_fuente_81():

    try:
        df = pd.read_excel(
            ARCHIVO_FUENTE_81,
            sheet_name=HOJA_FUENTE_81,
            engine="openpyxl"
        )
    except Exception as error:
        st.error(
            f"8.1 ERROR al cargar {HOJA_FUENTE_81}: {error}"
        )
        return None

    columnas = [
        "Restriccion_ID",
        "Producto",
        "Tipo",
        "Precaución / Contraindicación",
        "Motivo",
        "Alternativas seguras"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:
        st.error(
            "8.1 ERROR: faltan columnas: "
            + ", ".join(faltantes)
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
        (df["Restriccion_ID"] != "")
        & (df["Producto"] != "")
        & (df["Tipo"] != "")
        & (df["Precaución / Contraindicación"] != "")
        & (df["Motivo"] != "")
        & (df["Alternativas seguras"] != "")
    ].copy()

    df["Fuente_ID"] = [
        f"RX-{i:06d}"
        for i in range(1, len(df) + 1)
    ]

    return df.reset_index(drop=True)


def cargar_banco_81():

    try:
        return pd.read_excel(
            ARCHIVO_BANCO_81,
            engine="openpyxl"
        )
    except Exception:
        return pd.DataFrame()


def obtener_fuentes_usadas_81(df_banco):

    usadas = set()

    if df_banco.empty:
        return usadas

    if "Fuente_ID" not in df_banco.columns:
        return usadas

    for valor in df_banco["Fuente_ID"].fillna(""):
        for fuente in str(valor).split(";"):
            fuente = fuente.strip()

            if fuente:
                usadas.add(fuente)

    return usadas


# ============================================================
# INTERFAZ 8.1
# ============================================================

st.markdown("## 8.1 Restricciones")

st.write(
    "Carga y control de la hoja Restricciones."
)

if st.button(
    "CARGAR Y VALIDAR FUENTE 8.1",
    key="cargar_fuentes_81"
):

    df_fuente_81 = cargar_fuente_81()

    if df_fuente_81 is None:
        st.stop()

    df_banco_81 = cargar_banco_81()

    fuentes_usadas_81 = obtener_fuentes_usadas_81(
        df_banco_81
    )

    df_disponible_81 = df_fuente_81[
        ~df_fuente_81["Fuente_ID"].isin(
            fuentes_usadas_81
        )
    ].copy()

    st.session_state["df_fuente_81"] = (
        df_fuente_81.copy()
    )

    st.session_state["df_banco_81"] = (
        df_banco_81.copy()
    )

    st.session_state["df_disponible_81"] = (
        df_disponible_81.reset_index(drop=True)
    )

    st.session_state["fuentes_usadas_81"] = (
        fuentes_usadas_81
    )


# ============================================================
# MOSTRAR CONTROL
# ============================================================

if "df_fuente_81" in st.session_state:

    df_fuente_81 = st.session_state["df_fuente_81"]

    df_disponible_81 = st.session_state[
        "df_disponible_81"
    ]

    fuentes_usadas_81 = st.session_state[
        "fuentes_usadas_81"
    ]

    st.success(
        "8.1 cargó correctamente la hoja Restricciones."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            "Restricciones fuente",
            len(df_fuente_81)
        )

    with col2:
        st.metric(
            "Fuentes utilizadas",
            len(
                fuentes_usadas_81
                & set(df_fuente_81["Fuente_ID"])
            )
        )

    with col3:
        st.metric(
            "Restricciones disponibles",
            len(df_disponible_81)
        )

    st.markdown("### Estructura cargada")

    st.dataframe(
        df_disponible_81[
            [
                "Fuente_ID",
                "Restriccion_ID",
                "Producto",
                "Tipo",
                "Precaución / Contraindicación",
                "Motivo",
                "Alternativas seguras"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
# ============================================================
# 8.1 - PARTE 2
# GENERADOR PRODUCTO - PRECAUCIÓN / CONTRAINDICACIÓN
# NIVEL 1
# ============================================================

def siguiente_id_81():

    df_banco = st.session_state.get(
        "df_banco_81",
        pd.DataFrame()
    )

    mayor = 0

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco["Pregunta_ID"].fillna(""):

            texto = str(valor).strip()

            if texto.startswith("PTRX-"):

                try:
                    numero = int(
                        texto.replace("PTRX-", "")
                    )
                    mayor = max(mayor, numero)

                except ValueError:
                    pass

    preguntas = st.session_state.get(
        "preguntas_generadas_81",
        []
    )

    for pregunta in preguntas:

        texto = str(
            pregunta.get("Pregunta_ID", "")
        ).strip()

        if texto.startswith("PTRX-"):

            try:
                numero = int(
                    texto.replace("PTRX-", "")
                )
                mayor = max(mayor, numero)

            except ValueError:
                pass

    return f"PTRX-{mayor + 1:06d}"


def generar_preguntas_81(cantidad):

    df = st.session_state.get(
        "df_disponible_81",
        pd.DataFrame()
    )

    if df.empty:
        return []

    columnas = [
        "Fuente_ID",
        "Producto",
        "Precaución / Contraindicación"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "8.1 ERROR: faltan columnas: "
            + ", ".join(faltantes)
        )

        return []

    consumidas = st.session_state.get(
        "fuentes_consumidas_precaucion_81",
        set()
    ).copy()

    candidatos = df[
        ~df["Fuente_ID"].isin(consumidas)
    ].copy()

    candidatos = candidatos[
        (candidatos["Producto"].astype(str).str.strip() != "")
        &
        (
            candidatos[
                "Precaución / Contraindicación"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
    ].copy()

    if len(candidatos) < 4:

        st.warning(
            "No hay suficientes relaciones disponibles "
            "para generar preguntas de precaución."
        )

        return []

    candidatos = candidatos.sample(
        frac=1
    ).reset_index(drop=True)

    preguntas = []

    for _, fila in candidatos.iterrows():

        producto = str(
            fila["Producto"]
        ).strip()

        correcta = str(
            fila[
                "Precaución / Contraindicación"
            ]
        ).strip()

        fuente = str(
            fila["Fuente_ID"]
        ).strip()

        falsas = candidatos[
            candidatos["Fuente_ID"] != fuente
        ].copy()

        falsas["respuesta"] = (
            falsas[
                "Precaución / Contraindicación"
            ]
            .astype(str)
            .str.strip()
        )

        falsas = falsas[
            falsas["respuesta"].str.lower()
            != correcta.lower()
        ]

        falsas = falsas.drop_duplicates(
            subset="respuesta"
        )

        if len(falsas) < 3:
            continue

        falsas = falsas.sample(
            n=3
        )

        opciones = [
            correcta,
            str(falsas.iloc[0]["respuesta"]).strip(),
            str(falsas.iloc[1]["respuesta"]).strip(),
            str(falsas.iloc[2]["respuesta"]).strip()
        ]

        np.random.shuffle(opciones)

        correcta_numero = (
            opciones.index(correcta) + 1
        )

        pregunta = {

            "Pregunta_ID":
                siguiente_id_81(),

            "Modulo":
                "Restricciones",

            "Tema":
                "Precaución / Contraindicación",

            "Nivel":
                "Nivel 1",

            "Tipo_Relacion":
                "Producto-Precaución/Contraindicación",

            "Pregunta":
                "Para el producto "
                f"{producto}, "
                "¿cuál de las siguientes corresponde "
                "a una precaución o contraindicación "
                "para su uso?",

            "Respuesta_1":
                opciones[0],

            "Respuesta_2":
                opciones[1],

            "Respuesta_3":
                opciones[2],

            "Respuesta_4":
                opciones[3],

            "Respuesta_Correcta":
                str(correcta_numero),

            "Estado":
                "PENDIENTE",

            "Observacion_Administrador":
                "",

            "Fecha_Generacion":
                pd.Timestamp.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

            "Fuente_ID":
                fuente
        }

        preguntas.append(pregunta)

        consumidas.add(fuente)

        if len(preguntas) >= int(cantidad):
            break

    st.session_state[
        "fuentes_consumidas_precaucion_81"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ DEL GENERADOR
# ============================================================

if "df_disponible_81" in st.session_state:

    st.markdown(
        "### 8.1 - Generador de preguntas"
    )

    st.info(
        "Nivel 1: Producto → Precaución / "
        "Contraindicación. La celda completa se "
        "conserva como una sola respuesta."
    )

    cantidad_81 = st.number_input(
        "Cantidad de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_precaucion_81"
    )

    if st.button(
        "GENERAR PREGUNTAS 8.1",
        key="generar_precaucion_81"
    ):

        nuevas_81 = generar_preguntas_81(
            int(cantidad_81)
        )

        st.session_state[
            "preguntas_generadas_81"
        ] = nuevas_81

        if nuevas_81:

            st.success(
                f"Se generaron {len(nuevas_81)} preguntas."
            )

        else:

            st.warning(
                "No fue posible generar preguntas "
                "con las relaciones disponibles."
            )


# ============================================================
# MOSTRAR PREGUNTAS
# ============================================================

preguntas_81 = st.session_state.get(
    "preguntas_generadas_81",
    []
)

if preguntas_81:

    st.markdown(
        "### Preguntas generadas 8.1"
    )

    for pregunta in preguntas_81:

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
# 8.2 - PARTE 3
# VALIDADOR DE PREGUNTAS
# PRODUCTO / CONTRAINDICACIÓN - MOTIVO
# NIVEL 1
#
# ENTRADA:
#     preguntas_generadas_83
#
# SALIDA / BANCO:
#     df_banco_83
# ============================================================


preguntas_83 = st.session_state.get(
    "preguntas_generadas_83",
    []
)


if preguntas_83:

    st.markdown(
        "### 8.2 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta y apruebe o rechace "
        "individualmente."
    )

    for i, pregunta in enumerate(preguntas_83):

        pregunta_id = pregunta.get(
            "Pregunta_ID",
            f"PTRM-{i + 1:06d}"
        )

        st.markdown(
            f"#### {pregunta_id}"
        )

        # ----------------------------------------------------
        # PREGUNTA
        # ----------------------------------------------------

        st.write(
            pregunta.get(
                "Pregunta",
                ""
            )
        )

        # ----------------------------------------------------
        # OPCIONES
        # ----------------------------------------------------

        st.write(
            f"1. {pregunta.get('Respuesta_1', '')}"
        )

        st.write(
            f"2. {pregunta.get('Respuesta_2', '')}"
        )

        st.write(
            f"3. {pregunta.get('Respuesta_3', '')}"
        )

        st.write(
            f"4. {pregunta.get('Respuesta_4', '')}"
        )

        # ----------------------------------------------------
        # RESPUESTA CORRECTA
        # ----------------------------------------------------

        st.caption(
            "Respuesta correcta: "
            + str(
                pregunta.get(
                    "Respuesta_Correcta",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # FUENTE
        # ----------------------------------------------------

        st.caption(
            "Fuente: "
            + str(
                pregunta.get(
                    "Fuente_ID",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # ESTADO
        # ----------------------------------------------------

        estado = pregunta.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"Estado: {estado}"
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
            key=f"observacion_restricciones_83_{i}"
        )

        col1, col2 = st.columns(2)

        # ====================================================
        # APROBAR
        # ====================================================

        with col1:

            if st.button(
                "APROBAR",
                key=f"aprobar_restricciones_83_{i}"
            ):

                preguntas_83[i]["Estado"] = (
                    "APROBADA"
                )

                preguntas_83[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_83"
                ] = preguntas_83

                # Actualizar DataFrame 8.3
                st.session_state[
                    "df_banco_83"
                ] = pd.DataFrame(
                    preguntas_83
                )

                st.rerun()

        # ====================================================
        # RECHAZAR
        # ====================================================

        with col2:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_restricciones_83_{i}"
            ):

                preguntas_83[i]["Estado"] = (
                    "RECHAZADA"
                )

                preguntas_83[i][
                    "Observacion_Administrador"
                ] = observacion

                st.session_state[
                    "preguntas_generadas_83"
                ] = preguntas_83

                # Actualizar DataFrame 8.3
                st.session_state[
                    "df_banco_83"
                ] = pd.DataFrame(
                    preguntas_83
                )

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 8.2
# ============================================================

if preguntas_83:

    aprobadas_83 = sum(
        1
        for pregunta in preguntas_83
        if pregunta.get(
            "Estado"
        ) == "APROBADA"
    )

    rechazadas_83 = sum(
        1
        for pregunta in preguntas_83
        if pregunta.get(
            "Estado"
        ) == "RECHAZADA"
    )

    pendientes_83 = sum(
        1
        for pregunta in preguntas_83
        if pregunta.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 8.2"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Aprobadas",
            aprobadas_83
        )

    with col2:

        st.metric(
            "Rechazadas",
            rechazadas_83
        )

    with col3:

        st.metric(
            "Pendientes",
            pendientes_83
        )

    if pendientes_83 == 0:

        st.success(
            "Todas las preguntas fueron revisadas."
        )
# ============================================================
# ============================================================
# 8.1 - PARTE 4
# SINCRONIZADOR
# ============================================================

GITHUB_USUARIO_81 = "franquiciasauces"
GITHUB_REPOSITORIO_81 = "Asesores"
GITHUB_RAMA_81 = "main"
GITHUB_ARCHIVO_81 = "BANCO_PREGUNTAS_GENERALES.xlsx"

URL_GITHUB_81 = (
    "https://api.github.com/repos/"
    + GITHUB_USUARIO_81
    + "/"
    + GITHUB_REPOSITORIO_81
    + "/contents/"
    + GITHUB_ARCHIVO_81
)


def sincronizar_81():

    preguntas_81 = st.session_state.get(
        "preguntas_generadas_81",
        []
    )

    aprobadas_81 = [
        pregunta
        for pregunta in preguntas_81
        if pregunta.get("Estado") == "APROBADA"
    ]

    if not aprobadas_81:

        st.warning(
            "No hay preguntas aprobadas para sincronizar."
        )

        return

    headers_81 = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    try:

        # ----------------------------------------------------
        # LEER BANCO ACTUAL DE GITHUB
        # ----------------------------------------------------

        solicitud_81 = urllib.request.Request(
            URL_GITHUB_81,
            headers=headers_81,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud_81,
            timeout=30
        ) as respuesta_81:

            datos_81 = json.loads(
                respuesta_81.read().decode("utf-8")
            )

        if "sha" not in datos_81:

            st.error(
                "8.1 ERROR: GitHub no devolvió el SHA "
                "del archivo."
            )

            return

        sha_81 = datos_81["sha"]

        if "content" not in datos_81:

            st.error(
                "8.1 ERROR: GitHub no devolvió el "
                "contenido del banco."
            )

            return

        contenido_81 = base64.b64decode(
            datos_81["content"].replace(
                "\n",
                ""
            )
        )

        memoria_81 = io.BytesIO(
            contenido_81
        )

        df_banco_81 = pd.read_excel(
            memoria_81,
            engine="openpyxl"
        )

        total_antes_81 = len(
            df_banco_81
        )

        # ----------------------------------------------------
        # MOSTRAR QUE EL BANCO SI FUE CARGADO
        # ----------------------------------------------------

        st.success(
            "BANCO_PREGUNTAS_GENERALES.xlsx "
            "cargado correctamente desde GitHub."
        )

        st.info(
            f"Preguntas en el banco antes de sincronizar: "
            f"{total_antes_81:,}"
        )

        # ----------------------------------------------------
        # COLUMNAS REQUERIDAS
        # ----------------------------------------------------

        columnas_81 = [
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

        faltantes_banco_81 = [
            columna
            for columna in columnas_81
            if columna not in df_banco_81.columns
        ]

        if faltantes_banco_81:

            st.error(
                "8.1 ERROR: faltan columnas en "
                "BANCO_PREGUNTAS_GENERALES.xlsx: "
                + ", ".join(
                    faltantes_banco_81
                )
            )

            return

        df_nuevas_81 = pd.DataFrame(
            aprobadas_81
        )

        faltantes_nuevas_81 = [
            columna
            for columna in columnas_81
            if columna not in df_nuevas_81.columns
        ]

        if faltantes_nuevas_81:

            st.error(
                "8.1 ERROR: faltan columnas en "
                "las preguntas aprobadas: "
                + ", ".join(
                    faltantes_nuevas_81
                )
            )

            return

        df_nuevas_81 = df_nuevas_81[
            columnas_81
        ].copy()

        # ----------------------------------------------------
        # EVITAR DUPLICADOS POR PREGUNTA_ID
        # ----------------------------------------------------

        ids_existentes_81 = set(
            df_banco_81[
                "Pregunta_ID"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        filas_nuevas_81 = []

        for _, fila_81 in df_nuevas_81.iterrows():

            pregunta_id_81 = str(
                fila_81["Pregunta_ID"]
            ).strip()

            if not pregunta_id_81:

                continue

            if pregunta_id_81 in ids_existentes_81:

                continue

            filas_nuevas_81.append(
                fila_81
            )

            ids_existentes_81.add(
                pregunta_id_81
            )

        nuevas_81 = len(
            filas_nuevas_81
        )

        # ----------------------------------------------------
        # SI NO HAY NADA NUEVO
        # ----------------------------------------------------

        if nuevas_81 == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                f"El banco continúa con "
                f"{total_antes_81:,} preguntas."
            )

            return

        # ----------------------------------------------------
        # AGREGAR PREGUNTAS
        # ----------------------------------------------------

        df_agregar_81 = pd.DataFrame(
            filas_nuevas_81
        )

        df_final_81 = pd.concat(
            [
                df_banco_81,
                df_agregar_81
            ],
            ignore_index=True
        )

        total_despues_81 = len(
            df_final_81
        )

        # ----------------------------------------------------
        # CREAR EXCEL
        # ----------------------------------------------------

        memoria_salida_81 = io.BytesIO()

        with pd.ExcelWriter(
            memoria_salida_81,
            engine="openpyxl"
        ) as escritor_81:

            df_final_81.to_excel(
                escritor_81,
                index=False,
                sheet_name="Banco"
            )

        contenido_nuevo_81 = base64.b64encode(
            memoria_salida_81.getvalue()
        ).decode("utf-8")

        # ----------------------------------------------------
        # ACTUALIZAR GITHUB
        # ----------------------------------------------------

        datos_actualizacion_81 = {
            "message":
                "Agregar preguntas 8.1 - Restricciones",
            "content":
                contenido_nuevo_81,
            "branch":
                GITHUB_RAMA_81,
            "sha":
                sha_81
        }

        cuerpo_81 = json.dumps(
            datos_actualizacion_81
        ).encode("utf-8")

        solicitud_actualizacion_81 = (
            urllib.request.Request(
                URL_GITHUB_81,
                data=cuerpo_81,
                headers={
                    **headers_81,
                    "Content-Type":
                        "application/json"
                },
                method="PUT"
            )
        )

        with urllib.request.urlopen(
            solicitud_actualizacion_81,
            timeout=30
        ) as respuesta_actualizacion_81:

            respuesta_actualizacion_81.read()

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        st.success(
            "8.1 sincronizado correctamente con GitHub."
        )

        st.info(
            f"Preguntas en el banco antes: "
            f"{total_antes_81:,}"
        )

        st.info(
            f"Preguntas nuevas incorporadas: "
            f"{nuevas_81:,}"
        )

        st.info(
            f"Preguntas en el banco después: "
            f"{total_despues_81:,}"
        )

    except Exception as error_81:

        st.error(
            "8.1 ERROR: no fue posible cargar o "
            "actualizar BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(
            error_81
        )


# ============================================================
# BOTÓN DE SINCRONIZACIÓN 8.1
# ============================================================

preguntas_81 = st.session_state.get(
    "preguntas_generadas_81",
    []
)

aprobadas_81 = sum(
    1
    for pregunta in preguntas_81
    if pregunta.get("Estado") == "APROBADA"
)

if aprobadas_81 > 0:

    st.markdown(
        "### 8.1 - Sincronización"
    )

    st.info(
        f"Preguntas aprobadas listas para sincronizar: "
        f"{aprobadas_81}"
    )

    if st.button(
        "SINCRONIZAR 8.1 CON BANCO DE PREGUNTAS",
        key="boton_sincronizar_81"
    ):

        sincronizar_81()

# ============================================================
# 8.2 - PARTE 2
# GENERADOR DE PREGUNTAS
# PRODUCTO / CONTRAINDICACIÓN - MOTIVO
# NIVEL 1
#
# FUENTE:
#     df_disponible_81
#
# SALIDA:
#     df_banco_83
# ============================================================


def siguiente_id_83():

    df_banco = st.session_state.get(
        "df_banco_83",
        pd.DataFrame()
    )

    mayor = 0

    if (
        not df_banco.empty
        and "Pregunta_ID" in df_banco.columns
    ):

        for valor in df_banco["Pregunta_ID"].fillna(""):

            texto = str(valor).strip()

            if texto.startswith("PTRM-"):

                try:

                    numero = int(
                        texto.replace("PTRM-", "")
                    )

                    mayor = max(
                        mayor,
                        numero
                    )

                except ValueError:
                    pass

    preguntas = st.session_state.get(
        "preguntas_generadas_83",
        []
    )

    for pregunta in preguntas:

        texto = str(
            pregunta.get("Pregunta_ID", "")
        ).strip()

        if texto.startswith("PTRM-"):

            try:

                numero = int(
                    texto.replace("PTRM-", "")
                )

                mayor = max(
                    mayor,
                    numero
                )

            except ValueError:
                pass

    return f"PTRM-{mayor + 1:06d}"


# ============================================================
# GENERADOR 8.2 restriccion producto motivo
# ============================================================

def generar_preguntas_83(cantidad):

    # --------------------------------------------------------
    # LA FUENTE ES LA MISMA DEL GENERADOR 8.1
    # --------------------------------------------------------

    df = st.session_state.get(
        "df_disponible_81",
        pd.DataFrame()
    )

    if df.empty:
        return []

    # --------------------------------------------------------
    # COLUMNAS REQUERIDAS PARA ESTE GENERADOR
    # --------------------------------------------------------

    columnas = [
        "Restriccion_ID",
        "Producto",
        "Precaución / Contraindicación",
        "Motivo",
        "Alternativas seguras"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "8.2 ERROR: faltan columnas: "
            + ", ".join(faltantes)
        )

        return []

    # --------------------------------------------------------
    # FUENTES YA UTILIZADAS
    # --------------------------------------------------------

    consumidas = st.session_state.get(
        "fuentes_consumidas_83",
        set()
    ).copy()

    # --------------------------------------------------------
    # CANDIDATOS
    # --------------------------------------------------------

    candidatos = df[
        ~df["Restriccion_ID"].astype(str).isin(
            {
                str(x)
                for x in consumidas
            }
        )
    ].copy()

    # --------------------------------------------------------
    # VALIDACIÓN MÍNIMA
    #
    # NO se modifica el contenido.
    # Solo se eliminan filas realmente vacías.
    # --------------------------------------------------------

    candidatos = candidatos[
        (candidatos["Restriccion_ID"].astype(str).str.strip() != "")
        &
        (candidatos["Producto"].astype(str).str.strip() != "")
        &
        (
            candidatos[
                "Precaución / Contraindicación"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
        &
        (
            candidatos["Motivo"]
            .astype(str)
            .str.strip()
            != ""
        )
    ].copy()

    if candidatos.empty:

        st.warning(
            "No hay relaciones de restricciones "
            "disponibles para generar preguntas."
        )

        return []

    # --------------------------------------------------------
    # ALEATORIZAR
    # --------------------------------------------------------

    candidatos = candidatos.sample(
        frac=1
    ).reset_index(drop=True)

    preguntas = []

    # ========================================================
    # GENERACIÓN
    # ========================================================

    for _, fila in candidatos.iterrows():

        fuente = str(
            fila["Restriccion_ID"]
        ).strip()

        producto = str(
            fila["Producto"]
        ).strip()

        # ----------------------------------------------------
        # IMPORTANTE:
        # SE CONSERVA TODA LA CELDA.
        #
        # NO se divide por ;
        # NO se resume;
        # NO se elimina información.
        # ----------------------------------------------------

        contraindicacion = str(
            fila[
                "Precaución / Contraindicación"
            ]
        ).strip()

        correcta = str(
            fila["Motivo"]
        ).strip()

        # ----------------------------------------------------
        # BUSCAR MOTIVOS FALSOS
        # ----------------------------------------------------

        falsas = candidatos[
            candidatos["Restriccion_ID"].astype(str)
            != fuente
        ].copy()

        falsas["respuesta"] = (
            falsas["Motivo"]
            .astype(str)
            .str.strip()
        )

        # Quitar el motivo correcto
        falsas = falsas[
            falsas["respuesta"].str.lower()
            != correcta.lower()
        ]

        # Quitar motivos duplicados
        falsas = falsas.drop_duplicates(
            subset="respuesta"
        )

        if len(falsas) < 3:
            continue

        falsas = falsas.sample(
            n=3
        )

        # ----------------------------------------------------
        # CUATRO OPCIONES
        # ----------------------------------------------------

        opciones = [
            correcta,
            str(
                falsas.iloc[0]["respuesta"]
            ).strip(),
            str(
                falsas.iloc[1]["respuesta"]
            ).strip(),
            str(
                falsas.iloc[2]["respuesta"]
            ).strip()
        ]

        np.random.shuffle(
            opciones
        )

        correcta_numero = (
            opciones.index(correcta) + 1
        )

        # ====================================================
        # PREGUNTA
        # ====================================================

        pregunta = {

            "Pregunta_ID":
                siguiente_id_83(),

            "Modulo":
                "Restricciones",

            "Tema":
                "Producto / Contraindicación - Motivo",

            "Nivel":
                "Nivel 1",

            "Tipo_Relacion":
                "Producto-Precaución/Contraindicación-Motivo",

            "Pregunta":
                f"{producto} puede generar "
                "contraindicaciones asociadas a "
                f"{contraindicacion} "
                "debido a que:",

            "Respuesta_1":
                opciones[0],

            "Respuesta_2":
                opciones[1],

            "Respuesta_3":
                opciones[2],

            "Respuesta_4":
                opciones[3],

            "Respuesta_Correcta":
                str(correcta_numero),

            "Estado":
                "PENDIENTE",

            "Observacion_Administrador":
                "",

            "Fecha_Generacion":
                pd.Timestamp.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

            "Fuente_ID":
                fuente
        }

        preguntas.append(
            pregunta
        )

        consumidas.add(
            fuente
        )

        if len(preguntas) >= int(cantidad):
            break

    # --------------------------------------------------------
    # PERSISTIR FUENTES CONSUMIDAS
    # --------------------------------------------------------

    st.session_state[
        "fuentes_consumidas_83"
    ] = consumidas

    return preguntas


# ============================================================
# INTERFAZ 8.2
# ============================================================

if "df_disponible_81" in st.session_state:

    st.markdown(
        "### 8.2 - Generador de preguntas "
        "Producto / Contraindicación - Motivo"
    )

    st.info(
        "Nivel 1: se conserva completa la "
        "Precaución / Contraindicación de la fuente "
        "y se pregunta por el Motivo."
    )

    cantidad_83 = st.number_input(
        "Cantidad de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_83"
    )

    if st.button(
        "GENERAR PREGUNTAS 8.2",
        key="generar_preguntas_83"
    ):

        nuevas_83 = generar_preguntas_83(
            int(cantidad_83)
        )

        st.session_state[
            "preguntas_generadas_83"
        ] = nuevas_83

        # ----------------------------------------------------
        # CREAR EL DATAFRAME 8.3
        # ----------------------------------------------------

        if nuevas_83:

            st.session_state[
                "df_banco_83"
            ] = pd.DataFrame(
                nuevas_83
            )

            st.success(
                f"Se generaron "
                f"{len(nuevas_83)} "
                "preguntas."
            )

        else:

            st.warning(
                "No fue posible generar preguntas "
                "con las relaciones disponibles."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_83 = st.session_state.get(
    "preguntas_generadas_83",
    []
)

if preguntas_83:

    st.markdown(
        "### Preguntas generadas 8.2"
    )

    for pregunta in preguntas_83:

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
# 8.2 - PARTE 4
# SINCRONIZADOR
# PRODUCTO / CONTRAINDICACIÓN - MOTIVO
# NIVEL 1
#
# ENTRADA:
#     preguntas_generadas_83
#
# BANCO:
#     BANCO_PREGUNTAS_GENERALES.xlsx
#
# DATAFRAME:
#     df_banco_83
# ============================================================


GITHUB_USUARIO_83 = "franquiciasauces"
GITHUB_REPOSITORIO_83 = "Asesores"
GITHUB_RAMA_83 = "main"
GITHUB_ARCHIVO_83 = "BANCO_PREGUNTAS_GENERALES.xlsx"

URL_GITHUB_83 = (
    "https://api.github.com/repos/"
    + GITHUB_USUARIO_83
    + "/"
    + GITHUB_REPOSITORIO_83
    + "/contents/"
    + GITHUB_ARCHIVO_83
)


# ============================================================
# FUNCIÓN DE SINCRONIZACIÓN
# ============================================================

def sincronizar_83():

    preguntas_83 = st.session_state.get(
        "preguntas_generadas_83",
        []
    )

    # --------------------------------------------------------
    # SOLO PREGUNTAS APROBADAS
    # --------------------------------------------------------

    aprobadas_83 = [
        pregunta
        for pregunta in preguntas_83
        if pregunta.get("Estado") == "APROBADA"
    ]

    if not aprobadas_83:

        st.warning(
            "No hay preguntas aprobadas para sincronizar."
        )

        return

    headers_83 = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    try:

        # ====================================================
        # LEER BANCO GENERAL ACTUAL DESDE GITHUB
        # ====================================================

        solicitud_83 = urllib.request.Request(
            URL_GITHUB_83,
            headers=headers_83,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud_83,
            timeout=30
        ) as respuesta_83:

            datos_83 = json.loads(
                respuesta_83.read().decode("utf-8")
            )

        # ----------------------------------------------------
        # VERIFICAR SHA
        # ----------------------------------------------------

        if "sha" not in datos_83:

            st.error(
                "8.2 ERROR: GitHub no devolvió el SHA "
                "del archivo."
            )

            return

        sha_83 = datos_83["sha"]

        # ----------------------------------------------------
        # VERIFICAR CONTENIDO
        # ----------------------------------------------------

        if "content" not in datos_83:

            st.error(
                "8.2 ERROR: GitHub no devolvió el "
                "contenido del banco."
            )

            return

        # ====================================================
        # DECODIFICAR BANCO
        # ====================================================

        contenido_83 = base64.b64decode(
            datos_83["content"].replace(
                "\n",
                ""
            )
        )

        memoria_83 = io.BytesIO(
            contenido_83
        )

        df_banco_83_github = pd.read_excel(
            memoria_83,
            engine="openpyxl"
        )

        total_antes_83 = len(
            df_banco_83_github
        )

        # ----------------------------------------------------
        # CONFIRMACIÓN DE CARGA
        # ----------------------------------------------------

        st.success(
            "BANCO_PREGUNTAS_GENERALES.xlsx "
            "cargado correctamente desde GitHub."
        )

        st.info(
            "Preguntas en el banco antes de sincronizar: "
            f"{total_antes_83:,}"
        )

        # ====================================================
        # COLUMNAS REQUERIDAS
        # ====================================================

        columnas_83 = [
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
        # VALIDAR BANCO GENERAL
        # ----------------------------------------------------

        faltantes_banco_83 = [
            columna
            for columna in columnas_83
            if columna not in df_banco_83_github.columns
        ]

        if faltantes_banco_83:

            st.error(
                "8.2 ERROR: faltan columnas en "
                "BANCO_PREGUNTAS_GENERALES.xlsx: "
                + ", ".join(
                    faltantes_banco_83
                )
            )

            return

        # ====================================================
        # PREPARAR PREGUNTAS APROBADAS
        # ====================================================

        df_nuevas_83 = pd.DataFrame(
            aprobadas_83
        )

        faltantes_nuevas_83 = [
            columna
            for columna in columnas_83
            if columna not in df_nuevas_83.columns
        ]

        if faltantes_nuevas_83:

            st.error(
                "8.2 ERROR: faltan columnas en "
                "las preguntas aprobadas: "
                + ", ".join(
                    faltantes_nuevas_83
                )
            )

            return

        # ----------------------------------------------------
        # RESPETAR ORDEN DEL BANCO GENERAL
        # ----------------------------------------------------

        df_nuevas_83 = df_nuevas_83[
            columnas_83
        ].copy()

        # ====================================================
        # EVITAR DUPLICADOS POR PREGUNTA_ID
        # ====================================================

        ids_existentes_83 = set(
            df_banco_83_github[
                "Pregunta_ID"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        filas_nuevas_83 = []

        for _, fila_83 in df_nuevas_83.iterrows():

            pregunta_id_83 = str(
                fila_83["Pregunta_ID"]
            ).strip()

            if not pregunta_id_83:
                continue

            if pregunta_id_83 in ids_existentes_83:
                continue

            filas_nuevas_83.append(
                fila_83
            )

            ids_existentes_83.add(
                pregunta_id_83
            )

        nuevas_83 = len(
            filas_nuevas_83
        )

        # ====================================================
        # SI NO HAY PREGUNTAS NUEVAS
        # ====================================================

        if nuevas_83 == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                "El banco continúa con "
                f"{total_antes_83:,} preguntas."
            )

            return

        # ====================================================
        # AGREGAR AL BANCO GENERAL
        # ====================================================

        df_agregar_83 = pd.DataFrame(
            filas_nuevas_83
        )

        df_final_83 = pd.concat(
            [
                df_banco_83_github,
                df_agregar_83
            ],
            ignore_index=True
        )

        total_despues_83 = len(
            df_final_83
        )

        # ----------------------------------------------------
        # ACTUALIZAR DATAFRAME DEL 8.3
        # ----------------------------------------------------

        st.session_state[
            "df_banco_83"
        ] = df_final_83.copy()

        # ====================================================
        # CREAR EXCEL ACTUALIZADO
        # ====================================================

        memoria_salida_83 = io.BytesIO()

        with pd.ExcelWriter(
            memoria_salida_83,
            engine="openpyxl"
        ) as escritor_83:

            df_final_83.to_excel(
                escritor_83,
                index=False,
                sheet_name="Banco"
            )

        # ====================================================
        # CODIFICAR ARCHIVO
        # ====================================================

        contenido_nuevo_83 = base64.b64encode(
            memoria_salida_83.getvalue()
        ).decode("utf-8")

        # ====================================================
        # ACTUALIZAR GITHUB
        # ====================================================

        datos_actualizacion_83 = {

            "message":
                "Agregar preguntas 8.2 - "
                "Producto Contraindicación Motivo",

            "content":
                contenido_nuevo_83,

            "branch":
                GITHUB_RAMA_83,

            "sha":
                sha_83
        }

        cuerpo_83 = json.dumps(
            datos_actualizacion_83
        ).encode("utf-8")

        solicitud_actualizacion_83 = (
            urllib.request.Request(
                URL_GITHUB_83,
                data=cuerpo_83,
                headers={
                    **headers_83,
                    "Content-Type":
                        "application/json"
                },
                method="PUT"
            )
        )

        with urllib.request.urlopen(
            solicitud_actualizacion_83,
            timeout=30
        ) as respuesta_actualizacion_83:

            respuesta_actualizacion_83.read()

        # ====================================================
        # RESULTADO
        # ====================================================

        st.success(
            "8.2 sincronizado correctamente con "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.info(
            "Preguntas en el banco antes: "
            f"{total_antes_83:,}"
        )

        st.info(
            "Preguntas nuevas incorporadas: "
            f"{nuevas_83:,}"
        )

        st.info(
            "Preguntas en el banco después: "
            f"{total_despues_83:,}"
        )

    except Exception as error_83:

        st.error(
            "8.2 ERROR: no fue posible cargar o "
            "actualizar BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(
            error_83
        )


# ============================================================
# BOTÓN DE SINCRONIZACIÓN 8.2
# ============================================================

preguntas_83 = st.session_state.get(
    "preguntas_generadas_83",
    []
)

aprobadas_83 = sum(
    1
    for pregunta in preguntas_83
    if pregunta.get("Estado") == "APROBADA"
)

if aprobadas_83 > 0:

    st.markdown(
        "### 8.2 - Sincronización"
    )

    st.info(
        "Preguntas aprobadas listas para sincronizar: "
        f"{aprobadas_83}"
    )

    if st.button(
        "SINCRONIZAR 8.2 CON BANCO DE PREGUNTAS",
        key="boton_sincronizar_83"
    ):

        sincronizar_83()
# ============================================================
# 8.3 - PARTE 2
# GENERADOR DE PREGUNTAS
# PRODUCTO / ALTERNATIVAS SEGURAS
# NIVEL 1
#
# FUENTE:
#     df_disponible_81
#
# SALIDA:
#     preguntas_generadas_85
#     df_banco_85
# ============================================================


def siguiente_id_85():

    df_banco_85 = st.session_state.get(
        "df_banco_85",
        pd.DataFrame()
    )

    mayor_85 = 0

    if (
        not df_banco_85.empty
        and "Pregunta_ID" in df_banco_85.columns
    ):

        for valor_85 in df_banco_85[
            "Pregunta_ID"
        ].fillna(""):

            texto_85 = str(
                valor_85
            ).strip()

            if texto_85.startswith("PTRS-"):

                try:

                    numero_85 = int(
                        texto_85.replace(
                            "PTRS-",
                            ""
                        )
                    )

                    mayor_85 = max(
                        mayor_85,
                        numero_85
                    )

                except ValueError:
                    pass

    preguntas_85 = st.session_state.get(
        "preguntas_generadas_85",
        []
    )

    for pregunta_85 in preguntas_85:

        texto_85 = str(
            pregunta_85.get(
                "Pregunta_ID",
                ""
            )
        ).strip()

        if texto_85.startswith("PTRS-"):

            try:

                numero_85 = int(
                    texto_85.replace(
                        "PTRS-",
                        ""
                    )
                )

                mayor_85 = max(
                    mayor_85,
                    numero_85
                )

            except ValueError:
                pass

    return f"PTRS-{mayor_85 + 1:06d}"


# ============================================================
# FUNCIÓN GENERADORA
# ============================================================

def generar_preguntas_85(cantidad_85):

    # --------------------------------------------------------
    # FUENTE COMÚN DE LOS TRES GENERADORES
    # --------------------------------------------------------

    df_85 = st.session_state.get(
        "df_disponible_81",
        pd.DataFrame()
    )

    if df_85.empty:

        st.warning(
            "No existen restricciones disponibles "
            "para generar preguntas 8.3."
        )

        return []

    # --------------------------------------------------------
    # COLUMNAS REQUERIDAS
    # --------------------------------------------------------

    columnas_85 = [
        "Restriccion_ID",
        "Producto",
        "Alternativas seguras"
    ]

    faltantes_85 = [
        columna_85
        for columna_85 in columnas_85
        if columna_85 not in df_85.columns
    ]

    if faltantes_85:

        st.error(
            "8.3 ERROR: faltan columnas: "
            + ", ".join(faltantes_85)
        )

        return []

    # --------------------------------------------------------
    # FUENTES YA UTILIZADAS POR 8.3
    #
    # ES INDEPENDIENTE DE 8.1 Y 8.2
    # --------------------------------------------------------

    consumidas_85 = st.session_state.get(
        "fuentes_consumidas_85",
        set()
    ).copy()

    # --------------------------------------------------------
    # CANDIDATOS
    # --------------------------------------------------------

    candidatos_85 = df_85[
        ~df_85[
            "Restriccion_ID"
        ].astype(str).isin(
            {
                str(x)
                for x in consumidas_85
            }
        )
    ].copy()

    # --------------------------------------------------------
    # VALIDAR FILAS
    #
    # NO SE RESUME NI SE MODIFICA EL CONTENIDO.
    # --------------------------------------------------------

    candidatos_85 = candidatos_85[
        (
            candidatos_85[
                "Restriccion_ID"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
        &
        (
            candidatos_85[
                "Producto"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
        &
        (
            candidatos_85[
                "Alternativas seguras"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
    ].copy()

    if candidatos_85.empty:

        st.warning(
            "No hay relaciones disponibles para "
            "generar preguntas 8.3."
        )

        return []

    candidatos_85 = candidatos_85.sample(
        frac=1
    ).reset_index(drop=True)

    preguntas_85 = []

    # ========================================================
    # GENERAR
    # ========================================================

    for _, fila_85 in candidatos_85.iterrows():

        fuente_85 = str(
            fila_85[
                "Restriccion_ID"
            ]
        ).strip()

        producto_85 = str(
            fila_85[
                "Producto"
            ]
        ).strip()

        correcta_85 = str(
            fila_85[
                "Alternativas seguras"
            ]
        ).strip()

        # ----------------------------------------------------
        # BUSCAR DISTRACTORES
        # ----------------------------------------------------

        falsas_85 = candidatos_85[
            candidatos_85[
                "Restriccion_ID"
            ].astype(str)
            != fuente_85
        ].copy()

        falsas_85["respuesta_85"] = (
            falsas_85[
                "Alternativas seguras"
            ]
            .astype(str)
            .str.strip()
        )

        # ----------------------------------------------------
        # REGLA FUNDAMENTAL:
        #
        # SI LA ALTERNATIVA SEGURA DEL DISTRACTOR ES IGUAL
        # A LA CORRECTA, NO SE PUEDE UTILIZAR.
        # ----------------------------------------------------

        falsas_85 = falsas_85[
            falsas_85[
                "respuesta_85"
            ].str.casefold()
            != correcta_85.casefold()
        ]

        # ----------------------------------------------------
        # UNA MISMA ALTERNATIVA NO PUEDE REPETIRSE
        # COMO DISTRACTOR
        # ----------------------------------------------------

        falsas_85 = falsas_85.drop_duplicates(
            subset="respuesta_85"
        )

        # ----------------------------------------------------
        # NECESITAMOS 3 DISTRACTORES
        # ----------------------------------------------------

        if len(falsas_85) < 3:
            continue

        falsas_85 = falsas_85.sample(
            n=3
        )

        opciones_85 = [
            correcta_85,
            str(
                falsas_85.iloc[0][
                    "respuesta_85"
                ]
            ).strip(),
            str(
                falsas_85.iloc[1][
                    "respuesta_85"
                ]
            ).strip(),
            str(
                falsas_85.iloc[2][
                    "respuesta_85"
                ]
            ).strip()
        ]

        # ----------------------------------------------------
        # VERIFICACIÓN FINAL DE OPCIONES
        # ----------------------------------------------------

        opciones_normalizadas_85 = [
            opcion.casefold()
            for opcion in opciones_85
        ]

        if len(
            set(opciones_normalizadas_85)
        ) != 4:

            continue

        # ----------------------------------------------------
        # ALEATORIZAR
        # ----------------------------------------------------

        np.random.shuffle(
            opciones_85
        )

        correcta_numero_85 = (
            opciones_85.index(
                correcta_85
            ) + 1
        )

        # ====================================================
        # CREAR PREGUNTA
        # ====================================================

        pregunta_85 = {

            "Pregunta_ID":
                siguiente_id_85(),

            "Modulo":
                "Restricciones",

            "Tema":
                "Producto / Alternativas seguras",

            "Nivel":
                "Nivel 1",

            "Tipo_Relacion":
                "Producto-Alternativas seguras",

            "Pregunta":
                f"Para el producto "
                f"{producto_85}, "
                "¿cuál de las siguientes corresponde "
                "a una alternativa segura?",

            "Respuesta_1":
                opciones_85[0],

            "Respuesta_2":
                opciones_85[1],

            "Respuesta_3":
                opciones_85[2],

            "Respuesta_4":
                opciones_85[3],

            "Respuesta_Correcta":
                str(
                    correcta_numero_85
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
                fuente_85
        }

        preguntas_85.append(
            pregunta_85
        )

        consumidas_85.add(
            fuente_85
        )

        if len(
            preguntas_85
        ) >= int(cantidad_85):

            break

    # --------------------------------------------------------
    # GUARDAR FUENTES CONSUMIDAS
    # --------------------------------------------------------

    st.session_state[
        "fuentes_consumidas_85"
    ] = consumidas_85

    return preguntas_85


# ============================================================
# INTERFAZ 8.3
# ============================================================

if "df_disponible_81" in st.session_state:

    st.markdown(
        "## 8.3 - Generador de preguntas"
    )

    st.info(
        "Nivel 1: Producto → Alternativas seguras. "
        "La información completa de la columna "
        "'Alternativas seguras' se conserva sin resumir."
    )

    cantidad_85 = st.number_input(
        "Cantidad de preguntas 8.3",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_85"
    )

    if st.button(
        "GENERAR PREGUNTAS 8.3",
        key="generar_preguntas_85"
    ):

        nuevas_85 = generar_preguntas_85(
            int(cantidad_85)
        )

        st.session_state[
            "preguntas_generadas_85"
        ] = nuevas_85

        if nuevas_85:

            st.session_state[
                "df_banco_85"
            ] = pd.DataFrame(
                nuevas_85
            )

            st.success(
                f"Se generaron "
                f"{len(nuevas_85)} "
                "preguntas 8.3."
            )

        else:

            st.warning(
                "No fue posible generar preguntas 8.3 "
                "con las relaciones disponibles."
            )


# ============================================================
# MOSTRAR PREGUNTAS
# ============================================================

preguntas_85 = st.session_state.get(
    "preguntas_generadas_85",
    []
)

if preguntas_85:

    st.markdown(
        "### Preguntas generadas 8.3"
    )

    for pregunta_85 in preguntas_85:

        st.markdown(
            f"**{pregunta_85['Pregunta_ID']} — "
            f"{pregunta_85['Nivel']}**"
        )

        st.write(
            pregunta_85["Pregunta"]
        )

        st.write(
            f"1. {pregunta_85['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta_85['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta_85['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta_85['Respuesta_4']}"
        )

        st.caption(
            "Respuesta correcta: "
            f"{pregunta_85['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta_85['Fuente_ID']}"
        )

        st.divider()
# ============================================================
# 8.3 - PARTE 3
# VALIDADOR DE PREGUNTAS
# PRODUCTO / ALTERNATIVAS SEGURAS
# NIVEL 1
#
# ENTRADA:
#     preguntas_generadas_85
#
# SALIDA:
#     df_banco_85
# ============================================================


preguntas_85 = st.session_state.get(
    "preguntas_generadas_85",
    []
)


# ============================================================
# VALIDACIÓN INDIVIDUAL
# ============================================================

if preguntas_85:

    st.markdown(
        "### 8.3 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta y apruebe o rechace "
        "individualmente."
    )

    for i, pregunta_85 in enumerate(
        preguntas_85
    ):

        pregunta_id_85 = pregunta_85.get(
            "Pregunta_ID",
            f"PTRS-{i + 1:06d}"
        )

        st.markdown(
            f"#### {pregunta_id_85}"
        )

        # ----------------------------------------------------
        # PREGUNTA
        # ----------------------------------------------------

        st.write(
            pregunta_85.get(
                "Pregunta",
                ""
            )
        )

        # ----------------------------------------------------
        # OPCIONES
        # ----------------------------------------------------

        st.write(
            f"1. {pregunta_85.get('Respuesta_1', '')}"
        )

        st.write(
            f"2. {pregunta_85.get('Respuesta_2', '')}"
        )

        st.write(
            f"3. {pregunta_85.get('Respuesta_3', '')}"
        )

        st.write(
            f"4. {pregunta_85.get('Respuesta_4', '')}"
        )

        # ----------------------------------------------------
        # RESPUESTA CORRECTA
        # ----------------------------------------------------

        st.caption(
            "Respuesta correcta: "
            + str(
                pregunta_85.get(
                    "Respuesta_Correcta",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # FUENTE
        # ----------------------------------------------------

        st.caption(
            "Fuente: "
            + str(
                pregunta_85.get(
                    "Fuente_ID",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # ESTADO
        # ----------------------------------------------------

        estado_85 = pregunta_85.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"Estado: {estado_85}"
        )

        # ----------------------------------------------------
        # OBSERVACIÓN DEL ADMINISTRADOR
        # ----------------------------------------------------

        observacion_85 = st.text_input(
            "Observación del administrador",
            value=pregunta_85.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_alternativas_85_{i}"
        )

        col1_85, col2_85 = st.columns(2)

        # ====================================================
        # APROBAR
        # ====================================================

        with col1_85:

            if st.button(
                "APROBAR",
                key=f"aprobar_alternativas_85_{i}"
            ):

                preguntas_85[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_85[i][
                    "Observacion_Administrador"
                ] = observacion_85

                # --------------------------------------------
                # ACTUALIZAR LISTA
                # --------------------------------------------

                st.session_state[
                    "preguntas_generadas_85"
                ] = preguntas_85

                # --------------------------------------------
                # ACTUALIZAR DATAFRAME 8.3
                # --------------------------------------------

                st.session_state[
                    "df_banco_85"
                ] = pd.DataFrame(
                    preguntas_85
                )

                st.rerun()

        # ====================================================
        # RECHAZAR
        # ====================================================

        with col2_85:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_alternativas_85_{i}"
            ):

                preguntas_85[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_85[i][
                    "Observacion_Administrador"
                ] = observacion_85

                # --------------------------------------------
                # ACTUALIZAR LISTA
                # --------------------------------------------

                st.session_state[
                    "preguntas_generadas_85"
                ] = preguntas_85

                # --------------------------------------------
                # ACTUALIZAR DATAFRAME 8.3
                # --------------------------------------------

                st.session_state[
                    "df_banco_85"
                ] = pd.DataFrame(
                    preguntas_85
                )

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 8.3
# ============================================================

if preguntas_85:

    aprobadas_85 = sum(
        1
        for pregunta_85 in preguntas_85
        if pregunta_85.get(
            "Estado"
        ) == "APROBADA"
    )

    rechazadas_85 = sum(
        1
        for pregunta_85 in preguntas_85
        if pregunta_85.get(
            "Estado"
        ) == "RECHAZADA"
    )

    pendientes_85 = sum(
        1
        for pregunta_85 in preguntas_85
        if pregunta_85.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 8.3"
    )

    col1_85, col2_85, col3_85 = st.columns(3)

    with col1_85:

        st.metric(
            "Aprobadas",
            aprobadas_85
        )

    with col2_85:

        st.metric(
            "Rechazadas",
            rechazadas_85
        )

    with col3_85:

        st.metric(
            "Pendientes",
            pendientes_85
        )

    if pendientes_85 == 0:

        st.success(
            "Todas las preguntas 8.3 "
            "fueron revisadas."
        )
# ============================================================
# 8.3 - PARTE 4
# SINCRONIZADOR
# PRODUCTO / RESTRICCIÓN - ALTERNATIVAS SEGURAS
# NIVEL 1
#
# ENTRADA:
#     preguntas_generadas_85
#
# BANCO:
#     BANCO_PREGUNTAS_GENERALES.xlsx
#
# SALIDA:
#     BANCO_PREGUNTAS_GENERALES.xlsx
#
# DATAFRAME LOCAL:
#     df_banco_85
# ============================================================


GITHUB_USUARIO_85 = "franquiciasauces"
GITHUB_REPOSITORIO_85 = "Asesores"
GITHUB_RAMA_85 = "main"
GITHUB_ARCHIVO_85 = "BANCO_PREGUNTAS_GENERALES.xlsx"


URL_GITHUB_85 = (
    "https://api.github.com/repos/"
    + GITHUB_USUARIO_85
    + "/"
    + GITHUB_REPOSITORIO_85
    + "/contents/"
    + GITHUB_ARCHIVO_85
)


# ============================================================
# FUNCIÓN DE SINCRONIZACIÓN
# ============================================================

def sincronizar_85():

    preguntas_85 = st.session_state.get(
        "preguntas_generadas_85",
        []
    )

    # --------------------------------------------------------
    # SOLO PREGUNTAS APROBADAS
    # --------------------------------------------------------

    aprobadas_85 = [
        pregunta
        for pregunta in preguntas_85
        if pregunta.get("Estado") == "APROBADA"
    ]

    if not aprobadas_85:

        st.warning(
            "No hay preguntas aprobadas para sincronizar."
        )

        return

    headers_85 = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    try:

        # ====================================================
        # LEER BANCO GENERAL DESDE GITHUB
        # ====================================================

        solicitud_85 = urllib.request.Request(
            URL_GITHUB_85,
            headers=headers_85,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud_85,
            timeout=30
        ) as respuesta_85:

            datos_85 = json.loads(
                respuesta_85.read().decode("utf-8")
            )

        # ----------------------------------------------------
        # VERIFICAR SHA
        # ----------------------------------------------------

        if "sha" not in datos_85:

            st.error(
                "8.3 ERROR: GitHub no devolvió el SHA "
                "del archivo."
            )

            return

        sha_85 = datos_85["sha"]

        # ----------------------------------------------------
        # VERIFICAR CONTENIDO
        # ----------------------------------------------------

        if "content" not in datos_85:

            st.error(
                "8.3 ERROR: GitHub no devolvió el "
                "contenido del banco."
            )

            return

        # ====================================================
        # DECODIFICAR BANCO
        # ====================================================

        contenido_85 = base64.b64decode(
            datos_85["content"].replace(
                "\n",
                ""
            )
        )

        memoria_85 = io.BytesIO(
            contenido_85
        )

        df_banco_85 = pd.read_excel(
            memoria_85,
            engine="openpyxl"
        )

        total_antes_85 = len(
            df_banco_85
        )

        # ====================================================
        # CONFIRMAR CARGA DEL BANCO
        # ====================================================

        st.success(
            "BANCO_PREGUNTAS_GENERALES.xlsx "
            "cargado correctamente desde GitHub."
        )

        st.info(
            "Preguntas en el banco antes de sincronizar: "
            f"{total_antes_85:,}"
        )

        # ====================================================
        # COLUMNAS OBLIGATORIAS DEL BANCO
        # ====================================================

        columnas_85 = [
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
        # VALIDAR BANCO
        # ----------------------------------------------------

        faltantes_banco_85 = [
            columna
            for columna in columnas_85
            if columna not in df_banco_85.columns
        ]

        if faltantes_banco_85:

            st.error(
                "8.3 ERROR: faltan columnas en "
                "BANCO_PREGUNTAS_GENERALES.xlsx: "
                + ", ".join(
                    faltantes_banco_85
                )
            )

            return

        # ====================================================
        # CONVERTIR PREGUNTAS APROBADAS A DATAFRAME
        # ====================================================

        df_nuevas_85 = pd.DataFrame(
            aprobadas_85
        )

        # ----------------------------------------------------
        # VALIDAR COLUMNAS DE LAS PREGUNTAS
        # ----------------------------------------------------

        faltantes_nuevas_85 = [
            columna
            for columna in columnas_85
            if columna not in df_nuevas_85.columns
        ]

        if faltantes_nuevas_85:

            st.error(
                "8.3 ERROR: faltan columnas en "
                "las preguntas aprobadas: "
                + ", ".join(
                    faltantes_nuevas_85
                )
            )

            return

        # ----------------------------------------------------
        # RESPETAR EXACTAMENTE EL ORDEN DEL BANCO
        # ----------------------------------------------------

        df_nuevas_85 = df_nuevas_85[
            columnas_85
        ].copy()

        # ====================================================
        # EVITAR DUPLICADOS POR PREGUNTA_ID
        # ====================================================

        ids_existentes_85 = set(
            df_banco_85[
                "Pregunta_ID"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        filas_nuevas_85 = []

        for _, fila_85 in df_nuevas_85.iterrows():

            pregunta_id_85 = str(
                fila_85["Pregunta_ID"]
            ).strip()

            if not pregunta_id_85:
                continue

            if pregunta_id_85 in ids_existentes_85:
                continue

            filas_nuevas_85.append(
                fila_85
            )

            ids_existentes_85.add(
                pregunta_id_85
            )

        nuevas_85 = len(
            filas_nuevas_85
        )

        # ====================================================
        # SI NO HAY PREGUNTAS NUEVAS
        # ====================================================

        if nuevas_85 == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                "El banco continúa con "
                f"{total_antes_85:,} preguntas."
            )

            return

        # ====================================================
        # AGREGAR AL BANCO GENERAL
        # ====================================================

        df_agregar_85 = pd.DataFrame(
            filas_nuevas_85
        )

        df_final_85 = pd.concat(
            [
                df_banco_85,
                df_agregar_85
            ],
            ignore_index=True
        )

        total_despues_85 = len(
            df_final_85
        )

        # ====================================================
        # ACTUALIZAR DATAFRAME LOCAL
        # ====================================================

        st.session_state[
            "df_banco_85"
        ] = df_final_85.copy()

        # ====================================================
        # CREAR EXCEL ACTUALIZADO
        # ====================================================

        memoria_salida_85 = io.BytesIO()

        with pd.ExcelWriter(
            memoria_salida_85,
            engine="openpyxl"
        ) as escritor_85:

            df_final_85.to_excel(
                escritor_85,
                index=False,
                sheet_name="Banco"
            )

        # ====================================================
        # CODIFICAR ARCHIVO
        # ====================================================

        contenido_nuevo_85 = base64.b64encode(
            memoria_salida_85.getvalue()
        ).decode("utf-8")

        # ====================================================
        # ACTUALIZAR GITHUB
        # ====================================================

        datos_actualizacion_85 = {

            "message":
                "Agregar preguntas 8.3 - "
                "Producto Restricción Alternativas Seguras",

            "content":
                contenido_nuevo_85,

            "branch":
                GITHUB_RAMA_85,

            "sha":
                sha_85
        }

        cuerpo_85 = json.dumps(
            datos_actualizacion_85
        ).encode("utf-8")

        solicitud_actualizacion_85 = (
            urllib.request.Request(
                URL_GITHUB_85,
                data=cuerpo_85,
                headers={
                    **headers_85,
                    "Content-Type":
                        "application/json"
                },
                method="PUT"
            )
        )

        with urllib.request.urlopen(
            solicitud_actualizacion_85,
            timeout=30
        ) as respuesta_actualizacion_85:

            respuesta_actualizacion_85.read()

        # ====================================================
        # RESULTADO
        # ====================================================

        st.success(
            "8.3 sincronizado correctamente con "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.info(
            "Preguntas en el banco antes: "
            f"{total_antes_85:,}"
        )

        st.info(
            "Preguntas nuevas incorporadas: "
            f"{nuevas_85:,}"
        )

        st.info(
            "Preguntas en el banco después: "
            f"{total_despues_85:,}"
        )

    except Exception as error_85:

        st.error(
            "8.3 ERROR: no fue posible cargar o "
            "actualizar BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(
            error_85
        )


# ============================================================
# BOTÓN DE SINCRONIZACIÓN 8.3
# ============================================================

preguntas_85 = st.session_state.get(
    "preguntas_generadas_85",
    []
)

aprobadas_85 = sum(
    1
    for pregunta in preguntas_85
    if pregunta.get("Estado") == "APROBADA"
)

if aprobadas_85 > 0:

    st.markdown(
        "### 8.3 - Sincronización"
    )

    st.info(
        "Preguntas aprobadas listas para sincronizar: "
        f"{aprobadas_85}"
    )

    if st.button(
        "SINCRONIZAR 8.3 CON BANCO DE PREGUNTAS",
        key="boton_sincronizar_85"
    ):

        sincronizar_85()
# ============================================================
# 9.1 - COMPLEMENTARIOS
# PARTE 1 - CARGA Y CONTROL DE FUENTE
# ============================================================

ARCHIVO_FUENTE_91 = "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
HOJA_FUENTE_91 = "Complementarios"
ARCHIVO_BANCO_91 = "BANCO_PREGUNTAS_GENERALES.xlsx"


# ============================================================
# CARGAR FUENTE 9.1
# ============================================================

def cargar_fuente_91():

    try:

        df = pd.read_excel(
            ARCHIVO_FUENTE_91,
            sheet_name=HOJA_FUENTE_91,
            engine="openpyxl"
        )

    except Exception as error:

        st.error(
            f"9.1 ERROR al cargar la hoja "
            f"{HOJA_FUENTE_91}: {error}"
        )

        return None

    # --------------------------------------------------------
    # COLUMNAS EXACTAS DE LA HOJA COMPLEMENTARIOS
    # --------------------------------------------------------

    columnas = [
        "Producto",
        "Categoría principal",
        "Indicaciones / Escenarios",
        "Modo de acción resumido",
        "Combinaciones estratégicas"
    ]

    faltantes = [
        columna
        for columna in columnas
        if columna not in df.columns
    ]

    if faltantes:

        st.error(
            "9.1 ERROR: faltan columnas en la hoja "
            "Complementarios: "
            + ", ".join(faltantes)
        )

        return None

    # --------------------------------------------------------
    # CONSERVAR COLUMNAS
    # --------------------------------------------------------

    df = df[columnas].copy()

    # --------------------------------------------------------
    # LIMPIEZA MÍNIMA
    #
    # NO SE DIVIDEN CELDAS
    # NO SE RESUMEN TEXTOS
    # NO SE ALTERAN LAS INDICACIONES
    # NO SE ALTERAN LAS COMBINACIONES
    # --------------------------------------------------------

    for columna in columnas:

        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # ELIMINAR SOLO REGISTROS SIN LOS DATOS NECESARIOS
    # PARA ESTE GENERADOR
    # --------------------------------------------------------

    df = df[
        (df["Producto"] != "")
        &
        (df["Indicaciones / Escenarios"] != "")
        &
        (df["Combinaciones estratégicas"] != "")
    ].copy()

    # --------------------------------------------------------
    # CREAR IDENTIFICADOR ÚNICO DE FUENTE
    # --------------------------------------------------------

    df["Fuente_ID"] = [
        f"CP-{i:06d}"
        for i in range(1, len(df) + 1)
    ]

    return df.reset_index(drop=True)


# ============================================================
# CARGAR BANCO GENERAL
# ============================================================

def cargar_banco_91():

    try:

        return pd.read_excel(
            ARCHIVO_BANCO_91,
            engine="openpyxl"
        )

    except Exception:

        return pd.DataFrame()


# ============================================================
# OBTENER FUENTES YA UTILIZADAS
# ============================================================

def obtener_fuentes_usadas_91(df_banco):

    usadas = set()

    if df_banco.empty:

        return usadas

    if "Fuente_ID" not in df_banco.columns:

        return usadas

    for valor in df_banco["Fuente_ID"].fillna(""):

        for fuente in str(valor).split(";"):

            fuente = fuente.strip()

            if fuente:

                usadas.add(fuente)

    return usadas


# ============================================================
# INTERFAZ 9.1
# ============================================================

st.markdown(
    "## 9.1 Complementarios"
)

st.write(
    "Carga y control de la hoja Complementarios."
)


if st.button(
    "CARGAR Y VALIDAR FUENTE 9.1",
    key="cargar_fuentes_91"
):

    # --------------------------------------------------------
    # CARGAR HOJA COMPLEMENTARIOS
    # --------------------------------------------------------

    df_fuente_91 = cargar_fuente_91()

    if df_fuente_91 is None:

        st.stop()

    # --------------------------------------------------------
    # CARGAR BANCO GENERAL
    # --------------------------------------------------------

    df_banco_91 = cargar_banco_91()

    # --------------------------------------------------------
    # IDENTIFICAR FUENTES YA UTILIZADAS
    # --------------------------------------------------------

    fuentes_usadas_91 = (
        obtener_fuentes_usadas_91(
            df_banco_91
        )
    )

    # --------------------------------------------------------
    # CREAR DATAFRAME DISPONIBLE
    #
    # SOLO SE EXCLUYEN LAS FUENTES YA UTILIZADAS
    # --------------------------------------------------------

    df_disponible_91 = df_fuente_91[
        ~df_fuente_91["Fuente_ID"].isin(
            fuentes_usadas_91
        )
    ].copy()

    df_disponible_91 = (
        df_disponible_91
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # GUARDAR EN SESSION STATE
    # --------------------------------------------------------

    st.session_state[
        "df_fuente_91"
    ] = df_fuente_91.copy()

    st.session_state[
        "df_banco_91"
    ] = df_banco_91.copy()

    st.session_state[
        "df_disponible_91"
    ] = df_disponible_91.copy()

    st.session_state[
        "fuentes_usadas_91"
    ] = fuentes_usadas_91


# ============================================================
# MOSTRAR CONTROL DE CARGA
# ============================================================

if "df_fuente_91" in st.session_state:

    df_fuente_91 = st.session_state[
        "df_fuente_91"
    ]

    df_disponible_91 = st.session_state[
        "df_disponible_91"
    ]

    fuentes_usadas_91 = st.session_state[
        "fuentes_usadas_91"
    ]

    st.success(
        "9.1 cargó correctamente la hoja "
        "Complementarios."
    )

    # --------------------------------------------------------
    # INDICADORES
    # --------------------------------------------------------

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Registros fuente",
            len(df_fuente_91)
        )

    with col2:

        st.metric(
            "Fuentes utilizadas",
            len(
                fuentes_usadas_91
                &
                set(
                    df_fuente_91[
                        "Fuente_ID"
                    ]
                )
            )
        )

    with col3:

        st.metric(
            "Registros disponibles",
            len(df_disponible_91)
        )

    # --------------------------------------------------------
    # MOSTRAR ESTRUCTURA
    # --------------------------------------------------------

    st.markdown(
        "### Estructura disponible"
    )

    st.dataframe(
        df_disponible_91[
            [
                "Fuente_ID",
                "Producto",
                "Categoría principal",
                "Indicaciones / Escenarios",
                "Modo de acción resumido",
                "Combinaciones estratégicas"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
# ============================================================
# 9.2 - PARTE 2
# GENERADOR DE PREGUNTAS
# PRODUCTO - USO REGULAR / COMBINACIONES ESTRATÉGICAS
# NIVEL 1
#
# FUENTE:
#     df_disponible_91
#
# SALIDA:
#     df_banco_92
# ============================================================


# ============================================================
# GENERAR SIGUIENTE ID
# ============================================================

def siguiente_id_92():

    df_banco_92 = st.session_state.get(
        "df_banco_92",
        pd.DataFrame()
    )

    mayor_92 = 0

    # --------------------------------------------------------
    # REVISAR DATAFRAME 9.2
    # --------------------------------------------------------

    if (
        not df_banco_92.empty
        and "Pregunta_ID" in df_banco_92.columns
    ):

        for valor_92 in df_banco_92[
            "Pregunta_ID"
        ].fillna(""):

            texto_92 = str(
                valor_92
            ).strip()

            if texto_92.startswith("PTCP-"):

                try:

                    numero_92 = int(
                        texto_92.replace(
                            "PTCP-",
                            ""
                        )
                    )

                    mayor_92 = max(
                        mayor_92,
                        numero_92
                    )

                except ValueError:

                    pass

    # --------------------------------------------------------
    # REVISAR PREGUNTAS GENERADAS EN SESIÓN
    # --------------------------------------------------------

    preguntas_92 = st.session_state.get(
        "preguntas_generadas_92",
        []
    )

    for pregunta_92 in preguntas_92:

        texto_92 = str(
            pregunta_92.get(
                "Pregunta_ID",
                ""
            )
        ).strip()

        if texto_92.startswith("PTCP-"):

            try:

                numero_92 = int(
                    texto_92.replace(
                        "PTCP-",
                        ""
                    )
                )

                mayor_92 = max(
                    mayor_92,
                    numero_92
                )

            except ValueError:

                pass

    return (
        f"PTCP-{mayor_92 + 1:06d}"
    )


# ============================================================
# GENERADOR
# ============================================================

def generar_preguntas_92(cantidad):

    # --------------------------------------------------------
    # LA FUENTE ES df_disponible_91
    # --------------------------------------------------------

    df_92 = st.session_state.get(
        "df_disponible_91",
        pd.DataFrame()
    )

    if df_92.empty:

        st.warning(
            "9.2: no hay registros disponibles "
            "en df_disponible_91."
        )

        return []

    # --------------------------------------------------------
    # COLUMNAS NECESARIAS
    # --------------------------------------------------------

    columnas_92 = [
        "Fuente_ID",
        "Producto",
        "Indicaciones / Escenarios",
        "Combinaciones estratégicas"
    ]

    faltantes_92 = [
        columna_92
        for columna_92 in columnas_92
        if columna_92 not in df_92.columns
    ]

    if faltantes_92:

        st.error(
            "9.2 ERROR: faltan columnas en "
            "df_disponible_91: "
            + ", ".join(
                faltantes_92
            )
        )

        return []

    # --------------------------------------------------------
    # FUENTES YA CONSUMIDAS DURANTE ESTA SESIÓN
    # --------------------------------------------------------

    consumidas_92 = st.session_state.get(
        "fuentes_consumidas_92",
        set()
    ).copy()

    # --------------------------------------------------------
    # CANDIDATOS
    # --------------------------------------------------------

    candidatos_92 = df_92[
        ~df_92["Fuente_ID"].astype(str).isin(
            {
                str(x)
                for x in consumidas_92
            }
        )
    ].copy()

    # --------------------------------------------------------
    # VALIDAR CAMPOS NECESARIOS
    # --------------------------------------------------------

    candidatos_92 = candidatos_92[
        (
            candidatos_92[
                "Fuente_ID"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
        &
        (
            candidatos_92[
                "Producto"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
        &
        (
            candidatos_92[
                "Indicaciones / Escenarios"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
        &
        (
            candidatos_92[
                "Combinaciones estratégicas"
            ]
            .astype(str)
            .str.strip()
            != ""
        )
    ].copy()

    if candidatos_92.empty:

        st.warning(
            "9.2: no hay registros completos "
            "disponibles para generar preguntas."
        )

        return []

    # --------------------------------------------------------
    # ALEATORIZAR CANDIDATOS
    # --------------------------------------------------------

    candidatos_92 = (
        candidatos_92
        .sample(frac=1)
        .reset_index(drop=True)
    )

    preguntas_92 = []

    # ========================================================
    # GENERACIÓN DE PREGUNTAS
    # ========================================================

    for _, fila_92 in candidatos_92.iterrows():

        fuente_92 = str(
            fila_92[
                "Fuente_ID"
            ]
        ).strip()

        producto_92 = str(
            fila_92[
                "Producto"
            ]
        ).strip()

        indicaciones_92 = str(
            fila_92[
                "Indicaciones / Escenarios"
            ]
        ).strip()

        correcta_92 = str(
            fila_92[
                "Combinaciones estratégicas"
            ]
        ).strip()

        # ----------------------------------------------------
        # BUSCAR DISTRACTORES
        #
        # Un distractor NO puede tener:
        # 1. La misma fuente
        # 2. La misma combinación estratégica
        #
        # Esto es indispensable porque dos productos pueden
        # compartir exactamente la misma combinación.
        # ----------------------------------------------------

        falsas_92 = candidatos_92[
            candidatos_92[
                "Fuente_ID"
            ].astype(str)
            != fuente_92
        ].copy()

        falsas_92["respuesta"] = (
            falsas_92[
                "Combinaciones estratégicas"
            ]
            .astype(str)
            .str.strip()
        )

        # ----------------------------------------------------
        # EXCLUIR LA COMBINACIÓN CORRECTA
        # ----------------------------------------------------

        falsas_92 = falsas_92[
            falsas_92[
                "respuesta"
            ].str.casefold()
            != correcta_92.casefold()
        ].copy()

        # ----------------------------------------------------
        # ELIMINAR RESPUESTAS DUPLICADAS
        #
        # Si varios productos tienen la misma combinación,
        # esa combinación cuenta como una sola alternativa.
        # ----------------------------------------------------

        falsas_92 = falsas_92.drop_duplicates(
            subset="respuesta"
        )

        # ----------------------------------------------------
        # SE NECESITAN 3 DISTRACTORES DIFERENTES
        # ----------------------------------------------------

        if len(falsas_92) < 3:

            continue

        falsas_92 = falsas_92.sample(
            n=3
        )

        # ----------------------------------------------------
        # CONSTRUIR LAS 4 OPCIONES
        # ----------------------------------------------------

        opciones_92 = [
            correcta_92,
            str(
                falsas_92.iloc[0][
                    "respuesta"
                ]
            ).strip(),
            str(
                falsas_92.iloc[1][
                    "respuesta"
                ]
            ).strip(),
            str(
                falsas_92.iloc[2][
                    "respuesta"
                ]
            ).strip()
        ]

        # ----------------------------------------------------
        # MEZCLAR OPCIONES
        # ----------------------------------------------------

        np.random.shuffle(
            opciones_92
        )

        correcta_numero_92 = (
            opciones_92.index(
                correcta_92
            ) + 1
        )

        # ====================================================
        # CREAR PREGUNTA
        # ====================================================

        pregunta_92 = {

            "Pregunta_ID":
                siguiente_id_92(),

            "Modulo":
                "Complementarios",

            "Tema":
                "Producto - "
                "Uso regular / "
                "Combinaciones estratégicas",

            "Nivel":
                "Nivel 1",

            "Tipo_Relacion":
                "Producto-Uso regular-Combinaciones estratégicas",

            "Pregunta":
                "Para el producto "
                f"{producto_92}, "
                "cuyo uso regular se relaciona con "
                f"{indicaciones_92}, "
                "¿cuál de las siguientes corresponde "
                "a sus combinaciones estratégicas?",

            "Respuesta_1":
                opciones_92[0],

            "Respuesta_2":
                opciones_92[1],

            "Respuesta_3":
                opciones_92[2],

            "Respuesta_4":
                opciones_92[3],

            "Respuesta_Correcta":
                str(
                    correcta_numero_92
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
                fuente_92
        }

        preguntas_92.append(
            pregunta_92
        )

        # ----------------------------------------------------
        # MARCAR FUENTE COMO CONSUMIDA
        # ----------------------------------------------------

        consumidas_92.add(
            fuente_92
        )

        if len(preguntas_92) >= int(cantidad):

            break

    # --------------------------------------------------------
    # PERSISTIR FUENTES CONSUMIDAS
    # --------------------------------------------------------

    st.session_state[
        "fuentes_consumidas_92"
    ] = consumidas_92

    return preguntas_92


# ============================================================
# INTERFAZ DEL GENERADOR 9.2
# ============================================================

if "df_disponible_91" in st.session_state:

    st.markdown(
        "### 9.2 - Generador de preguntas"
    )

    st.info(
        "Nivel 1: Producto + uso regular "
        "(Indicaciones / Escenarios) → "
        "Combinaciones estratégicas."
    )

    cantidad_92 = st.number_input(
        "Cantidad de preguntas",
        min_value=1,
        max_value=500,
        value=10,
        step=1,
        key="cantidad_generar_92"
    )

    if st.button(
        "GENERAR PREGUNTAS 9.2",
        key="generar_preguntas_92"
    ):

        nuevas_92 = generar_preguntas_92(
            int(cantidad_92)
        )

        st.session_state[
            "preguntas_generadas_92"
        ] = nuevas_92

        # ----------------------------------------------------
        # CREAR DATAFRAME DE SALIDA 9.2
        # ----------------------------------------------------

        if nuevas_92:

            st.session_state[
                "df_banco_92"
            ] = pd.DataFrame(
                nuevas_92
            )

            st.success(
                f"Se generaron "
                f"{len(nuevas_92)} "
                "preguntas de Nivel 1."
            )

        else:

            st.warning(
                "No fue posible generar preguntas "
                "con las relaciones disponibles."
            )


# ============================================================
# MOSTRAR PREGUNTAS GENERADAS
# ============================================================

preguntas_92 = st.session_state.get(
    "preguntas_generadas_92",
    []
)

if preguntas_92:

    st.markdown(
        "### Preguntas generadas 9.2"
    )

    for pregunta_92 in preguntas_92:

        st.markdown(
            f"**{pregunta_92['Pregunta_ID']} — "
            f"{pregunta_92['Nivel']}**"
        )

        st.write(
            pregunta_92["Pregunta"]
        )

        st.write(
            f"1. {pregunta_92['Respuesta_1']}"
        )

        st.write(
            f"2. {pregunta_92['Respuesta_2']}"
        )

        st.write(
            f"3. {pregunta_92['Respuesta_3']}"
        )

        st.write(
            f"4. {pregunta_92['Respuesta_4']}"
        )

        st.caption(
            "Respuesta correcta: "
            f"{pregunta_92['Respuesta_Correcta']}"
        )

        st.caption(
            "Fuente: "
            f"{pregunta_92['Fuente_ID']}"
        )

        st.divider()
# ============================================================
# 9.3 - PARTE 3
# VALIDADOR DE PREGUNTAS
# COMPLEMENTARIOS
# PRODUCTO - USO REGULAR / COMBINACIONES ESTRATÉGICAS
# NIVEL 1
#
# ENTRADA:
#     preguntas_generadas_92
#
# SALIDA:
#     df_banco_92
# ============================================================


preguntas_92 = st.session_state.get(
    "preguntas_generadas_92",
    []
)


# ============================================================
# VALIDADOR
# ============================================================

if preguntas_92:

    st.markdown(
        "### 9.3 - Validación de preguntas"
    )

    st.info(
        "Revise cada pregunta y apruebe o rechace "
        "individualmente."
    )

    for i, pregunta_92 in enumerate(
        preguntas_92
    ):

        pregunta_id_92 = pregunta_92.get(
            "Pregunta_ID",
            f"PTCP-{i + 1:06d}"
        )

        st.markdown(
            f"#### {pregunta_id_92}"
        )

        # ----------------------------------------------------
        # PREGUNTA
        # ----------------------------------------------------

        st.write(
            pregunta_92.get(
                "Pregunta",
                ""
            )
        )

        # ----------------------------------------------------
        # OPCIONES
        # ----------------------------------------------------

        st.write(
            "1. "
            + str(
                pregunta_92.get(
                    "Respuesta_1",
                    ""
                )
            )
        )

        st.write(
            "2. "
            + str(
                pregunta_92.get(
                    "Respuesta_2",
                    ""
                )
            )
        )

        st.write(
            "3. "
            + str(
                pregunta_92.get(
                    "Respuesta_3",
                    ""
                )
            )
        )

        st.write(
            "4. "
            + str(
                pregunta_92.get(
                    "Respuesta_4",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # RESPUESTA CORRECTA
        # ----------------------------------------------------

        st.caption(
            "Respuesta correcta: "
            + str(
                pregunta_92.get(
                    "Respuesta_Correcta",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # FUENTE
        # ----------------------------------------------------

        st.caption(
            "Fuente: "
            + str(
                pregunta_92.get(
                    "Fuente_ID",
                    ""
                )
            )
        )

        # ----------------------------------------------------
        # ESTADO ACTUAL
        # ----------------------------------------------------

        estado_92 = pregunta_92.get(
            "Estado",
            "PENDIENTE"
        )

        st.write(
            f"Estado: {estado_92}"
        )

        # ----------------------------------------------------
        # OBSERVACIÓN DEL ADMINISTRADOR
        # ----------------------------------------------------

        observacion_92 = st.text_input(
            "Observación del administrador",
            value=pregunta_92.get(
                "Observacion_Administrador",
                ""
            ),
            key=f"observacion_complementarios_92_{i}"
        )

        col1_92, col2_92 = st.columns(2)

        # ====================================================
        # APROBAR
        # ====================================================

        with col1_92:

            if st.button(
                "APROBAR",
                key=f"aprobar_complementarios_92_{i}"
            ):

                preguntas_92[i][
                    "Estado"
                ] = "APROBADA"

                preguntas_92[i][
                    "Observacion_Administrador"
                ] = observacion_92

                # --------------------------------------------
                # ACTUALIZAR PREGUNTAS
                # --------------------------------------------

                st.session_state[
                    "preguntas_generadas_92"
                ] = preguntas_92

                # --------------------------------------------
                # ACTUALIZAR DATAFRAME 9.2
                # --------------------------------------------

                st.session_state[
                    "df_banco_92"
                ] = pd.DataFrame(
                    preguntas_92
                )

                st.rerun()

        # ====================================================
        # RECHAZAR
        # ====================================================

        with col2_92:

            if st.button(
                "RECHAZAR",
                key=f"rechazar_complementarios_92_{i}"
            ):

                preguntas_92[i][
                    "Estado"
                ] = "RECHAZADA"

                preguntas_92[i][
                    "Observacion_Administrador"
                ] = observacion_92

                # --------------------------------------------
                # ACTUALIZAR PREGUNTAS
                # --------------------------------------------

                st.session_state[
                    "preguntas_generadas_92"
                ] = preguntas_92

                # --------------------------------------------
                # ACTUALIZAR DATAFRAME 9.2
                # --------------------------------------------

                st.session_state[
                    "df_banco_92"
                ] = pd.DataFrame(
                    preguntas_92
                )

                st.rerun()

        st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN 9.3
# ============================================================

if preguntas_92:

    aprobadas_92 = sum(
        1
        for pregunta_92 in preguntas_92
        if pregunta_92.get(
            "Estado"
        ) == "APROBADA"
    )

    rechazadas_92 = sum(
        1
        for pregunta_92 in preguntas_92
        if pregunta_92.get(
            "Estado"
        ) == "RECHAZADA"
    )

    pendientes_92 = sum(
        1
        for pregunta_92 in preguntas_92
        if pregunta_92.get(
            "Estado",
            "PENDIENTE"
        ) == "PENDIENTE"
    )

    st.markdown(
        "### Resumen de validación 9.3"
    )

    col1_92, col2_92, col3_92 = st.columns(3)

    with col1_92:

        st.metric(
            "Aprobadas",
            aprobadas_92
        )

    with col2_92:

        st.metric(
            "Rechazadas",
            rechazadas_92
        )

    with col3_92:

        st.metric(
            "Pendientes",
            pendientes_92
        )

    if pendientes_92 == 0:

        st.success(
            "Todas las preguntas fueron revisadas."
        )
# ============================================================
# 9.4 - PARTE 4
# SINCRONIZADOR
# COMPLEMENTARIOS
# PRODUCTO - USO REGULAR / COMBINACIONES ESTRATÉGICAS
# NIVEL 1
#
# ENTRADA:
#     preguntas_generadas_92
#
# BANCO:
#     BANCO_PREGUNTAS_GENERALES.xlsx
#
# DATAFRAME:
#     df_banco_92
# ============================================================


GITHUB_USUARIO_92 = "franquiciasauces"
GITHUB_REPOSITORIO_92 = "Asesores"
GITHUB_RAMA_92 = "main"
GITHUB_ARCHIVO_92 = "BANCO_PREGUNTAS_GENERALES.xlsx"


URL_GITHUB_92 = (
    "https://api.github.com/repos/"
    + GITHUB_USUARIO_92
    + "/"
    + GITHUB_REPOSITORIO_92
    + "/contents/"
    + GITHUB_ARCHIVO_92
)


# ============================================================
# FUNCIÓN DE SINCRONIZACIÓN
# ============================================================

def sincronizar_92():

    # --------------------------------------------------------
    # OBTENER PREGUNTAS GENERADAS
    # --------------------------------------------------------

    preguntas_92 = st.session_state.get(
        "preguntas_generadas_92",
        []
    )

    # --------------------------------------------------------
    # SOLO PREGUNTAS APROBADAS
    # --------------------------------------------------------

    aprobadas_92 = [
        pregunta_92
        for pregunta_92 in preguntas_92
        if pregunta_92.get(
            "Estado"
        ) == "APROBADA"
    ]

    if not aprobadas_92:

        st.warning(
            "No hay preguntas aprobadas para sincronizar."
        )

        return

    # --------------------------------------------------------
    # TOKEN
    # --------------------------------------------------------

    headers_92 = {
        "Authorization":
            f"Bearer {GITHUB_TOKEN}",

        "Accept":
            "application/vnd.github+json"
    }

    try:

        # ====================================================
        # LEER BANCO ACTUAL DESDE GITHUB
        # ====================================================

        solicitud_92 = urllib.request.Request(
            URL_GITHUB_92,
            headers=headers_92,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud_92,
            timeout=30
        ) as respuesta_92:

            datos_92 = json.loads(
                respuesta_92
                .read()
                .decode("utf-8")
            )

        # ----------------------------------------------------
        # SHA
        # ----------------------------------------------------

        if "sha" not in datos_92:

            st.error(
                "9.4 ERROR: GitHub no devolvió "
                "el SHA del archivo."
            )

            return

        sha_92 = datos_92["sha"]

        # ----------------------------------------------------
        # CONTENIDO
        # ----------------------------------------------------

        if "content" not in datos_92:

            st.error(
                "9.4 ERROR: GitHub no devolvió "
                "el contenido del banco."
            )

            return

        # ====================================================
        # DECODIFICAR ARCHIVO
        # ====================================================

        contenido_92 = base64.b64decode(
            datos_92["content"].replace(
                "\n",
                ""
            )
        )

        memoria_92 = io.BytesIO(
            contenido_92
        )

        df_banco_general_92 = pd.read_excel(
            memoria_92,
            engine="openpyxl"
        )

        total_antes_92 = len(
            df_banco_general_92
        )

        # ----------------------------------------------------
        # CONFIRMACIÓN DE CARGA
        # ----------------------------------------------------

        st.success(
            "BANCO_PREGUNTAS_GENERALES.xlsx "
            "cargado correctamente desde GitHub."
        )

        st.info(
            "Preguntas en el banco antes de sincronizar: "
            f"{total_antes_92:,}"
        )

        # ====================================================
        # COLUMNAS DEL BANCO GENERAL
        # ====================================================

        columnas_92 = [
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
        # VALIDAR BANCO
        # ----------------------------------------------------

        faltantes_banco_92 = [
            columna_92
            for columna_92 in columnas_92
            if columna_92
            not in df_banco_general_92.columns
        ]

        if faltantes_banco_92:

            st.error(
                "9.4 ERROR: faltan columnas en "
                "BANCO_PREGUNTAS_GENERALES.xlsx: "
                + ", ".join(
                    faltantes_banco_92
                )
            )

            return

        # ====================================================
        # CREAR DATAFRAME DE PREGUNTAS APROBADAS
        # ====================================================

        df_nuevas_92 = pd.DataFrame(
            aprobadas_92
        )

        # ----------------------------------------------------
        # VALIDAR PREGUNTAS
        # ----------------------------------------------------

        faltantes_nuevas_92 = [
            columna_92
            for columna_92 in columnas_92
            if columna_92
            not in df_nuevas_92.columns
        ]

        if faltantes_nuevas_92:

            st.error(
                "9.4 ERROR: faltan columnas en "
                "las preguntas aprobadas: "
                + ", ".join(
                    faltantes_nuevas_92
                )
            )

            return

        # ----------------------------------------------------
        # RESPETAR ORDEN DEL BANCO
        # ----------------------------------------------------

        df_nuevas_92 = df_nuevas_92[
            columnas_92
        ].copy()

        # ====================================================
        # EVITAR DUPLICADOS POR PREGUNTA_ID
        # ====================================================

        ids_existentes_92 = set(
            df_banco_general_92[
                "Pregunta_ID"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        filas_nuevas_92 = []

        for _, fila_92 in df_nuevas_92.iterrows():

            pregunta_id_92 = str(
                fila_92[
                    "Pregunta_ID"
                ]
            ).strip()

            # ------------------------------------------------
            # IGNORAR ID VACÍO
            # ------------------------------------------------

            if not pregunta_id_92:

                continue

            # ------------------------------------------------
            # IGNORAR DUPLICADO
            # ------------------------------------------------

            if pregunta_id_92 in ids_existentes_92:

                continue

            filas_nuevas_92.append(
                fila_92
            )

            ids_existentes_92.add(
                pregunta_id_92
            )

        nuevas_92 = len(
            filas_nuevas_92
        )

        # ====================================================
        # SI NO HAY NADA NUEVO
        # ====================================================

        if nuevas_92 == 0:

            st.info(
                "No hay preguntas nuevas para agregar."
            )

            st.info(
                "El banco continúa con "
                f"{total_antes_92:,} preguntas."
            )

            return

        # ====================================================
        # AGREGAR AL BANCO GENERAL
        # ====================================================

        df_agregar_92 = pd.DataFrame(
            filas_nuevas_92
        )

        df_final_92 = pd.concat(
            [
                df_banco_general_92,
                df_agregar_92
            ],
            ignore_index=True
        )

        total_despues_92 = len(
            df_final_92
        )

        # ====================================================
        # ACTUALIZAR DATAFRAME LOCAL 9.2
        # ====================================================

        st.session_state[
            "df_banco_92"
        ] = df_final_92.copy()

        # ====================================================
        # CREAR EXCEL
        # ====================================================

        memoria_salida_92 = io.BytesIO()

        with pd.ExcelWriter(
            memoria_salida_92,
            engine="openpyxl"
        ) as escritor_92:

            df_final_92.to_excel(
                escritor_92,
                index=False,
                sheet_name="Banco"
            )

        # ====================================================
        # CODIFICAR ARCHIVO
        # ====================================================

        contenido_nuevo_92 = base64.b64encode(
            memoria_salida_92.getvalue()
        ).decode("utf-8")

        # ====================================================
        # ACTUALIZAR GITHUB
        # ====================================================

        datos_actualizacion_92 = {

            "message":
                "Agregar preguntas 9.2 - "
                "Complementarios",

            "content":
                contenido_nuevo_92,

            "branch":
                GITHUB_RAMA_92,

            "sha":
                sha_92
        }

        cuerpo_92 = json.dumps(
            datos_actualizacion_92
        ).encode("utf-8")

        solicitud_actualizacion_92 = (
            urllib.request.Request(
                URL_GITHUB_92,
                data=cuerpo_92,
                headers={
                    **headers_92,
                    "Content-Type":
                        "application/json"
                },
                method="PUT"
            )
        )

        with urllib.request.urlopen(
            solicitud_actualizacion_92,
            timeout=30
        ) as respuesta_actualizacion_92:

            respuesta_actualizacion_92.read()

        # ====================================================
        # RESULTADO
        # ====================================================

        st.success(
            "9.4 sincronizado correctamente con "
            "BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.info(
            "Preguntas en el banco antes: "
            f"{total_antes_92:,}"
        )

        st.info(
            "Preguntas nuevas incorporadas: "
            f"{nuevas_92:,}"
        )

        st.info(
            "Preguntas en el banco después: "
            f"{total_despues_92:,}"
        )

    except Exception as error_92:

        st.error(
            "9.4 ERROR: no fue posible cargar o "
            "actualizar BANCO_PREGUNTAS_GENERALES.xlsx."
        )

        st.exception(
            error_92
        )


# ============================================================
# BOTÓN DE SINCRONIZACIÓN 9.4
# ============================================================

preguntas_92 = st.session_state.get(
    "preguntas_generadas_92",
    []
)

aprobadas_92 = sum(
    1
    for pregunta_92 in preguntas_92
    if pregunta_92.get(
        "Estado"
    ) == "APROBADA"
)


if aprobadas_92 > 0:

    st.markdown(
        "### 9.4 - Sincronización"
    )

    st.info(
        "Preguntas aprobadas listas para sincronizar: "
        f"{aprobadas_92}"
    )

    if st.button(
        "SINCRONIZAR 9.2 CON BANCO DE PREGUNTAS",
        key="boton_sincronizar_92"
    ):

        sincronizar_92()

# ============================================================
# 10.1 - CARGA Y CONTROL DEL BANCO GENERAL
# ============================================================
# FUENTE:
#     BANCO_PREGUNTAS_GENERALES.xlsx
#
# UBICACIÓN:
#     GitHub / Asesores / main
#
# SALIDA:
#     df_banco_101
#
# FUNCIÓN:
#     Cargar el banco general una sola vez y dejarlo
#     disponible para el análisis y generación de evaluaciones.
#
# NO UTILIZA:
#     Tema
# ============================================================


GITHUB_USUARIO_101 = "franquiciasauces"
GITHUB_REPOSITORIO_101 = "Asesores"
GITHUB_RAMA_101 = "main"
GITHUB_ARCHIVO_101 = "BANCO_PREGUNTAS_GENERALES.xlsx"

URL_GITHUB_101 = (
    "https://api.github.com/repos/"
    + GITHUB_USUARIO_101
    + "/"
    + GITHUB_REPOSITORIO_101
    + "/contents/"
    + GITHUB_ARCHIVO_101
)


# ============================================================
# FUNCIÓN DE CARGA
# ============================================================

def cargar_banco_general_101():

    headers_101 = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    try:

        solicitud_101 = urllib.request.Request(
            URL_GITHUB_101,
            headers=headers_101,
            method="GET"
        )

        with urllib.request.urlopen(
            solicitud_101,
            timeout=30
        ) as respuesta_101:

            datos_101 = json.loads(
                respuesta_101.read().decode("utf-8")
            )

        # ----------------------------------------------------
        # VERIFICAR SHA
        # ----------------------------------------------------

        if "sha" not in datos_101:

            st.error(
                "10.1 ERROR: GitHub no devolvió el SHA "
                "del BANCO_PREGUNTAS_GENERALES.xlsx."
            )

            return None

        # ----------------------------------------------------
        # VERIFICAR CONTENIDO
        # ----------------------------------------------------

        if "content" not in datos_101:

            st.error(
                "10.1 ERROR: GitHub no devolvió el contenido "
                "del BANCO_PREGUNTAS_GENERALES.xlsx."
            )

            return None

        sha_101 = datos_101["sha"]

        # ----------------------------------------------------
        # DECODIFICAR ARCHIVO
        # ----------------------------------------------------

        contenido_101 = base64.b64decode(
            datos_101["content"].replace(
                "\n",
                ""
            )
        )

        memoria_101 = io.BytesIO(
            contenido_101
        )

        df_101 = pd.read_excel(
            memoria_101,
            engine="openpyxl"
        )

        # ----------------------------------------------------
        # COLUMNAS OBLIGATORIAS
        #
        # TEMA NO SE UTILIZA
        # ----------------------------------------------------

        columnas_101 = [
            "Pregunta_ID",
            "Modulo",
            "Tipo_Relacion",
            "Nivel",
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

        faltantes_101 = [
            columna
            for columna in columnas_101
            if columna not in df_101.columns
        ]

        if faltantes_101:

            st.error(
                "10.1 ERROR: faltan columnas en "
                "BANCO_PREGUNTAS_GENERALES.xlsx: "
                + ", ".join(faltantes_101)
            )

            return None

        # ----------------------------------------------------
        # RESPETAR EL ORDEN DEL BANCO
        # ----------------------------------------------------

        df_101 = df_101[
            columnas_101
        ].copy()

        # ----------------------------------------------------
        # NORMALIZACIÓN DE CAMPOS DE CONTROL
        # ----------------------------------------------------

        for columna in [
            "Pregunta_ID",
            "Modulo",
            "Tipo_Relacion",
            "Nivel",
            "Estado",
            "Fuente_ID"
        ]:

            df_101[columna] = (
                df_101[columna]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        # ----------------------------------------------------
        # GUARDAR SHA
        # ----------------------------------------------------

        st.session_state[
            "sha_banco_general_101"
        ] = sha_101

        return df_101.reset_index(
            drop=True
        )

    except Exception as error_101:

        st.error(
            "10.1 ERROR: no fue posible cargar "
            "BANCO_PREGUNTAS_GENERALES.xlsx desde GitHub."
        )

        st.exception(error_101)

        return None


# ============================================================
# INTERFAZ 10.1
# ============================================================

st.markdown(
    "## 10.1 - Banco general de preguntas"
)

st.info(
    "Cargue el banco general actualizado desde GitHub "
    "antes de analizar y generar evaluaciones."
)


if st.button(
    "CARGAR BANCO GENERAL",
    key="cargar_banco_general_101"
):

    df_banco_cargado_101 = (
        cargar_banco_general_101()
    )

    if df_banco_cargado_101 is not None:

        # ----------------------------------------------------
        # GUARDAR FUENTE PARA TODO EL MÓDULO 10
        # ----------------------------------------------------

        st.session_state[
            "df_banco_101"
        ] = df_banco_cargado_101.copy()

        st.session_state[
            "banco_101_cargado"
        ] = True

        # ----------------------------------------------------
        # LIMPIAR ESTADOS ANTERIORES DE ANÁLISIS
        # ----------------------------------------------------

        st.session_state.pop(
            "resumen_disponibilidad_102",
            None
        )

        st.session_state.pop(
            "df_aprobadas_102",
            None
        )

        st.success(
            "BANCO_PREGUNTAS_GENERALES.xlsx "
            "cargado correctamente."
        )


# ============================================================
# MOSTRAR CONTROL DE CARGA
# ============================================================

if st.session_state.get(
    "banco_101_cargado",
    False
):

    df_banco_101 = st.session_state[
        "df_banco_101"
    ]

    # --------------------------------------------------------
    # TOTAL GENERAL
    # --------------------------------------------------------

    total_101 = len(
        df_banco_101
    )

    # --------------------------------------------------------
    # APROBADAS
    # --------------------------------------------------------

    aprobadas_101 = int(
        (
            df_banco_101["Estado"]
            .str.upper()
            == "APROBADA"
        ).sum()
    )

    # --------------------------------------------------------
    # RECHAZADAS
    # --------------------------------------------------------

    rechazadas_101 = int(
        (
            df_banco_101["Estado"]
            .str.upper()
            == "RECHAZADA"
        ).sum()
    )

    # --------------------------------------------------------
    # PENDIENTES
    # --------------------------------------------------------

    pendientes_101 = int(
        (
            df_banco_101["Estado"]
            .str.upper()
            == "PENDIENTE"
        ).sum()
    )

    # --------------------------------------------------------
    # MOSTRAR CONTROL
    # --------------------------------------------------------

    st.markdown(
        "### Control del banco cargado"
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.metric(
            "Preguntas totales",
            f"{total_101:,}"
        )

    with col2:

        st.metric(
            "Aprobadas",
            f"{aprobadas_101:,}"
        )

    with col3:

        st.metric(
            "Rechazadas",
            f"{rechazadas_101:,}"
        )

    with col4:

        st.metric(
            "Pendientes",
            f"{pendientes_101:,}"
        )

    # --------------------------------------------------------
    # CONTROL DE ESTRUCTURA
    # --------------------------------------------------------

    st.success(
        "Fuente disponible para 10.2."
    )

    st.caption(
        "La estructuración de las evaluaciones utilizará "
        "únicamente Módulo + Tipo_Relacion + Nivel. "
        "La columna Tema queda fuera del proceso."
    )
# ============================================================
# 10.2 - ANÁLISIS Y DISPONIBILIDAD DEL BANCO
# ============================================================
# ENTRADA:
#     df_banco_101
#
# FUENTE:
#     BANCO_PREGUNTAS_GENERALES.xlsx
#     cargado por 10.1
#
# CLASIFICACIÓN:
#     Módulo
#     Tipo_Relacion
#     Nivel
#
# NO UTILIZA:
#     Tema
#
# REGLA:
#     1 evaluación = 10 preguntas
#
# SALIDAS:
#     df_aprobadas_102
#     resumen_disponibilidad_102
#
# 10.2 NO consume preguntas.
# 10.2 NO modifica el banco.
# ============================================================


# ============================================================
# VERIFICAR QUE 10.1 HAYA CARGADO EL BANCO
# ============================================================

if not st.session_state.get(
    "banco_101_cargado",
    False
):

    st.warning(
        "Primero debe cargar el Banco General "
        "desde 10.1."
    )

else:

    # ========================================================
    # RECUPERAR FUENTE DE 10.1
    # ========================================================

    df_banco_101 = st.session_state.get(
        "df_banco_101",
        pd.DataFrame()
    ).copy()


    # ========================================================
    # FUNCIÓN DE ANÁLISIS
    # ========================================================

    def analizar_disponibilidad_102(
        df_banco_101
    ):

        # ----------------------------------------------------
        # COLUMNAS NECESARIAS
        # ----------------------------------------------------

        columnas_requeridas_102 = [
            "Pregunta_ID",
            "Modulo",
            "Tipo_Relacion",
            "Nivel",
            "Estado"
        ]

        faltantes_102 = [
            columna
            for columna in columnas_requeridas_102
            if columna not in df_banco_101.columns
        ]

        if faltantes_102:

            st.error(
                "10.2 ERROR: faltan columnas en "
                "df_banco_101: "
                + ", ".join(faltantes_102)
            )

            return None, None


        # ----------------------------------------------------
        # SOLO APROBADAS
        # ----------------------------------------------------

        df_aprobadas_102 = df_banco_101[
            df_banco_101["Estado"]
            .astype(str)
            .str.strip()
            .str.upper()
            == "APROBADA"
        ].copy()


        # ----------------------------------------------------
        # VALIDAR ESTRUCTURA
        # ----------------------------------------------------

        df_aprobadas_102 = df_aprobadas_102[
            (
                df_aprobadas_102["Pregunta_ID"]
                .astype(str)
                .str.strip()
                != ""
            )
            &
            (
                df_aprobadas_102["Modulo"]
                .astype(str)
                .str.strip()
                != ""
            )
            &
            (
                df_aprobadas_102["Tipo_Relacion"]
                .astype(str)
                .str.strip()
                != ""
            )
            &
            (
                df_aprobadas_102["Nivel"]
                .astype(str)
                .str.strip()
                != ""
            )
        ].copy()


        # ----------------------------------------------------
        # TOTALIZAR POR:
        #
        # MODULO
        # TIPO_RELACION
        # NIVEL
        # ----------------------------------------------------

        if df_aprobadas_102.empty:

            resumen_102 = pd.DataFrame(
                columns=[
                    "Modulo",
                    "Tipo_Relacion",
                    "Nivel",
                    "Preguntas_Disponibles",
                    "Evaluaciones_Completas",
                    "Preguntas_Sobrantes"
                ]
            )

        else:

            resumen_102 = (
                df_aprobadas_102
                .groupby(
                    [
                        "Modulo",
                        "Tipo_Relacion",
                        "Nivel"
                    ],
                    dropna=False
                )
                .size()
                .reset_index(
                    name="Preguntas_Disponibles"
                )
            )


            # ------------------------------------------------
            # EVALUACIONES COMPLETAS DE 10
            # ------------------------------------------------

            resumen_102[
                "Evaluaciones_Completas"
            ] = (
                resumen_102[
                    "Preguntas_Disponibles"
                ]
                // 10
            )


            # ------------------------------------------------
            # SOBRANTES
            # ------------------------------------------------

            resumen_102[
                "Preguntas_Sobrantes"
            ] = (
                resumen_102[
                    "Preguntas_Disponibles"
                ]
                % 10
            )


            # ------------------------------------------------
            # ORDEN
            # ------------------------------------------------

            resumen_102 = (
                resumen_102
                .sort_values(
                    [
                        "Modulo",
                        "Tipo_Relacion",
                        "Nivel"
                    ]
                )
                .reset_index(drop=True)
            )


        return (
            df_aprobadas_102.reset_index(
                drop=True
            ),
            resumen_102
        )


    # ========================================================
    # INTERFAZ 10.2
    # ========================================================

    st.markdown(
        "## 10.2 - Análisis y disponibilidad"
    )

    st.info(
        "El banco se clasifica exclusivamente por "
        "Módulo + Tipo_Relacion + Nivel. "
        "La columna Tema no interviene."
    )


    if st.button(
        "ANALIZAR DISPONIBILIDAD DEL BANCO",
        key="analizar_disponibilidad_102"
    ):

        resultado_102 = (
            analizar_disponibilidad_102(
                df_banco_101
            )
        )

        if resultado_102[0] is not None:

            (
                df_aprobadas_102,
                resumen_102
            ) = resultado_102


            # ------------------------------------------------
            # GUARDAR RESULTADOS
            # ------------------------------------------------

            st.session_state[
                "df_aprobadas_102"
            ] = df_aprobadas_102.copy()

            st.session_state[
                "resumen_disponibilidad_102"
            ] = resumen_102.copy()

            st.session_state[
                "analisis_102_realizado"
            ] = True


    # ========================================================
    # MOSTRAR RESULTADOS
    # ========================================================

    if st.session_state.get(
        "analisis_102_realizado",
        False
    ):

        df_aprobadas_102 = st.session_state.get(
            "df_aprobadas_102",
            pd.DataFrame()
        )

        resumen_102 = st.session_state.get(
            "resumen_disponibilidad_102",
            pd.DataFrame()
        )


        # ====================================================
        # CONTADORES GENERALES
        # ====================================================

        total_aprobadas_102 = len(
            df_aprobadas_102
        )

        grupos_102 = len(
            resumen_102
        )

        evaluaciones_102 = (
            int(
                resumen_102[
                    "Evaluaciones_Completas"
                ].sum()
            )
            if not resumen_102.empty
            else 0
        )

        sobrantes_102 = (
            int(
                resumen_102[
                    "Preguntas_Sobrantes"
                ].sum()
            )
            if not resumen_102.empty
            else 0
        )


        # ====================================================
        # MÉTRICAS
        # ====================================================

        st.markdown(
            "### Estado actual del banco"
        )

        col1, col2, col3, col4 = st.columns(4)

        with col1:

            st.metric(
                "Preguntas aprobadas",
                f"{total_aprobadas_102:,}"
            )

        with col2:

            st.metric(
                "Grupos",
                f"{grupos_102:,}"
            )

        with col3:

            st.metric(
                "Evaluaciones posibles",
                f"{evaluaciones_102:,}"
            )

        with col4:

            st.metric(
                "Preguntas sobrantes",
                f"{sobrantes_102:,}"
            )


        # ====================================================
        # TABLA DE DISPONIBILIDAD
        # ====================================================

        st.markdown(
            "### Disponibilidad por Módulo / Relación / Nivel"
        )

        if resumen_102.empty:

            st.warning(
                "No hay preguntas aprobadas disponibles "
                "para estructurar evaluaciones."
            )

        else:

            st.dataframe(
                resumen_102[
                    [
                        "Modulo",
                        "Tipo_Relacion",
                        "Nivel",
                        "Preguntas_Disponibles",
                        "Evaluaciones_Completas",
                        "Preguntas_Sobrantes"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )


        # ====================================================
        # DETALLE POR MÓDULO
        # ====================================================

        if not resumen_102.empty:

            st.markdown(
                "### Resumen por módulo"
            )

            resumen_modulo_102 = (
                resumen_102
                .groupby(
                    "Modulo",
                    as_index=False
                )
                .agg(
                    Preguntas_Disponibles=(
                        "Preguntas_Disponibles",
                        "sum"
                    ),
                    Evaluaciones_Completas=(
                        "Evaluaciones_Completas",
                        "sum"
                    ),
                    Preguntas_Sobrantes=(
                        "Preguntas_Sobrantes",
                        "sum"
                    )
                )
            )

            st.dataframe(
                resumen_modulo_102,
                use_container_width=True,
                hide_index=True
            )


        # ====================================================
        # DETALLE POR RELACIÓN
        # ====================================================

        if not resumen_102.empty:

            st.markdown(
                "### Resumen por Tipo_Relacion"
            )

            resumen_relacion_102 = (
                resumen_102
                .groupby(
                    "Tipo_Relacion",
                    as_index=False
                )
                .agg(
                    Preguntas_Disponibles=(
                        "Preguntas_Disponibles",
                        "sum"
                    ),
                    Evaluaciones_Completas=(
                        "Evaluaciones_Completas",
                        "sum"
                    ),
                    Preguntas_Sobrantes=(
                        "Preguntas_Sobrantes",
                        "sum"
                    )
                )
            )

            st.dataframe(
                resumen_relacion_102,
                use_container_width=True,
                hide_index=True
            )


        # ====================================================
        # REGLAS PARA 10.3
        # ====================================================

        st.markdown(
            "### Reglas para la generación de evaluaciones"
        )

        st.write(
            "• Cada evaluación tendrá exactamente "
            "10 preguntas."
        )

        st.write(
            "• Las preguntas se agrupan por "
            "Módulo + Tipo_Relacion + Nivel."
        )

        st.write(
            "• Una pregunta utilizada en una evaluación "
            "no podrá reutilizarse en otra."
        )

        st.write(
            "• Las preguntas sobrantes permanecen "
            "disponibles."
        )

        st.write(
            "• Si un grupo tiene menos de 10 preguntas "
            "disponibles, no genera una evaluación completa."
        )

        st.write(
            "• 10.2 solamente analiza; no consume ni "
            "modifica preguntas."
        )



# ============================================================
# 10.3 - PREPARACIÓN DE EVALUACIONES
# SELECCIÓN DE EVALUACIÓN
#
# FUENTE:
#     df_banco_101
#
# NO RECARGA EL EXCEL
# NO DEPENDE DE 10.2
#
# ESTRUCTURA:
#     Módulo
#     Tipo_Relacion
#     Nivel
#
# TEMA:
#     NO SE UTILIZA
# ============================================================


# ============================================================
# CARGAR EL BANCO DESDE LA MEMORIA DE 10.1
# ============================================================

df_banco_101 = st.session_state.get(
    "df_banco_101",
    pd.DataFrame()
)


# ============================================================
# VALIDAR QUE 10.1 CARGÓ EL BANCO
# ============================================================

if df_banco_101.empty:

    st.error(
        "10.3 ERROR: no se encontró el banco general "
        "cargado por 10.1."
    )

else:

    # ========================================================
    # COLUMNAS NECESARIAS
    # ========================================================

    columnas_103 = [
        "Pregunta_ID",
        "Modulo",
        "Tipo_Relacion",
        "Nivel",
        "Pregunta",
        "Respuesta_1",
        "Respuesta_2",
        "Respuesta_3",
        "Respuesta_4",
        "Respuesta_Correcta",
        "Estado",
        "Fuente_ID"
    ]

    faltantes_103 = [
        columna
        for columna in columnas_103
        if columna not in df_banco_101.columns
    ]

    if faltantes_103:

        st.error(
            "10.3 ERROR: faltan columnas en "
            "BANCO_PREGUNTAS_GENERALES.xlsx: "
            + ", ".join(faltantes_103)
        )

    else:

        # ====================================================
        # CONSERVAR SOLO PREGUNTAS APROBADAS
        # ====================================================

        df_disponibles_103 = df_banco_101[
            df_banco_101["Estado"]
            .astype(str)
            .str.strip()
            .str.upper()
            == "APROBADA"
        ].copy()

        # ====================================================
        # PREGUNTAS YA UTILIZADAS EN EVALUACIONES
        #
        # Estas preguntas NO pueden volver a utilizarse.
        # ====================================================

        preguntas_usadas_103 = st.session_state.get(
            "preguntas_usadas_evaluaciones_103",
            set()
        )

        if not isinstance(
            preguntas_usadas_103,
            set
        ):

            preguntas_usadas_103 = set(
                preguntas_usadas_103
            )

        # ====================================================
        # ELIMINAR PREGUNTAS YA UTILIZADAS
        # ====================================================

        df_disponibles_103 = df_disponibles_103[
            ~df_disponibles_103[
                "Pregunta_ID"
            ]
            .astype(str)
            .str.strip()
            .isin(
                preguntas_usadas_103
            )
        ].copy()

        # ====================================================
        # GUARDAR DISPONIBILIDAD ACTUAL
        # ====================================================

        st.session_state[
            "df_disponibles_103"
        ] = df_disponibles_103.reset_index(
            drop=True
        )

        st.session_state[
            "preguntas_usadas_evaluaciones_103"
        ] = preguntas_usadas_103

        # ====================================================
        # ENCABEZADO
        # ====================================================

        st.markdown(
            "## 10.3 - Preparación de evaluaciones"
        )

        st.info(
            "Las evaluaciones se estructuran únicamente "
            "por Módulo, Tipo_Relacion y Nivel. "
            "El campo Tema no se utiliza."
        )

        # ====================================================
        # RESUMEN GENERAL
        # ====================================================

        total_banco_103 = len(
            df_banco_101
        )

        total_aprobadas_103 = len(
            df_banco_101[
                df_banco_101["Estado"]
                .astype(str)
                .str.strip()
                .str.upper()
                == "APROBADA"
            ]
        )

        total_disponibles_103 = len(
            df_disponibles_103
        )

        total_usadas_103 = len(
            preguntas_usadas_103
        )

        col1, col2, col3, col4 = st.columns(4)

        with col1:

            st.metric(
                "Preguntas en banco",
                total_banco_103
            )

        with col2:

            st.metric(
                "Preguntas aprobadas",
                total_aprobadas_103
            )

        with col3:

            st.metric(
                "Preguntas utilizadas",
                total_usadas_103
            )

        with col4:

            st.metric(
                "Preguntas disponibles",
                total_disponibles_103
            )

        # ====================================================
        # SI NO HAY DISPONIBLES
        # ====================================================

        if df_disponibles_103.empty:

            st.warning(
                "No hay preguntas aprobadas disponibles "
                "para generar nuevas evaluaciones."
            )

        else:

            # =================================================
            # ANÁLISIS POR MÓDULO / RELACIÓN / NIVEL
            # =================================================

            resumen_103 = (
                df_disponibles_103
                .groupby(
                    [
                        "Modulo",
                        "Tipo_Relacion",
                        "Nivel"
                    ],
                    dropna=False
                )
                .size()
                .reset_index(
                    name="Preguntas_disponibles"
                )
            )

            resumen_103[
                "Evaluaciones_maximas"
            ] = (
                resumen_103[
                    "Preguntas_disponibles"
                ]
                // 10
            )

            # Si quedan entre 1 y 9 preguntas,
            # también puede generarse una evaluación
            # parcial con esas preguntas.

            resumen_103[
                "Preguntas_restantes"
            ] = (
                resumen_103[
                    "Preguntas_disponibles"
                ]
                % 10
            )

            resumen_103[
                "Puede_generar_evaluacion"
            ] = (
                resumen_103[
                    "Preguntas_disponibles"
                ] > 0
            )

            # ================================================
            # MOSTRAR TABLA
            # ================================================

            st.markdown(
                "### Disponibilidad por módulo, relación y nivel"
            )

            st.dataframe(
                resumen_103[
                    [
                        "Modulo",
                        "Tipo_Relacion",
                        "Nivel",
                        "Preguntas_disponibles",
                        "Evaluaciones_maximas",
                        "Preguntas_restantes",
                        "Puede_generar_evaluacion"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

            # =================================================
            # SELECCIÓN DE EVALUACIÓN
            # =================================================

            st.markdown(
                "### Seleccionar evaluación a generar"
            )

            modulos_103 = sorted(
                resumen_103[
                    "Modulo"
                ]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            if modulos_103:

                modulo_seleccionado_103 = st.selectbox(
                    "Módulo",
                    modulos_103,
                    key="modulo_evaluacion_103"
                )

                resumen_modulo_103 = resumen_103[
                    resumen_103[
                        "Modulo"
                    ].astype(str)
                    == str(
                        modulo_seleccionado_103
                    )
                ].copy()

                relaciones_103 = sorted(
                    resumen_modulo_103[
                        "Tipo_Relacion"
                    ]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )

                if relaciones_103:

                    relacion_seleccionada_103 = (
                        st.selectbox(
                            "Tipo de relación",
                            relaciones_103,
                            key="relacion_evaluacion_103"
                        )
                    )

                    resumen_relacion_103 = (
                        resumen_modulo_103[
                            resumen_modulo_103[
                                "Tipo_Relacion"
                            ].astype(str)
                            ==
                            str(
                                relacion_seleccionada_103
                            )
                        ]
                        .copy()
                    )

                    niveles_103 = sorted(
                        resumen_relacion_103[
                            "Nivel"
                        ]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )

                    if niveles_103:

                        nivel_seleccionado_103 = (
                            st.selectbox(
                                "Nivel",
                                niveles_103,
                                key="nivel_evaluacion_103"
                            )
                        )

                        fila_103 = (
                            resumen_relacion_103[
                                resumen_relacion_103[
                                    "Nivel"
                                ].astype(str)
                                ==
                                str(
                                    nivel_seleccionado_103
                                )
                            ]
                        )

                        if not fila_103.empty:

                            disponibles_103 = int(
                                fila_103.iloc[0][
                                    "Preguntas_disponibles"
                                ]
                            )

                            max_evaluaciones_103 = (
                                disponibles_103 // 10
                            )

                            restantes_103 = (
                                disponibles_103 % 10
                            )

                            st.info(
                                f"Disponibles para esta "
                                f"combinación: "
                                f"{disponibles_103}"
                            )

                            st.info(
                                f"Evaluaciones completas "
                                f"posibles de 10 preguntas: "
                                f"{max_evaluaciones_103}"
                            )

                            if restantes_103 > 0:

                                st.warning(
                                    f"Quedan {restantes_103} "
                                    "preguntas que no completan "
                                    "otro bloque de 10. "
                                    "Permanecerán disponibles "
                                    "para una evaluación posterior."
                                )

                            # =================================
                            # CANTIDAD DE EVALUACIONES
                            # =================================

                            if disponibles_103 > 0:

                                cantidad_maxima_103 = max(
                                    1,
                                    (disponibles_103 + 9) // 10
                                )

                                cantidad_evaluaciones_103 = (
                                    st.number_input(
                                        "Cantidad de evaluaciones "
                                        "a preparar",
                                        min_value=1,
                                        max_value=
                                        cantidad_maxima_103,
                                        value=1,
                                        step=1,
                                        key=
                                        "cantidad_evaluaciones_103"
                                    )
                                )

                                preguntas_necesarias_103 = (
                                    min(
                                        disponibles_103,
                                        int(
                                            cantidad_evaluaciones_103
                                        ) * 10
                                    )
                                )

                                st.info(
                                    f"Se utilizarán "
                                    f"{preguntas_necesarias_103} "
                                    "preguntas. "
                                    f"Quedarán "
                                    f"{disponibles_103 - preguntas_necesarias_103} "
                                    "disponibles después."
                                )

                                # =================================
                                # GUARDAR SELECCIÓN
                                # =================================

                                if st.button(
                                    "PREPARAR EVALUACIONES",
                                    key=
                                    "preparar_evaluaciones_103"
                                ):

                                    st.session_state[
                                        "configuracion_evaluacion_103"
                                    ] = {

                                        "Modulo":
                                            modulo_seleccionado_103,

                                        "Tipo_Relacion":
                                            relacion_seleccionada_103,

                                        "Nivel":
                                            nivel_seleccionado_103,

                                        "Cantidad":
                                            int(
                                                cantidad_evaluaciones_103
                                            )
                                    }

                                    st.success(
                                        "Configuración de "
                                        "evaluación preparada."
                                    )

                                    st.info(
                                        "El siguiente módulo "
                                        "10.4 utilizará esta "
                                        "configuración para "
                                        "seleccionar las preguntas."
                                    )

# ============================================================
# 10.4 - PREPARACIÓN Y VALIDACIÓN INDIVIDUAL DE EVALUACIONES
#
# REGLAS:
#   - Trabaja únicamente con:
#       Módulo + Tipo_Relacion + Nivel
#   - No utiliza Tema para estructurar evaluaciones.
#   - Cada evaluación tiene código automático.
#   - Las preguntas se muestran individualmente.
#   - APROBADA       -> queda reservada para la evaluación.
#   - RECHAZADA      -> queda descartada.
#   - NO APLICA AÚN  -> vuelve a quedar disponible.
#   - No modifica todavía BANCO_PREGUNTAS_GENERALES.xlsx.
#   - La sincronización se hará posteriormente en /evaluaciones/
#
# ENTRADA PRINCIPAL:
#   Banco de preguntas general cargado por 10.1
#
# SALIDA:
#   Evaluaciones preparadas en session_state
# ============================================================


# ============================================================
# CONFIGURACIÓN
# ============================================================

COLUMNAS_EVALUACION_104 = [
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


# ============================================================
# OBTENER BANCO GENERAL
#
# 10.1 puede haber utilizado diferentes nombres internos.
# Buscamos primero el banco que normalmente utiliza el módulo.
# ============================================================

def obtener_banco_104():

    posibles_claves_104 = [
        "df_banco_general_101",
        "df_banco_101",
        "df_banco_102",
        "df_banco_103",
        "df_banco_104",
        "df_banco_general",
        "banco_preguntas_general",
        "df_banco"
    ]

    for clave_104 in posibles_claves_104:

        valor_104 = st.session_state.get(
            clave_104
        )

        if isinstance(valor_104, pd.DataFrame):

            if not valor_104.empty:

                if "Pregunta_ID" in valor_104.columns:

                    return valor_104.copy()

    return pd.DataFrame()


# ============================================================
# NORMALIZAR TEXTO
# ============================================================

def normalizar_texto_104(valor):

    if pd.isna(valor):

        return ""

    return (
        str(valor)
        .strip()
    )


# ============================================================
# CÓDIGO DE MÓDULO
# ============================================================

def codigo_modulo_104(modulo):

    texto = normalizar_texto_104(
        modulo
    ).lower()

    if "patolog" in texto:
        return "PAT"

    if "producto" in texto:
        return "PRO"

    if "restric" in texto:
        return "RES"

    if "complement" in texto:
        return "COM"

    # Código genérico para módulos futuros
    limpio = (
        texto
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

    if limpio:

        return limpio[:3].upper()

    return "MOD"


# ============================================================
# CÓDIGO DEL TIPO DE RELACIÓN
#
# No se utiliza Tema.
# El código identifica el tipo de evaluación.
# ============================================================

def codigo_relacion_104(tipo_relacion):

    texto = normalizar_texto_104(
        tipo_relacion
    ).lower()

    if "defin" in texto:
        return "DEF"

    if "causa" in texto:
        return "CAU"

    if "sintom" in texto:
        return "SIN"

    if (
        "descripcion" in texto
        or "descripción" in texto
    ):
        return "DES"

    if "componente" in texto:
        return "COM"

    if "accion" in texto or "acción" in texto:
        return "ACC"

    if "categoria" in texto or "categoría" in texto:
        return "CAT"

    if "alternativa" in texto:
        return "ALT"

    if "motivo" in texto:
        return "MOT"

    if (
        "precauc" in texto
        or "contraindic" in texto
        or "restric" in texto
    ):
        return "RES"

    if "complement" in texto:
        return "CMP"

    limpio = (
        texto
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

    if limpio:

        return limpio[:3].upper()

    return "REL"


# ============================================================
# 10.4 - PREPARACIÓN Y VALIDACIÓN DE EVALUACIONES
#
# REGLAS:
#
# APROBADA:
#   - La pregunta queda consumida.
#
# RECHAZADA:
#   - La pregunta sale de la evaluación.
#   - Se reemplaza inmediatamente por otra del Banco.
#   - La pregunta rechazada NO vuelve a utilizarse.
#
# NO APLICA AÚN:
#   - La pregunta sale de la evaluación.
#   - Se reemplaza inmediatamente por otra del Banco.
#   - La pregunta original vuelve a quedar disponible.
#
# VALIDAR BLOQUE:
#   - Aprueba todas las preguntas pendientes.
#
# RECHAZAR BLOQUE:
#   - Rechaza todo el bloque.
#   - Genera un nuevo bloque completo.
#
# ============================================================


# ============================================================
# SIGUIENTE CÓDIGO DE EVALUACIÓN
# ============================================================

def siguiente_codigo_evaluacion_104(
    modulo,
    tipo_relacion,
    nivel
):

    codigo_modulo = codigo_modulo_104(modulo)

    codigo_relacion = codigo_relacion_104(
        tipo_relacion
    )

    nivel_texto = normalizar_texto_104(
        nivel
    ).upper()

    if "NIVEL 1" in nivel_texto:
        codigo_nivel = "N1"

    elif "NIVEL 2" in nivel_texto:
        codigo_nivel = "N2"

    else:
        codigo_nivel = (
            nivel_texto
            .replace(" ", "")
            .replace("-", "")
        )

        if not codigo_nivel:
            codigo_nivel = "NX"

    prefijo = (
        f"EVAL-{codigo_modulo}-"
        f"{codigo_relacion}-"
        f"{codigo_nivel}-"
    )

    evaluaciones_104 = st.session_state.get(
        "evaluaciones_preparadas_104",
        []
    )

    mayor = 0

    for evaluacion in evaluaciones_104:

        codigo = normalizar_texto_104(
            evaluacion.get(
                "Evaluacion_ID",
                ""
            )
        )

        if not codigo.startswith(prefijo):
            continue

        numero = codigo.replace(
            prefijo,
            ""
        )

        try:
            mayor = max(
                mayor,
                int(numero)
            )
        except ValueError:
            pass

    return f"{prefijo}{mayor + 1:04d}"


# ============================================================
# OBTENER PREGUNTAS QUE ESTÁN REALMENTE CONSUMIDAS
#
# APROBADA:
#   consumida.
#
# PENDIENTE:
#   reservada mientras pertenece a una evaluación activa.
#
# NO APLICA AÚN:
#   NO consumida.
#
# RECHAZADA:
#   normalmente tampoco aparece aquí porque se registra
#   en el historial de reemplazos.
# ============================================================

def preguntas_no_disponibles_104():

    no_disponibles = set()

    evaluaciones_104 = st.session_state.get(
        "evaluaciones_preparadas_104",
        []
    )

    for evaluacion in evaluaciones_104:

        preguntas = evaluacion.get(
            "Preguntas",
            []
        )

        for pregunta in preguntas:

            pregunta_id = normalizar_texto_104(
                pregunta.get(
                    "Pregunta_ID",
                    ""
                )
            )

            estado = normalizar_texto_104(
                pregunta.get(
                    "Estado_Evaluacion",
                    ""
                )
            ).upper()

            if not pregunta_id:
                continue

            if estado in [
                "APROBADA",
                "PENDIENTE"
            ]:

                no_disponibles.add(
                    pregunta_id
                )

    return no_disponibles


# ============================================================
# OBTENER PREGUNTAS RECHAZADAS
#
# Las rechazadas NO deben volver a aparecer.
# ============================================================

def preguntas_rechazadas_104():

    rechazadas = set()

    evaluaciones_104 = st.session_state.get(
        "evaluaciones_preparadas_104",
        []
    )

    for evaluacion in evaluaciones_104:

        historial = evaluacion.get(
            "Historial_Reemplazos",
            []
        )

        for registro in historial:

            estado = normalizar_texto_104(
                registro.get(
                    "Resultado",
                    ""
                )
            ).upper()

            if estado != "RECHAZADA":
                continue

            pregunta_id = normalizar_texto_104(
                registro.get(
                    "Pregunta_ID",
                    ""
                )
            )

            if pregunta_id:
                rechazadas.add(
                    pregunta_id
                )

    return rechazadas


# ============================================================
# OBTENER IDS QUE NO DEBEN SER SELECCIONADOS
#
# Incluye:
#   - aprobadas
#   - pendientes
#   - rechazadas
# ============================================================

def preguntas_bloqueadas_104():

    bloqueadas = set()

    bloqueadas.update(
        preguntas_no_disponibles_104()
    )

    bloqueadas.update(
        preguntas_rechazadas_104()
    )

    return bloqueadas


# ============================================================
# FILTRAR CANDIDATOS
# ============================================================

def candidatos_disponibles_104(
    banco,
    modulo,
    tipo_relacion,
    nivel,
    ids_extra_excluir=None
):

    if banco.empty:
        return pd.DataFrame()

    candidatos = banco.copy()

    # --------------------------------------------------------
    # MÓDULO
    # --------------------------------------------------------

    candidatos = candidatos[
        candidatos["Modulo"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        str(modulo)
        .strip()
        .lower()
    ].copy()

    # --------------------------------------------------------
    # TIPO DE RELACIÓN
    # --------------------------------------------------------

    candidatos = candidatos[
        candidatos["Tipo_Relacion"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        str(tipo_relacion)
        .strip()
        .lower()
    ].copy()

    # --------------------------------------------------------
    # NIVEL
    # --------------------------------------------------------

    candidatos = candidatos[
        candidatos["Nivel"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        str(nivel)
        .strip()
        .lower()
    ].copy()

    # --------------------------------------------------------
    # SOLO APROBADAS DEL BANCO
    # --------------------------------------------------------

    candidatos = candidatos[
        candidatos["Estado"]
        .astype(str)
        .str.strip()
        .str.upper()
        ==
        "APROBADA"
    ].copy()

    # --------------------------------------------------------
    # IDS BLOQUEADOS
    # --------------------------------------------------------

    bloqueadas = preguntas_bloqueadas_104()

    if ids_extra_excluir:
        bloqueadas.update(
            ids_extra_excluir
        )

    if bloqueadas:

        candidatos = candidatos[
            ~candidatos["Pregunta_ID"]
            .astype(str)
            .str.strip()
            .isin(bloqueadas)
        ].copy()

    return candidatos


# ============================================================
# CONTAR DISPONIBILIDAD
# ============================================================

def disponibilidad_104(
    banco,
    modulo,
    tipo_relacion,
    nivel
):

    candidatos = candidatos_disponibles_104(
        banco,
        modulo,
        tipo_relacion,
        nivel
    )

    return len(candidatos)


# ============================================================
# CONSTRUIR DICCIONARIO DE PREGUNTA
# ============================================================

def construir_pregunta_104(fila):

    pregunta = {}

    for columna in COLUMNAS_EVALUACION_104:

        if columna in fila.index:

            valor = fila[columna]

            if pd.isna(valor):
                valor = ""

            pregunta[columna] = str(
                valor
            ).strip()

        else:

            pregunta[columna] = ""

    pregunta[
        "Estado_Evaluacion"
    ] = "PENDIENTE"

    pregunta[
        "Observacion_Evaluacion"
    ] = ""

    return pregunta


# ============================================================
# REGISTRAR REEMPLAZO
# ============================================================

def registrar_reemplazo_104(
    evaluacion,
    pregunta_original,
    resultado,
    observacion=""
):

    if "Historial_Reemplazos" not in evaluacion:

        evaluacion[
            "Historial_Reemplazos"
        ] = []

    evaluacion[
        "Historial_Reemplazos"
    ].append({

        "Fecha":
            pd.Timestamp.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

        "Pregunta_ID":
            pregunta_original.get(
                "Pregunta_ID",
                ""
            ),

        "Resultado":
            resultado,

        "Observacion":
            observacion
    })


# ============================================================
# OBTENER UNA PREGUNTA DE REEMPLAZO
# ============================================================

def obtener_reemplazo_104(
    banco,
    evaluacion,
    indice_pregunta
):

    ids_actuales = set()

    for indice, pregunta in enumerate(
        evaluacion.get("Preguntas", [])
    ):

        if indice == indice_pregunta:
            continue

        pregunta_id = normalizar_texto_104(
            pregunta.get(
                "Pregunta_ID",
                ""
            )
        )

        if pregunta_id:
            ids_actuales.add(
                pregunta_id
            )

    candidatos = candidatos_disponibles_104(
        banco,
        evaluacion["Modulo"],
        evaluacion["Tipo_Relacion"],
        evaluacion["Nivel"],
        ids_extra_excluir=ids_actuales
    )

    if candidatos.empty:
        return None

    candidatos = candidatos.sample(
        frac=1
    ).reset_index(drop=True)

    fila = candidatos.iloc[0]

    return construir_pregunta_104(
        fila
    )


# ============================================================
# REEMPLAZAR UNA PREGUNTA
# ============================================================

def reemplazar_pregunta_104(
    evaluacion,
    indice_pregunta,
    resultado,
    observacion=""
):

    banco = obtener_banco_104()

    preguntas = evaluacion.get(
        "Preguntas",
        []
    )

    if (
        indice_pregunta < 0
        or indice_pregunta >= len(preguntas)
    ):
        return False

    pregunta_original = preguntas[
        indice_pregunta
    ].copy()

    nueva_pregunta = obtener_reemplazo_104(
        banco,
        evaluacion,
        indice_pregunta
    )

    # --------------------------------------------------------
    # NO HAY REEMPLAZO
    # --------------------------------------------------------

    if nueva_pregunta is None:

        # Se conserva la pregunta original
        # para no reducir silenciosamente el bloque.

        pregunta_original[
            "Estado_Evaluacion"
        ] = resultado

        pregunta_original[
            "Observacion_Evaluacion"
        ] = observacion

        preguntas[
            indice_pregunta
        ] = pregunta_original

        registrar_reemplazo_104(
            evaluacion,
            pregunta_original,
            resultado,
            observacion
        )

        evaluacion[
            "Estado"
        ] = "SIN_REEMPLAZO_DISPONIBLE"

        return False

    # --------------------------------------------------------
    # GUARDAR HISTORIAL
    # --------------------------------------------------------

    registrar_reemplazo_104(
        evaluacion,
        pregunta_original,
        resultado,
        observacion
    )

    # --------------------------------------------------------
    # SI ES RECHAZADA:
    # queda registrada en historial y NO vuelve.
    #
    # SI ES NO APLICA:
    # queda registrada, pero NO queda bloqueada.
    # --------------------------------------------------------

    preguntas[
        indice_pregunta
    ] = nueva_pregunta

    evaluacion[
        "Preguntas"
    ] = preguntas

    evaluacion[
        "Cantidad_Generada"
    ] = len(preguntas)

    return True


# ============================================================
# PREPARAR UNA EVALUACIÓN
# ============================================================

def preparar_evaluacion_104(
    modulo,
    tipo_relacion,
    nivel,
    cantidad
):

    banco = obtener_banco_104()

    if banco.empty:

        st.error(
            "10.4 ERROR: no se encontró el Banco "
            "de Preguntas General en memoria."
        )

        return None

    faltantes = [
        columna
        for columna in [
            "Pregunta_ID",
            "Modulo",
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
        if columna not in banco.columns
    ]

    if faltantes:

        st.error(
            "10.4 ERROR: faltan columnas en el "
            "Banco General: "
            + ", ".join(faltantes)
        )

        return None

    candidatos = candidatos_disponibles_104(
        banco,
        modulo,
        tipo_relacion,
        nivel
    )

    candidatos = candidatos.sample(
        frac=1
    ).reset_index(drop=True)

    disponibles_104 = len(
        candidatos
    )

    if disponibles_104 == 0:

        st.warning(
            "No hay preguntas disponibles para "
            "esta combinación de Módulo + "
            "Tipo_Relacion + Nivel."
        )

        return None

    cantidad_real = min(
        int(cantidad),
        disponibles_104
    )

    seleccionadas = candidatos.head(
        cantidad_real
    )

    codigo_evaluacion = (
        siguiente_codigo_evaluacion_104(
            modulo,
            tipo_relacion,
            nivel
        )
    )

    preguntas = []

    for _, fila in seleccionadas.iterrows():

        preguntas.append(
            construir_pregunta_104(
                fila
            )
        )

    evaluacion = {

        "Evaluacion_ID":
            codigo_evaluacion,

        "Modulo":
            normalizar_texto_104(
                modulo
            ),

        "Tipo_Relacion":
            normalizar_texto_104(
                tipo_relacion
            ),

        "Nivel":
            normalizar_texto_104(
                nivel
            ),

        "Cantidad_Solicitada":
            int(cantidad),

        "Cantidad_Generada":
            len(preguntas),

        "Estado":
            "EN_VALIDACION",

        "Fecha_Creacion":
            pd.Timestamp.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

        "Preguntas":
            preguntas,

        "Historial_Reemplazos":
            []
    }

    evaluaciones = st.session_state.get(
        "evaluaciones_preparadas_104",
        []
    ).copy()

    evaluaciones.append(
        evaluacion
    )

    st.session_state[
        "evaluaciones_preparadas_104"
    ] = evaluaciones

    return evaluacion


# ============================================================
# VALIDAR BLOQUE COMPLETO
# ============================================================

def validar_bloque_completo_104(
    evaluacion
):

    for pregunta in evaluacion.get(
        "Preguntas",
        []
    ):

        if pregunta.get(
            "Estado_Evaluacion",
            "PENDIENTE"
        ) == "PENDIENTE":

            pregunta[
                "Estado_Evaluacion"
            ] = "APROBADA"

            pregunta[
                "Observacion_Evaluacion"
            ] = (
                "Validación del bloque completo."
            )

    evaluacion[
        "Estado"
    ] = "VALIDADA"


# ============================================================
# RECHAZAR BLOQUE COMPLETO
#
# Se genera un bloque nuevo completo.
# ============================================================

def rechazar_bloque_completo_104(
    indice_evaluacion
):

    evaluaciones = st.session_state.get(
        "evaluaciones_preparadas_104",
        []
    )

    if (
        indice_evaluacion < 0
        or indice_evaluacion >= len(evaluaciones)
    ):
        return

    evaluacion_original = evaluaciones[
        indice_evaluacion
    ]

    banco = obtener_banco_104()

    cantidad = evaluacion_original[
        "Cantidad_Solicitada"
    ]

    # --------------------------------------------------------
    # REGISTRAR TODO EL BLOQUE COMO RECHAZADO
    # --------------------------------------------------------

    for pregunta in evaluacion_original.get(
        "Preguntas",
        []
    ):

        registrar_reemplazo_104(
            evaluacion_original,
            pregunta,
            "RECHAZADA",
            "Bloque completo rechazado."
        )

    evaluacion_original[
        "Estado"
    ] = "BLOQUE_RECHAZADO"

    # --------------------------------------------------------
    # CREAR NUEVO BLOQUE
    # --------------------------------------------------------

    nuevo_bloque = preparar_evaluacion_104(
        evaluacion_original["Modulo"],
        evaluacion_original["Tipo_Relacion"],
        evaluacion_original["Nivel"],
        cantidad
    )

    return nuevo_bloque


# ============================================================
# INTERFAZ DE PREPARACIÓN
# ============================================================

banco_104 = obtener_banco_104()

if not banco_104.empty:

    st.markdown(
        "## 10.4 - Preparación y validación "
        "de evaluación"
    )

    st.info(
        "Las evaluaciones se construyen únicamente "
        "con preguntas APROBADAS del Banco General. "
        "Las preguntas rechazadas o marcadas como "
        "NO APLICA AÚN pueden ser reemplazadas."
    )

    modulos_104 = sorted(
        banco_104["Modulo"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[
            lambda x: x != ""
        ]
        .unique()
        .tolist()
    )

    if modulos_104:

        modulo_104 = st.selectbox(
            "Módulo",
            modulos_104,
            key="seleccion_modulo_104"
        )

        relaciones_104 = sorted(
            banco_104[
                banco_104["Modulo"]
                .astype(str)
                .str.strip()
                ==
                modulo_104
            ]["Tipo_Relacion"]
            .dropna()
            .astype(str)
            .str.strip()
            .loc[
                lambda x: x != ""
            ]
            .unique()
            .tolist()
        )

        if relaciones_104:

            tipo_relacion_104 = st.selectbox(
                "Tipo de relación",
                relaciones_104,
                key="seleccion_relacion_104"
            )

            niveles_104 = sorted(
                banco_104[
                    (
                        banco_104["Modulo"]
                        .astype(str)
                        .str.strip()
                        ==
                        modulo_104
                    )
                    &
                    (
                        banco_104["Tipo_Relacion"]
                        .astype(str)
                        .str.strip()
                        ==
                        tipo_relacion_104
                    )
                ]["Nivel"]
                .dropna()
                .astype(str)
                .str.strip()
                .loc[
                    lambda x: x != ""
                ]
                .unique()
                .tolist()
            )

            if niveles_104:

                nivel_104 = st.selectbox(
                    "Nivel",
                    niveles_104,
                    key="seleccion_nivel_104"
                )

                disponibles_actuales_104 = (
                    disponibilidad_104(
                        banco_104,
                        modulo_104,
                        tipo_relacion_104,
                        nivel_104
                    )
                )

                st.metric(
                    "Preguntas disponibles",
                    disponibles_actuales_104
                )

                cantidad_104 = st.number_input(
                    "Cantidad de preguntas "
                    "solicitadas",
                    min_value=1,
                    max_value=10,
                    value=min(
                        10,
                        max(
                            1,
                            disponibles_actuales_104
                        )
                    ),
                    step=1,
                    key="cantidad_evaluacion_104"
                )

                if st.button(
                    "PREPARAR EVALUACIÓN",
                    key="preparar_evaluacion_104"
                ):

                    preparar_evaluacion_104(
                        modulo_104,
                        tipo_relacion_104,
                        nivel_104,
                        int(cantidad_104)
                    )

            else:

                st.warning(
                    "No existen niveles disponibles "
                    "para este tipo de relación."
                )

        else:

            st.warning(
                "No existen tipos de relación "
                "disponibles para este módulo."
            )

else:

    st.warning(
        "10.4: no hay Banco General disponible "
        "en memoria. Ejecute 10.1 antes de utilizar "
        "esta sección."
    )


# ============================================================
# VALIDACIÓN DE EVALUACIONES
# ============================================================

evaluaciones_104 = st.session_state.get(
    "evaluaciones_preparadas_104",
    []
)

if evaluaciones_104:

    st.markdown(
        "## Evaluaciones preparadas"
    )

    for indice_eval_104, evaluacion_104 in enumerate(
        evaluaciones_104
    ):

        evaluacion_id_104 = evaluacion_104[
            "Evaluacion_ID"
        ]

        st.markdown(
            f"### {evaluacion_id_104}"
        )

        st.write(
            f"**Módulo:** "
            f"{evaluacion_104['Modulo']}"
        )

        st.write(
            f"**Tipo de relación:** "
            f"{evaluacion_104['Tipo_Relacion']}"
        )

        st.write(
            f"**Nivel:** "
            f"{evaluacion_104['Nivel']}"
        )

        preguntas_eval_104 = evaluacion_104[
            "Preguntas"
        ]

        # ----------------------------------------------------
        # CONTADORES
        # ----------------------------------------------------

        aprobadas_eval_104 = sum(
            1
            for pregunta in preguntas_eval_104
            if pregunta.get(
                "Estado_Evaluacion"
            ) == "APROBADA"
        )

        pendientes_eval_104 = sum(
            1
            for pregunta in preguntas_eval_104
            if pregunta.get(
                "Estado_Evaluacion",
                "PENDIENTE"
            ) == "PENDIENTE"
        )

        no_aplica_eval_104 = sum(
            1
            for pregunta in preguntas_eval_104
            if pregunta.get(
                "Estado_Evaluacion"
            ) == "NO APLICA AÚN"
        )

        col1_104, col2_104, col3_104 = (
            st.columns(3)
        )

        with col1_104:

            st.metric(
                "Aprobadas",
                aprobadas_eval_104
            )

        with col2_104:

            st.metric(
                "No aplica aún",
                no_aplica_eval_104
            )

        with col3_104:

            st.metric(
                "Pendientes",
                pendientes_eval_104
            )

        # ====================================================
        # ACCIONES SOBRE BLOQUE COMPLETO
        # ====================================================

        st.markdown(
            "#### Validación del bloque completo"
        )

        col_b1_104, col_b2_104 = st.columns(2)

        with col_b1_104:

            if st.button(
                "VALIDAR BLOQUE COMPLETO",
                key=(
                    f"validar_bloque_104_"
                    f"{indice_eval_104}"
                ),
                disabled=(
                    evaluacion_104.get(
                        "Estado"
                    ) == "VALIDADA"
                )
            ):

                validar_bloque_completo_104(
                    evaluacion_104
                )

                st.session_state[
                    "evaluaciones_preparadas_104"
                ] = evaluaciones_104

                st.success(
                    f"{evaluacion_id_104}: "
                    "bloque completo validado."
                )

                st.rerun()

        with col_b2_104:

            if st.button(
                "RECHAZAR BLOQUE COMPLETO "
                "Y GENERAR UNO NUEVO",
                key=(
                    f"rechazar_bloque_104_"
                    f"{indice_eval_104}"
                ),
                disabled=(
                    evaluacion_104.get(
                        "Estado"
                    ) == "VALIDADA"
                )
            ):

                nuevo_bloque_104 = (
                    rechazar_bloque_completo_104(
                        indice_eval_104
                    )
                )

                st.session_state[
                    "evaluaciones_preparadas_104"
                ] = evaluaciones_104

                if nuevo_bloque_104:

                    st.success(
                        "Bloque rechazado. "
                        "Se generó una nueva evaluación "
                        "con preguntas diferentes."
                    )

                else:

                    st.warning(
                        "El bloque fue rechazado, "
                        "pero no hay suficientes "
                        "preguntas disponibles para "
                        "generar un bloque nuevo."
                    )

                st.rerun()

        st.divider()

        # ====================================================
        # PREGUNTAS INDIVIDUALES
        # ====================================================

        for indice_pregunta_104, pregunta_104 in enumerate(
            preguntas_eval_104
        ):

            estado_actual_104 = pregunta_104.get(
                "Estado_Evaluacion",
                "PENDIENTE"
            )

            pregunta_id_104 = pregunta_104.get(
                "Pregunta_ID",
                ""
            )

            st.markdown(
                f"#### Pregunta "
                f"{indice_pregunta_104 + 1} "
                f"— {pregunta_id_104}"
            )

            st.write(
                pregunta_104.get(
                    "Pregunta",
                    ""
                )
            )

            st.write(
                "1. "
                + pregunta_104.get(
                    "Respuesta_1",
                    ""
                )
            )

            st.write(
                "2. "
                + pregunta_104.get(
                    "Respuesta_2",
                    ""
                )
            )

            st.write(
                "3. "
                + pregunta_104.get(
                    "Respuesta_3",
                    ""
                )
            )

            st.write(
                "4. "
                + pregunta_104.get(
                    "Respuesta_4",
                    ""
                )
            )

            st.caption(
                "Respuesta correcta: "
                + pregunta_104.get(
                    "Respuesta_Correcta",
                    ""
                )
            )

            st.caption(
                "Fuente: "
                + pregunta_104.get(
                    "Fuente_ID",
                    ""
                )
            )

            observacion_104 = st.text_input(
                "Observación",
                value=pregunta_104.get(
                    "Observacion_Evaluacion",
                    ""
                ),
                key=(
                    f"obs_104_"
                    f"{indice_eval_104}_"
                    f"{indice_pregunta_104}"
                )
            )

            st.write(
                f"Estado actual: "
                f"**{estado_actual_104}**"
            )

            col_a_104, col_r_104, col_n_104 = (
                st.columns(3)
            )

            # ------------------------------------------------
            # APROBAR
            # ------------------------------------------------

            with col_a_104:

                if st.button(
                    "APROBAR",
                    key=(
                        f"aprobar_104_"
                        f"{indice_eval_104}_"
                        f"{indice_pregunta_104}"
                    )
                ):

                    pregunta_104[
                        "Estado_Evaluacion"
                    ] = "APROBADA"

                    pregunta_104[
                        "Observacion_Evaluacion"
                    ] = observacion_104

                    evaluacion_104[
                        "Estado"
                    ] = "EN_VALIDACION"

                    st.session_state[
                        "evaluaciones_preparadas_104"
                    ] = evaluaciones_104

                    st.rerun()

            # ------------------------------------------------
            # RECHAZAR + REEMPLAZAR
            # ------------------------------------------------

            with col_r_104:

                if st.button(
                    "RECHAZAR Y REEMPLAZAR",
                    key=(
                        f"rechazar_104_"
                        f"{indice_eval_104}_"
                        f"{indice_pregunta_104}"
                    )
                ):

                    reemplazado = (
                        reemplazar_pregunta_104(
                            evaluacion_104,
                            indice_pregunta_104,
                            "RECHAZADA",
                            observacion_104
                        )
                    )

                    st.session_state[
                        "evaluaciones_preparadas_104"
                    ] = evaluaciones_104

                    if reemplazado:

                        st.success(
                            "Pregunta rechazada y "
                            "reemplazada por otra del Banco."
                        )

                    else:

                        st.warning(
                            "La pregunta fue rechazada, "
                            "pero no existe otra pregunta "
                            "disponible para reemplazarla."
                        )

                    st.rerun()

            # ------------------------------------------------
            # NO APLICA + REEMPLAZAR
            # ------------------------------------------------

            with col_n_104:

                if st.button(
                    "NO APLICA AÚN + REEMPLAZAR",
                    key=(
                        f"no_aplica_104_"
                        f"{indice_eval_104}_"
                        f"{indice_pregunta_104}"
                    )
                ):

                    reemplazado = (
                        reemplazar_pregunta_104(
                            evaluacion_104,
                            indice_pregunta_104,
                            "NO APLICA AÚN",
                            observacion_104
                        )
                    )

                    st.session_state[
                        "evaluaciones_preparadas_104"
                    ] = evaluaciones_104

                    if reemplazado:

                        st.success(
                            "Pregunta marcada como "
                            "NO APLICA AÚN y reemplazada "
                            "por otra del Banco."
                        )

                    else:

                        st.warning(
                            "No hay otra pregunta disponible "
                            "para realizar el reemplazo."
                        )

                    st.rerun()

            st.divider()

        # ====================================================
        # ESTADO FINAL
        # ====================================================

        preguntas_eval_104 = evaluacion_104[
            "Preguntas"
        ]

        pendientes_finales_104 = sum(
            1
            for pregunta in preguntas_eval_104
            if pregunta.get(
                "Estado_Evaluacion",
                "PENDIENTE"
            ) == "PENDIENTE"
        )

        aprobadas_finales_104 = sum(
            1
            for pregunta in preguntas_eval_104
            if pregunta.get(
                "Estado_Evaluacion"
            ) == "APROBADA"
        )

        if (
            pendientes_finales_104 == 0
            and aprobadas_finales_104
            == len(preguntas_eval_104)
        ):

            evaluacion_104[
                "Estado"
            ] = "VALIDADA"

            st.success(
                f"{evaluacion_id_104}: "
                "todas las preguntas actuales "
                "están aprobadas."
            )

        elif evaluacion_104.get(
            "Estado"
        ) != "BLOQUE_RECHAZADO":

            evaluacion_104[
                "Estado"
            ] = "EN_VALIDACION"

        # ====================================================
        # DISPONIBILIDAD
        # ====================================================

        banco_actual_104 = obtener_banco_104()

        disponibles_despues_104 = (
            disponibilidad_104(
                banco_actual_104,
                evaluacion_104["Modulo"],
                evaluacion_104["Tipo_Relacion"],
                evaluacion_104["Nivel"]
            )
        )

        st.info(
            "Preguntas disponibles para otra "
            "evaluación de esta combinación: "
            f"{disponibles_despues_104:,}"
        )

        st.divider()

# ============================================================
# FIN 10.4
# ============================================================




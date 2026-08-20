# ============================================================
# FITOASISTE
# APLICATIVO DE EVALUACIÓN
# NORMALIZACIÓN DE PRODUCTO - COMPONENTE - ACCIÓN GENERAL
# ============================================================

from pathlib import Path
import pandas as pd
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

        # Por ahora NO suponemos el nombre de la hoja.
        # El usuario selecciona la hoja real.
        hoja_fuente = st.selectbox(
            "Seleccione la hoja de la matriz que contiene la información:",
            libro.sheet_names,
            key="hoja_matriz_normalizacion"
        )

        df_fuente = pd.read_excel(
            ARCHIVO_MATRIZ,
            sheet_name=hoja_fuente
        )

        # Eliminamos únicamente columnas completamente vacías.
        df_fuente = df_fuente.dropna(
            axis=1,
            how="all"
        )

        st.success(
            f"✅ Hoja cargada correctamente: **{hoja_fuente}**"
        )

        st.info(
            f"Registros encontrados: **{len(df_fuente)}**  |  "
            f"Columnas encontradas: **{len(df_fuente.columns)}**"
        )

        st.write("### Columnas REALES encontradas en la hoja")

        columnas_reales = pd.DataFrame({
            "N.º": range(1, len(df_fuente.columns) + 1),
            "Nombre real de la columna": [
                str(col) for col in df_fuente.columns
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
        f"🔴 5.1 ERROR al leer la matriz: {type(e).__name__}: {e}"
    )

# ============================================================
# 5.2 NORMALIZACIÓN INICIAL DE ACCIONES GENERALES
# FUENTE EXCLUSIVA:
# A = Producto
# E = Acciones generales
# ============================================================

st.markdown("### 5.2 Normalización inicial de acciones generales")

try:

    if "df_fuente" not in locals():

        st.error(
            "🔴 5.2 ERROR: No está disponible la matriz "
            "cargada en 5.1."
        )

    elif len(df_fuente.columns) < 5:

        st.error(
            "🔴 5.2 ERROR: La hoja seleccionada tiene menos "
            "de 5 columnas. Se necesitan las columnas A y E."
        )

    else:

        # ----------------------------------------------------
        # IDENTIFICACIÓN FIJA DE LAS COLUMNAS DE ORIGEN
        # ----------------------------------------------------

        columna_producto = df_fuente.columns[0]
        columna_acciones = df_fuente.columns[4]

        st.info(
            f"📌 Columna A utilizada como PRODUCTO: "
            f"**{columna_producto}**"
        )

        st.info(
            f"📌 Columna E utilizada como ACCIONES GENERALES: "
            f"**{columna_acciones}**"
        )

        # ----------------------------------------------------
        # COPIA DE TRABAJO:
        # SOLO A Y E
        # ----------------------------------------------------

        df_trabajo = df_fuente[
            [columna_producto, columna_acciones]
        ].copy()

        registros = []

        # ----------------------------------------------------
        # PROCESAR CADA PRODUCTO
        # ----------------------------------------------------

        for _, fila in df_trabajo.iterrows():

            producto = fila[columna_producto]
            acciones = fila[columna_acciones]

            # -----------------------------------------------
            # Validar producto
            # -----------------------------------------------

            if pd.isna(producto):
                continue

            producto = str(producto).strip()

            if not producto:
                continue

            # -----------------------------------------------
            # Validar acciones
            # -----------------------------------------------

            if pd.isna(acciones):
                continue

            acciones = str(acciones).strip()

            if not acciones:
                continue

            # -----------------------------------------------
            # Separadores de acciones
            #
            # NO se utiliza ninguna otra columna.
            # -----------------------------------------------

            texto_acciones = acciones

            for separador in [
                ";",
                "|",
                "•",
                "\n"
            ]:

                texto_acciones = texto_acciones.replace(
                    separador,
                    "\n"
                )

            partes = texto_acciones.split("\n")

            # -----------------------------------------------
            # UNA ACCIÓN = UNA FILA
            # -----------------------------------------------

            for parte in partes:

                accion = str(parte).strip()

                if not accion:
                    continue

                registros.append(
                    {
                        "Nombre del producto": producto,
                        "Acción": accion
                    }
                )

        # ----------------------------------------------------
        # CREAR DATAFRAME NORMALIZADO
        # ----------------------------------------------------

        df_normalizado = pd.DataFrame(
            registros,
            columns=[
                "Nombre del producto",
                "Acción"
            ]
        )

        # ----------------------------------------------------
        # GENERAR CÓDIGO AUTOMÁTICO
        # ----------------------------------------------------

        if not df_normalizado.empty:

            df_normalizado.insert(
                0,
                "Código",
                [
                    f"AG{numero:06d}"
                    for numero in range(
                        1,
                        len(df_normalizado) + 1
                    )
                ]
            )

        # ----------------------------------------------------
        # VALIDAR RESULTADO
        # ----------------------------------------------------

        if df_normalizado.empty:

            st.error(
                "🔴 5.2 ERROR: No se encontraron registros "
                "para normalizar utilizando A y E."
            )

        else:

            cantidad_productos = (
                df_normalizado[
                    "Nombre del producto"
                ]
                .nunique()
            )

            cantidad_acciones = len(
                df_normalizado
            )

            st.success(
                f"🟢 5.2 OK: "
                f"{cantidad_productos} productos | "
                f"{cantidad_acciones} acciones independientes."
            )

            # ------------------------------------------------
            # MOSTRAR SOLO LAS 3 COLUMNAS DEFINIDAS
            # ------------------------------------------------

            st.write(
                "### Matriz de normalización provisional"
            )

            st.dataframe(
                df_normalizado,
                use_container_width=True,
                hide_index=True
            )

            # ------------------------------------------------
            # VALIDACIÓN DE PROSTENFIT
            # ------------------------------------------------

            ejemplo_prostenfit = df_normalizado[
                df_normalizado[
                    "Nombre del producto"
                ]
                .str.contains(
                    "PROSTENFIT",
                    case=False,
                    na=False
                )
            ]

            if not ejemplo_prostenfit.empty:

                st.success(
                    "🟢 VALIDACIÓN PROSTENFIT: "
                    f"{len(ejemplo_prostenfit)} "
                    "acciones independientes."
                )

                st.dataframe(
                    ejemplo_prostenfit,
                    use_container_width=True,
                    hide_index=True
                )

            # ------------------------------------------------
            # CONFIRMACIÓN DE ESTRUCTURA
            # ------------------------------------------------

            columnas_salida = list(
                df_normalizado.columns
            )

            if columnas_salida == [
                "Código",
                "Nombre del producto",
                "Acción"
            ]:

                st.success(
                    "✅ Estructura correcta: "
                    "Código | Nombre del producto | Acción"
                )

            else:

                st.error(
                    "🔴 ERROR: La estructura de salida "
                    "no coincide con las tres columnas requeridas."
                )

            st.info(
                "ℹ️ 5.2 todavía NO realiza la depuración "
                "semántica de las acciones. En esta etapa "
                "solo se extraen A y E, se separan las acciones "
                "y se genera una fila y un código por acción."
            )

except Exception as e:

    st.error(
        f"🔴 5.2 ERROR: {type(e).__name__}: {e}"
    )

# ============================================================
# 5.3 DEPURACIÓN FINAL DE ACCIONES
# Trabaja exclusivamente sobre df_normalizado generado en 5.2
# ============================================================

st.markdown("### 5.3 Depuración final de acciones")

try:

    if "df_normalizado" not in locals() or df_normalizado.empty:

        st.error(
            "🔴 5.3 ERROR: No existe una matriz normalizada "
            "provisional válida proveniente del 5.2."
        )

    else:

        # ----------------------------------------------------
        # MARCADORES QUE INDICAN QUE TERMINÓ LA ACCIÓN
        # Y COMIENZA INFORMACIÓN QUE NO DEBE PERTENECER
        # A LA ACCIÓN GENERAL.
        # ----------------------------------------------------

        marcadores_corte = [
            "COMBINACIONES:",
            "COMBINACION:",
            "FRASE DE VENTA:",
            "FRASE VENTA:",
            "MODO DE ACCIÓN:",
            "MODO DE ACCION:",
            "RECOMENDACIÓN:",
            "RECOMENDACIONES:",
            "RECOMENDACION:",
            "RECOMENDACIONES:",
            "POSOLOGÍA:",
            "POSOLOGIA:",
            "DOSIS:",
            "FORMA DE USO:",
            "MODO DE USO:",
            "INSTRUCCIONES DE USO:"
        ]

        # ----------------------------------------------------
        # MARCADORES DE TEXTO COMERCIAL / EXPLICATIVO
        # ----------------------------------------------------

        frases_comerciales = [
            "ideal para",
            "perfecto para",
            "perfecta para",
            "excelente para",
            "una excelente opción",
            "una excelente opcion",
            "recomendado para",
            "recomendada para",
            "te ayuda a",
            "ayuda a",
            "disfruta de",
            "descubre",
            "conoce",
            "lleva tu",
            "potencia tu",
            "cuida tu",
            "obtén",
            "obten",
            "logra",
            "consigue"
        ]

        # ----------------------------------------------------
        # FUNCIÓN DE DEPURACIÓN
        # ----------------------------------------------------

        def depurar_accion(texto):

            if texto is None:
                return ""

            accion_original = str(texto).strip()

            if not accion_original:
                return ""

            accion = accion_original

            # ------------------------------------------------
            # 1. Cortar todo lo que aparezca después de
            #    cualquier marcador estructural.
            # ------------------------------------------------

            posiciones = []

            accion_mayusculas = accion.upper()

            for marcador in marcadores_corte:

                posicion = accion_mayusculas.find(
                    marcador.upper()
                )

                if posicion >= 0:
                    posiciones.append(posicion)

            if posiciones:

                accion = accion[:min(posiciones)].strip()

            # ------------------------------------------------
            # 2. Eliminar espacios repetidos.
            # ------------------------------------------------

            accion = " ".join(
                accion.split()
            ).strip()

            # ------------------------------------------------
            # 3. Eliminar puntos o separadores sobrantes
            #    al final.
            # ------------------------------------------------

            while accion.endswith((".", ";", "|", "-", ":", ",")):

                accion = accion[:-1].strip()

            # ------------------------------------------------
            # 4. Si queda una frase claramente comercial,
            #    intentar conservar únicamente el texto anterior.
            #
            #    NO se elimina toda la acción automáticamente.
            # ------------------------------------------------

            accion_minusculas = accion.lower()

            posiciones_comerciales = []

            for frase in frases_comerciales:

                posicion = accion_minusculas.find(
                    frase.lower()
                )

                if posicion > 0:

                    posiciones_comerciales.append(
                        posicion
                    )

            if posiciones_comerciales:

                accion_cortada = accion[
                    :min(posiciones_comerciales)
                ].strip()

                if len(accion_cortada) >= 8:

                    accion = accion_cortada

            # ------------------------------------------------
            # 5. Limpieza final.
            # ------------------------------------------------

            accion = " ".join(
                accion.split()
            ).strip()

            while accion.endswith((".", ";", "|", "-", ":", ",")):

                accion = accion[:-1].strip()

            return accion

        # ----------------------------------------------------
        # PROCESAR TODAS LAS ACCIONES DEL 5.2
        # ----------------------------------------------------

        df_depurado = df_normalizado.copy()

        df_depurado["Acción original"] = (
            df_depurado["Acción"]
        )

        df_depurado["Acción"] = (
            df_depurado["Acción"]
            .apply(depurar_accion)
        )

        # ----------------------------------------------------
        # IDENTIFICAR CAMBIOS
        # ----------------------------------------------------

        df_depurado["_cambio"] = (
            df_depurado["Acción original"].astype(str).str.strip()
            !=
            df_depurado["Acción"].astype(str).str.strip()
        )

        cantidad_total = len(
            df_depurado
        )

        cantidad_modificadas = int(
            df_depurado["_cambio"].sum()
        )

        cantidad_sin_cambio = (
            cantidad_total
            - cantidad_modificadas
        )

        # ----------------------------------------------------
        # NO CONSERVAR ACCIONES VACÍAS
        # ----------------------------------------------------

        df_depurado = df_depurado[
            df_depurado["Acción"].astype(str).str.strip() != ""
        ].copy()

        cantidad_eliminadas = (
            cantidad_total
            - len(df_depurado)
        )

        # ----------------------------------------------------
        # REGENERAR CÓDIGOS
        # ----------------------------------------------------

        df_depurado = df_depurado[
            [
                "Nombre del producto",
                "Acción"
            ]
        ].copy()

        df_depurado.insert(
            0,
            "Código",
            [
                f"AG{numero:06d}"
                for numero in range(
                    1,
                    len(df_depurado) + 1
                )
            ]
        )

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        st.success(
            f"🟢 5.3 TERMINADO: "
            f"**{len(df_depurado)} acciones procesadas**."
        )

        st.info(
            f"Sin cambios: **{cantidad_sin_cambio}** | "
            f"Depuradas: **{cantidad_modificadas}** | "
            f"Eliminadas por quedar vacías: **{cantidad_eliminadas}**"
        )

        # ----------------------------------------------------
        # MOSTRAR MATRIZ FINAL DE ESTA ETAPA
        # ----------------------------------------------------

        st.write(
            "### Matriz después de la depuración"
        )

        st.dataframe(
            df_depurado,
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # MOSTRAR SOLO UNA MUESTRA DE LOS CAMBIOS
        # NO SE REVISA UNA POR UNA.
        # ----------------------------------------------------

        cambios = df_normalizado.copy()

        cambios["Acción depurada"] = (
            df_depurado["Acción"]
            if len(df_depurado) == len(df_normalizado)
            else ""
        )

        # Crear nuevamente la comparación directamente
        cambios["Acción depurada"] = (
            cambios["Acción"]
            .apply(depurar_accion)
        )

        cambios = cambios[
            cambios["Acción"].astype(str).str.strip()
            !=
            cambios["Acción depurada"].astype(str).str.strip()
        ]

        if not cambios.empty:

            st.write(
                "### Muestra automática de acciones modificadas"
            )

            muestra = cambios[
                [
                    "Nombre del producto",
                    "Acción",
                    "Acción depurada"
                ]
            ].head(15).copy()

            muestra.columns = [
                "Producto",
                "Antes",
                "Después"
            ]

            st.dataframe(
                muestra,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "ℹ️ No se detectaron modificaciones "
                "en esta etapa de depuración."
            )

        st.success(
            "✅ La depuración se realizó automáticamente "
            "sobre todas las acciones generadas por 5.2. "
            "No se modificó la matriz original."
        )

        st.info(
            "ℹ️ Esta etapa trabaja exclusivamente sobre "
            "Código | Nombre del producto | Acción."
        )

except Exception as e:

    st.error(
        f"🔴 5.3 ERROR: {type(e).__name__}: {e}"
    )

# ============================================================
# 5.4 CLASIFICACIÓN SEMÁNTICA DE ACCIONES
# ============================================================

st.markdown("### 5.4 Clasificación semántica de acciones")

try:

    import numpy as np
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    # --------------------------------------------------------
    # VALIDAR RESULTADO DEL 5.3
    # --------------------------------------------------------

    if "df_depurado" not in locals() or df_depurado.empty:

        st.error(
            "🔴 5.4 ERROR: No existe df_depurado "
            "proveniente del 5.3."
        )

    else:

        # ----------------------------------------------------
        # DATAFRAME BASE
        # ----------------------------------------------------

        df_ml = df_depurado[
            [
                "Código",
                "Nombre del producto",
                "Acción"
            ]
        ].copy()

        df_ml["Código"] = (
            df_ml["Código"]
            .astype(str)
            .str.strip()
        )

        df_ml["Nombre del producto"] = (
            df_ml["Nombre del producto"]
            .astype(str)
            .str.strip()
        )

        df_ml["Acción"] = (
            df_ml["Acción"]
            .astype(str)
            .str.strip()
        )

        df_ml = df_ml[
            df_ml["Acción"] != ""
        ].drop_duplicates(
            subset=["Código"]
        ).reset_index(drop=True)

        # ----------------------------------------------------
        # CATEGORÍAS
        # ----------------------------------------------------

        categorias_54 = [
            "ACCIÓN GENERAL",
            "COMPONENTE + FUNCIÓN",
            "RECOMENDACIÓN / COMPLEMENTO",
            "USO / POSOLOGÍA / PRECAUCIÓN",
            "COMERCIAL"
        ]

        # ----------------------------------------------------
        # MEMORIA DEL ENTRENAMIENTO
        # ----------------------------------------------------

        if "entrenamiento_acciones_54" not in st.session_state:

            st.session_state.entrenamiento_acciones_54 = {}

        etiquetas = (
            st.session_state.entrenamiento_acciones_54
        )

        # ----------------------------------------------------
        # MÁXIMO 50
        # ----------------------------------------------------

        total_etiquetados = len(etiquetas)

        if total_etiquetados < 50:

            pendientes = df_ml[
                ~df_ml["Código"].isin(
                    etiquetas.keys()
                )
            ].copy()

            cupos = min(
                50 - total_etiquetados,
                len(pendientes)
            )

            if cupos > 0:

                posiciones = np.linspace(
                    0,
                    len(pendientes) - 1,
                    cupos,
                    dtype=int
                )

                muestra = (
                    pendientes
                    .iloc[posiciones]
                    .drop_duplicates(
                        subset=["Código"]
                    )
                )

            else:

                muestra = pd.DataFrame()

        else:

            muestra = pd.DataFrame()

        # ----------------------------------------------------
        # ENTRENAMIENTO MANUAL
        # ----------------------------------------------------

        st.write(
            "### Ejemplos para entrenamiento"
        )

        st.info(
            "El sistema utilizará entre 10 y 50 ejemplos. "
            "Cuando llegue a 50 no mostrará más."
        )

        if not muestra.empty:

            for _, fila in muestra.iterrows():

                codigo = str(
                    fila["Código"]
                )

                producto = str(
                    fila["Nombre del producto"]
                )

                accion = str(
                    fila["Acción"]
                )

                st.markdown(
                    f"**{codigo} — {producto}**"
                )

                st.write(
                    f"Acción: **{accion}**"
                )

                opcion = st.radio(
                    "¿Qué representa este texto?",
                    categorias_54,
                    index=None,
                    key=f"clasificacion_54_{codigo}"
                )

                if opcion is not None:

                    st.session_state.entrenamiento_acciones_54[
                        codigo
                    ] = opcion

                st.divider()

        # ----------------------------------------------------
        # ACTUALIZAR CONTADORES
        # ----------------------------------------------------

        total_etiquetados = len(
            st.session_state.entrenamiento_acciones_54
        )

        conteo = {
            categoria: 0
            for categoria in categorias_54
        }

        for valor in (
            st.session_state
            .entrenamiento_acciones_54
            .values()
        ):

            if valor in conteo:

                conteo[valor] += 1

        st.write(
            f"**Ejemplos clasificados: "
            f"{total_etiquetados}/50**"
        )

        st.write(
            " | ".join(
                [
                    f"{categoria}: **{conteo[categoria]}**"
                    for categoria in categorias_54
                ]
            )
        )

        # ----------------------------------------------------
        # ENTRENAR DESDE 10 EJEMPLOS
        # ----------------------------------------------------

        if total_etiquetados >= 10:

            # -----------------------------------------------
            # PREPARAR ENTRENAMIENTO
            # -----------------------------------------------

            codigos_entrenamiento = list(
                st.session_state
                .entrenamiento_acciones_54
                .keys()
            )

            df_entrenamiento = df_ml[
                df_ml["Código"].isin(
                    codigos_entrenamiento
                )
            ].copy()

            df_entrenamiento["Etiqueta"] = (
                df_entrenamiento["Código"]
                .map(
                    st.session_state
                    .entrenamiento_acciones_54
                )
            )

            # -----------------------------------------------
            # TEXTO DE ENTRENAMIENTO
            #
            # IMPORTANTE:
            # producto + acción
            #
            # Esto permite diferenciar, por ejemplo:
            #
            # COLÁGENO PLUS
            # Biotina (cabello y uñas)
            #
            # de
            #
            # BIOTIN
            # Soporte estructural del cabello y uñas
            # -----------------------------------------------

            df_entrenamiento["Texto ML"] = (
                "PRODUCTO: "
                + df_entrenamiento[
                    "Nombre del producto"
                ]
                + " ACCIÓN: "
                + df_entrenamiento[
                    "Acción"
                ]
            )

            df_ml["Texto ML"] = (
                "PRODUCTO: "
                + df_ml[
                    "Nombre del producto"
                ]
                + " ACCIÓN: "
                + df_ml[
                    "Acción"
                ]
            )

            clases_presentes = (
                df_entrenamiento[
                    "Etiqueta"
                ].nunique()
            )

            # -----------------------------------------------
            # VERIFICAR QUE EXISTAN AL MENOS 2 CATEGORÍAS
            # -----------------------------------------------

            if clases_presentes < 2:

                st.warning(
                    "⚠️ Ya hay 10 ejemplos, pero todos "
                    "pertenecen a la misma categoría. "
                    "Clasifique al menos un ejemplo "
                    "de otra categoría para poder entrenar."
                )

            else:

                # -------------------------------------------
                # MODELO
                # -------------------------------------------

                modelo_acciones_54 = Pipeline(
                    [
                        (
                            "tfidf",
                            TfidfVectorizer(
                                lowercase=True,
                                strip_accents="unicode",
                                ngram_range=(1, 2),
                                max_features=8000,
                                sublinear_tf=True
                            )
                        ),
                        (
                            "clasificador",
                            LogisticRegression(
                                max_iter=3000,
                                class_weight="balanced"
                            )
                        )
                    ]
                )

                # -------------------------------------------
                # ENTRENAR
                # -------------------------------------------

                modelo_acciones_54.fit(
                    df_entrenamiento["Texto ML"],
                    df_entrenamiento["Etiqueta"]
                )

                # -------------------------------------------
                # CLASIFICAR TODAS LAS ACCIONES
                # -------------------------------------------

                probabilidades = (
                    modelo_acciones_54
                    .predict_proba(
                        df_ml["Texto ML"]
                    )
                )

                predicciones = (
                    modelo_acciones_54
                    .predict(
                        df_ml["Texto ML"]
                    )
                )

                confianza = np.max(
                    probabilidades,
                    axis=1
                )

                df_resultado_54 = df_ml.copy()

                df_resultado_54[
                    "Clasificación IA"
                ] = predicciones

                df_resultado_54[
                    "Confianza IA"
                ] = confianza

                # -------------------------------------------
                # ESTADO
                # -------------------------------------------

                df_resultado_54[
                    "Estado IA"
                ] = np.where(
                    confianza >= 0.70,
                    "ALTA",
                    np.where(
                        confianza >= 0.50,
                        "MEDIA",
                        "REVISAR"
                    )
                )

                # -------------------------------------------
                # GUARDAR RESULTADO
                # -------------------------------------------

                st.session_state[
                    "df_resultado_54"
                ] = df_resultado_54

                # -------------------------------------------
                # ESTADÍSTICAS
                # -------------------------------------------

                total = len(
                    df_resultado_54
                )

                alta = int(
                    (
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "ALTA"
                    ).sum()
                )

                media = int(
                    (
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "MEDIA"
                    ).sum()
                )

                revisar = int(
                    (
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "REVISAR"
                    ).sum()
                )

                st.success(
                    "🟢 Modelo entrenado correctamente."
                )

                st.info(
                    f"Acciones procesadas: **{total}** | "
                    f"Alta confianza: **{alta}** | "
                    f"Confianza media: **{media}** | "
                    f"Revisar: **{revisar}**"
                )

                # -------------------------------------------
                # DISTRIBUCIÓN
                # -------------------------------------------

                st.write(
                    "### Distribución de clasificación"
                )

                for categoria in categorias_54:

                    cantidad = int(
                        (
                            df_resultado_54[
                                "Clasificación IA"
                            ]
                            == categoria
                        ).sum()
                    )

                    st.write(
                        f"- {categoria}: **{cantidad}**"
                    )

                # -------------------------------------------
                # SOLO MOSTRAR LOS CASOS MÁS DUDOSOS
                # -------------------------------------------

                df_revision_54 = (
                    df_resultado_54[
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "REVISAR"
                    ]
                    .sort_values(
                        "Confianza IA"
                    )
                    .head(30)
                )

                if not df_revision_54.empty:

                    st.warning(
                        "⚠️ Estos son solamente los "
                        "30 casos de menor confianza."
                    )

                    st.dataframe(
                        df_revision_54[
                            [
                                "Código",
                                "Nombre del producto",
                                "Acción",
                                "Clasificación IA",
                                "Confianza IA"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                # -------------------------------------------
                # MATRIZ FINAL DEL PRIMER PROCESO
                # SOLO 3 COLUMNAS
                # -------------------------------------------

                st.write(
                    "### Matriz de normalización"
                )

                df_normalizacion_54 = (
                    df_resultado_54[
                        [
                            "Código",
                            "Nombre del producto",
                            "Acción"
                        ]
                    ].copy()
                )

                st.dataframe(
                    df_normalizacion_54,
                    use_container_width=True,
                    hide_index=True
                )

        else:

            faltan = (
                10 - total_etiquetados
            )

            st.warning(
                f"⚠️ Faltan **{faltan}** ejemplos "
                "para iniciar el entrenamiento."
            )

except Exception as e:

    st.error(
        f"🔴 5.4 ERROR: {type(e).__name__}: {e}"
    )


# ============================================================
# 5.4 ENTRENAMIENTO SEMÁNTICO REPRESENTATIVO
# ============================================================

st.markdown("### 5.4 Entrenamiento semántico de acciones")

try:

    import numpy as np
    import pandas as pd

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    # ========================================================
    # 1. VALIDAR DATAFRAME DEL 5.3
    # ========================================================

    if "df_depurado" not in locals() or df_depurado.empty:

        st.error(
            "🔴 5.4 ERROR: No existe df_depurado "
            "proveniente del 5.3."
        )

    else:

        # ====================================================
        # 2. CREAR DATAFRAME DE TRABAJO
        # ====================================================

        df_ml = df_depurado[
            [
                "Código",
                "Nombre del producto",
                "Acción"
            ]
        ].copy()

        df_ml["Código"] = (
            df_ml["Código"]
            .astype(str)
            .str.strip()
        )

        df_ml["Nombre del producto"] = (
            df_ml["Nombre del producto"]
            .astype(str)
            .str.strip()
        )

        df_ml["Acción"] = (
            df_ml["Acción"]
            .astype(str)
            .str.strip()
        )

        df_ml = df_ml[
            df_ml["Acción"] != ""
        ].copy()

        df_ml = (
            df_ml
            .drop_duplicates(subset=["Código"])
            .reset_index(drop=True)
        )

        # ====================================================
        # 3. CATEGORÍAS
        # ====================================================

        CATEGORIAS_54 = [
            "ACCIÓN GENERAL",
            "COMPONENTE + FUNCIÓN",
            "RECOMENDACIÓN / COMPLEMENTO",
            "USO / POSOLOGÍA / PRECAUCIÓN",
            "COMERCIAL"
        ]

        # ====================================================
        # 4. MEMORIA DE ENTRENAMIENTO
        # ====================================================

        if "entrenamiento_54" not in st.session_state:

            st.session_state.entrenamiento_54 = {}

        entrenamiento = (
            st.session_state.entrenamiento_54
        )

        # ====================================================
        # 5. MÁXIMO ABSOLUTO DE 50
        # ====================================================

        total_entrenamiento = len(
            entrenamiento
        )

        # ====================================================
        # 6. SELECCIÓN DE EJEMPLOS
        #
        # No se seleccionan simplemente los primeros 50.
        #
        # Se intenta obtener una muestra variada de:
        # - productos diferentes
        # - acciones diferentes
        # - textos cortos y largos
        # ====================================================

        if total_entrenamiento < 50:

            disponibles = df_ml[
                ~df_ml["Código"].isin(
                    entrenamiento.keys()
                )
            ].copy()

            cupos = min(
                50 - total_entrenamiento,
                len(disponibles)
            )

            if cupos > 0:

                # --------------------------------------------
                # QUITAR DUPLICADOS SEMÁNTICOS SIMPLES
                # --------------------------------------------

                disponibles["_texto_normal"] = (
                    disponibles["Acción"]
                    .str.lower()
                    .str.replace(
                        r"\s+",
                        " ",
                        regex=True
                    )
                    .str.strip()
                )

                disponibles = (
                    disponibles
                    .drop_duplicates(
                        subset=["_texto_normal"]
                    )
                    .drop(columns=["_texto_normal"])
                )

                # --------------------------------------------
                # DISTRIBUCIÓN A LO LARGO DEL DATAFRAME
                # --------------------------------------------

                if len(disponibles) <= cupos:

                    muestra = disponibles.copy()

                else:

                    posiciones = np.linspace(
                        0,
                        len(disponibles) - 1,
                        cupos,
                        dtype=int
                    )

                    muestra = (
                        disponibles
                        .iloc[posiciones]
                        .copy()
                    )

            else:

                muestra = pd.DataFrame()

        else:

            muestra = pd.DataFrame()

        # ====================================================
        # 7. MOSTRAR EJEMPLOS PARA CLASIFICACIÓN
        # ====================================================

        st.write(
            "### Entrenamiento inicial"
        )

        st.info(
            "Máximo: **50 ejemplos**. "
            "El modelo clasifica cada texto según "
            "la categoría que corresponda."
        )

        if not muestra.empty:

            for _, fila in muestra.iterrows():

                codigo = str(
                    fila["Código"]
                )

                producto = str(
                    fila["Nombre del producto"]
                )

                accion = str(
                    fila["Acción"]
                )

                st.markdown(
                    f"**{codigo} — {producto}**"
                )

                st.write(
                    f"**Texto:** {accion}"
                )

                seleccion = st.selectbox(
                    "Clasifique este contenido",
                    [
                        "Seleccione...",
                        *CATEGORIAS_54
                    ],
                    key=f"select_54_{codigo}"
                )

                if seleccion != "Seleccione...":

                    st.session_state.entrenamiento_54[
                        codigo
                    ] = seleccion

                st.divider()

        # ====================================================
        # 8. CONTADORES
        # ====================================================

        total_entrenamiento = len(
            st.session_state.entrenamiento_54
        )

        conteo_categorias = {
            categoria: 0
            for categoria in CATEGORIAS_54
        }

        for valor in (
            st.session_state
            .entrenamiento_54
            .values()
        ):

            if valor in conteo_categorias:

                conteo_categorias[
                    valor
                ] += 1

        st.write(
            f"**Ejemplos clasificados: "
            f"{total_entrenamiento}/50**"
        )

        for categoria in CATEGORIAS_54:

            st.write(
                f"- {categoria}: "
                f"**{conteo_categorias[categoria]}**"
            )

        # ====================================================
        # 9. ENTRENAR
        # ====================================================

        if total_entrenamiento >= 10:

            codigos = list(
                st.session_state
                .entrenamiento_54
                .keys()
            )

            df_train = df_ml[
                df_ml["Código"].isin(
                    codigos
                )
            ].copy()

            df_train["Etiqueta"] = (
                df_train["Código"]
                .map(
                    st.session_state.entrenamiento_54
                )
            )

            # -----------------------------------------------
            # VERIFICAR DIVERSIDAD
            # -----------------------------------------------

            clases = (
                df_train["Etiqueta"]
                .nunique()
            )

            if clases < 2:

                st.warning(
                    "⚠️ El modelo necesita ejemplos "
                    "de al menos dos categorías diferentes."
                )

            else:

                # -------------------------------------------
                # TEXTO UTILIZADO POR EL MODELO
                # -------------------------------------------

                df_train["Texto_Modelo"] = (
                    "PRODUCTO: "
                    + df_train[
                        "Nombre del producto"
                    ]
                    + " | ACCIÓN: "
                    + df_train[
                        "Acción"
                    ]
                )

                df_ml["Texto_Modelo"] = (
                    "PRODUCTO: "
                    + df_ml[
                        "Nombre del producto"
                    ]
                    + " | ACCIÓN: "
                    + df_ml[
                        "Acción"
                    ]
                )

                # -------------------------------------------
                # MODELO
                # -------------------------------------------

                modelo_54 = Pipeline(
                    [
                        (
                            "tfidf",
                            TfidfVectorizer(
                                lowercase=True,
                                strip_accents="unicode",
                                ngram_range=(1, 2),
                                max_features=10000,
                                sublinear_tf=True
                            )
                        ),
                        (
                            "clasificador",
                            LogisticRegression(
                                max_iter=4000,
                                class_weight="balanced"
                            )
                        )
                    ]
                )

                # -------------------------------------------
                # ENTRENAR
                # -------------------------------------------

                modelo_54.fit(
                    df_train["Texto_Modelo"],
                    df_train["Etiqueta"]
                )

                # -------------------------------------------
                # CLASIFICAR TODO
                # -------------------------------------------

                predicciones = (
                    modelo_54.predict(
                        df_ml["Texto_Modelo"]
                    )
                )

                probabilidades = (
                    modelo_54.predict_proba(
                        df_ml["Texto_Modelo"]
                    )
                )

                confianza = np.max(
                    probabilidades,
                    axis=1
                )

                # -------------------------------------------
                # SEGUNDA MEJOR CATEGORÍA
                #
                # Esto es importante para detectar casos
                # realmente ambiguos.
                # -------------------------------------------

                orden_probabilidades = np.argsort(
                    probabilidades,
                    axis=1
                )

                mejor = (
                    orden_probabilidades[:, -1]
                )

                segunda = (
                    orden_probabilidades[:, -2]
                )

                diferencia = (
                    probabilidades[
                        np.arange(
                            len(probabilidades)
                        ),
                        mejor
                    ]
                    -
                    probabilidades[
                        np.arange(
                            len(probabilidades)
                        ),
                        segunda
                    ]
                )

                # -------------------------------------------
                # CREAR RESULTADO
                # -------------------------------------------

                df_resultado_54 = df_ml[
                    [
                        "Código",
                        "Nombre del producto",
                        "Acción"
                    ]
                ].copy()

                df_resultado_54[
                    "Clasificación IA"
                ] = predicciones

                df_resultado_54[
                    "Confianza IA"
                ] = confianza

                df_resultado_54[
                    "Diferencia IA"
                ] = diferencia

                # -------------------------------------------
                # ESTADO
                #
                # No usamos solamente 70%.
                #
                # Si la categoría ganadora está claramente
                # por encima de la segunda, se acepta.
                # -------------------------------------------

                estados = []

                for conf, dif in zip(
                    confianza,
                    diferencia
                ):

                    if conf >= 0.55 and dif >= 0.15:

                        estados.append(
                            "ACEPTADA_IA"
                        )

                    elif conf >= 0.45 and dif >= 0.10:

                        estados.append(
                            "ACEPTADA_IA"
                        )

                    else:

                        estados.append(
                            "REVISAR"
                        )

                df_resultado_54[
                    "Estado IA"
                ] = estados

                # -------------------------------------------
                # GUARDAR
                # -------------------------------------------

                st.session_state[
                    "df_resultado_54"
                ] = df_resultado_54

                # =================================================
                # 10. RESULTADOS
                # =================================================

                total = len(
                    df_resultado_54
                )

                aceptadas = int(
                    (
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "ACEPTADA_IA"
                    ).sum()
                )

                revisar = int(
                    (
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "REVISAR"
                    ).sum()
                )

                st.success(
                    "🟢 Modelo entrenado y aplicado."
                )

                st.info(
                    f"Procesadas: **{total}** | "
                    f"Aceptadas automáticamente: "
                    f"**{aceptadas}** | "
                    f"Revisión: **{revisar}**"
                )

                # =================================================
                # 11. DISTRIBUCIÓN
                # =================================================

                st.write(
                    "### Distribución por categoría"
                )

                for categoria in CATEGORIAS_54:

                    cantidad = int(
                        (
                            df_resultado_54[
                                "Clasificación IA"
                            ]
                            == categoria
                        ).sum()
                    )

                    st.write(
                        f"- {categoria}: **{cantidad}**"
                    )

                # =================================================
                # 12. CASOS DUDOSOS
                # =================================================

                df_revision = (
                    df_resultado_54[
                        df_resultado_54[
                            "Estado IA"
                        ]
                        == "REVISAR"
                    ]
                    .sort_values(
                        [
                            "Diferencia IA",
                            "Confianza IA"
                        ]
                    )
                    .head(30)
                )

                if not df_revision.empty:

                    st.warning(
                        "⚠️ Muestra de los casos "
                        "realmente ambiguos."
                    )

                    st.dataframe(
                        df_revision[
                            [
                                "Código",
                                "Nombre del producto",
                                "Acción",
                                "Clasificación IA",
                                "Confianza IA",
                                "Diferencia IA"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                # =================================================
                # 13. MATRIZ DE SALIDA
                # =================================================

                st.write(
                    "### Matriz de normalización"
                )

                st.dataframe(
                    df_resultado_54[
                        [
                            "Código",
                            "Nombre del producto",
                            "Acción"
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True
                )

        else:

            st.warning(
                f"⚠️ Clasifique "
                f"{10 - total_entrenamiento} "
                "ejemplos más para iniciar el entrenamiento."
            )

except Exception as e:

    st.error(
        f"🔴 5.4 ERROR: {type(e).__name__}: {e}"
    )


# ============================================================
# 5.6 APRENDIZAJE ACTIVO ITERATIVO
# ============================================================

st.markdown("### 5.6 Aprendizaje activo iterativo")

try:

    import numpy as np
    import pandas as pd

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    # ========================================================
    # CONFIGURACIÓN
    # ========================================================

    CATEGORIAS_56 = [
        "ACCIÓN GENERAL",
        "COMPONENTE + FUNCIÓN",
        "RECOMENDACIÓN / COMPLEMENTO",
        "USO / POSOLOGÍA / PRECAUCIÓN",
        "COMERCIAL"
    ]

    MAX_EJEMPLOS_56 = 50
    CASOS_POR_CICLO_56 = 10

    # ========================================================
    # RECUPERAR DATAFRAME DEL PASO ANTERIOR
    # ========================================================

    if (
        "df_resultado_55" in st.session_state
        and not st.session_state[
            "df_resultado_55"
        ].empty
    ):

        df_base_56 = (
            st.session_state[
                "df_resultado_55"
            ].copy()
        )

    elif (
        "df_resultado_54" in st.session_state
        and not st.session_state[
            "df_resultado_54"
        ].empty
    ):

        df_base_56 = (
            st.session_state[
                "df_resultado_54"
            ].copy()
        )

    else:

        st.error(
            "🔴 5.6 ERROR: No existe información "
            "proveniente de 5.4 o 5.5."
        )

        st.stop()

    # ========================================================
    # ASEGURAR COLUMNAS BÁSICAS
    # ========================================================

    columnas_base = [
        "Código",
        "Nombre del producto",
        "Acción"
    ]

    faltantes = [
        c for c in columnas_base
        if c not in df_base_56.columns
    ]

    if faltantes:

        st.error(
            "🔴 5.6 ERROR: Faltan columnas: "
            + ", ".join(faltantes)
        )

        st.stop()

    # ========================================================
    # MEMORIA ACUMULADA
    # ========================================================

    if "aprendizaje_56" not in st.session_state:

        st.session_state.aprendizaje_56 = {}

    aprendizaje_56 = (
        st.session_state.aprendizaje_56
    )

    # ========================================================
    # RECUPERAR ENTRENAMIENTO DE 5.4
    # ========================================================

    entrenamiento_total = {}

    if "entrenamiento_54" in st.session_state:

        entrenamiento_total.update(
            st.session_state.entrenamiento_54
        )

    # ========================================================
    # RECUPERAR ENTRENAMIENTO DE 5.5
    # ========================================================

    if "aprendizaje_activo_55" in st.session_state:

        entrenamiento_total.update(
            st.session_state.aprendizaje_activo_55
        )

    # ========================================================
    # AGREGAR APRENDIZAJE NUEVO DE 5.6
    # ========================================================

    entrenamiento_total.update(
        aprendizaje_56
    )

    # ========================================================
    # LIMITAR A 50 EJEMPLOS
    # ========================================================

    if len(entrenamiento_total) > MAX_EJEMPLOS_56:

        claves = list(
            entrenamiento_total.keys()
        )[:MAX_EJEMPLOS_56]

        entrenamiento_total = {
            k: entrenamiento_total[k]
            for k in claves
        }

    # ========================================================
    # CONTADOR DE CICLO
    # ========================================================

    if "ciclo_56" not in st.session_state:

        st.session_state.ciclo_56 = 1

    ciclo_actual = (
        st.session_state.ciclo_56
    )

    # ========================================================
    # ENCABEZADO DEL CICLO
    # ========================================================

    st.write(
        f"## 🔄 Ciclo de aprendizaje {ciclo_actual}"
    )

    st.info(
        f"Registros analizados: **{len(df_base_56)}** | "
        f"Ejemplos acumulados: "
        f"**{len(entrenamiento_total)}/{MAX_EJEMPLOS_56}**"
    )

    # ========================================================
    # CREAR DATAFRAME DE ENTRENAMIENTO
    # ========================================================

    df_train = df_base_56[
        df_base_56["Código"]
        .astype(str)
        .isin(
            entrenamiento_total.keys()
        )
    ].copy()

    if not df_train.empty:

        df_train["Etiqueta"] = (
            df_train["Código"]
            .astype(str)
            .map(
                entrenamiento_total
            )
        )

        df_train = df_train[
            df_train["Etiqueta"]
            .isin(
                CATEGORIAS_56
            )
        ].copy()

        df_train["Texto_Modelo"] = (
            "PRODUCTO: "
            + df_train[
                "Nombre del producto"
            ].astype(str)
            + " | INFORMACIÓN: "
            + df_train[
                "Acción"
            ].astype(str)
        )

    # ========================================================
    # VERIFICAR ENTRENAMIENTO
    # ========================================================

    if (
        len(df_train) < 10
        or df_train["Etiqueta"].nunique() < 2
    ):

        st.warning(
            "⚠️ Todavía no hay suficientes ejemplos "
            "diversos para realizar el aprendizaje."
        )

        st.write(
            f"Ejemplos disponibles: "
            f"**{len(df_train)}**"
        )

        st.stop()

    # ========================================================
    # ENTRENAR MODELO
    # ========================================================

    modelo_56 = Pipeline(
        [
            (
                "tfidf",
                TfidfVectorizer(
                    lowercase=True,
                    strip_accents="unicode",
                    ngram_range=(1, 2),
                    max_features=12000,
                    sublinear_tf=True
                )
            ),
            (
                "clasificador",
                LogisticRegression(
                    max_iter=5000,
                    class_weight="balanced"
                )
            )
        ]
    )

    modelo_56.fit(
        df_train["Texto_Modelo"],
        df_train["Etiqueta"]
    )

    # ========================================================
    # CLASIFICAR TODA LA MATRIZ
    # ========================================================

    df_pred = df_base_56.copy()

    df_pred["Texto_Modelo"] = (
        "PRODUCTO: "
        + df_pred[
            "Nombre del producto"
        ].astype(str)
        + " | INFORMACIÓN: "
        + df_pred[
            "Acción"
        ].astype(str)
    )

    probabilidades = (
        modelo_56.predict_proba(
            df_pred["Texto_Modelo"]
        )
    )

    predicciones = (
        modelo_56.predict(
            df_pred["Texto_Modelo"]
        )
    )

    clases = (
        modelo_56
        .named_steps[
            "clasificador"
        ]
        .classes_
    )

    # ========================================================
    # PRIMERA Y SEGUNDA CATEGORÍA
    # ========================================================

    orden = np.argsort(
        probabilidades,
        axis=1
    )

    mejor = orden[:, -1]
    segunda = orden[:, -2]

    confianza_1 = (
        probabilidades[
            np.arange(
                len(probabilidades)
            ),
            mejor
        ]
    )

    confianza_2 = (
        probabilidades[
            np.arange(
                len(probabilidades)
            ),
            segunda
        ]
    )

    diferencia = (
        confianza_1
        - confianza_2
    )

    segunda_categoria = [
        clases[i]
        for i in segunda
    ]

    df_pred[
        "Clasificación IA"
    ] = predicciones

    df_pred[
        "Confianza IA"
    ] = confianza_1

    df_pred[
        "Segunda opción IA"
    ] = segunda_categoria

    df_pred[
        "Confianza segunda"
    ] = confianza_2

    df_pred[
        "Diferencia IA"
    ] = diferencia

    # ========================================================
    # DETECTAR AMBIGÜEDAD
    #
    # Se considera ambiguo cuando:
    # - las dos categorías están muy próximas, O
    # - la confianza máxima sigue siendo baja.
    # ========================================================

    df_pred[
        "Ambiguo"
    ] = (
        (
            df_pred[
                "Diferencia IA"
            ] < 0.15
        )
        |
        (
            df_pred[
                "Confianza IA"
            ] < 0.45
        )
    )

    # ========================================================
    # NO VOLVER A MOSTRAR EJEMPLOS YA APRENDIDOS
    # ========================================================

    df_pred.loc[
        df_pred["Código"]
        .astype(str)
        .isin(
            entrenamiento_total.keys()
        ),
        "Ambiguo"
    ] = False

    # ========================================================
    # GUARDAR RESULTADO DEL CICLO
    # ========================================================

    st.session_state[
        "df_resultado_56"
    ] = df_pred.copy()

    # ========================================================
    # CONTADORES
    # ========================================================

    total_56 = len(
        df_pred
    )

    ambiguos_56 = int(
        df_pred[
            "Ambiguo"
        ].sum()
    )

    # ========================================================
    # PANEL DE ESTADO
    # ========================================================

    st.markdown(
        "### 📊 Estado del aprendizaje"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Analizados",
            total_56
        )

    with col2:

        st.metric(
            "Ejemplos aprendidos",
            f"{len(entrenamiento_total)}/50"
        )

    with col3:

        st.metric(
            "Ambiguos actuales",
            ambiguos_56
        )

    # ========================================================
    # SELECCIONAR SOLO LOS 10 MÁS AMBIGUOS
    # ========================================================

    df_ambiguos_56 = (
        df_pred[
            df_pred[
                "Ambiguo"
            ]
        ]
        .sort_values(
            [
                "Diferencia IA",
                "Confianza IA"
            ]
        )
        .head(
            CASOS_POR_CICLO_56
        )
        .copy()
    )

    # ========================================================
    # BANDEJA DE APRENDIZAJE
    # ========================================================

    st.markdown(
        "### 🎯 Casos para este ciclo"
    )

    if not df_ambiguos_56.empty:

        st.warning(
            f"Hay **{ambiguos_56} casos ambiguos**. "
            f"En este ciclo se muestran solamente "
            f"**{len(df_ambiguos_56)}**."
        )

        for _, fila in (
            df_ambiguos_56.iterrows()
        ):

            codigo = str(
                fila["Código"]
            )

            producto = str(
                fila["Nombre del producto"]
            )

            accion = str(
                fila["Acción"]
            )

            propuesta = str(
                fila["Clasificación IA"]
            )

            segunda = str(
                fila["Segunda opción IA"]
            )

            confianza = float(
                fila["Confianza IA"]
            )

            confianza_2 = float(
                fila["Confianza segunda"]
            )

            st.markdown(
                f"**{codigo} — {producto}**"
            )

            st.write(
                f"**Información:** {accion}"
            )

            st.write(
                f"**Primera opción:** "
                f"{propuesta} "
                f"({confianza:.1%})"
            )

            st.write(
                f"**Segunda opción:** "
                f"{segunda} "
                f"({confianza_2:.1%})"
            )

            decision = st.selectbox(
                "Seleccione la categoría correcta",
                [
                    "Seleccione...",
                    *CATEGORIAS_56
                ],
                key=f"categoria_56_{codigo}"
            )

            if decision != "Seleccione...":

                aprendizaje_56[
                    codigo
                ] = decision

            st.divider()

    else:

        st.success(
            "🟢 No quedan casos ambiguos."
        )

    # ========================================================
    # CONTADOR DE DECISIONES DEL CICLO
    # ========================================================

    decisiones_ciclo = len(
        aprendizaje_56
    )

    st.info(
        f"Decisiones nuevas guardadas: "
        f"**{decisiones_ciclo}/{CASOS_POR_CICLO_56}**"
    )

    # ========================================================
    # BOTÓN DE APRENDIZAJE
    # ========================================================

    if decisiones_ciclo > 0:

        if st.button(
            "🧠 APRENDER ESTAS DECISIONES Y VOLVER A ANALIZAR",
            key="aprender_56"
        ):

            # -----------------------------------------------
            # Las decisiones YA están en aprendizaje_56.
            # Se conservan y se incorporan en el siguiente
            # entrenamiento.
            # -----------------------------------------------

            st.session_state.ciclo_56 += 1

            st.rerun()

    # ========================================================
    # MATRIZ DE ACCIONES GENERALES
    #
    # SOLO SE MUESTRA COMO RESULTADO,
    # NO COMO BANDEJA DE REVISIÓN.
    # ========================================================

    st.markdown(
        "### 📋 Resultado actual — Acciones generales"
    )

    df_acciones_56 = df_pred[
        df_pred[
            "Clasificación IA"
        ]
        == "ACCIÓN GENERAL"
    ].copy()

    df_acciones_56 = df_acciones_56[
        [
            "Código",
            "Nombre del producto",
            "Acción"
        ]
    ].copy()

    st.session_state[
        "df_normalizacion_final"
    ] = df_acciones_56.copy()

    st.dataframe(
        df_acciones_56,
        use_container_width=True,
        hide_index=True
    )

    st.success(
        f"Acciones generales actuales: "
        f"**{len(df_acciones_56)}**"
    )

    # ========================================================
    # RESUMEN DE CATEGORÍAS
    # ========================================================

    st.markdown(
        "### Distribución actual"
    )

    for categoria in CATEGORIAS_56:

        cantidad = int(
            (
                df_pred[
                    "Clasificación IA"
                ]
                == categoria
            ).sum()
        )

        st.write(
            f"- {categoria}: **{cantidad}**"
        )

except Exception as e:

    st.error(
        f"🔴 5.6 ERROR: {type(e).__name__}: {e}"
    )


# ============================================================
# 5.7 APRENDIZAJE ACTIVO - CICLOS DE 5 AMBIGUOS
# ============================================================

st.markdown("### 5.7 Aprendizaje activo")

try:

    import numpy as np
    import pandas as pd

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    # ========================================================
    # CONFIGURACIÓN
    # ========================================================

    CATEGORIAS_57 = [
        "ACCIÓN GENERAL",
        "COMPONENTE + FUNCIÓN",
        "RECOMENDACIÓN / COMPLEMENTO",
        "USO / POSOLOGÍA / PRECAUCIÓN",
        "COMERCIAL"
    ]

    CASOS_POR_CICLO = 5
    MAX_ENTRENAMIENTO = 100

    # ========================================================
    # 1. RECUPERAR MATRIZ
    # ========================================================

    if (
        "df_resultado_56" in st.session_state
        and not st.session_state["df_resultado_56"].empty
    ):

        df_base = st.session_state["df_resultado_56"].copy()

    elif (
        "df_resultado_55" in st.session_state
        and not st.session_state["df_resultado_55"].empty
    ):

        df_base = st.session_state["df_resultado_55"].copy()

    elif (
        "df_resultado_54" in st.session_state
        and not st.session_state["df_resultado_54"].empty
    ):

        df_base = st.session_state["df_resultado_54"].copy()

    else:

        st.error(
            "🔴 5.7 ERROR: No existe información "
            "proveniente de 5.4, 5.5 o 5.6."
        )

        st.stop()

    # ========================================================
    # 2. VALIDAR COLUMNAS
    # ========================================================

    requeridas = [
        "Código",
        "Nombre del producto",
        "Acción"
    ]

    faltantes = [
        c for c in requeridas
        if c not in df_base.columns
    ]

    if faltantes:

        st.error(
            "🔴 5.7 ERROR: Faltan columnas: "
            + ", ".join(faltantes)
        )

        st.stop()

    # ========================================================
    # 3. NORMALIZAR
    # ========================================================

    df_base["Código"] = (
        df_base["Código"]
        .astype(str)
        .str.strip()
    )

    df_base["Nombre del producto"] = (
        df_base["Nombre del producto"]
        .astype(str)
        .str.strip()
    )

    df_base["Acción"] = (
        df_base["Acción"]
        .astype(str)
        .str.strip()
    )

    df_base["Texto_Modelo_57"] = (
        "PRODUCTO: "
        + df_base["Nombre del producto"]
        + " | INFORMACIÓN: "
        + df_base["Acción"]
    )

    # ========================================================
    # 4. RECUPERAR TODO EL APRENDIZAJE ANTERIOR
    # ========================================================

    aprendizaje_total = {}

    if "entrenamiento_54" in st.session_state:

        aprendizaje_total.update(
            st.session_state.entrenamiento_54
        )

    if "aprendizaje_activo_55" in st.session_state:

        aprendizaje_total.update(
            st.session_state.aprendizaje_activo_55
        )

    if "aprendizaje_56" in st.session_state:

        aprendizaje_total.update(
            st.session_state.aprendizaje_56
        )

    # ========================================================
    # 5. MEMORIA PROPIA DE 5.7
    # ========================================================

    if "aprendizaje_57" not in st.session_state:

        st.session_state.aprendizaje_57 = {}

    aprendizaje_57 = (
        st.session_state.aprendizaje_57
    )

    aprendizaje_total.update(
        aprendizaje_57
    )

    # ========================================================
    # 6. VALIDAR CATEGORÍAS
    # ========================================================

    aprendizaje_total = {
        str(codigo): categoria
        for codigo, categoria
        in aprendizaje_total.items()
        if categoria in CATEGORIAS_57
    }

    # ========================================================
    # 7. ENTRENAMIENTO
    # ========================================================

    df_train = df_base[
        df_base["Código"].isin(
            aprendizaje_total.keys()
        )
    ].copy()

    df_train["Etiqueta"] = (
        df_train["Código"]
        .map(aprendizaje_total)
    )

    df_train = df_train[
        df_train["Etiqueta"].isin(
            CATEGORIAS_57
        )
    ].copy()

    # ========================================================
    # 8. VERIFICAR SUFICIENTE INFORMACIÓN
    # ========================================================

    if (
        len(df_train) < 10
        or df_train["Etiqueta"].nunique() < 2
    ):

        st.warning(
            "⚠️ Se necesitan al menos 10 ejemplos "
            "de entrenamiento distribuidos en "
            "mínimo 2 categorías."
        )

        st.write(
            f"Ejemplos actuales: {len(df_train)}"
        )

        st.stop()

    # ========================================================
    # 9. ENTRENAR
    # ========================================================

    modelo = Pipeline(
        [
            (
                "tfidf",
                TfidfVectorizer(
                    lowercase=True,
                    strip_accents="unicode",
                    ngram_range=(1, 2),
                    max_features=15000,
                    sublinear_tf=True
                )
            ),
            (
                "clasificador",
                LogisticRegression(
                    max_iter=5000,
                    class_weight="balanced"
                )
            )
        ]
    )

    modelo.fit(
        df_train["Texto_Modelo_57"],
        df_train["Etiqueta"]
    )

    # ========================================================
    # 10. CLASIFICAR TODA LA MATRIZ
    # ========================================================

    probabilidades = (
        modelo.predict_proba(
            df_base["Texto_Modelo_57"]
        )
    )

    predicciones = (
        modelo.predict(
            df_base["Texto_Modelo_57"]
        )
    )

    clases = (
        modelo
        .named_steps["clasificador"]
        .classes_
    )

    # ========================================================
    # 11. PRIMERA Y SEGUNDA OPCIÓN
    # ========================================================

    orden = np.argsort(
        probabilidades,
        axis=1
    )

    mejor = orden[:, -1]
    segunda = orden[:, -2]

    confianza_1 = (
        probabilidades[
            np.arange(
                len(probabilidades)
            ),
            mejor
        ]
    )

    confianza_2 = (
        probabilidades[
            np.arange(
                len(probabilidades)
            ),
            segunda
        ]
    )

    diferencia = (
        confianza_1
        - confianza_2
    )

    segunda_categoria = [
        clases[i]
        for i in segunda
    ]

    # ========================================================
    # 12. ENTROPÍA
    # ========================================================

    p = np.clip(
        probabilidades,
        1e-12,
        1
    )

    entropia = -np.sum(
        p * np.log(p),
        axis=1
    )

    # ========================================================
    # 13. RESULTADO
    # ========================================================

    df_resultado = df_base[
        [
            "Código",
            "Nombre del producto",
            "Acción"
        ]
    ].copy()

    df_resultado[
        "Clasificación IA"
    ] = predicciones

    df_resultado[
        "Confianza IA"
    ] = confianza_1

    df_resultado[
        "Segunda opción IA"
    ] = segunda_categoria

    df_resultado[
        "Confianza segunda"
    ] = confianza_2

    df_resultado[
        "Diferencia IA"
    ] = diferencia

    df_resultado[
        "Entropía IA"
    ] = entropia

    # ========================================================
    # 14. PUNTAJE DE INCERTIDUMBRE
    # ========================================================

    max_entropia = np.log(
        len(clases)
    )

    entropia_norm = (
        entropia / max_entropia
    )

    incertidumbre = (
        (1 - diferencia) * 0.50
        +
        (1 - confianza_1) * 0.30
        +
        entropia_norm * 0.20
    )

    df_resultado[
        "Incertidumbre IA"
    ] = incertidumbre

    # ========================================================
    # 15. NO CONSIDERAR AMBIGUOS LOS YA APRENDIDOS
    # ========================================================

    df_resultado[
        "Aprendido"
    ] = (
        df_resultado["Código"]
        .isin(
            aprendizaje_total.keys()
        )
    )

    # ========================================================
    # 16. DEFINIR AMBIGUOS
    #
    # IMPORTANTE:
    # No aprendidos NO significa ambiguos.
    #
    # Se toman solamente los registros con mayor
    # incertidumbre real.
    # ========================================================

    candidatos = df_resultado[
        ~df_resultado["Aprendido"]
    ].copy()

    # ========================================================
    # 17. UMBRAL DINÁMICO
    #
    # Se selecciona aproximadamente el 20% más incierto
    # como universo candidato, pero nunca se muestran
    # más de 5 por ciclo.
    # ========================================================

    if not candidatos.empty:

        limite_dinamico = candidatos[
            "Incertidumbre IA"
        ].quantile(0.80)

        candidatos[
            "Es_ambiguo"
        ] = (
            candidatos[
                "Incertidumbre IA"
            ]
            >= limite_dinamico
        )

    else:

        candidatos[
            "Es_ambiguo"
        ] = False

    # ========================================================
    # 18. LISTA REAL DE AMBIGUOS
    # ========================================================

    df_ambiguos = candidatos[
        candidatos["Es_ambiguo"]
    ].copy()

    df_ambiguos = (
        df_ambiguos
        .sort_values(
            "Incertidumbre IA",
            ascending=False
        )
        .copy()
    )

    cantidad_ambiguos = len(
        df_ambiguos
    )

    # ========================================================
    # 19. CONTADOR DE CICLO
    # ========================================================

    if "ciclo_57" not in st.session_state:

        st.session_state.ciclo_57 = 1

    # ========================================================
    # 20. MOSTRAR ESTADO
    # ========================================================

    st.markdown(
        f"## 🔄 Ciclo {st.session_state.ciclo_57}"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Total analizado",
            len(df_resultado)
        )

    with col2:

        st.metric(
            "Entrenamiento",
            len(df_train)
        )

    with col3:

        st.metric(
            "AMBIGUOS",
            cantidad_ambiguos
        )

    # ========================================================
    # 21. GUARDAR CANTIDAD ANTERIOR
    # ========================================================

    if (
        "cantidad_ambiguos_anterior_57"
        not in st.session_state
    ):

        st.session_state[
            "cantidad_ambiguos_anterior_57"
        ] = cantidad_ambiguos

    anterior = (
        st.session_state[
            "cantidad_ambiguos_anterior_57"
        ]
    )

    if st.session_state.ciclo_57 > 1:

        reduccion = (
            anterior
            - cantidad_ambiguos
        )

        if reduccion > 0:

            st.success(
                f"📉 El aprendizaje redujo "
                f"**{reduccion} ambiguos**."
            )

        elif reduccion == 0:

            st.info(
                "El número de ambiguos "
                "no cambió en este ciclo."
            )

        else:

            st.warning(
                f"El número aumentó en "
                f"{abs(reduccion)}. "
                "El modelo está encontrando nuevas "
                "incertidumbres."
            )

    # ========================================================
    # 22. SELECCIONAR SOLAMENTE 5
    # ========================================================

    muestra = (
        df_ambiguos
        .head(CASOS_POR_CICLO)
        .copy()
    )

    # ========================================================
    # 23. MOSTRAR LOS 5
    # ========================================================

    if not muestra.empty:

        st.markdown(
            "### 🎯 Revise estos 5 casos ambiguos"
        )

        st.write(
            f"Hay **{cantidad_ambiguos} ambiguos "
            f"en total**."
        )

        st.write(
            "El sistema seleccionó los 5 "
            "con mayor incertidumbre."
        )

        for _, fila in muestra.iterrows():

            codigo = str(
                fila["Código"]
            )

            producto = str(
                fila[
                    "Nombre del producto"
                ]
            )

            accion = str(
                fila["Acción"]
            )

            propuesta = str(
                fila[
                    "Clasificación IA"
                ]
            )

            segunda = str(
                fila[
                    "Segunda opción IA"
                ]
            )

            confianza = float(
                fila[
                    "Confianza IA"
                ]
            )

            confianza_2 = float(
                fila[
                    "Confianza segunda"
                ]
            )

            st.markdown(
                f"**{codigo} — {producto}**"
            )

            st.write(
                f"**Texto:** {accion}"
            )

            st.write(
                f"**IA propone:** "
                f"{propuesta} "
                f"({confianza:.1%})"
            )

            st.write(
                f"**Segunda opción:** "
                f"{segunda} "
                f"({confianza_2:.1%})"
            )

            decision = st.selectbox(
                "Seleccione la categoría correcta",
                [
                    "Seleccione...",
                    *CATEGORIAS_57
                ],
                key=(
                    "decision_57_"
                    + codigo
                )
            )

            if decision != "Seleccione...":

                aprendizaje_57[
                    codigo
                ] = decision

            st.divider()

        # ====================================================
        # 24. BOTÓN APRENDER
        # ====================================================

        decisiones = 0

        for codigo in muestra[
            "Código"
        ].astype(str):

            if codigo in aprendizaje_57:

                decisiones += 1

        st.write(
            f"Casos clasificados: "
            f"**{decisiones}/{len(muestra)}**"
        )

        if decisiones == len(muestra):

            if st.button(
                "🧠 APRENDER Y GENERAR NUEVOS AMBIGUOS",
                key="boton_aprender_57"
            ):

                # --------------------------------------------
                # GUARDAR EL ESTADO ANTERIOR
                # --------------------------------------------

                st.session_state[
                    "cantidad_ambiguos_anterior_57"
                ] = cantidad_ambiguos

                # --------------------------------------------
                # SIGUIENTE CICLO
                # --------------------------------------------

                st.session_state.ciclo_57 += 1

                st.rerun()

        else:

            st.info(
                "Clasifique los 5 casos para continuar."
            )

    else:

        st.success(
            "🎉 No quedan casos ambiguos "
            "según el criterio actual."
        )

    # ========================================================
    # 25. GUARDAR RESULTADO
    # ========================================================

    st.session_state[
        "df_resultado_57"
    ] = df_resultado.copy()

    # ========================================================
    # 26. MATRIZ FINAL DE ACCIONES GENERALES
    # ========================================================

    df_acciones = df_resultado[
        df_resultado[
            "Clasificación IA"
        ]
        == "ACCIÓN GENERAL"
    ][
        [
            "Código",
            "Nombre del producto",
            "Acción"
        ]
    ].copy()

    st.session_state[
        "df_normalizacion_final"
    ] = df_acciones.copy()

    st.markdown(
        "### 📋 Acciones generales identificadas"
    )

    st.dataframe(
        df_acciones,
        use_container_width=True,
        hide_index=True
    )

    st.success(
        f"Total acciones generales: "
        f"**{len(df_acciones)}**"
    )

except Exception as e:

    st.error(
        f"🔴 5.7 ERROR: "
        f"{type(e).__name__}: {e}"
    )

# ============================================================
# 5.8 APRENDIZAJE ACTIVO ITERATIVO
# CICLOS CERRADOS DE 5 CASOS
# ============================================================

st.markdown("### 5.8 Aprendizaje activo iterativo")

try:

    import numpy as np
    import pandas as pd

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    # ========================================================
    # CONFIGURACIÓN
    # ========================================================

    CATEGORIAS_58 = [
        "ACCIÓN GENERAL",
        "COMPONENTE + FUNCIÓN",
        "RECOMENDACIÓN / COMPLEMENTO",
        "USO / POSOLOGÍA / PRECAUCIÓN",
        "COMERCIAL"
    ]

    CASOS_POR_CICLO = 5

    # ========================================================
    # 1. RECUPERAR MATRIZ
    # ========================================================

    df_base_58 = None

    for nombre_estado in [
        "df_resultado_57",
        "df_resultado_56",
        "df_resultado_55",
        "df_resultado_54"
    ]:

        if nombre_estado in st.session_state:

            posible = st.session_state[
                nombre_estado
            ]

            if (
                isinstance(posible, pd.DataFrame)
                and not posible.empty
            ):

                df_base_58 = posible.copy()
                break

    if df_base_58 is None:

        st.error(
            "🔴 5.8 ERROR: No se encontró la matriz "
            "generada por los módulos anteriores."
        )

        st.stop()

    # ========================================================
    # 2. VALIDAR COLUMNAS
    # ========================================================

    requeridas_58 = [
        "Código",
        "Nombre del producto",
        "Acción"
    ]

    faltantes_58 = [
        c
        for c in requeridas_58
        if c not in df_base_58.columns
    ]

    if faltantes_58:

        st.error(
            "🔴 5.8 ERROR: Faltan columnas: "
            + ", ".join(faltantes_58)
        )

        st.stop()

    # ========================================================
    # 3. NORMALIZAR
    # ========================================================

    for columna in requeridas_58:

        df_base_58[columna] = (
            df_base_58[columna]
            .astype(str)
            .str.strip()
        )

    df_base_58["Texto_58"] = (
        "PRODUCTO: "
        + df_base_58[
            "Nombre del producto"
        ]
        + " | INFORMACIÓN: "
        + df_base_58[
            "Acción"
        ]
    )

    # ========================================================
    # 4. CREAR MEMORIAS PERSISTENTES DE 5.8
    # ========================================================

    if "historico_58" not in st.session_state:

        st.session_state.historico_58 = {}

    if "lote_58" not in st.session_state:

        st.session_state.lote_58 = []

    if "decisiones_lote_58" not in st.session_state:

        st.session_state.decisiones_lote_58 = {}

    if "ciclo_58" not in st.session_state:

        st.session_state.ciclo_58 = 1

    if "resultado_ciclo_58" not in st.session_state:

        st.session_state.resultado_ciclo_58 = None

    # ========================================================
    # 5. RECUPERAR APRENDIZAJE ANTERIOR
    # ========================================================

    historico = {}

    if "entrenamiento_54" in st.session_state:

        historico.update(
            st.session_state.entrenamiento_54
        )

    if "aprendizaje_activo_55" in st.session_state:

        historico.update(
            st.session_state.aprendizaje_activo_55
        )

    if "aprendizaje_56" in st.session_state:

        historico.update(
            st.session_state.aprendizaje_56
        )

    if "aprendizaje_57" in st.session_state:

        historico.update(
            st.session_state.aprendizaje_57
        )

    # ========================================================
    # 6. AGREGAR APRENDIZAJE PROPIO DE 5.8
    # ========================================================

    historico.update(
        st.session_state.historico_58
    )

    # ========================================================
    # 7. LIMPIAR CATEGORÍAS INVÁLIDAS
    # ========================================================

    historico = {
        str(codigo): categoria
        for codigo, categoria in historico.items()
        if categoria in CATEGORIAS_58
    }

    st.session_state.historico_58 = historico

    # ========================================================
    # 8. ENTRENAMIENTO ACTUAL
    # ========================================================

    df_train_58 = df_base_58[
        df_base_58["Código"].isin(
            historico.keys()
        )
    ].copy()

    df_train_58["Etiqueta_58"] = (
        df_train_58["Código"]
        .map(historico)
    )

    df_train_58 = df_train_58[
        df_train_58[
            "Etiqueta_58"
        ].isin(
            CATEGORIAS_58
        )
    ].copy()

    # ========================================================
    # 9. VALIDAR MODELO
    # ========================================================

    if (
        len(df_train_58) < 10
        or df_train_58[
            "Etiqueta_58"
        ].nunique() < 2
    ):

        st.warning(
            "⚠️ El modelo necesita al menos "
            "10 ejemplos de entrenamiento "
            "distribuidos en mínimo 2 categorías."
        )

        st.write(
            f"Ejemplos actuales: {len(df_train_58)}"
        )

        st.stop()

    # ========================================================
    # 10. ENTRENAR MODELO DESDE CERO
    # ========================================================

    modelo_58 = Pipeline(
        [
            (
                "tfidf",
                TfidfVectorizer(
                    lowercase=True,
                    strip_accents="unicode",
                    ngram_range=(1, 2),
                    max_features=15000,
                    sublinear_tf=True
                )
            ),
            (
                "clasificador",
                LogisticRegression(
                    max_iter=5000,
                    class_weight="balanced"
                )
            )
        ]
    )

    modelo_58.fit(
        df_train_58[
            "Texto_58"
        ],
        df_train_58[
            "Etiqueta_58"
        ]
    )

    # ========================================================
    # 11. ANALIZAR TODA LA MATRIZ
    # ========================================================

    probabilidades_58 = (
        modelo_58.predict_proba(
            df_base_58[
                "Texto_58"
            ]
        )
    )

    predicciones_58 = (
        modelo_58.predict(
            df_base_58[
                "Texto_58"
            ]
        )
    )

    clases_58 = (
        modelo_58
        .named_steps[
            "clasificador"
        ]
        .classes_
    )

    # ========================================================
    # 12. CALCULAR INCERTIDUMBRE
    # ========================================================

    orden_58 = np.argsort(
        probabilidades_58,
        axis=1
    )

    mejor_58 = orden_58[:, -1]

    if probabilidades_58.shape[1] > 1:

        segundo_58 = orden_58[:, -2]

    else:

        segundo_58 = mejor_58

    confianza_58 = (
        probabilidades_58[
            np.arange(
                len(probabilidades_58)
            ),
            mejor_58
        ]
    )

    confianza_2_58 = (
        probabilidades_58[
            np.arange(
                len(probabilidades_58)
            ),
            segundo_58
        ]
    )

    diferencia_58 = (
        confianza_58
        - confianza_2_58
    )

    segunda_categoria_58 = [
        clases_58[i]
        for i in segundo_58
    ]

    # ========================================================
    # 13. ENTROPÍA
    # ========================================================

    p_58 = np.clip(
        probabilidades_58,
        1e-12,
        1
    )

    entropia_58 = -np.sum(
        p_58 * np.log(p_58),
        axis=1
    )

    max_entropia_58 = np.log(
        len(clases_58)
    )

    entropia_normalizada_58 = (
        entropia_58
        / max_entropia_58
        if max_entropia_58 > 0
        else entropia_58
    )

    # ========================================================
    # 14. ÍNDICE DE INCERTIDUMBRE
    # ========================================================

    incertidumbre_58 = (
        (1 - diferencia_58) * 0.50
        +
        (1 - confianza_58) * 0.30
        +
        entropia_normalizada_58 * 0.20
    )

    # ========================================================
    # 15. CONSTRUIR RESULTADO
    # ========================================================

    df_resultado_58 = df_base_58[
        [
            "Código",
            "Nombre del producto",
            "Acción"
        ]
    ].copy()

    df_resultado_58[
        "Clasificación IA"
    ] = predicciones_58

    df_resultado_58[
        "Confianza IA"
    ] = confianza_58

    df_resultado_58[
        "Segunda opción IA"
    ] = segunda_categoria_58

    df_resultado_58[
        "Confianza segunda"
    ] = confianza_2_58

    df_resultado_58[
        "Diferencia IA"
    ] = diferencia_58

    df_resultado_58[
        "Incertidumbre IA"
    ] = incertidumbre_58

    # ========================================================
    # 16. EXCLUIR TODO LO YA APRENDIDO
    # ========================================================

    df_no_aprendido_58 = (
        df_resultado_58[
            ~df_resultado_58[
                "Código"
            ].isin(
                historico.keys()
            )
        ].copy()
    )

    # ========================================================
    # 17. DEFINIR AMBIGUOS
    #
    # NO son simplemente "pendientes".
    #
    # Se toman los registros con mayor incertidumbre.
    # ========================================================

    if not df_no_aprendido_58.empty:

        # Umbral adaptativo basado en distribución.
        # Se consideran candidatos los registros
        # ubicados en el 20% superior de incertidumbre.

        umbral_58 = (
            df_no_aprendido_58[
                "Incertidumbre IA"
            ].quantile(
                0.80
            )
        )

        df_ambiguos_58 = (
            df_no_aprendido_58[
                df_no_aprendido_58[
                    "Incertidumbre IA"
                ] >= umbral_58
            ]
            .sort_values(
                "Incertidumbre IA",
                ascending=False
            )
            .copy()
        )

    else:

        df_ambiguos_58 = (
            df_no_aprendido_58.copy()
        )

    # ========================================================
    # 18. IDENTIFICAR EL LOTE ACTUAL
    # ========================================================

    lote_actual = (
        st.session_state.lote_58
    )

    # ========================================================
    # 19. SI NO HAY LOTE, CREAR EXACTAMENTE 5
    # ========================================================

    if len(lote_actual) == 0:

        nuevos_codigos = (
            df_ambiguos_58[
                "Código"
            ]
            .astype(str)
            .head(
                CASOS_POR_CICLO
            )
            .tolist()
        )

        st.session_state.lote_58 = (
            nuevos_codigos
        )

        st.session_state[
            "decisiones_lote_58"
        ] = {}

        lote_actual = nuevos_codigos

    # ========================================================
    # 20. FILTRAR LOTE
    # ========================================================

    df_lote_58 = df_resultado_58[
        df_resultado_58[
            "Código"
        ].isin(
            lote_actual
        )
    ].copy()

    # ========================================================
    # 21. MOSTRAR RESUMEN
    # ========================================================

    cantidad_ambiguos = len(
        df_ambiguos_58
    )

    st.markdown(
        f"## 🔄 Ciclo {st.session_state.ciclo_58}"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Total analizado",
            len(df_resultado_58)
        )

    with col2:

        st.metric(
            "Ejemplos aprendidos",
            len(historico)
        )

    with col3:

        st.metric(
            "AMBIGUOS ACTUALES",
            cantidad_ambiguos
        )

    # ========================================================
    # 22. RESULTADO DEL CICLO ANTERIOR
    # ========================================================

    resultado_anterior = (
        st.session_state[
            "resultado_ciclo_58"
        ]
    )

    if resultado_anterior is not None:

        st.markdown(
            "### 📉 Resultado del aprendizaje anterior"
        )

        antes = resultado_anterior[
            "antes"
        ]

        despues = resultado_anterior[
            "despues"
        ]

        aprendidos = resultado_anterior[
            "aprendidos"
        ]

        reduccion = (
            antes - despues
        )

        col_a, col_b, col_c, col_d = st.columns(4)

        with col_a:

            st.metric(
                "Antes",
                antes
            )

        with col_b:

            st.metric(
                "Aprendidos",
                aprendidos
            )

        with col_c:

            st.metric(
                "Después",
                despues
            )

        with col_d:

            st.metric(
                "Reducción",
                reduccion
            )

    # ========================================================
    # 23. MOSTRAR LOTE CONGELADO
    # ========================================================

    st.markdown(
        "### 🎯 Casos ambiguos a revisar"
    )

    st.info(
        f"Hay **{cantidad_ambiguos} casos ambiguos** "
        f"en total."
    )

    st.warning(
        f"Este ciclo tiene exactamente "
        f"**{len(df_lote_58)} casos**."
    )

    # ========================================================
    # 24. FORMULARIO
    #
    # MUY IMPORTANTE:
    # El formulario impide que una selección parcial
    # se convierta en aprendizaje.
    # ========================================================

    with st.form(
        "formulario_lote_58"
    ):

        decisiones_formulario = {}

        for _, fila in df_lote_58.iterrows():

            codigo = str(
                fila["Código"]
            )

            st.markdown(
                f"**{codigo} — "
                f"{fila['Nombre del producto']}**"
            )

            st.write(
                f"**Texto:** "
                f"{fila['Acción']}"
            )

            st.write(
                f"**IA propone:** "
                f"{fila['Clasificación IA']} "
                f"({fila['Confianza IA']:.1%})"
            )

            st.write(
                f"**Segunda opción:** "
                f"{fila['Segunda opción IA']} "
                f"({fila['Confianza segunda']:.1%})"
            )

            decision = st.selectbox(
                "Clasificación correcta",
                [
                    "Seleccione...",
                    *CATEGORIAS_58
                ],
                key=(
                    "form_58_"
                    + codigo
                )
            )

            decisiones_formulario[
                codigo
            ] = decision

            st.divider()

        enviar = st.form_submit_button(
            "🧠 APRENDER ESTOS 5 Y RECLASIFICAR"
        )

    # ========================================================
    # 25. PROCESAR EL CICLO
    # ========================================================

    if enviar:

        incompletos = [
            codigo
            for codigo, decision
            in decisiones_formulario.items()
            if decision == "Seleccione..."
        ]

        if incompletos:

            st.error(
                "Debe clasificar los 5 casos "
                "antes de aprender."
            )

        else:

            # -----------------------------------------------
            # GUARDAR LOS 5
            # -----------------------------------------------

            for codigo, decision in (
                decisiones_formulario.items()
            ):

                st.session_state[
                    "historico_58"
                ][codigo] = decision

            # -----------------------------------------------
            # CANTIDAD DE AMBIGUOS ANTES
            # -----------------------------------------------

            ambiguos_antes = (
                cantidad_ambiguos
            )

            # -----------------------------------------------
            # CERRAR LOTE
            # -----------------------------------------------

            st.session_state[
                "lote_58"
            ] = []

            st.session_state[
                "decisiones_lote_58"
            ] = {}

            # -----------------------------------------------
            # SIGUIENTE CICLO
            # -----------------------------------------------

            st.session_state[
                "ciclo_58"
            ] += 1

            # -----------------------------------------------
            # Guardar temporalmente el valor anterior.
            # En el siguiente rerun el modelo se entrena
            # nuevamente con los 5 nuevos ejemplos y se
            # calcula la nueva cantidad de ambiguos.
            # -----------------------------------------------

            st.session_state[
                "ambiguos_antes_58"
            ] = ambiguos_antes

            st.session_state[
                "aprendidos_ultimo_ciclo_58"
            ] = len(
                decisiones_formulario
            )

            st.rerun()

    # ========================================================
    # 26. SI ACABA DE HABER UN APRENDIZAJE,
    #     REGISTRAR EL RESULTADO REAL
    # ========================================================

    if (
        "ambiguos_antes_58"
        in st.session_state
    ):

        antes = st.session_state[
            "ambiguos_antes_58"
        ]

        aprendidos = st.session_state[
            "aprendidos_ultimo_ciclo_58"
        ]

        # En este punto ya se hizo el nuevo análisis.
        despues = cantidad_ambiguos

        st.session_state[
            "resultado_ciclo_58"
        ] = {
            "antes": antes,
            "despues": despues,
            "aprendidos": aprendidos
        }

        del st.session_state[
            "ambiguos_antes_58"
        ]

        del st.session_state[
            "aprendidos_ultimo_ciclo_58"
        ]

        st.rerun()

    # ========================================================
    # 27. GUARDAR RESULTADO COMPLETO
    # ========================================================

    st.session_state[
        "df_resultado_58"
    ] = df_resultado_58.copy()

    # ========================================================
    # 28. MATRIZ DE ACCIONES GENERALES
    # ========================================================

    st.markdown(
        "### 📋 Acciones generales"
    )

    df_acciones_58 = df_resultado_58[
        df_resultado_58[
            "Clasificación IA"
        ]
        == "ACCIÓN GENERAL"
    ][
        [
            "Código",
            "Nombre del producto",
            "Acción"
        ]
    ].copy()

    st.session_state[
        "df_normalizacion_final"
    ] = df_acciones_58.copy()

    st.dataframe(
        df_acciones_58,
        use_container_width=True,
        hide_index=True
    )

    st.success(
        f"Acciones generales actuales: "
        f"{len(df_acciones_58)}"
    )

except Exception as e:

    st.error(
        f"🔴 5.8 ERROR: "
        f"{type(e).__name__}: {e}"
    )
# ============================================================
# 5.9 PRIMER DATAFRAME DE ACCIONES GENERALES
#
# Usa:
#   - df_resultado_58
#   - aprendizaje acumulado 5.4 → 5.8
#
# NO modifica 5.2, 5.3, 5.4, 5.5, 5.6, 5.7 ni 5.8
#
# Salida:
#   Código | Nombre del producto | Acción general
# ============================================================

st.markdown("### 5.9 Primer DataFrame de acciones generales")

try:

    import io
    import pandas as pd
    import numpy as np

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    # ========================================================
    # 1. TOMAR EL RESULTADO REAL DEL 5.8
    # ========================================================

    if "df_resultado_58" not in locals():

        df_resultado_58 = st.session_state.get(
            "df_resultado_58",
            None
        )

    if (
        df_resultado_58 is None
        or not isinstance(
            df_resultado_58,
            pd.DataFrame
        )
        or df_resultado_58.empty
    ):

        st.error(
            "🔴 5.9 ERROR: No existe df_resultado_58. "
            "Debe ejecutarse primero el 5.8."
        )

        st.stop()

    # ========================================================
    # 2. VALIDAR ESTRUCTURA
    # ========================================================

    columnas_base_59 = [
        "Código",
        "Nombre del producto",
        "Acción"
    ]

    faltantes_59 = [
        c
        for c in columnas_base_59
        if c not in df_resultado_58.columns
    ]

    if faltantes_59:

        st.error(
            "🔴 5.9 ERROR: Faltan columnas en "
            "df_resultado_58: "
            + ", ".join(faltantes_59)
        )

        st.stop()

    # ========================================================
    # 3. RECUPERAR TODO EL APRENDIZAJE ACUMULADO REAL
    #
    # NO se crea una memoria nueva.
    # Se utilizan exactamente las memorias existentes.
    # ========================================================

    entrenamiento_total_59 = {}

    nombres_entrenamiento_59 = [
        "entrenamiento_54",
        "entrenamiento_acciones_54",
        "aprendizaje_activo_55",
        "aprendizaje_55",
        "aprendizaje_56",
        "aprendizaje_57",
        "historico_58",
        "historico_57"
    ]

    for nombre_memoria in nombres_entrenamiento_59:

        if nombre_memoria in st.session_state:

            memoria = st.session_state[
                nombre_memoria
            ]

            if isinstance(
                memoria,
                dict
            ):

                entrenamiento_total_59.update(
                    memoria
                )

    # ========================================================
    # 4. TAMBIÉN RECUPERAR EL ENTRENAMIENTO ACUMULADO
    # QUE 5.6/5.7/5.8 PUEDAN HABER GUARDADO COMO
    # "entrenamiento_total"
    # ========================================================

    posibles_totales_59 = [
        "entrenamiento_total",
        "entrenamiento_acumulado",
        "aprendizaje_acumulado"
    ]

    for nombre_total in posibles_totales_59:

        if nombre_total in st.session_state:

            memoria_total = st.session_state[
                nombre_total
            ]

            if isinstance(
                memoria_total,
                dict
            ):

                entrenamiento_total_59.update(
                    memoria_total
                )

    # ========================================================
    # 5. CATEGORÍAS VÁLIDAS
    # ========================================================

    categorias_59 = [
        "ACCIÓN GENERAL",
        "COMPONENTE + FUNCIÓN",
        "RECOMENDACIÓN / COMPLEMENTO",
        "USO / POSOLOGÍA / PRECAUCIÓN",
        "COMERCIAL"
    ]

    entrenamiento_total_59 = {

        str(codigo): str(categoria)

        for codigo, categoria
        in entrenamiento_total_59.items()

        if str(categoria).strip()
        in categorias_59
    }

    # ========================================================
    # 6. CREAR BASE PARA EL MODELO
    # ========================================================

    df_modelo_59 = df_resultado_58[
        columnas_base_59
    ].copy()

    df_modelo_59["Código"] = (
        df_modelo_59["Código"]
        .astype(str)
        .str.strip()
    )

    df_modelo_59["Nombre del producto"] = (
        df_modelo_59["Nombre del producto"]
        .astype(str)
        .str.strip()
    )

    df_modelo_59["Acción"] = (
        df_modelo_59["Acción"]
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # 7. TEXTO QUE APRENDE EL MODELO
    #
    # Producto + acción.
    #
    # La salida continuará siendo únicamente la acción.
    # ========================================================

    df_modelo_59["Texto_modelo"] = (
        "PRODUCTO: "
        + df_modelo_59["Nombre del producto"]
        + " ACCIÓN: "
        + df_modelo_59["Acción"]
    )

    # ========================================================
    # 8. CONSTRUIR CONJUNTO DE ENTRENAMIENTO
    # ========================================================

    df_train_59 = df_modelo_59[
        df_modelo_59["Código"].isin(
            entrenamiento_total_59.keys()
        )
    ].copy()

    df_train_59["Etiqueta"] = (
        df_train_59["Código"]
        .map(
            entrenamiento_total_59
        )
    )

    df_train_59 = df_train_59[
        df_train_59["Etiqueta"].isin(
            categorias_59
        )
    ].copy()

    # ========================================================
    # 9. MOSTRAR CUÁNTO APRENDIZAJE REAL SE RECUPERÓ
    # ========================================================

    st.info(
        "🧠 Aprendizaje acumulado recuperado: "
        f"**{len(df_train_59)} ejemplos**."
    )

    if len(df_train_59) < 10:

        st.error(
            "🔴 5.9 ERROR: El aprendizaje acumulado "
            "recuperado tiene menos de 10 ejemplos."
        )

        st.stop()

    if (
        df_train_59["Etiqueta"]
        .nunique()
        < 2
    ):

        st.error(
            "🔴 5.9 ERROR: El aprendizaje acumulado "
            "solo contiene una categoría."
        )

        st.stop()

    # ========================================================
    # 10. ENTRENAR CON TODO LO APRENDIDO
    # ========================================================

    modelo_final_59 = Pipeline(
        [
            (
                "tfidf",
                TfidfVectorizer(
                    lowercase=True,
                    strip_accents="unicode",
                    ngram_range=(1, 2),
                    max_features=20000,
                    sublinear_tf=True
                )
            ),
            (
                "clasificador",
                LogisticRegression(
                    max_iter=5000,
                    class_weight="balanced"
                )
            )
        ]
    )

    modelo_final_59.fit(
        df_train_59["Texto_modelo"],
        df_train_59["Etiqueta"]
    )

    # ========================================================
    # 11. CLASIFICAR TODO
    # ========================================================

    predicciones_finales_59 = (
        modelo_final_59.predict(
            df_modelo_59["Texto_modelo"]
        )
    )

    probabilidades_finales_59 = (
        modelo_final_59.predict_proba(
            df_modelo_59["Texto_modelo"]
        )
    )

    confianza_final_59 = (
        np.max(
            probabilidades_finales_59,
            axis=1
        )
    )

    # ========================================================
    # 12. RESULTADO COMPLETO
    # ========================================================

    df_clasificacion_final_59 = (
        df_modelo_59[
            columnas_base_59
        ].copy()
    )

    df_clasificacion_final_59[
        "Clasificación final"
    ] = predicciones_finales_59

    df_clasificacion_final_59[
        "Confianza"
    ] = confianza_final_59

    # ========================================================
    # 13. GUARDAR RESULTADO COMPLETO
    # ========================================================

    st.session_state[
        "df_clasificacion_final_59"
    ] = df_clasificacion_final_59.copy()

    # ========================================================
    # 14. EXTRAER ACCIONES GENERALES
    # ========================================================

    df_acciones_generales_59 = (
        df_clasificacion_final_59[
            df_clasificacion_final_59[
                "Clasificación final"
            ]
            == "ACCIÓN GENERAL"
        ]
        [
            [
                "Código",
                "Nombre del producto",
                "Acción"
            ]
        ]
        .copy()
    )

    # ========================================================
    # 15. RENOMBRAR COLUMNA
    # ========================================================

    df_acciones_generales_59.rename(
        columns={
            "Acción": "Acción general"
        },
        inplace=True
    )

    # ========================================================
    # 16. GUARDAR PRIMER DATAFRAME
    # ========================================================

    st.session_state[
        "df_acciones_generales_59"
    ] = df_acciones_generales_59.copy()

    # ========================================================
    # 17. RESUMEN
    # ========================================================

    st.success(
        "🟢 5.9 TERMINADO"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Registros analizados",
            len(df_clasificacion_final_59)
        )

    with col2:

        st.metric(
            "Ejemplos de aprendizaje",
            len(df_train_59)
        )

    with col3:

        st.metric(
            "Acciones generales",
            len(df_acciones_generales_59)
        )

    # ========================================================
    # 18. DISTRIBUCIÓN
    # ========================================================

    st.write(
        "### Distribución de la clasificación final"
    )

    distribucion_59 = (
        df_clasificacion_final_59[
            "Clasificación final"
        ]
        .value_counts()
        .rename_axis(
            "Clasificación"
        )
        .reset_index(
            name="Cantidad"
        )
    )

    st.dataframe(
        distribucion_59,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # 19. MOSTRAR PRIMER DATAFRAME
    # ========================================================

    st.write(
        "### 📋 PRIMER DATAFRAME — ACCIONES GENERALES"
    )

    st.dataframe(
        df_acciones_generales_59,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # 20. PREPARAR EXCEL
    # ========================================================

    buffer_excel_59 = io.BytesIO()

    with pd.ExcelWriter(
        buffer_excel_59,
        engine="openpyxl"
    ) as writer:

        # -----------------------------------------------
        # HOJA 1
        # -----------------------------------------------

        df_acciones_generales_59.to_excel(
            writer,
            sheet_name="Acciones_Generales",
            index=False
        )

        # -----------------------------------------------
        # HOJA 2
        # -----------------------------------------------

        df_clasificacion_final_59.to_excel(
            writer,
            sheet_name="Clasificacion_Completa",
            index=False
        )

        # -----------------------------------------------
        # HOJA 3
        # -----------------------------------------------

        df_train_59[
            [
                "Código",
                "Nombre del producto",
                "Acción",
                "Etiqueta"
            ]
        ].to_excel(
            writer,
            sheet_name="Aprendizaje_Acumulado",
            index=False
        )

    buffer_excel_59.seek(0)

    # ========================================================
    # 21. DESCARGA
    # ========================================================

    st.download_button(
        label="⬇️ DESCARGAR PRIMER DATAFRAME",
        data=buffer_excel_59,
        file_name=(
            "PRIMER_DATAFRAME_ACCIONES_GENERALES.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        key="descargar_dataframe_59"
    )

    st.success(
        "✅ El Excel contiene el primer DataFrame "
        "y, adicionalmente, la clasificación completa "
        "y el aprendizaje acumulado utilizado."
    )

except Exception as e:

    st.error(
        f"🔴 5.9 ERROR: "
        f"{type(e).__name__}: {e}"
    )

# ============================================================
# 5.10 PERSISTENCIA Y SINCRONIZACIÓN DEL PRIMER DATAFRAME
#
# Toma:
#   df_acciones_generales_59
#
# Conserva:
#   - registros existentes
#   - códigos AG existentes
#   - aprendizaje realizado
#
# Agrega:
#   - únicamente productos + acciones nuevas
#
# Sincroniza:
#   GitHub -> franquiciasauces / Asesores
#
# NO modifica 5.9
# ============================================================

st.markdown("### 5.10 Persistencia y sincronización")

try:

    import io
    import base64
    import requests
    import pandas as pd

    # ========================================================
    # 1. CONFIGURACIÓN REAL DEL REPOSITORIO
    # ========================================================

    GITHUB_USUARIO_510 = "franquiciasauces"
    GITHUB_REPOSITORIO_510 = "Asesores"
    GITHUB_RAMA_510 = "main"

    GITHUB_ARCHIVO_510 = (
        "ACCIONES_GENERALES.xlsx"
    )

    # ========================================================
    # 2. RUTA LOCAL
    # ========================================================

    if "BASE_DIR" in globals():

        ARCHIVO_ACCIONES_510 = (
            BASE_DIR /
            GITHUB_ARCHIVO_510
        )

    else:

        from pathlib import Path

        ARCHIVO_ACCIONES_510 = (
            Path(__file__).resolve().parent.parent /
            GITHUB_ARCHIVO_510
        )

    # ========================================================
    # 3. RECUPERAR RESULTADO DE 5.9
    # ========================================================

    df_nuevo_510 = st.session_state.get(
        "df_acciones_generales_59",
        None
    )

    if (
        df_nuevo_510 is None
        or not isinstance(
            df_nuevo_510,
            pd.DataFrame
        )
        or df_nuevo_510.empty
    ):

        st.error(
            "🔴 5.10 ERROR: No existe "
            "df_acciones_generales_59 generado por 5.9."
        )

        st.stop()

    # ========================================================
    # 4. COLUMNAS OBLIGATORIAS
    # ========================================================

    columnas_510 = [
        "Código",
        "Nombre del producto",
        "Acción general"
    ]

    faltantes_510 = [
        columna
        for columna in columnas_510
        if columna not in df_nuevo_510.columns
    ]

    if faltantes_510:

        st.error(
            "🔴 5.10 ERROR: Faltan columnas: "
            + ", ".join(faltantes_510)
        )

        st.stop()

    df_nuevo_510 = df_nuevo_510[
        columnas_510
    ].copy()

    # ========================================================
    # 5. LIMPIAR
    # ========================================================

    for columna in columnas_510:

        df_nuevo_510[columna] = (
            df_nuevo_510[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df_nuevo_510 = df_nuevo_510[
        (df_nuevo_510["Nombre del producto"] != "")
        &
        (df_nuevo_510["Acción general"] != "")
    ].copy()

    # ========================================================
    # 6. CARGAR ARCHIVO LOCAL EXISTENTE
    # ========================================================

    if ARCHIVO_ACCIONES_510.exists():

        try:

            df_existente_510 = pd.read_excel(
                ARCHIVO_ACCIONES_510
            )

        except Exception:

            df_existente_510 = pd.DataFrame(
                columns=columnas_510
            )

    else:

        df_existente_510 = pd.DataFrame(
            columns=columnas_510
        )

    # ========================================================
    # 7. NORMALIZAR ARCHIVO EXISTENTE
    # ========================================================

    for columna in columnas_510:

        if columna not in df_existente_510.columns:

            df_existente_510[columna] = ""

        df_existente_510[columna] = (
            df_existente_510[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df_existente_510 = df_existente_510[
        columnas_510
    ].copy()

    # ========================================================
    # 8. CLAVE ÚNICA
    #
    # Producto + Acción general
    #
    # El código NO determina si es duplicado.
    # ========================================================

    def clave_510(fila):

        producto = (
            str(
                fila["Nombre del producto"]
            )
            .strip()
            .upper()
        )

        accion = (
            str(
                fila["Acción general"]
            )
            .strip()
            .upper()
        )

        return (
            producto
            + "|||"
            + accion
        )

    if not df_existente_510.empty:

        df_existente_510[
            "_CLAVE_510"
        ] = df_existente_510.apply(
            clave_510,
            axis=1
        )

    else:

        df_existente_510[
            "_CLAVE_510"
        ] = ""

    df_nuevo_510[
        "_CLAVE_510"
    ] = df_nuevo_510.apply(
        clave_510,
        axis=1
    )

    # ========================================================
    # 9. DETECTAR NUEVOS
    # ========================================================

    claves_existentes_510 = set(
        df_existente_510[
            "_CLAVE_510"
        ]
        .tolist()
    )

    df_nuevos_510 = (
        df_nuevo_510[
            ~df_nuevo_510[
                "_CLAVE_510"
            ]
            .isin(
                claves_existentes_510
            )
        ]
        .copy()
    )

    # ========================================================
    # 10. OBTENER ÚLTIMO CÓDIGO EXISTENTE
    # ========================================================

    numeros_510 = []

    for codigo in (
        df_existente_510["Código"]
        .astype(str)
        .tolist()
    ):

        codigo = codigo.strip().upper()

        if codigo.startswith("AG"):

            try:

                numeros_510.append(
                    int(
                        codigo[2:]
                    )
                )

            except Exception:

                pass

    siguiente_510 = (
        max(
            numeros_510,
            default=0
        )
        + 1
    )

    # ========================================================
    # 11. ASIGNAR CÓDIGOS SOLO A LOS NUEVOS
    # ========================================================

    nuevos_codigos_510 = []

    for _ in range(
        len(df_nuevos_510)
    ):

        nuevos_codigos_510.append(
            f"AG{siguiente_510:06d}"
        )

        siguiente_510 += 1

    if not df_nuevos_510.empty:

        df_nuevos_510[
            "Código"
        ] = nuevos_codigos_510

    # ========================================================
    # 12. UNIR SIN BORRAR
    # ========================================================

    df_final_510 = pd.concat(
        [
            df_existente_510[
                columnas_510
            ],
            df_nuevos_510[
                columnas_510
            ]
        ],
        ignore_index=True
    )

    # ========================================================
    # 13. ELIMINAR DUPLICADOS
    # ========================================================

    df_final_510[
        "_CLAVE_510"
    ] = df_final_510.apply(
        clave_510,
        axis=1
    )

    df_final_510 = (
        df_final_510
        .drop_duplicates(
            subset=[
                "_CLAVE_510"
            ],
            keep="first"
        )
        .drop(
            columns=[
                "_CLAVE_510"
            ]
        )
        .reset_index(
            drop=True
        )
    )

    # ========================================================
    # 14. GUARDAR LOCALMENTE
    # ========================================================

    with pd.ExcelWriter(
        ARCHIVO_ACCIONES_510,
        engine="openpyxl"
    ) as writer:

        df_final_510.to_excel(
            writer,
            sheet_name="Acciones_Generales",
            index=False
        )

    # ========================================================
    # 15. GUARDAR EN SESSION STATE
    # ========================================================

    st.session_state[
        "df_acciones_generales_persistente"
    ] = df_final_510.copy()

    # ========================================================
    # 16. RESUMEN
    # ========================================================

    cantidad_existente_510 = (
        len(df_existente_510)
    )

    cantidad_nueva_510 = (
        len(df_nuevos_510)
    )

    cantidad_final_510 = (
        len(df_final_510)
    )

    st.success(
        "🟢 5.10 PROCESO DE PERSISTENCIA TERMINADO"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Existentes conservados",
            cantidad_existente_510
        )

    with col2:

        st.metric(
            "Nuevos agregados",
            cantidad_nueva_510
        )

    with col3:

        st.metric(
            "Total acumulado",
            cantidad_final_510
        )

    # ========================================================
    # 17. MOSTRAR DATAFRAME
    # ========================================================

    st.write(
        "### DataFrame persistente"
    )

    st.dataframe(
        df_final_510,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # 18. PREPARAR EXCEL
    # ========================================================

    buffer_510 = io.BytesIO()

    with pd.ExcelWriter(
        buffer_510,
        engine="openpyxl"
    ) as writer:

        df_final_510.to_excel(
            writer,
            sheet_name="Acciones_Generales",
            index=False
        )

    buffer_510.seek(0)

    # ========================================================
    # 19. DESCARGA
    # ========================================================

    st.download_button(
        label="⬇️ Descargar ACCIONES_GENERALES.xlsx",
        data=buffer_510,
        file_name="ACCIONES_GENERALES.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        key="descargar_acciones_510"
    )

    # ========================================================
    # 20. OBTENER TOKEN DE GITHUB
    #
    # NO se escribe el token en el código.
    #
    # Se intenta encontrar en:
    #
    # st.secrets["GITHUB_TOKEN"]
    # st.secrets["github"]["token"]
    #
    # o variable de entorno.
    # ========================================================

    import os

    github_token_510 = None

    # --------------------------------------------------------
    # Opción 1
    # --------------------------------------------------------

    try:

        github_token_510 = st.secrets.get(
            "GITHUB_TOKEN"
        )

    except Exception:

        pass

    # --------------------------------------------------------
    # Opción 2
    # --------------------------------------------------------

    if not github_token_510:

        try:

            github_token_510 = (
                st.secrets
                .get(
                    "github",
                    {}
                )
                .get(
                    "token"
                )
            )

        except Exception:

            pass

    # --------------------------------------------------------
    # Opción 3
    # --------------------------------------------------------

    if not github_token_510:

        github_token_510 = os.getenv(
            "GITHUB_TOKEN"
        )

    # ========================================================
    # 21. BOTÓN DE SINCRONIZACIÓN
    # ========================================================

    st.markdown(
        "### ☁️ Guardar en GitHub"
    )

    if not github_token_510:

        st.warning(
            "⚠️ No se encontró el token de GitHub "
            "en los secretos de la aplicación."
        )

        st.info(
            "El DataFrame ya quedó guardado localmente "
            "y puede descargarse. La sincronización "
            "con GitHub queda pendiente de disponer "
            "del token configurado."
        )

    else:

        if st.button(
            "☁️ GUARDAR Y SINCRONIZAR CON GITHUB",
            key="guardar_github_510"
        ):

            try:

                # --------------------------------------------
                # CONVERTIR EXCEL A BASE64
                # --------------------------------------------

                contenido_510 = (
                    buffer_510.getvalue()
                )

                contenido_base64_510 = (
                    base64.b64encode(
                        contenido_510
                    )
                    .decode(
                        "utf-8"
                    )
                )

                # --------------------------------------------
                # URL REAL DEL REPOSITORIO
                # --------------------------------------------

                url_github_510 = (
                    "https://api.github.com/repos/"
                    f"{GITHUB_USUARIO_510}/"
                    f"{GITHUB_REPOSITORIO_510}/contents/"
                    f"{GITHUB_ARCHIVO_510}"
                )

                headers_510 = {
                    "Authorization":
                        f"Bearer {github_token_510}",
                    "Accept":
                        "application/vnd.github+json",
                    "X-GitHub-Api-Version":
                        "2022-11-28"
                }

                # --------------------------------------------
                # CONSULTAR ARCHIVO
                # --------------------------------------------

                respuesta_get_510 = (
                    requests.get(
                        url_github_510,
                        headers=headers_510,
                        params={
                            "ref":
                                GITHUB_RAMA_510
                        },
                        timeout=30
                    )
                )

                sha_510 = None

                if (
                    respuesta_get_510.status_code
                    == 200
                ):

                    datos_github_510 = (
                        respuesta_get_510.json()
                    )

                    sha_510 = (
                        datos_github_510
                        .get("sha")
                    )

                elif (
                    respuesta_get_510.status_code
                    != 404
                ):

                    st.error(
                        "🔴 GitHub devolvió "
                        f"{respuesta_get_510.status_code}"
                    )

                    st.code(
                        respuesta_get_510.text
                    )

                    st.stop()

                # --------------------------------------------
                # PAYLOAD
                # --------------------------------------------

                payload_510 = {
                    "message":
                        "Actualizar "
                        "ACCIONES_GENERALES.xlsx",
                    "content":
                        contenido_base64_510,
                    "branch":
                        GITHUB_RAMA_510
                }

                if sha_510:

                    payload_510[
                        "sha"
                    ] = sha_510

                # --------------------------------------------
                # SUBIR A GITHUB
                # --------------------------------------------

                respuesta_put_510 = (
                    requests.put(
                        url_github_510,
                        headers=headers_510,
                        json=payload_510,
                        timeout=30
                    )
                )

                if (
                    respuesta_put_510.status_code
                    in [200, 201]
                ):

                    st.success(
                        "🟢 SINCRONIZACIÓN EXITOSA"
                    )

                    st.info(
                        "ACCIONES_GENERALES.xlsx "
                        "quedó actualizado en GitHub."
                    )

                else:

                    st.error(
                        "🔴 ERROR AL GUARDAR EN GITHUB: "
                        f"{respuesta_put_510.status_code}"
                    )

                    st.code(
                        respuesta_put_510.text
                    )

            except Exception as error_510:

                st.error(
                    "🔴 ERROR DE SINCRONIZACIÓN: "
                    f"{type(error_510).__name__}: "
                    f"{error_510}"
                )

except Exception as e:

    st.error(
        f"🔴 5.10 ERROR: "
        f"{type(e).__name__}: {e}"
    )
# ============================================================

# ============================================================
# 5.11 NORMALIZACIÓN DE COMPONENTES
#
# FUENTE:
#   df_fuente generado en 5.1
#
# OBJETIVO:
#   Construir la base que utilizará 5.12 para interpretar
#   las relaciones:
#
#   PRODUCTO → COMPONENTE → ACCIÓN CANDIDATA
#
# IMPORTANTE:
#   5.11 NO decide todavía si una acción pertenece
#   a un componente.
#
#   5.11 solamente:
#   - identifica productos
#   - identifica componentes
#   - separa componentes cuando existen separadores explícitos
#   - conserva el texto funcional asociado al componente
#   - conserva la acción general disponible para 5.12
#
# NO UTILIZA:
#   - Categoría principal
#   - Categorías complementarias
#   - Precio
#   - Foto
#   - Posología
#   - Frases comerciales
# ============================================================

st.markdown(
    "### 5.11 Normalización de componentes"
)

try:

    import re
    import pandas as pd

    # ========================================================
    # 1. VALIDAR MATRIZ DE 5.1
    # ========================================================

    if (
        "df_fuente" not in locals()
        or df_fuente is None
        or not isinstance(df_fuente, pd.DataFrame)
        or df_fuente.empty
    ):

        st.error(
            "🔴 5.11 ERROR: No existe una matriz válida "
            "proveniente de 5.1."
        )

    else:

        # ====================================================
        # 2. NORMALIZAR NOMBRES DE COLUMNAS PARA BUSCARLAS
        # ====================================================

        def normalizar_nombre_columna_511(texto):

            texto = str(texto).strip().lower()

            reemplazos = {
                "á": "a",
                "é": "e",
                "í": "i",
                "ó": "o",
                "ú": "u",
                "ü": "u",
                "ñ": "n"
            }

            for origen, destino in reemplazos.items():
                texto = texto.replace(
                    origen,
                    destino
                )

            texto = re.sub(
                r"[^a-z0-9]+",
                " ",
                texto
            )

            return " ".join(
                texto.split()
            )

        columnas_normalizadas_511 = {
            columna:
                normalizar_nombre_columna_511(
                    columna
                )
            for columna in df_fuente.columns
        }

        # ====================================================
        # 3. IDENTIFICAR PRODUCTO
        # ====================================================

        columna_producto_511 = None

        candidatos_producto_511 = {
            "producto",
            "productos",
            "nombre producto",
            "nombre del producto"
        }

        for columna, nombre_normalizado in (
            columnas_normalizadas_511.items()
        ):

            if nombre_normalizado in candidatos_producto_511:

                columna_producto_511 = columna
                break

        # Si no encuentra el nombre, utiliza la columna A,
        # que ya fue definida como producto en 5.2.

        if columna_producto_511 is None:

            if len(df_fuente.columns) >= 1:

                columna_producto_511 = (
                    df_fuente.columns[0]
                )

            else:

                st.error(
                    "🔴 5.11 ERROR: No existe columna "
                    "de producto."
                )

        # ====================================================
        # 4. IDENTIFICAR COLUMNAS DE COMPONENTES
        # ====================================================

        palabras_componente_511 = [
            "componente",
            "componentes",
            "ingrediente",
            "ingredientes",
            "principio activo",
            "principios activos"
        ]

        columnas_componente_511 = []

        for columna, nombre_normalizado in (
            columnas_normalizadas_511.items()
        ):

            if any(
                palabra in nombre_normalizado
                for palabra in palabras_componente_511
            ):

                columnas_componente_511.append(
                    columna
                )

        # ====================================================
        # 5. IDENTIFICAR ACCIONES GENERALES
        #
        # Esta columna NO se utiliza para decidir todavía.
        # Se conserva para que 5.12 pueda interpretar.
        # ====================================================

        columna_acciones_511 = None

        candidatos_acciones_511 = {
            "accion general",
            "acciones generales",
            "accion",
            "acciones"
        }

        for columna, nombre_normalizado in (
            columnas_normalizadas_511.items()
        ):

            if nombre_normalizado in candidatos_acciones_511:

                columna_acciones_511 = columna
                break

        # ====================================================
        # 6. VALIDAR COMPONENTES
        # ====================================================

        if not columnas_componente_511:

            st.error(
                "🔴 5.11 ERROR: No se encontró ninguna "
                "columna relacionada con componentes."
            )

            st.write(
                "Columnas disponibles:"
            )

            st.write(
                list(df_fuente.columns)
            )

        else:

            st.success(
                "🟢 5.11: Estructura de componentes "
                "identificada correctamente."
            )

            st.info(
                f"Producto: **{columna_producto_511}**"
            )

            st.info(
                "Componentes: **"
                +
                " | ".join(
                    str(columna)
                    for columna
                    in columnas_componente_511
                )
                +
                "**"
            )

            if columna_acciones_511 is not None:

                st.info(
                    "Acciones generales disponibles para "
                    f"5.12: **{columna_acciones_511}**"
                )

            else:

                st.warning(
                    "⚠️ 5.11: No se identificó una columna "
                    "de acciones generales. "
                    "5.12 trabajará únicamente con la "
                    "información funcional disponible."
                )

            # =================================================
            # 7. FUNCIÓN PARA LIMPIAR TEXTO
            # =================================================

            def limpiar_texto_511(texto):

                if texto is None:
                    return ""

                if pd.isna(texto):
                    return ""

                texto = str(texto).strip()

                if not texto:
                    return ""

                texto = re.sub(
                    r"\s+",
                    " ",
                    texto
                )

                return texto.strip()

            # =================================================
            # 8. FUNCIÓN PARA SEPARAR COMPONENTES
            #
            # SOLO separadores explícitos.
            # NO se divide por comas porque una coma puede
            # formar parte del nombre o descripción.
            # =================================================

            def separar_componentes_511(texto):

                texto = limpiar_texto_511(texto)

                if not texto:
                    return []

                partes = re.split(
                    r";|\n|\||•",
                    texto
                )

                resultado = []

                for parte in partes:

                    parte = limpiar_texto_511(
                        parte
                    )

                    if parte:
                        resultado.append(
                            parte
                        )

                return resultado

            # =================================================
            # 9. EXTRAER COMPONENTE Y TEXTO FUNCIONAL
            #
            # Ejemplos:
            #
            # Zinc (refuerzo inmunológico)
            #
            # produce:
            #   Componente = Zinc
            #   Texto funcional = refuerzo inmunológico
            #
            # Biotina - cabello y uñas
            #
            # produce:
            #   Componente = Biotina
            #   Texto funcional = cabello y uñas
            #
            # NO se inventa información.
            # =================================================

            def interpretar_componente_511(
                texto
            ):

                texto = limpiar_texto_511(
                    texto
                )

                if not texto:
                    return "", ""

                componente = texto
                texto_funcional = ""

                # ---------------------------------------------
                # Componente (función)
                # ---------------------------------------------

                coincidencia_parentesis = re.match(
                    r"^\s*(.+?)\s*"
                    r"\(([^()]+)\)\s*$",
                    texto
                )

                if coincidencia_parentesis:

                    componente = (
                        coincidencia_parentesis
                        .group(1)
                        .strip()
                    )

                    texto_funcional = (
                        coincidencia_parentesis
                        .group(2)
                        .strip()
                    )

                else:

                    # -----------------------------------------
                    # Componente - función
                    # Componente – función
                    # Componente — función
                    # Componente : función
                    # -----------------------------------------

                    coincidencia_separador = re.match(
                        r"^\s*(.+?)\s+"
                        r"[-–—:]\s+"
                        r"(.+?)\s*$",
                        texto
                    )

                    if coincidencia_separador:

                        componente = (
                            coincidencia_separador
                            .group(1)
                            .strip()
                        )

                        texto_funcional = (
                            coincidencia_separador
                            .group(2)
                            .strip()
                        )

                componente = limpiar_texto_511(
                    componente
                )

                texto_funcional = limpiar_texto_511(
                    texto_funcional
                )

                return (
                    componente,
                    texto_funcional
                )

            # =================================================
            # 10. CREAR REGISTROS
            # =================================================

            registros_511 = []

            for indice_511, fila_511 in (
                df_fuente.iterrows()
            ):

                producto_511 = limpiar_texto_511(
                    fila_511[
                        columna_producto_511
                    ]
                )

                if not producto_511:
                    continue

                # ---------------------------------------------
                # RECORRER TODAS LAS COLUMNAS DE COMPONENTES
                # ---------------------------------------------

                for columna_componente_511 in (
                    columnas_componente_511
                ):

                    texto_componentes_511 = (
                        limpiar_texto_511(
                            fila_511[
                                columna_componente_511
                            ]
                        )
                    )

                    if not texto_componentes_511:
                        continue

                    componentes_511 = (
                        separar_componentes_511(
                            texto_componentes_511
                        )
                    )

                    # -----------------------------------------
                    # CADA COMPONENTE = UNA FILA
                    # -----------------------------------------

                    for componente_original_511 in (
                        componentes_511
                    ):

                        (
                            componente_511,
                            texto_funcional_511
                        ) = interpretar_componente_511(
                            componente_original_511
                        )

                        if not componente_511:
                            continue

                        # -------------------------------------
                        # ACCIONES GENERALES
                        #
                        # Se conservan como texto de contexto.
                        # NO se asignan todavía al componente.
                        # -------------------------------------

                        acciones_generales_511 = ""

                        if (
                            columna_acciones_511
                            is not None
                        ):

                            acciones_generales_511 = (
                                limpiar_texto_511(
                                    fila_511[
                                        columna_acciones_511
                                    ]
                                )
                            )

                        registros_511.append(
                            {
                                "Producto":
                                    producto_511,

                                "Componente":
                                    componente_511,

                                "Texto funcional":
                                    texto_funcional_511,

                                "Acciones generales":
                                    acciones_generales_511,

                                "Fuente":
                                    str(
                                        columna_componente_511
                                    ),

                                "Fila origen":
                                    int(
                                        indice_511 + 2
                                    )
                            }
                        )

            # =================================================
            # 11. CREAR DATAFRAME
            # =================================================

            df_normalizado_componentes_511 = (
                pd.DataFrame(
                    registros_511,
                    columns=[
                        "Producto",
                        "Componente",
                        "Texto funcional",
                        "Acciones generales",
                        "Fuente",
                        "Fila origen"
                    ]
                )
            )

            # =================================================
            # 12. LIMPIEZA FINAL
            # =================================================

            if not (
                df_normalizado_componentes_511
                .empty
            ):

                for columna in [
                    "Producto",
                    "Componente",
                    "Texto funcional",
                    "Acciones generales",
                    "Fuente"
                ]:

                    df_normalizado_componentes_511[
                        columna
                    ] = (
                        df_normalizado_componentes_511[
                            columna
                        ]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                    )

                # ---------------------------------------------
                # ELIMINAR DUPLICADOS REALES
                # ---------------------------------------------

                df_normalizado_componentes_511 = (
                    df_normalizado_componentes_511
                    .drop_duplicates(
                        subset=[
                            "Producto",
                            "Componente",
                            "Texto funcional",
                            "Acciones generales"
                        ]
                    )
                    .reset_index(
                        drop=True
                    )
                )

            # =================================================
            # 13. VALIDAR RESULTADO
            # =================================================

            if (
                df_normalizado_componentes_511
                .empty
            ):

                st.error(
                    "🔴 5.11 ERROR: No se encontraron "
                    "registros de Producto + Componente."
                )

            else:

                # ---------------------------------------------
                # CÓDIGO PROVISIONAL DEL REGISTRO
                # ---------------------------------------------

                df_normalizado_componentes_511.insert(
                    0,
                    "Código",
                    [
                        f"CF{numero:06d}"
                        for numero in range(
                            1,
                            len(
                                df_normalizado_componentes_511
                            ) + 1
                        )
                    ]
                )

                # =================================================
                # 14. GUARDAR PARA 5.12
                # =================================================

                st.session_state[
                    "df_normalizado_componentes_511"
                ] = (
                    df_normalizado_componentes_511.copy()
                )

                # =================================================
                # 15. CONTADORES
                # =================================================

                cantidad_registros_511 = len(
                    df_normalizado_componentes_511
                )

                cantidad_productos_511 = (
                    df_normalizado_componentes_511[
                        "Producto"
                    ]
                    .nunique()
                )

                cantidad_componentes_511 = (
                    df_normalizado_componentes_511[
                        "Componente"
                    ]
                    .nunique()
                )

                cantidad_funciones_511 = int(
                    (
                        df_normalizado_componentes_511[
                            "Texto funcional"
                        ]
                        .astype(str)
                        .str.strip()
                        != ""
                    ).sum()
                )

                cantidad_sin_funcion_511 = (
                    cantidad_registros_511
                    - cantidad_funciones_511
                )

                # =================================================
                # 16. RESULTADO
                # =================================================

                st.success(
                    f"🟢 5.11 TERMINADO: "
                    f"**{cantidad_registros_511} registros** | "
                    f"**{cantidad_productos_511} productos** | "
                    f"**{cantidad_componentes_511} componentes**."
                )

                st.info(
                    f"Con texto funcional explícito: "
                    f"**{cantidad_funciones_511}** | "
                    f"Sin texto funcional explícito: "
                    f"**{cantidad_sin_funcion_511}**"
                )

                # =================================================
                # 17. MOSTRAR RESULTADO
                # =================================================

                st.write(
                    "### 5.11 — Base normalizada de componentes"
                )

                st.dataframe(
                    df_normalizado_componentes_511[
                        [
                            "Código",
                            "Producto",
                            "Componente",
                            "Texto funcional",
                            "Acciones generales"
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True
                )

                # =================================================
                # 18. VALIDACIÓN PROSTENFIT
                # =================================================

                ejemplo_prostenfit_511 = (
                    df_normalizado_componentes_511[
                        df_normalizado_componentes_511[
                            "Producto"
                        ]
                        .str.contains(
                            "PROSTENFIT",
                            case=False,
                            na=False
                        )
                    ]
                )

                if not ejemplo_prostenfit_511.empty:

                    st.write(
                        "### Validación FITO PROSTENFIT"
                    )

                    st.dataframe(
                        ejemplo_prostenfit_511[
                            [
                                "Código",
                                "Producto",
                                "Componente",
                                "Texto funcional",
                                "Acciones generales"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                # =================================================
                # 19. MENSAJE DE CONTROL
                # =================================================

                st.success(
                    "✅ 5.11 listo. "
                    "La interpretación de las acciones "
                    "componente se realizará en 5.12."
                )

                st.info(
                    "ℹ️ 5.11 NO decide todavía qué acción "
                    "pertenece a cada componente. "
                    "Entrega a 5.12 el componente, su texto "
                    "funcional y las acciones generales "
                    "disponibles como contexto."
                )

except Exception as e:

    st.error(
        f"🔴 5.11 ERROR: "
        f"{type(e).__name__}: {e}"
    )

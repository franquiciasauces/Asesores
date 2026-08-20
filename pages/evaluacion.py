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
# 5.4 ENTRENAMIENTO SEMÁNTICO INICIAL
# Clasificador específico para ACCIONES GENERALES
# ============================================================

st.markdown("### 5.4 Entrenamiento del clasificador de acciones")

try:

    if "df_depurado" not in locals() or df_depurado.empty:

        st.error(
            "🔴 5.4 ERROR: No existe df_depurado proveniente del 5.3."
        )

    else:

        # ----------------------------------------------------
        # IMPORTACIONES LOCALES
        # ----------------------------------------------------

        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.linear_model import LogisticRegression
        from sklearn.pipeline import Pipeline

        # ----------------------------------------------------
        # PREPARAR DATOS
        # ----------------------------------------------------

        df_ml = df_depurado[
            [
                "Código",
                "Nombre del producto",
                "Acción"
            ]
        ].copy()

        df_ml["Acción"] = (
            df_ml["Acción"]
            .astype(str)
            .str.strip()
        )

        df_ml = df_ml[
            df_ml["Acción"] != ""
        ].copy()

        # ----------------------------------------------------
        # ESTADO DEL ENTRENAMIENTO
        # Se conserva durante la sesión de Streamlit
        # ----------------------------------------------------

        if "entrenamiento_acciones" not in st.session_state:

            st.session_state.entrenamiento_acciones = {}

        # ----------------------------------------------------
        # SELECCIONAR MUESTRA INICIAL
        # MÁXIMO 50
        # ----------------------------------------------------

        etiquetados = (
            st.session_state
            .entrenamiento_acciones
        )

        pendientes = df_ml[
            ~df_ml["Código"].astype(str).isin(
                etiquetados.keys()
            )
        ].copy()

        # ----------------------------------------------------
        # SELECCIÓN REPRESENTATIVA
        #
        # Se distribuyen los ejemplos a lo largo del
        # conjunto, en lugar de tomar simplemente los
        # primeros 50.
        # ----------------------------------------------------

        limite_muestra = min(
            50,
            len(pendientes)
        )

        if limite_muestra > 0:

            posiciones = np.linspace(
                0,
                len(pendientes) - 1,
                limite_muestra,
                dtype=int
            )

            muestra = (
                pendientes
                .iloc[posiciones]
                .drop_duplicates(
                    subset=["Código"]
                )
                .copy()
            )

        else:

            muestra = pd.DataFrame()

        # ----------------------------------------------------
        # INTERFAZ DE ETIQUETADO
        # ----------------------------------------------------

        st.write(
            "El sistema selecciona hasta **50 ejemplos** "
            "para construir el entrenamiento inicial."
        )

        st.info(
            "Clasifique solamente los ejemplos mostrados. "
            "No necesita revisar las ~500 acciones."
        )

        if not muestra.empty:

            for _, fila in muestra.iterrows():

                codigo = str(
                    fila["Código"]
                )

                accion = str(
                    fila["Acción"]
                )

                producto = str(
                    fila["Nombre del producto"]
                )

                st.markdown(
                    f"**{codigo} — {producto}**"
                )

                st.write(
                    f"Acción: **{accion}**"
                )

                opcion = st.radio(
                    "Clasificación",
                    [
                        "Acción general",
                        "No es acción"
                    ],
                    index=None,
                    key=f"clasificacion_54_{codigo}"
                )

                if opcion is not None:

                    st.session_state.entrenamiento_acciones[
                        codigo
                    ] = (
                        1
                        if opcion == "Acción general"
                        else 0
                    )

                st.divider()

        # ----------------------------------------------------
        # ESTADO ACTUAL
        # ----------------------------------------------------

        total_etiquetados = len(
            st.session_state.entrenamiento_acciones
        )

        positivos = sum(
            1
            for valor
            in st.session_state.entrenamiento_acciones.values()
            if valor == 1
        )

        negativos = (
            total_etiquetados
            - positivos
        )

        st.write(
            f"**Ejemplos clasificados:** "
            f"{total_etiquetados}/50"
        )

        st.write(
            f"Acción general: **{positivos}** | "
            f"No es acción: **{negativos}**"
        )

        # ----------------------------------------------------
        # ENTRENAR CUANDO HAYA SUFICIENTES EJEMPLOS
        # ----------------------------------------------------

        if total_etiquetados >= 20:

            codigos_entrenamiento = list(
                st.session_state
                .entrenamiento_acciones
                .keys()
            )

            df_entrenamiento = df_ml[
                df_ml["Código"].astype(str).isin(
                    codigos_entrenamiento
                )
            ].copy()

            df_entrenamiento["Etiqueta"] = (
                df_entrenamiento["Código"]
                .astype(str)
                .map(
                    st.session_state
                    .entrenamiento_acciones
                )
            )

            clases = (
                df_entrenamiento["Etiqueta"]
                .nunique()
            )

            if clases < 2:

                st.warning(
                    "⚠️ El entrenamiento necesita ejemplos "
                    "de las dos clases: Acción general y "
                    "No es acción."
                )

            else:

                # --------------------------------------------
                # MODELO
                # --------------------------------------------

                modelo_acciones = Pipeline(
                    [
                        (
                            "vectorizador",
                            TfidfVectorizer(
                                lowercase=True,
                                strip_accents="unicode",
                                ngram_range=(1, 2),
                                min_df=1,
                                max_features=5000
                            )
                        ),
                        (
                            "clasificador",
                            LogisticRegression(
                                max_iter=2000,
                                class_weight="balanced"
                            )
                        )
                    ]
                )

                modelo_acciones.fit(
                    df_entrenamiento["Acción"],
                    df_entrenamiento["Etiqueta"]
                )

                # --------------------------------------------
                # PREDICCIÓN SOBRE TODAS LAS ACCIONES
                # --------------------------------------------

                df_resultado_ml = df_ml.copy()

                probabilidades = (
                    modelo_acciones.predict_proba(
                        df_resultado_ml["Acción"]
                    )
                )

                clases_modelo = (
                    modelo_acciones
                    .named_steps[
                        "clasificador"
                    ]
                    .classes_
                )

                indice_positivo = list(
                    clases_modelo
                ).index(1)

                df_resultado_ml[
                    "Confianza acción general"
                ] = (
                    probabilidades[
                        :,
                        indice_positivo
                    ]
                )

                df_resultado_ml[
                    "Clasificación IA"
                ] = np.where(
                    df_resultado_ml[
                        "Confianza acción general"
                    ] >= 0.70,
                    "Acción general",
                    "Revisar"
                )

                # --------------------------------------------
                # GUARDAR RESULTADO EN SESIÓN
                # --------------------------------------------

                st.session_state[
                    "df_resultado_ml"
                ] = df_resultado_ml

                # --------------------------------------------
                # ESTADÍSTICAS
                # --------------------------------------------

                alta_confianza = int(
                    (
                        df_resultado_ml[
                            "Clasificación IA"
                        ]
                        == "Acción general"
                    ).sum()
                )

                revisar = int(
                    (
                        df_resultado_ml[
                            "Clasificación IA"
                        ]
                        == "Revisar"
                    ).sum()
                )

                st.success(
                    "🟢 Modelo entrenado y aplicado."
                )

                st.info(
                    f"Acciones procesadas: **{len(df_resultado_ml)}** | "
                    f"Alta confianza: **{alta_confianza}** | "
                    f"Para revisar: **{revisar}**"
                )

                # --------------------------------------------
                # MOSTRAR SOLO CASOS DE REVISIÓN
                # --------------------------------------------

                df_revision = (
                    df_resultado_ml[
                        df_resultado_ml[
                            "Clasificación IA"
                        ]
                        == "Revisar"
                    ]
                    .sort_values(
                        "Confianza acción general"
                    )
                    .head(30)
                    .copy()
                )

                if not df_revision.empty:

                    st.warning(
                        f"⚠️ Se muestran los "
                        f"{len(df_revision)} casos "
                        "de menor confianza."
                    )

                    st.dataframe(
                        df_revision[
                            [
                                "Código",
                                "Nombre del producto",
                                "Acción",
                                "Confianza acción general"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                # --------------------------------------------
                # MATRIZ FINAL PROVISIONAL
                # SOLO LAS 3 COLUMNAS REQUERIDAS
                # --------------------------------------------

                st.write(
                    "### Resultado del clasificador"
                )

                st.dataframe(
                    df_resultado_ml[
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

            faltan = (
                20 - total_etiquetados
            )

            st.warning(
                f"⚠️ Faltan **{faltan}** ejemplos "
                "clasificados para entrenar el primer modelo."
            )

except Exception as e:

    st.error(
        f"🔴 5.4 ERROR: {type(e).__name__}: {e}"
    )

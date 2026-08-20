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
# 5.4 ENTRENAMIENTO SEMÁNTICO DE ACCIONES
# ============================================================

st.markdown("### 5.4 Clasificación semántica de acciones")

try:

    import numpy as np
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline

    if "df_depurado" not in locals() or df_depurado.empty:

        st.error(
            "🔴 5.4 ERROR: No existe df_depurado proveniente del 5.3."
        )

    else:

        # ----------------------------------------------------
        # PREPARAR DATAFRAME
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
        # CATEGORÍAS DE ENTRENAMIENTO
        # ----------------------------------------------------

        categorias_54 = {
            "ACCIÓN GENERAL": 0,
            "COMPONENTE + FUNCIÓN": 1,
            "RECOMENDACIÓN / COMPLEMENTO": 2,
            "USO / POSOLOGÍA / PRECAUCIÓN": 3,
            "COMERCIAL": 4
        }

        # ----------------------------------------------------
        # MEMORIA DE CLASIFICACIONES
        # ----------------------------------------------------

        if "entrenamiento_acciones_54" not in st.session_state:

            st.session_state.entrenamiento_acciones_54 = {}

        etiquetas = (
            st.session_state.entrenamiento_acciones_54
        )

        # ----------------------------------------------------
        # MÁXIMO ABSOLUTO DE 50 EJEMPLOS
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
        # INTERFAZ DE ENTRENAMIENTO
        # ----------------------------------------------------

        st.write(
            "### Entrenamiento inicial"
        )

        st.info(
            "El entrenamiento utiliza como máximo "
            "**50 ejemplos**. Cuando llegue a 50, "
            "no se mostrarán más ejemplos."
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
                    f"Texto: **{accion}**"
                )

                opcion = st.radio(
                    "Clasificación",
                    list(
                        categorias_54.keys()
                    ),
                    index=None,
                    key=f"clasificacion_54_{codigo}"
                )

                if opcion is not None:

                    st.session_state.entrenamiento_acciones_54[
                        codigo
                    ] = categorias_54[
                        opcion
                    ]

                st.divider()

        # ----------------------------------------------------
        # ACTUALIZAR CONTADOR
        # ----------------------------------------------------

        total_etiquetados = len(
            st.session_state.entrenamiento_acciones_54
        )

        conteo = {
            nombre: 0
            for nombre in categorias_54
        }

        for valor in (
            st.session_state
            .entrenamiento_acciones_54
            .values()
        ):

            for nombre, numero in categorias_54.items():

                if valor == numero:
                    conteo[nombre] += 1

        st.write(
            f"**Ejemplos clasificados: "
            f"{total_etiquetados}/50**"
        )

        st.write(
            " | ".join(
                [
                    f"{nombre}: **{conteo[nombre]}**"
                    for nombre in categorias_54
                ]
            )
        )

        # ----------------------------------------------------
        # ENTRENAMIENTO
        # ----------------------------------------------------

        if total_etiquetados >= 20:

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

            clases_presentes = (
                df_entrenamiento["Etiqueta"]
                .nunique()
            )

            if clases_presentes < 2:

                st.warning(
                    "⚠️ Se necesitan al menos "
                    "dos categorías diferentes para entrenar."
                )

            else:

                # --------------------------------------------
                # MODELO
                # --------------------------------------------

                modelo_acciones_54 = Pipeline(
                    [
                        (
                            "tfidf",
                            TfidfVectorizer(
                                lowercase=True,
                                strip_accents="unicode",
                                ngram_range=(1, 2),
                                max_features=5000,
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

                modelo_acciones_54.fit(
                    df_entrenamiento["Acción"],
                    df_entrenamiento["Etiqueta"]
                )

                # --------------------------------------------
                # PREDICCIÓN
                # --------------------------------------------

                probabilidades = (
                    modelo_acciones_54
                    .predict_proba(
                        df_ml["Acción"]
                    )
                )

                predicciones = (
                    modelo_acciones_54
                    .predict(
                        df_ml["Acción"]
                    )
                )

                clases_modelo = (
                    modelo_acciones_54
                    .named_steps[
                        "clasificador"
                    ]
                    .classes_
                )

                confianza = (
                    np.max(
                        probabilidades,
                        axis=1
                    )
                )

                nombres_categorias = {
                    numero: nombre
                    for nombre, numero
                    in categorias_54.items()
                }

                df_resultado_54 = df_ml.copy()

                df_resultado_54[
                    "Clasificación IA"
                ] = [
                    nombres_categorias.get(
                        int(valor),
                        "REVISAR"
                    )
                    for valor in predicciones
                ]

                df_resultado_54[
                    "Confianza IA"
                ] = confianza

                # --------------------------------------------
                # NIVEL DE CONFIANZA
                # --------------------------------------------

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

                # --------------------------------------------
                # GUARDAR
                # --------------------------------------------

                st.session_state[
                    "df_resultado_54"
                ] = df_resultado_54

                # --------------------------------------------
                # ESTADÍSTICAS
                # --------------------------------------------

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
                    "🟢 Modelo entrenado y aplicado "
                    "a todas las acciones."
                )

                st.info(
                    f"Total: **{total}** | "
                    f"Alta: **{alta}** | "
                    f"Media: **{media}** | "
                    f"Revisar: **{revisar}**"
                )

                # --------------------------------------------
                # DISTRIBUCIÓN
                # --------------------------------------------

                st.write(
                    "### Clasificación obtenida"
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

                # --------------------------------------------
                # SOLO LOS CASOS DE BAJA CONFIANZA
                # --------------------------------------------

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
                        "⚠️ Muestra de los casos "
                        "con menor confianza."
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

                # --------------------------------------------
                # MATRIZ DEL PRIMER DATAFRAME
                # NO SE AGREGAN COLUMNAS
                # --------------------------------------------

                st.write(
                    "### Matriz de acciones"
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
                f"⚠️ Faltan "
                f"{20 - total_etiquetados} "
                "ejemplos para iniciar el entrenamiento."
            )

except Exception as e:

    st.error(
        f"🔴 5.4 ERROR: {type(e).__name__}: {e}"
    )


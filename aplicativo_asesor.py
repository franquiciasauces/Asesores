# -*- coding: utf-8 -*-
"""Aplicativo Asesor - versión limpia para Streamlit.

Se conserva la lógica de búsqueda híbrida de síntomas del BLOQUE 3F,
que era la última versión de consulta de síntomas antes del BLOQUE 4.

Se eliminaron del aplicativo:
- Google Colab / Google Drive.
- Instalaciones de paquetes con !pip.
- Control de usuarios, login y roles.
- Banco de preguntas y evaluaciones.
- Resultados y avance de evaluaciones.
- Pruebas interactivas con input().
- Versiones anteriores/repetidas de consulta de síntomas.
"""

from pathlib import Path
import os
import glob
import numpy as np
import pandas as pd
import streamlit as st
from unidecode import unidecode
from rapidfuzz import process, fuzz
from sklearn.metrics.pairwise import cosine_similarity

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Aplicativo Asesor",
    page_icon="🔎",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
RUTA_PORTAFOLIO = Path(
    st.sidebar.text_input(
        "Carpeta PORTAFOLIO",
        str(BASE_DIR / "PORTAFOLIO")
    )
).expanduser()

RUTA_DATOS = RUTA_PORTAFOLIO / "DATOS_MATRIZ"
RUTA_MATRIZ = RUTA_DATOS / "MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx"
RUTA_SEMANTICA = RUTA_DATOS / "SEMANTICA"
RUTA_IMAGENES = RUTA_DATOS / "IMAGENESPRODUCTOS"
ARCHIVO_SEMANTICA = RUTA_SEMANTICA / "base_sintomas_semantica.csv"
ARCHIVO_EMBEDDINGS = RUTA_SEMANTICA / "embeddings_sintomas.npy"

MODELO_SEMANTICO = "SINAI/ALIA-MrBERT-es-biomedical-embeddings"
UMBRAL_SEMANTICO = 65.0
MAX_RESULTADOS_SEMANTICOS = 5

# ============================================================
# CARGA DE DATOS
# ============================================================

@st.cache_data(show_spinner=False)
def cargar_matriz(ruta):
    ruta = str(ruta)
    hojas = pd.read_excel(ruta, sheet_name=None)
    return hojas


@st.cache_data(show_spinner=False)
def cargar_semantica(archivo_base, archivo_embeddings):
    base = pd.read_csv(archivo_base, encoding="utf-8-sig")
    emb = np.load(archivo_embeddings)
    return base, emb


@st.cache_resource(show_spinner="Cargando modelo biomédico...")
def cargar_modelo():
    from sentence_transformers import SentenceTransformer
    return SentenceTransformer(MODELO_SEMANTICO)


def limpiar_texto(texto):
    texto = str(texto).lower()
    texto = unidecode(texto).strip()
    return texto


if not RUTA_MATRIZ.exists():
    st.error(
        "No se encontró la matriz principal. "
        f"Verifique la carpeta PORTAFOLIO: {RUTA_PORTAFOLIO}"
    )
    st.stop()

try:
    hojas = cargar_matriz(RUTA_MATRIZ)
except Exception as exc:
    st.error(f"No fue posible cargar la matriz Excel: {exc}")
    st.stop()

productos = hojas.get("Base_Productos", pd.DataFrame())
patologias = hojas.get("Patologias", pd.DataFrame())
condiciones = hojas.get("Condiciones", pd.DataFrame())
reglas = hojas.get("Reglas_Paquetes", pd.DataFrame())
restricciones = hojas.get("Restricciones", pd.DataFrame())
complementarios = hojas.get("Complementarios", pd.DataFrame())
entrevista = hojas.get("Entrevista", pd.DataFrame())

# ============================================================
# BASE SEMÁNTICA
# ============================================================

base_semantica = None
embeddings_sintomas = None
modelo_biomedico = None

if ARCHIVO_SEMANTICA.exists() and ARCHIVO_EMBEDDINGS.exists():
    try:
        base_semantica, embeddings_sintomas = cargar_semantica(
            str(ARCHIVO_SEMANTICA),
            str(ARCHIVO_EMBEDDINGS),
        )
        if len(base_semantica) != len(embeddings_sintomas):
            st.warning(
                "La base semántica y los embeddings no tienen la misma cantidad de registros."
            )
            base_semantica = None
            embeddings_sintomas = None
    except Exception as exc:
        st.warning(f"No fue posible cargar la base semántica: {exc}")
else:
    st.warning(
        "No se encontraron base_sintomas_semantica.csv y/o embeddings_sintomas.npy. "
        "La consulta semántica de síntomas no estará disponible."
    )

# ============================================================
# CONSULTA DE PATOLOGÍAS
# ============================================================

if not patologias.empty and "Patología" in patologias.columns:
    patologias = patologias.copy()
    patologias["Patologia_busqueda"] = patologias["Patología"].apply(limpiar_texto)


def buscar_patologia(texto, limite=5):
    if patologias.empty or "Patologia_busqueda" not in patologias.columns:
        return []

    texto = limpiar_texto(texto)
    if not texto:
        return []

    lista = patologias["Patologia_busqueda"].tolist()
    resultados = process.extract(texto, lista, scorer=fuzz.WRatio, limit=limite)

    salida = []
    for _, puntaje, indice in resultados:
        if puntaje >= 60:
            fila = patologias.iloc[indice]
            salida.append({
                "Patologia_ID": fila.get("Patologia_ID", ""),
                "Patología": fila.get("Patología", ""),
                "Coincidencia": round(float(puntaje), 2),
                "Indice": indice,
            })
    return salida


def ficha_patologia(fila):
    campos = [
        ("Patología", "Patología"),
        ("Descripción", "Descripción breve (para cliente)"),
        ("Causas frecuentes", "Causas frecuentes (resumen)"),
        ("Síntomas / señales clave", "Síntomas/Señales clave (checklist)"),
        ("Objetivo del paquete", "Objetivo del paquete"),
        ("Notas para asesor", "Notas (para asesor)"),
    ]
    for titulo, columna in campos:
        if columna in fila.index and pd.notna(fila[columna]):
            st.markdown(f"**{titulo}**")
            st.write(fila[columna])

# ============================================================
# CONSULTA DE SÍNTOMAS - BLOQUE 3F DEFINITIVO
# ============================================================

if base_semantica is not None:
    df_3f = base_semantica.copy()
    df_3f["Busqueda_3f"] = df_3f["Sintoma"].apply(limpiar_texto)
else:
    df_3f = pd.DataFrame()


def buscar_exacta_3f(consulta):
    if df_3f.empty:
        return []

    consulta_limpia = limpiar_texto(consulta)
    resultados = []

    for _, fila in df_3f.iterrows():
        if consulta_limpia == fila["Busqueda_3f"]:
            resultados.append({
                "Sintoma_consultado": consulta,
                "Sintoma_encontrado": fila["Sintoma"],
                "Patologia_ID": fila["Patologia_ID"],
                "Patologia": fila["Patologia"],
                "Tipo": "Exacta",
                "Puntaje": 100.0,
            })
    return resultados


def buscar_semantica_3f(consulta):
    if df_3f.empty or embeddings_sintomas is None:
        return []

    try:
        if modelo_biomedico is None:
            return []
        embedding_consulta = modelo_biomedico.encode(
            [consulta], normalize_embeddings=True
        )
        similitudes = cosine_similarity(
            embedding_consulta, embeddings_sintomas
        )[0]
    except Exception:
        return []

    indices = np.argsort(similitudes)[::-1][:MAX_RESULTADOS_SEMANTICOS]
    resultados = []

    for indice in indices:
        puntaje = float(similitudes[indice] * 100)
        if puntaje < UMBRAL_SEMANTICO:
            continue

        fila = df_3f.iloc[indice]
        resultados.append({
            "Sintoma_consultado": consulta,
            "Sintoma_encontrado": fila["Sintoma"],
            "Patologia_ID": fila["Patologia_ID"],
            "Patologia": fila["Patologia"],
            "Tipo": "Semántica",
            "Puntaje": round(puntaje, 2),
        })
    return resultados


def buscar_sintoma_3f(consulta):
    resultados = []
    resultados.extend(buscar_exacta_3f(consulta))
    resultados.extend(buscar_semantica_3f(consulta))
    return resultados


def eliminar_duplicados_3f(resultados):
    mejores = {}
    for resultado in resultados:
        clave = (
            resultado["Sintoma_consultado"],
            resultado["Patologia_ID"],
            resultado["Sintoma_encontrado"],
        )
        if clave not in mejores or resultado["Puntaje"] > mejores[clave]["Puntaje"]:
            mejores[clave] = resultado
    return list(mejores.values())


def analizar_sintomas(sintomas):
    sintomas = [x.strip() for x in sintomas if str(x).strip()]
    sintomas_limpios = []
    vistos = set()

    for sintoma in sintomas:
        clave = limpiar_texto(sintoma)
        if clave and clave not in vistos:
            vistos.add(clave)
            sintomas_limpios.append(sintoma)

    if not sintomas_limpios:
        return [], {}

    resultados_por_sintoma = {}
    for sintoma in sintomas_limpios:
        resultados_por_sintoma[sintoma] = eliminar_duplicados_3f(
            buscar_sintoma_3f(sintoma)
        )

    patologias_resultado = {}
    for sintoma, resultados in resultados_por_sintoma.items():
        for resultado in resultados:
            pid = resultado["Patologia_ID"]
            if pid not in patologias_resultado:
                patologias_resultado[pid] = {
                    "Patologia_ID": pid,
                    "Patologia": resultado["Patologia"],
                    "Por_sintoma": {},
                }
            patologias_resultado[pid]["Por_sintoma"].setdefault(sintoma, []).append(resultado)

    resultados_finales = []
    total_sintomas = len(sintomas_limpios)

    for pid, datos in patologias_resultado.items():
        coincidencias_validas = []
        sintomas_respaldo = 0

        for sintoma in sintomas_limpios:
            resultados = datos["Por_sintoma"].get(sintoma, [])
            if not resultados:
                continue

            mejor = sorted(resultados, key=lambda x: x["Puntaje"], reverse=True)[0]
            es_valida = (
                mejor["Tipo"] == "Exacta"
                or (mejor["Tipo"] == "Semántica" and mejor["Puntaje"] >= UMBRAL_SEMANTICO)
            )

            if es_valida:
                sintomas_respaldo += 1
                coincidencias_validas.append(mejor)

        if sintomas_respaldo == 0:
            continue

        cobertura = (sintomas_respaldo / total_sintomas) * 100
        puntajes = [x["Puntaje"] for x in coincidencias_validas]
        promedio = sum(puntajes) / len(puntajes)

        if sintomas_respaldo == total_sintomas and total_sintomas >= 2:
            nivel = "EVIDENCIA ACUMULADA"
        elif sintomas_respaldo >= 2:
            nivel = "EVIDENCIA MÚLTIPLE"
        elif sintomas_respaldo == 1 and promedio >= 70:
            nivel = "COINCIDENCIA FUERTE"
        else:
            nivel = "CANDIDATA - REQUIERE CONFIRMACIÓN"

        resultados_finales.append({
            "Patologia_ID": pid,
            "Patologia": datos["Patologia"],
            "Sintomas_respaldo": sintomas_respaldo,
            "Cobertura": round(cobertura, 2),
            "Promedio": round(promedio, 2),
            "Nivel": nivel,
            "Coincidencias": coincidencias_validas,
        })

    resultados_finales.sort(
        key=lambda x: (x["Sintomas_respaldo"], x["Cobertura"], x["Promedio"]),
        reverse=True,
    )

    return resultados_finales, resultados_por_sintoma

# ============================================================
# BÚSQUEDA DE PRODUCTOS
# ============================================================

def buscar_en_dataframe(df, consulta, columnas_preferidas=None, limite=20):
    if df.empty or not consulta.strip():
        return pd.DataFrame()

    consulta_limpia = limpiar_texto(consulta)
    columnas = columnas_preferidas or list(df.columns)
    columnas = [c for c in columnas if c in df.columns]
    if not columnas:
        columnas = list(df.columns)

    mascara = pd.Series(False, index=df.index)
    for columna in columnas:
        mascara = mascara | df[columna].fillna("").astype(str).apply(limpiar_texto).str.contains(
            consulta_limpia, regex=False, na=False
        )

    return df.loc[mascara].head(limite)

# ============================================================
# BÚSQUEDA DE RESTRICCIONES
# ============================================================

def buscar_restricciones(consulta, limite=20):
    if restricciones.empty:
        return pd.DataFrame()
    return buscar_en_dataframe(restricciones, consulta, limite=limite)

# ============================================================
# INTERFAZ
# ============================================================

st.title("🔎 Aplicativo Asesor")
st.caption("Consulta de productos, patologías, síntomas y restricciones")

with st.sidebar:
    st.markdown("### Estado")
    st.write(f"Patologías: {len(patologias):,}")
    st.write(f"Productos: {len(productos):,}")
    st.write(f"Restricciones: {len(restricciones):,}")
    if base_semantica is not None:
        st.write(f"Síntomas semánticos: {len(base_semantica):,}")
    else:
        st.write("Síntomas semánticos: no disponibles")

    if st.button("Recargar datos"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()

pestanas = st.tabs([
    "🩺 Patologías",
    "🔬 Síntomas",
    "📦 Productos",
    "⚠️ Restricciones",
])

# ------------------------------------------------------------
# PATOLOGÍAS
# ------------------------------------------------------------
with pestanas[0]:
    st.subheader("Consulta de patologías")
    consulta = st.text_input(
        "Ingrese el nombre o parte del nombre de la patología",
        key="consulta_patologia",
    )

    if consulta:
        resultados = buscar_patologia(consulta)
        if not resultados:
            st.warning("No se encontraron coincidencias suficientes.")
        else:
            opciones = [
                f"{r['Patología']} — {r['Coincidencia']:.0f}%"
                for r in resultados
            ]
            seleccion = st.selectbox("Seleccione una patología", opciones)
            indice = opciones.index(seleccion)
            fila = patologias.iloc[resultados[indice]["Indice"]]
            st.divider()
            ficha_patologia(fila)

# ------------------------------------------------------------
# SÍNTOMAS
# ------------------------------------------------------------
with pestanas[1]:
    st.subheader("Consulta de síntomas")
    st.write(
        "Ingrese uno o varios síntomas separados por comas. "
        "La consulta utiliza la versión híbrida semántica definitiva del archivo original."
    )

    entrada = st.text_area(
        "Síntomas",
        placeholder="Ejemplo: dificultad para orinar, mal olor",
        height=100,
    )

    if st.button("Buscar síntomas", type="primary", key="buscar_sintomas"):
        if not entrada.strip():
            st.warning("Ingrese al menos un síntoma.")
        elif base_semantica is None:
            st.error("La base semántica no está disponible.")
        else:
            if modelo_biomedico is None:
                try:
                    modelo_biomedico = cargar_modelo()
                except Exception as exc:
                    st.error(f"No fue posible cargar el modelo biomédico: {exc}")
                    st.stop()

            sintomas = entrada.split(",")
            resultados_finales, resultados_por_sintoma = analizar_sintomas(sintomas)

            if not resultados_finales:
                st.warning(
                    "No se encontró evidencia suficiente. "
                    "Ingrese otro síntoma o describa mejor la molestia."
                )
            else:
                st.markdown("### Resultados integrados por patología")
                for numero, resultado in enumerate(resultados_finales[:10], start=1):
                    with st.expander(
                        f"{numero}. {resultado['Patologia']} — {resultado['Nivel']}",
                        expanded=(numero == 1),
                    ):
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Síntomas respaldados", resultado["Sintomas_respaldo"])
                        c2.metric("Cobertura", f"{resultado['Cobertura']:.2f}%")
                        c3.metric("Promedio", f"{resultado['Promedio']:.2f}%")

                        st.markdown("**Coincidencias que respaldan:**")
                        for coincidencia in resultado["Coincidencias"]:
                            st.write(
                                f"• {coincidencia['Sintoma_consultado']} → "
                                f"{coincidencia['Sintoma_encontrado']} | "
                                f"{coincidencia['Tipo']} | "
                                f"{coincidencia['Puntaje']:.2f}%"
                            )

                mejor = resultados_finales[0]
                resueltos = {x["Sintoma_consultado"] for x in mejor["Coincidencias"]}
                sintomas_ingresados = [x.strip() for x in sintomas if x.strip()]
                sin_coincidencia = [x for x in sintomas_ingresados if x not in resueltos]

                if sin_coincidencia:
                    st.info(
                        "La patología principal está respaldada solo por una parte de los síntomas. "
                        f"Sin coincidencia suficiente: {', '.join(sin_coincidencia)}."
                    )
                else:
                    st.success(
                        "Todos los síntomas ingresados presentan evidencia para la patología principal."
                    )

# ------------------------------------------------------------
# PRODUCTOS
# ------------------------------------------------------------
with pestanas[2]:
    st.subheader("Consulta de productos")
    consulta_producto = st.text_input(
        "Buscar producto",
        key="consulta_producto",
    )

    if consulta_producto:
        columnas_producto = [
            c for c in productos.columns
            if any(x in limpiar_texto(c) for x in ["producto", "nombre", "codigo", "id"])
        ]
        resultados_productos = buscar_en_dataframe(
            productos,
            consulta_producto,
            columnas_preferidas=columnas_producto,
        )
        if resultados_productos.empty:
            st.warning("No se encontraron productos.")
        else:
            st.dataframe(resultados_productos, use_container_width=True, hide_index=True)

# ------------------------------------------------------------
# RESTRICCIONES
# ------------------------------------------------------------
with pestanas[3]:
    st.subheader("Consulta de restricciones")
    consulta_restriccion = st.text_input(
        "Buscar restricción por producto, contenido o texto",
        key="consulta_restriccion",
    )

    if consulta_restriccion:
        resultados_restricciones = buscar_restricciones(consulta_restriccion)
        if resultados_restricciones.empty:
            st.warning("No se encontraron restricciones.")
        else:
            st.dataframe(
                resultados_restricciones,
                use_container_width=True,
                hide_index=True,
            )

st.divider()
st.caption("Aplicativo Asesor — versión depurada sin usuarios ni evaluaciones")

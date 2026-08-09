import os
import re
import glob
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

st.set_page_config(page_title="Aplicativo Asesor", layout="wide")

# ---------- Rutas ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def buscar_archivo(nombre):
    candidatos = [
        os.path.join(BASE_DIR, nombre),
        os.path.join(BASE_DIR, "DATOS_MATRIZ", nombre),
        os.path.join(BASE_DIR, "DATOS_MATRIZ", "SEMANTICA", nombre),
        os.path.join(BASE_DIR, "SEMANTICA", nombre),
    ]
    for ruta in candidatos:
        if os.path.isfile(ruta):
            return ruta
    encontrados = glob.glob(os.path.join(BASE_DIR, "**", nombre), recursive=True)
    return encontrados[0] if encontrados else None


RUTA_MATRIZ = buscar_archivo("MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx")
RUTA_BASE_SEMANTICA = buscar_archivo("base_sintomas_semantica.csv")
RUTA_EMBEDDINGS = buscar_archivo("embeddings_sintomas.npy")


def cargar_matriz(ruta):
    if not ruta:
        raise FileNotFoundError("No se encontró MATRIZ_PRODUCTO_PATOLOGIAS_PAQUETES.xlsx.")
    xls = pd.ExcelFile(ruta)
    hojas = set(xls.sheet_names)
    requeridas = ["Base_Productos", "Patologias", "Condiciones", "Reglas_Paquetes", "Restricciones", "Complementarios", "Entrevista"]
    faltantes = [h for h in requeridas if h not in hojas]
    if faltantes:
        raise ValueError("Faltan hojas en la matriz: " + ", ".join(faltantes))
    return {h: pd.read_excel(ruta, sheet_name=h) for h in requeridas}


@st.cache_data(show_spinner=False)
def cargar_datos(ruta):
    return cargar_matriz(ruta)


try:
    datos = cargar_datos(RUTA_MATRIZ)
    productos = datos["Base_Productos"]
    patologias = datos["Patologias"]
    condiciones = datos["Condiciones"]
    reglas = datos["Reglas_Paquetes"]
    restricciones = datos["Restricciones"]
    complementarios = datos["Complementarios"]
    entrevista = datos["Entrevista"]
except Exception as e:
    st.error(str(e))
    st.stop()


@st.cache_data(show_spinner=False)
def cargar_semantica(ruta_base, ruta_embeddings):
    if not ruta_base or not ruta_embeddings:
        raise FileNotFoundError("No se encontraron base_sintomas_semantica.csv y/o embeddings_sintomas.npy.")
    df = pd.read_csv(ruta_base, encoding="utf-8-sig")
    emb = np.load(ruta_embeddings)
    if len(df) != len(emb):
        raise ValueError("La cantidad de síntomas y embeddings no coincide.")
    return df, emb


try:
    df_sintomas_semantica, embeddings_sintomas = cargar_semantica(RUTA_BASE_SEMANTICA, RUTA_EMBEDDINGS)
except Exception as e:
    st.error(str(e))
    st.stop()


@st.cache_resource(show_spinner="Cargando modelo semántico...")
def cargar_modelo():
    return SentenceTransformer("SINAI/ALIA-MrBERT-es-biomedical-embeddings")


modelo_biomedico = cargar_modelo()


def quitar_tildes(texto):
    texto = "" if texto is None else str(texto)
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def normalizar(texto):
    texto = quitar_tildes(texto).lower().strip()
    texto = re.sub(r"[-_/]", " ", texto)
    return " ".join(texto.split())


def normalizar_producto(nombre):
    if pd.isna(nombre):
        return ""
    nombre = str(nombre)
    nombre = re.sub(r" x \d+ (CAP|ML|GR|SOFGELS|PERLAS|TABLETAS|TABS|CAPS|COMPRIMIDOS)", "", nombre, flags=re.I)
    return re.sub(r"\s+", " ", nombre).strip().upper()


def buscar_patologias(consulta, limite=10):
    q = normalizar(consulta)
    if not q:
        return []
    resultados = []
    nombres = patologias["Patología"].fillna("").astype(str).tolist() if "Patología" in patologias.columns else []
    for match, score, idx in process.extract(q, [normalizar(x) for x in nombres], scorer=fuzz.WRatio, limit=limite):
        if score >= 60:
            fila = patologias.iloc[idx].to_dict()
            fila["Coincidencia"] = round(score, 1)
            fila["Indice"] = idx
            resultados.append(fila)
    return resultados


def buscar_exacta_3f(consulta):
    q = normalizar(consulta)
    resultados = []
    for _, fila in df_sintomas_semantica.iterrows():
        if q == normalizar(fila.get("Sintoma", "")):
            resultados.append({
                "Sintoma_consultado": consulta,
                "Sintoma_encontrado": fila.get("Sintoma", ""),
                "Patologia_ID": fila.get("Patologia_ID", ""),
                "Patologia": fila.get("Patologia", ""),
                "Tipo": "Exacta",
                "Puntaje": 100.0,
            })
    return resultados


def buscar_semantica_3f(consulta):
    embedding = modelo_biomedico.encode([consulta], normalize_embeddings=True)
    similitudes = cosine_similarity(embedding, embeddings_sintomas)[0]
    indices = np.argsort(similitudes)[::-1][:5]
    resultados = []
    for indice in indices:
        puntaje = float(similitudes[indice] * 100)
        if puntaje < 65:
            continue
        fila = df_sintomas_semantica.iloc[indice]
        resultados.append({
            "Sintoma_consultado": consulta,
            "Sintoma_encontrado": fila.get("Sintoma", ""),
            "Patologia_ID": fila.get("Patologia_ID", ""),
            "Patologia": fila.get("Patologia", ""),
            "Tipo": "Semántica",
            "Puntaje": round(puntaje, 2),
        })
    return resultados


def buscar_sintoma_3f(consulta):
    resultados = buscar_exacta_3f(consulta) + buscar_semantica_3f(consulta)
    mejores = {}
    for r in resultados:
        clave = (r["Sintoma_consultado"], r["Patologia_ID"], r["Sintoma_encontrado"])
        if clave not in mejores or r["Puntaje"] > mejores[clave]["Puntaje"]:
            mejores[clave] = r
    return list(mejores.values())


def analizar_sintomas(entrada):
    sintomas = []
    for parte in entrada.split(","):
        parte = parte.strip()
        if parte and normalizar(parte) not in [normalizar(x) for x in sintomas]:
            sintomas.append(parte)
    resultados = {s: buscar_sintoma_3f(s) for s in sintomas}
    integradas = {}
    for sintoma, filas in resultados.items():
        for r in filas:
            pid = r["Patologia_ID"]
            if pid not in integradas:
                integradas[pid] = {"Patologia_ID": pid, "Patologia": r["Patologia"], "Sintomas": {}, "Puntajes": []}
            integradas[pid]["Sintomas"].setdefault(sintoma, []).append(r)
            integradas[pid]["Puntajes"].append(r["Puntaje"])
    ranking = []
    for item in integradas.values():
        cobertura = len(item["Sintomas"]) / max(len(sintomas), 1) * 100
        promedio = float(np.mean(item["Puntajes"])) if item["Puntajes"] else 0
        ranking.append({"Patologia_ID": item["Patologia_ID"], "Patologia": item["Patologia"], "Cobertura": round(cobertura, 1), "Promedio": round(promedio, 1), "Sintomas": item["Sintomas"]})
    ranking.sort(key=lambda x: (x["Cobertura"], x["Promedio"]), reverse=True)
    return sintomas, resultados, ranking


def checklist_patologia(patologia_id):
    filas = df_sintomas_semantica[df_sintomas_semantica["Patologia_ID"].astype(str) == str(patologia_id)]
    return list(dict.fromkeys(filas["Sintoma"].dropna().astype(str).tolist()))


def evaluar_condicion(texto, estado):
    texto = str(texto).strip()
    if "=" in texto:
        c, esperado = texto.split("=", 1)
        c, esperado = c.strip(), esperado.strip().lower()
        r = estado.get(c)
        if r is None:
            return False
        if isinstance(r, list):
            vals = [str(x).lower() for x in r]
            if esperado == "sí":
                return bool(r) and "ninguno" not in vals
            if esperado == "no":
                return not bool(r) or ("ninguno" in vals and len(r) == 1)
            return esperado in vals
        val = str(r).lower()
        if esperado == "sí":
            return val in ("sí", "sospecha") or (c == "C011" and val in ("sobrepeso", "obesidad"))
        if esperado == "no":
            return val == "no"
        return val == esperado
    m = re.search(r"^(.+?)\s+(contiene|incluye)\s+(.+)$", texto, re.I)
    if m:
        r = estado.get(m.group(1).strip())
        valor = m.group(3).strip().lower()
        if r is None:
            return False
        if isinstance(r, list):
            return any(valor in str(x).lower() for x in r)
        return valor in str(r).lower()
    return False


def generar_recomendaciones(patologia_id, estado):
    reglas_actuales = reglas[reglas["Patologia_ID"].astype(str) == str(patologia_id)].copy()
    activadas = []
    for _, regla in reglas_actuales.iterrows():
        expresion = str(regla.get("Condiciones (lógica)", "")).strip()
        if not expresion or expresion.lower() == "nan":
            continue
        if re.search(r"\s+OR\s+", expresion, re.I):
            cumple = any(evaluar_condicion(x, estado) for x in re.split(r"\s+OR\s+", expresion, flags=re.I))
        elif re.search(r"\s+AND\s+", expresion, re.I):
            cumple = all(evaluar_condicion(x, estado) for x in re.split(r"\s+AND\s+", expresion, flags=re.I))
        else:
            cumple = evaluar_condicion(expresion, estado)
        if cumple:
            activadas.append(regla.to_dict())
    if not activadas:
        return pd.DataFrame(), pd.DataFrame()
    activadas_df = pd.DataFrame(activadas)
    productos_info = productos.copy()
    if "Producto" not in productos_info.columns:
        return activadas_df, pd.DataFrame()
    activadas_df["Producto_Normalizado"] = activadas_df["Producto principal"].apply(normalizar_producto)
    productos_info["Producto_Normalizado"] = productos_info["Producto"].apply(normalizar_producto)
    cols = [c for c in ["Producto_Normalizado", "Acciones generales", "Precio público", "Foto"] if c in productos_info.columns]
    activadas_df = activadas_df.merge(productos_info[cols], on="Producto_Normalizado", how="left")
    lista = []
    for _, f in activadas_df.iterrows():
        lista.append({"Producto": f.get("Producto principal"), "Tipo": "Principal", "Precio público": f.get("Precio público"), "Foto": f.get("Foto"), "Acciones generales": f.get("Acciones generales"), "Regla_ID": f.get("Regla_ID")})
        coad = f.get("Coadyuvantes sugeridos (1-3)")
        if pd.notna(coad):
            for c in re.split(r"[;+]", str(coad)):
                if c.strip():
                    lista.append({"Producto": c.strip(), "Tipo": "Coadyuvante", "Precio público": np.nan, "Foto": np.nan, "Acciones generales": np.nan, "Regla_ID": f.get("Regla_ID")})
    sugeridos = pd.DataFrame(lista)
    if not sugeridos.empty:
        sugeridos["_norm"] = sugeridos["Producto"].apply(normalizar_producto)
        sugeridos = sugeridos.drop_duplicates("_norm").drop(columns="_norm").reset_index(drop=True)
    return activadas_df, sugeridos


def calcular_total(df):
    if df.empty or "Precio público" not in df.columns:
        return 0.0
    vals = pd.to_numeric(df["Precio público"], errors="coerce").fillna(0)
    return float(vals.sum())


def buscar_restricciones(consulta):
    q = normalizar(consulta)
    if not q:
        return []
    base = restricciones.copy()
    for col in ["Producto", "Precaución / Contraindicación", "Motivo", "Restriccion_ID", "Alternativas seguras"]:
        if col not in base.columns:
            base[col] = ""
    base["_texto"] = base[["Producto", "Precaución / Contraindicación", "Motivo", "Restriccion_ID", "Alternativas seguras"]].fillna("").astype(str).agg(" ".join, axis=1).map(normalizar)
    resultados = process.extract(q, base["_texto"].tolist(), scorer=fuzz.WRatio, limit=10)
    salida = []
    for _, score, idx in resultados:
        if score >= 70:
            r = base.iloc[idx].drop(labels=["_texto"]).to_dict()
            r["Coincidencia"] = round(score, 1)
            salida.append(r)
    return salida


def resolver_imagen(valor):
    if pd.isna(valor) or not str(valor).strip():
        return None
    valor = str(valor).strip()
    candidatos = [
        os.path.join(BASE_DIR, valor),
        os.path.join(BASE_DIR, "DATOS_MATRIZ", "IMAGENESPRODUCTOS", valor),
        os.path.join(BASE_DIR, "IMAGENESPRODUCTOS", valor),
    ]
    for c in candidatos:
        if os.path.isfile(c):
            return c
    encontrados = glob.glob(os.path.join(BASE_DIR, "**", os.path.basename(valor)), recursive=True)
    return encontrados[0] if encontrados else None


if "historial" not in st.session_state:
    st.session_state.historial = []
if "estado_entrevista" not in st.session_state:
    st.session_state.estado_entrevista = {}

st.title("Aplicativo Asesor")

menu = st.sidebar.radio("Módulo", ["Inicio", "Síntomas", "Patologías", "Entrevista y recomendación", "Restricciones", "Cotización", "Listas de chequeo", "Historial"])

if menu == "Inicio":
    st.subheader("Consulta del aplicativo")
    c1, c2, c3 = st.columns(3)
    c1.metric("Productos", len(productos))
    c2.metric("Patologías", len(patologias))
    c3.metric("Síntomas semánticos", len(df_sintomas_semantica))
    st.info("Seleccione un módulo en el menú lateral.")

elif menu == "Síntomas":
    st.subheader("Consulta de síntomas")
    entrada = st.text_input("Ingrese uno o varios síntomas separados por coma", placeholder="dificultad para orinar, mal olor")
    if st.button("Consultar síntomas", type="primary") and entrada.strip():
        sintomas, individuales, ranking = analizar_sintomas(entrada)
        st.session_state.historial.append({"Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Tipo": "Síntomas", "Consulta": entrada, "Resultado": ranking[0]["Patologia"] if ranking else "Sin coincidencia"})
        if ranking:
            st.subheader("Patologías relacionadas")
            st.dataframe(pd.DataFrame(ranking)[["Patologia_ID", "Patologia", "Cobertura", "Promedio"]], use_container_width=True)
        else:
            st.warning("No se encontraron coincidencias suficientes.")
        with st.expander("Detalle por síntoma"):
            for sintoma in sintomas:
                st.markdown(f"**{sintoma}**")
                filas = sorted(individuales[sintoma], key=lambda x: x["Puntaje"], reverse=True)[:5]
                if filas:
                    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
                else:
                    st.write("Sin coincidencia suficiente.")

elif menu == "Patologías":
    st.subheader("Consulta de patologías")
    q = st.text_input("Código o nombre de la patología")
    resultados = buscar_patologias(q) if q else []
    if resultados:
        opciones = [f"{r.get('Patologia_ID', '')} — {r.get('Patología', '')} ({r['Coincidencia']}%)" for r in resultados]
        elegido = st.selectbox("Seleccione", opciones)
        r = resultados[opciones.index(elegido)]
        st.dataframe(pd.DataFrame([r]).drop(columns=["Indice"], errors="ignore"), use_container_width=True)

elif menu == "Entrevista y recomendación":
    st.subheader("Entrevista y recomendación")
    q = st.text_input("Código o nombre de la patología", key="patologia_entrevista")
    resultados = buscar_patologias(q) if q else []
    if resultados:
        opciones = [f"{r.get('Patologia_ID', '')} — {r.get('Patología', '')} ({r['Coincidencia']}%)" for r in resultados]
        elegido = st.selectbox("Patología", opciones)
        pat = resultados[opciones.index(elegido)]
        pid = pat.get("Patologia_ID")
        filas = entrevista[entrevista["Patologia_ID"].astype(str) == str(pid)].sort_values("Orden") if "Patologia_ID" in entrevista.columns else pd.DataFrame()
        estado = {}
        if not filas.empty:
            for _, fila in filas.iterrows():
                condicion = fila.get("Condicion_ID")
                pregunta = str(fila.get("Pregunta", ""))
                tipo = normalizar(fila.get("Tipo_Control", "texto"))
                opciones_fila = [] if pd.isna(fila.get("Opciones")) else [x.strip() for x in str(fila.get("Opciones")).split(";") if x.strip()]
                key = f"entrevista_{pid}_{condicion}"
                if tipo == "texto":
                    resp = st.text_input(pregunta, key=key)
                elif tipo == "numero":
                    resp = st.number_input(pregunta, value=0.0, key=key)
                elif tipo in ("si no", "si/no"):
                    resp = st.selectbox(pregunta, ["", "Sí", "No"], key=key)
                elif tipo == "lista":
                    resp = st.selectbox(pregunta, [""] + opciones_fila, key=key)
                elif "seleccion multiple" in tipo:
                    resp = st.multiselect(pregunta, opciones_fila, key=key)
                else:
                    resp = st.text_input(pregunta, key=key)
                estado[condicion] = resp if resp not in ("", []) else None
        st.session_state.estado_entrevista = estado
        if st.button("Generar recomendación", type="primary"):
            activadas, sugeridos = generar_recomendaciones(pid, estado)
            if sugeridos.empty:
                st.warning("No se activaron reglas para las respuestas registradas.")
            else:
                st.dataframe(sugeridos, use_container_width=True, hide_index=True)
                st.session_state.productos_cotizacion = sugeridos.copy()
                st.session_state.historial.append({"Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Tipo": "Recomendación", "Consulta": str(pid), "Resultado": f"{len(sugeridos)} productos"})

elif menu == "Restricciones":
    st.subheader("Consulta de restricciones")
    q = st.text_input("Producto, código, precaución o motivo")
    if st.button("Buscar restricciones", type="primary") and q.strip():
        resultados = buscar_restricciones(q)
        if resultados:
            st.dataframe(pd.DataFrame(resultados), use_container_width=True, hide_index=True)
            st.session_state.historial.append({"Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Tipo": "Restricciones", "Consulta": q, "Resultado": f"{len(resultados)} resultados"})
        else:
            st.warning("No se encontraron restricciones relacionadas.")

elif menu == "Cotización":
    st.subheader("Cotización")
    disponibles = st.session_state.get("productos_cotizacion", pd.DataFrame())
    if disponibles.empty:
        st.info("Primero genere una recomendación desde 'Entrevista y recomendación'.")
    else:
        nombres = disponibles["Producto"].astype(str).tolist()
        seleccion = st.multiselect("Seleccione los productos para cotizar", nombres, default=nombres)
        cot = disponibles[disponibles["Producto"].astype(str).isin(seleccion)].copy()
        total = calcular_total(cot)
        for _, fila in cot.iterrows():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"**{fila['Producto']}** — {fila['Tipo']}")
                if pd.notna(fila.get("Precio público")):
                    st.write(f"Precio: ${float(fila['Precio público']):,.0f}")
            with col2:
                img = resolver_imagen(fila.get("Foto"))
                if img:
                    st.image(img, width=130)
        st.divider()
        st.metric("Total de la cotización", f"${total:,.0f}")
        if st.button("Guardar cotización en historial"):
            st.session_state.historial.append({"Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Tipo": "Cotización", "Consulta": ", ".join(seleccion), "Resultado": f"${total:,.0f}"})
            st.success("Cotización guardada en el historial de esta sesión.")

elif menu == "Listas de chequeo":
    st.subheader("Lista de chequeo de síntomas")
    q = st.text_input("Código o nombre de la patología", key="patologia_check")
    resultados = buscar_patologias(q) if q else []
    if resultados:
        opciones = [f"{r.get('Patologia_ID', '')} — {r.get('Patología', '')}" for r in resultados]
        elegido = st.selectbox("Patología", opciones, key="check_pat")
        pid = resultados[opciones.index(elegido)].get("Patologia_ID")
        items = checklist_patologia(pid)
        if items:
            marcados = []
            for i, item in enumerate(items):
                if st.checkbox(item, key=f"check_{pid}_{i}"):
                    marcados.append(item)
            st.write(f"Elementos seleccionados: {len(marcados)} de {len(items)}")
        else:
            st.info("No hay elementos de checklist registrados para esta patología.")

elif menu == "Historial":
    st.subheader("Historial de consultas")
    if st.session_state.historial:
        st.dataframe(pd.DataFrame(st.session_state.historial), use_container_width=True, hide_index=True)
        if st.button("Limpiar historial"):
            st.session_state.historial = []
            st.rerun()
    else:
        st.info("No hay consultas registradas en esta sesión.")

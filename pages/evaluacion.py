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
# 5.3A CLASIFICACIÓN DETERMINISTA (LISTA MAESTRA) Y SINCRONIZACIÓN
# ============================================================
st.markdown("### 5.3 Clasificación Manual (Lista Maestra)")

# Asegurarse de que tenemos la data de la sección 5.2
if "df_acciones_52" not in st.session_state:
    st.error("❌ Primero debes ejecutar la sección 5.2 para tener datos que clasificar.")
else:
    df_acciones = st.session_state["df_acciones_52"]
    
    # 1. Configuración inicial del editor (solo si no existe)
    if "df_maestro_acciones" not in st.session_state:
        acciones_unicas = df_acciones["Acción general"].unique()
        st.session_state["df_maestro_acciones"] = pd.DataFrame({
            "Acción general": acciones_unicas,
            "¿Es General?": True 
        })

    st.info("💡 **Definición Maestra:** Marca con un check las que SÍ son Acciones Generales.")
    
    # 2. Editor interactivo
    df_maestro = st.data_editor(
        st.session_state["df_maestro_acciones"], 
        use_container_width=True, 
        hide_index=True
    )
    # Guardamos cambios en el editor inmediatamente
    st.session_state["df_maestro_acciones"] = df_maestro

    # 3. Botón para Finalizar Validación (Procesa la data)
    if st.button("✅ Finalizar Validación y Generar Archivo"):
        acciones_validas = df_maestro[df_maestro["¿Es General?"] == True]["Acción general"].tolist()
        df_limpio = df_acciones[df_acciones["Acción general"].isin(acciones_validas)].copy()
        
        # Guardamos en sesión para que NO se pierda nunca
        st.session_state["df_limpio"] = df_limpio
        st.success(f"✅ Validación lista: {len(df_limpio)} relaciones procesadas.")

    # 4. Zona de Sincronización (Solo aparece si la validación fue finalizada)
    if "df_limpio" in st.session_state:
        st.write("---")
        st.write("### 📋 Vista previa del resultado:")
        st.dataframe(st.session_state["df_limpio"], use_container_width=True)

        # Botón de Descarga
        csv_csv = st.session_state["df_limpio"].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar CSV Normalizado",
            data=csv_csv,
            file_name="RELACIONES_PRODUCTO_ACCION_GENERAL.csv",
            mime="text/csv"
        )

        # Botón de Sincronización con GitHub
        if st.button("🚀 Guardar y Sincronizar con GitHub"):
            
            # Definimos la función de sync dentro del botón para evitar problemas de alcance
            def ejecutar_sync():
                csv_str = st.session_state["df_limpio"].to_csv(index=False)
                try:
                    token = st.secrets["github"]["token"]
                    repo_full_name = st.secrets["github"]["repo_name"]
                    g = Github(token)
                    repo = g.get_repo(repo_full_name)
                    
                    try:
                        contents = repo.get_contents("RELACIONES_PRODUCTO_ACCION_GENERAL.csv")
                        repo.update_file(contents.path, "Actualización automática", csv_str, contents.sha)
                        return True, "Archivo actualizado exitosamente."
                    except:
                        repo.create_file("RELACIONES_PRODUCTO_ACCION_GENERAL.csv", "Creación automática", csv_str)
                        return True, "Archivo creado exitosamente."
                except Exception as e:
                    return False, str(e)

            with st.spinner("Sincronizando con GitHub..."):
                exito, mensaje = ejecutar_sync()
                if exito:
                    st.success(f"✅ {mensaje}")
                else:
                    st.error(f"❌ Error: {mensaje}")

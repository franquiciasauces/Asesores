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
# 5.3A CLASIFICACIÓN MANUAL - MODO FIJO (CERRADO)
# ============================================================
st.markdown("### 5.3 Clasificación Manual (Lista Maestra)")

if "df_acciones_52" not in st.session_state:
    st.error("❌ Primero ejecuta la sección 5.2.")
else:
    # --- INICIALIZACIÓN DE ESTADOS ---
    if "validacion_fija" not in st.session_state:
        st.session_state["validacion_fija"] = False
    
    # Preparamos el editor solo si no hemos validado aún
    if "df_maestro_acciones" not in st.session_state:
        acciones_unicas = st.session_state["df_acciones_52"]["Acción general"].unique()
        st.session_state["df_maestro_acciones"] = pd.DataFrame({
            "Acción general": acciones_unicas,
            "¿Es General?": True 
        })

    # --- FLUJO LÓGICO ---
    # Si NO está validado (Modo Edición Abierto)
    if not st.session_state["validacion_fija"]:
        st.info("💡 **Modo Edición:** Marca las que son Generales y pulsa 'Finalizar Validación'.")
        
        # El usuario edita aquí
        st.session_state["df_maestro_acciones"] = st.data_editor(
            st.session_state["df_maestro_acciones"], 
            use_container_width=True, 
            hide_index=True
        )

        if st.button("✅ Finalizar Validación Manual"):
            # Procesamos el limpio AHORA y lo fijamos
            acciones_validas = st.session_state["df_maestro_acciones"][st.session_state["df_maestro_acciones"]["¿Es General?"] == True]["Acción general"].tolist()
            st.session_state["df_limpio"] = st.session_state["df_acciones_52"][st.session_state["df_acciones_52"]["Acción general"].isin(acciones_validas)].copy()
            
            # ACTIVAMOS EL CERROJO
            st.session_state["validacion_fija"] = True
            st.rerun()

    # Si YA está validado (Modo Lectura/Sync)
    else:
        st.success("🔒 **Estado: Validación Finalizada y Lista Fija.**")
        
        # Mostrar el resultado fijo
        st.dataframe(st.session_state["df_limpio"], use_container_width=True)

        # Botones de Acción (Sync y Descarga)
        col1, col2 = st.columns(2)
        with col1:
            csv_csv = st.session_state["df_limpio"].to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar CSV", csv_csv, "RELACIONES_PRODUCTO_ACCION_GENERAL.csv", "text/csv")
        
        with col2:
            if st.button("🚀 Guardar y Sincronizar en GitHub"):
                # Función interna para no tener dependencias externas
                csv_str = st.session_state["df_limpio"].to_csv(index=False)
                try:
                    token = st.secrets["github"]["token"]
                    repo_full_name = st.secrets["github"]["repo_name"]
                    repo = Github(token).get_repo(repo_full_name)
                    
                    try:
                        contents = repo.get_contents("RELACIONES_PRODUCTO_ACCION_GENERAL.csv")
                        repo.update_file(contents.path, "Actualización manual desde app", csv_str, contents.sha)
                        st.success("✅ Sincronizado!")
                    except:
                        repo.create_file("RELACIONES_PRODUCTO_ACCION_GENERAL.csv", "Creación desde app", csv_str)
                        st.success("✅ Creado en GitHub!")
                except Exception as e:
                    st.error(f"Error: {e}")

        # BOTÓN DE SEGURIDAD PARA EDITAR
        if st.button("✏️ Desbloquear para editar de nuevo"):
            st.session_state["validacion_fija"] = False
            st.rerun()

import pandas as pd
import streamlit as st

# 1. Configuración de la página en modo ancho
st.set_page_config(
    page_title="Dashboard Ejecutivo - Seguimiento Operativo", layout="wide"
)

st.title("📊 Dashboard Ejecutivo - Seguimiento Operativo y Contractual")
st.markdown(
    "Visualización interactiva sincronizada con la estructura de datos del equipo."
)

# 2. Widget para subir el archivo Excel directamente desde la interfaz
st.sidebar.header("📁 Cargar Archivo de Datos")
uploaded_file = st.sidebar.file_uploader(
    "Sube tu archivo Excel (RoadMap)", type=["xlsx", "xls"]
)

# Validamos si se ha subido un archivo
if uploaded_file is not None:
    # Cargar los datos desde el archivo subido
    @st.cache_data
    def cargar_datos(file):
        df_roadmap = pd.read_excel(file, sheet_name="RoadMap")
        return df_roadmap


    df = cargar_datos(uploaded_file)

    # 3. Filtros interactivos en la barra lateral (Sidebar)
    st.sidebar.header("🔍 Filtros Globales")

    aliados = ["Todos"] + list(df["ALIADOS"].unique())
    sel_aliado = st.sidebar.selectbox("Aliado", options=aliados)
    if sel_aliado != "Todos":
        df = df[df["ALIADOS"] == sel_aliado]

    deptos = ["Todos"] + list(df["DEPARTAMENTO"].unique())
    sel_depto = st.sidebar.selectbox("Departamento", options=deptos)
    if sel_depto != "Todos":
        df = df[df["DEPARTAMENTO"] == sel_depto]

    estados_pgrs = ["Todos"] + list(df["ESTADO PQRS CON PDR"].unique())
    sel_estado = st.sidebar.selectbox("Estado PQRS con PDR", options=estados_pgrs)
    if sel_estado != "Todos":
        df = df[df["ESTADO PQRS CON PDR"] == sel_estado]

    viabilidades = ["Todos"] + list(df["VIABILIDAD DE MTTO"].unique())
    sel_viabilidad = st.sidebar.selectbox("Viabilidad de Mtto", options=viabilidades)
    if sel_viabilidad != "Todos":
        df = df[df["VIABILIDAD DE MTTO"] == sel_viabilidad]

    # 4. Indicadores Principales (KPIs)
    st.subheader("Indicadores Clave de Operación (KPIs)")

    total_casos = len(df)
    vencidos = len(df[df["ESTADO PQRS CON PDR"] == "VENCIDO"])
    vigentes = len(df[df["ESTADO PQRS CON PDR"] == "VIGENTE"])
    ejecutables = len(df[df["VIABILIDAD DE MTTO"] == "EJECUTABLE"])
    no_ejecutables = len(df[df["VIABILIDAD DE MTTO"] == "NO EJECUTABLE"])

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric(label="Total Casos / PQRSD", value=total_casos)
    with col2:
        st.metric(label="Casos Vencidos", value=vencidos)
    with col3:
        st.metric(label="Casos Vigentes", value=vigentes)
    with col4:
        st.metric(label="Mtto Ejecutable", value=ejecutables)
    with col5:
        st.metric(label="Mtto No Ejecutable", value=no_ejecutables)

    st.markdown("---")

    # 5. Replicar la Tabla Dinámica de Análisis (Estilo Tabla Dinámica)
    st.subheader("📈 Resumen Dinámico de Análisis (Estilo Tabla Dinámica)")

    if not df.empty:
        df["FECHA_VISITA_STR"] = (
            pd.to_datetime(df["FECHA VISITA"])
            .dt.strftime("%Y-%m-%d")
            .fillna("Sin Fecha")
        )

        pivot_df = pd.pivot_table(
            df,
            values="ID BENEFICIARIO",
            index=["PRIORIDAD", "FECHA_VISITA_STR"],
            columns="DEPARTAMENTO",
            aggfunc="count",
            fill_value=0,
        )

        pivot_df["Total General"] = pivot_df.sum(axis=1)
        st.dataframe(pivot_df, use_container_width=True)
    else:
        st.warning("No hay datos disponibles con los filtros seleccionados.")

    # 6. Tabla de Detalle Operativo Completo
    st.markdown("---")
    st.subheader("🔍 Detalle Operativo (Base Completa Filtrada)")
    st.dataframe(df.drop(columns=["FECHA_VISITA_STR"], errors="ignore"))

else:
    # Mensaje inicial cuando no se ha cargado ningún archivo
    st.info(
        "👈 Por favor, utiliza el panel de la izquierda para **subir tu archivo Excel** (`RoadMap_11.09.26.xlsx`) y comenzar."
    )
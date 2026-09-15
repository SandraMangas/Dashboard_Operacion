from datetime import datetime
from io import BytesIO
from pathlib import Path
import os

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Dashboard ejecutivo | RoadMap",
    page_icon=":material/monitoring:",
    layout="wide",
)


# ============================================================
# RUTAS DEL PROYECTO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TESERACT_LOGO = (
    BASE_DIR
    / "assets"
    / "Logo_Teseract.jpeg"
)

ARCHIVO_GUARDADO = (
    BASE_DIR
    / "data"
    / "archivo_actual.xlsx"
)


# ============================================================
# PALETA
# ============================================================

PALETTE = {
    "navy": "#17365D",
    "blue": "#3D6E9E",
    "teal": "#2F7E7B",
    "amber": "#C68A1D",
    "red": "#BF4E4E",
    "slate": "#5B6776",
    "light_blue": "#EAF1F7",
}


# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>

      /* =====================================================
         HEADER
         ===================================================== */

      .st-key-executive_header {
          background: #17365D;
          padding: 18px 22px 16px;
          border-radius: 14px;
          margin-bottom: 20px;
          box-shadow: 0 7px 18px rgba(23, 54, 93, .18);
      }

      .header-title {
          color: #FFFFFF;
          font-size: 1.34rem;
          font-weight: 750;
          line-height: 1.2;
      }

      .header-subtitle {
          color: #D9E7F5;
          margin-top: 5px;
          font-size: .84rem;
          line-height: 1.35;
      }

      .st-key-executive_header [data-testid="stCaptionContainer"] {
          color: #D9E7F5;
      }

      .header-status-detail {
          color: #D9E7F5;
          font-size: .73rem;
          margin-top: 3px;
          line-height: 1.35;
          text-align: right;
      }

      /* =====================================================
         KPI
         ===================================================== */

      .kpi-card {
          min-height: 82px;
          height: 82px;
          padding: 9px 12px;
          border-radius: 10px;
          border: 1px solid #D8E1EA;
          border-top: 4px solid #3D6E9E;
          background: #FFFFFF;
          box-sizing: border-box;
      }

      .kpi-card--blue {
          border-top-color: #3D6E9E;
          background: #F7FAFD;
      }

      .kpi-card--teal {
          border-top-color: #2F7E7B;
          background: #F4FAF9;
      }

      .kpi-card--amber {
          border-top-color: #C68A1D;
          background: #FFF9ED;
      }

      .kpi-card--red {
          border-top-color: #BF4E4E;
          background: #FFF5F4;
      }

      .kpi-label {
          display: block;
          color: #5B6776;
          font-size: 0.64rem;
          font-weight: 700;
          letter-spacing: .035em;
          text-transform: uppercase;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          text-align: center;
      }

      .kpi-value {
          display: block;
          margin-top: 5px;
          color: #17365D;
          font-size: 1.32rem;
          font-weight: 750;
          line-height: 1;
          text-align: center;
      }

      .kpi-detail {
          display: block;
          margin-top: 5px;
          color: #5B6776;
          font-size: .70rem;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          text-align: center;
      }

      .dashboard-space {
          height: 7px;
      }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# UTILIDADES
# ============================================================

def columna(
    df: pd.DataFrame,
    nombre: str,
) -> pd.Series:
    """Devuelve la serie de una columna o una serie vacía."""

    if nombre in df.columns:
        return df[nombre]

    return pd.Series(dtype="object")


def limpiar_texto(
    serie: pd.Series,
) -> pd.Series:
    """Normaliza textos para evitar problemas con vacíos y espacios."""

    return (
        serie
        .fillna("Sin información")
        .astype(str)
        .str.strip()
        .replace("", "Sin información")
    )


def buscar_columna(
    df: pd.DataFrame,
    candidatos: list[str],
) -> str | None:
    """
    Busca primero coincidencia exacta y luego parcial.
    """

    columnas = {
        str(col).strip().upper(): col
        for col in df.columns
    }

    # --------------------------------------------------------
    # COINCIDENCIA EXACTA
    # --------------------------------------------------------

    for candidato in candidatos:

        clave = candidato.strip().upper()

        if clave in columnas:
            return columnas[clave]

    # --------------------------------------------------------
    # COINCIDENCIA PARCIAL
    # --------------------------------------------------------

    for clave, original in columnas.items():

        for candidato in candidatos:

            candidato_normalizado = (
                candidato.strip().upper()
            )

            if candidato_normalizado in clave:
                return original

    return None


def calcular_porcentaje(
    valor: int | float,
    total: int,
) -> float:
    """Calcula un porcentaje evitando división por cero."""

    if total == 0:
        return 0.0

    return valor / total * 100


# ============================================================
# CARGA DE DATOS
# ============================================================

@st.cache_data(ttl="10m")
def cargar_datos(
    file_bytes: bytes,
) -> pd.DataFrame:
    """
    Lee la hoja RoadMap del archivo Excel.
    """

    return pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="RoadMap",
    )


# ============================================================
# CARGA DEL ROADMAP
# ============================================================

def cargar_roadmap() -> pd.DataFrame:
    """
    Si existe archivo_actual.xlsx, lo carga automáticamente.

    Si el usuario selecciona un nuevo archivo (solo local):
    - se guarda como archivo_actual.xlsx
    - se limpia la caché
    - se carga la nueva versión
    """

    ARCHIVO_GUARDADO.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Detección de entorno (Producción en Streamlit Cloud o Local)
    es_produccion = (
        st.secrets.get("modo_produccion", False) 
        or os.getenv("STREAMLIT_SHARING_MODE") is not None
    )

    archivo_nuevo = None

    with st.sidebar:

        st.header(
            "Datos y filtros",
            icon=":material/tune:",
        )

        # ----------------------------------------------------
        # ESTADO DEL ARCHIVO
        # ----------------------------------------------------

        if ARCHIVO_GUARDADO.exists():

            st.success(
                "Archivo cargado",
                icon=":material/check_circle:",
            )

            st.caption(
                f"Fuente actual: **{ARCHIVO_GUARDADO.name}**"
            )

        else:

            st.info(
                "Aún no hay un archivo guardado.",
                icon=":material/info:",
            )

        # ----------------------------------------------------
        # ACTUALIZAR ARCHIVO (Solo visible en entorno local)
        # ----------------------------------------------------

        if not es_produccion:
            archivo_nuevo = st.file_uploader(
                "Actualizar archivo",
                type=["xlsx", "xls"],
                key="roadmap_file",
                help=(
                    "Selecciona un nuevo Excel para actualizar "
                    "la información del dashboard."
                ),
            )

        st.caption(
            "La aplicación conserva el último archivo "
            "cargado para abrirlo automáticamente."
        )

    # ========================================================
    # NUEVO ARCHIVO
    # ========================================================

    if archivo_nuevo is not None:

        nuevos_bytes = archivo_nuevo.getvalue()

        ARCHIVO_GUARDADO.write_bytes(
            nuevos_bytes
        )

        cargar_datos.clear()

        datos = (
            cargar_datos(
                nuevos_bytes
            )
            .copy()
        )

    # ========================================================
    # ARCHIVO EXISTENTE
    # ========================================================

    elif ARCHIVO_GUARDADO.exists():

        datos = (
            cargar_datos(
                ARCHIVO_GUARDADO.read_bytes()
            )
            .copy()
        )

    else:

        return pd.DataFrame()

    # ========================================================
    # FILTROS
    # ========================================================

    filtros = [
        ("ALIADOS", "Aliado"),
        ("DEPARTAMENTO", "Departamento"),
        ("PRIORIDAD", "Prioridad"),
        ("ESTADO PQRS CON PDR", "Estado PQRS"),
        ("VIABILIDAD DE MTTO", "Viabilidad"),
    ]

    with st.sidebar:

        st.subheader(
            "Filtros principales",
            icon=":material/filter_list:",
        )

        for campo, etiqueta in filtros:

            if campo not in datos.columns:
                continue

            opciones = (
                datos[campo]
                .dropna()
                .astype(str)
                .str.strip()
                .replace(
                    "",
                    "Sin información",
                )
                .unique()
                .tolist()
            )

            opciones = [
                "Todos"
            ] + sorted(opciones)

            elegido = st.selectbox(
                etiqueta,
                opciones,
                key=f"filter_{campo}",
            )

            if elegido != "Todos":

                datos = datos[
                    datos[campo]
                    .astype(str)
                    .str.strip()
                    .eq(elegido)
                ]

    return datos


# ============================================================
# RELOJ
# ============================================================

def mostrar_reloj() -> None:

    ahora = datetime.now().astimezone()

    st.caption(
        f":material/schedule: "
        f"Actualizado desde el equipo: "
        f"**{ahora:%d/%m/%Y · %I:%M %p}**"
    )


@st.fragment(run_every="60s")
def reloj_actualizable() -> None:

    mostrar_reloj()


# ============================================================
# FIGURAS
# ============================================================

def figura_base(
    fig,
    height: int = 290,
):

    fig.update_layout(
        height=height,
        margin=dict(
            t=20,
            b=10,
            l=10,
            r=10,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            color="#25364A"
        ),
        coloraxis_colorbar=dict(
            title=None
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
    )

    fig.update_xaxes(
        showgrid=True,
        gridcolor="#E7EDF3",
        zeroline=False,
    )

    fig.update_yaxes(
        showgrid=False,
        zeroline=False,
    )

    return fig


# ============================================================
# TARJETA KPI
# ============================================================

def tarjeta_kpi(
    etiqueta: str,
    valor: str,
    detalle: str,
    color: str,
) -> None:

    st.markdown(
        f"""
        <div class="kpi-card kpi-card--{color}">
          <span class="kpi-label">{etiqueta}</span>
          <span class="kpi-value">{valor}</span>
          <span class="kpi-detail">{detalle}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# CLASIFICACIÓN DE VISITAS
# ============================================================

def clasificar_visitas(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Clasifica las visitas según la fecha disponible.

    Valores:
    - Sin fecha
    - Visita vencida
    - Visita hoy
    - Visita programada
    - Visita programada - caso vencido
    """

    resultado = df.copy()

    col_fecha = buscar_columna(
        resultado,
        [
            "FECHA VISITA",
            "FECHA DE VISITA",
            "FECHA PROGRAMADA",
            "PROGRAMACION VISITA",
            "PROGRAMACIÓN VISITA",
        ],
    )

    if col_fecha is None:

        resultado["_ESTADO_VISITA"] = (
            "Sin fecha"
        )

        return resultado

    fechas = (
        pd.to_datetime(
            resultado[col_fecha],
            errors="coerce",
        )
        .dt.normalize()
    )

    hoy = pd.Timestamp.now().normalize()

    resultado["_ESTADO_VISITA"] = (
        "Sin fecha"
    )

    # --------------------------------------------------------
    # VISITA VENCIDA
    # --------------------------------------------------------

    resultado.loc[
        fechas.notna()
        & (fechas < hoy),
        "_ESTADO_VISITA",
    ] = "Visita vencida"

    # --------------------------------------------------------
    # VISITA HOY
    # --------------------------------------------------------

    resultado.loc[
        fechas == hoy,
        "_ESTADO_VISITA",
    ] = "Visita hoy"

    # --------------------------------------------------------
    # VISITA PROGRAMADA
    # --------------------------------------------------------

    resultado.loc[
        fechas > hoy,
        "_ESTADO_VISITA",
    ] = "Visita programada"

    # --------------------------------------------------------
    # CASO VENCIDO + VISITA FUTURA
    # --------------------------------------------------------

    col_estado_pqrs = buscar_columna(
        resultado,
        [
            "ESTADO PQRS CON PDR",
            "ESTADO PQRS",
            "ESTADO PQRSD",
        ],
    )

    if col_estado_pqrs is not None:

        estados_pqrs = (
            limpiar_texto(
                resultado[col_estado_pqrs]
            )
            .str.upper()
        )

        caso_vencido = (
            estados_pqrs == "VENCIDO"
        )

        visita_programada = (
            resultado["_ESTADO_VISITA"]
            == "Visita programada"
        )

        resultado.loc[
            caso_vencido
            & visita_programada,
            "_ESTADO_VISITA",
        ] = (
            "Visita programada - caso vencido"
        )

    return resultado


# ============================================================
# ESCALADOS A PROYECTOS
# ============================================================

def mascara_escalado_proyectos(
    df: pd.DataFrame,
) -> pd.Series:
    """
    Identifica los casos escalados a proyectos
    a partir de la columna ESTADO.

    Se considera escalado cuando ESTADO contiene
    la palabra PROYECT.
    """

    if "ESTADO" not in df.columns:

        return pd.Series(
            False,
            index=df.index,
        )

    estados = (
        limpiar_texto(
            df["ESTADO"]
        )
        .str.upper()
    )

    return estados.str.contains(
        "PROYECT",
        na=False,
    )


def contar_escalados_proyectos(
    df: pd.DataFrame,
) -> tuple[int, str | None]:
    """
    Cuenta casos escalados a proyectos.
    """

    if "ESTADO" not in df.columns:

        return 0, None

    mascara = mascara_escalado_proyectos(
        df
    )

    return (
        int(mascara.sum()),
        "ESTADO",
    )


# ============================================================
# ESTADO GERENCIAL
# ============================================================

def calcular_estado_gerencial(
    df: pd.DataFrame,
) -> dict:
    """
    Determina el semáforo gerencial.

    Regla:

    🔴 REQUIERE ATENCIÓN
       Hay vencidos NO escalados a proyectos.

    🟠 ATENCIÓN PREVENTIVA
       No hay vencidos operativos, pero sí próximos
       a vencer NO escalados a proyectos.

    🟢 OPERACIÓN CONTROLADA
       No hay vencidos ni próximos a vencer operativos.

    Los casos escalados a proyectos NO generan alerta.
    """

    estado_pqrs = (
        columna(
            df,
            "ESTADO PQRS CON PDR",
        )
        .fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
    )

    escalado_proyectos = (
        mascara_escalado_proyectos(
            df
        )
    )

    # --------------------------------------------------------
    # VENCIDOS OPERATIVOS
    # --------------------------------------------------------

    vencidos_operativos = (
        estado_pqrs.eq("VENCIDO")
        & ~escalado_proyectos
    )

    # --------------------------------------------------------
    # PRÓXIMOS A VENCER OPERATIVOS
    # --------------------------------------------------------

    proximos_operativos = (
        estado_pqrs.isin(
            [
                "PRÓXIMO A VENCER",
                "PROXIMO A VENCER",
            ]
        )
        & ~escalado_proyectos
    )

    vencidos_n = int(
        vencidos_operativos.sum()
    )

    proximos_n = int(
        proximos_operativos.sum()
    )

    escalados_n = int(
        escalado_proyectos.sum()
    )

    # --------------------------------------------------------
    # SEMÁFORO
    # --------------------------------------------------------

    if vencidos_n > 0:

        estado = {
            "texto": "Estado: requiere atención",
            "icono": ":material/warning:",
            "color": "red",
            "detalle": (
                f"{vencidos_n:,} "
                f"vencidos operativos"
            ),
        }

    elif proximos_n > 0:

        estado = {
            "texto": "Estado: atención preventiva",
            "icono": ":material/priority_high:",
            "color": "orange",
            "detalle": (
                f"{proximos_n:,} "
                f"próximos a vencer operativos"
            ),
        }

    else:

        estado = {
            "texto": "Estado: operación controlada",
            "icono": ":material/check_circle:",
            "color": "green",
            "detalle": (
                "Sin vencidos ni próximos a vencer "
                "operativos"
            ),
        }

    return {
        "vencidos_operativos": vencidos_operativos,
        "proximos_operativos": proximos_operativos,
        "escalado_proyectos": escalado_proyectos,
        "vencidos_n": vencidos_n,
        "proximos_n": proximos_n,
        "escalados_n": escalados_n,
        "estado": estado,
    }


# ============================================================
# HEADER
# ============================================================

def mostrar_cabecera(
    estado_gerencial: dict,
) -> None:

    with st.container(
        key="executive_header",
        gap="small",
    ):

        logo, texto, estado = st.columns(
            [1.45, 5.85, 2.0],
            gap="small",
            vertical_alignment="center",
        )

        # ----------------------------------------------------
        # LOGO
        # ----------------------------------------------------

        with logo:

            if TESERACT_LOGO.exists():

                st.image(
                    TESERACT_LOGO,
                    width=150,
                )

        # ----------------------------------------------------
        # TÍTULO
        # ----------------------------------------------------

        with texto:

            st.markdown(
                """
                <div class="header-title">
                    Dashboard Ejecutivo —
                    Seguimiento Operativo y Contractual
                </div>

                <div class="header-subtitle">
                    Panorama gerencial de RoadMap ·
                    riesgos, vencimientos,
                    ejecutabilidad y capacidad operativa
                </div>
                """,
                unsafe_allow_html=True,
            )

        # ----------------------------------------------------
        # SEMÁFORO GERENCIAL
        # ----------------------------------------------------

        with estado:

            estado_info = (
                estado_gerencial["estado"]
            )

            st.badge(
                estado_info["texto"],
                icon=estado_info["icono"],
                color=estado_info["color"],
            )

            st.markdown(
                f"""
                <div class="header-status-detail">
                    {estado_info["detalle"]}
                </div>
                """,
                unsafe_allow_html=True,
            )

            # ------------------------------------------------
            # INFORMACIÓN DE ESCALADOS
            # ------------------------------------------------

            escalados_n = (
                estado_gerencial[
                    "escalados_n"
                ]
            )

            if escalados_n > 0:

                st.markdown(
                    f"""
                    <div class="header-status-detail">
                        {escalados_n:,} escalados a proyectos
                        · fuera de alerta operativa
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            reloj_actualizable()


# ============================================================
# DASHBOARD
# ============================================================

def dashboard(
    df: pd.DataFrame,
) -> None:

    # --------------------------------------------------------
    # PREPARACIÓN
    # --------------------------------------------------------

    df = clasificar_visitas(
        df
    )

    total_registros = len(df)

    # --------------------------------------------------------
    # ESTADO PQRSD
    # --------------------------------------------------------

    estado_pqrs = (
        columna(
            df,
            "ESTADO PQRS CON PDR",
        )
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # VIABILIDAD
    # --------------------------------------------------------

    viabilidad = (
        columna(
            df,
            "VIABILIDAD DE MTTO",
        )
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # ESTADO GERENCIAL
    # --------------------------------------------------------

    estado_gerencial = (
        calcular_estado_gerencial(
            df
        )
    )

    vencidos_n = (
        estado_gerencial[
            "vencidos_n"
        ]
    )

    proximos_n = (
        estado_gerencial[
            "proximos_n"
        ]
    )

    escalados_n = (
        estado_gerencial[
            "escalados_n"
        ]
    )

    # --------------------------------------------------------
    # MÉTRICAS GENERALES
    # --------------------------------------------------------

    ejecutables_n = int(
        viabilidad
        .str.upper()
        .eq("EJECUTABLE")
        .sum()
    )

    no_ejecutables_n = int(
        viabilidad
        .str.upper()
        .eq("NO EJECUTABLE")
        .sum()
    )

    # --------------------------------------------------------
    # CD
    # --------------------------------------------------------

    id_beneficiario = columna(
        df,
        "ID BENEFICIARIO",
    )

    if id_beneficiario.empty:

        total_cd = total_registros

    else:

        total_cd = (
            id_beneficiario
            .nunique()
        )

    # --------------------------------------------------------
    # PQRSD
    # --------------------------------------------------------

    pqrs = columna(
        df,
        "CANT. PQRSD",
    )

    if not pqrs.empty:

        total_pqrs = (
            pd.to_numeric(
                pqrs,
                errors="coerce",
            )
            .fillna(0)
            .sum()
        )

    else:

        total_pqrs = total_registros

    # ========================================================
    # HEADER
    # ========================================================

    mostrar_cabecera(
        estado_gerencial
    )

    # ========================================================
    # INDICADORES CLAVE
    # ========================================================

    st.subheader(
        "Indicadores clave",
        icon=":material/query_stats:",
    )

    # --------------------------------------------------------
    # FILA 1
    # --------------------------------------------------------

    (
        margen_izq,
        kpi_1,
        sep_1,
        kpi_2,
        sep_2,
        kpi_3,
        sep_3,
        kpi_4,
        margen_der,
    ) = st.columns(
        [
            0.35,
            1,
            0.10,
            1,
            0.10,
            1,
            0.10,
            1,
            0.35,
        ]
    )

    with kpi_1:

        tarjeta_kpi(
            "CD",
            f"{total_cd:,}",
            "unidades",
            "blue",
        )

    with kpi_2:

        tarjeta_kpi(
            "PQRSD",
            f"{int(total_pqrs):,}",
            "carga de casos",
            "blue",
        )

    with kpi_3:

        tarjeta_kpi(
            "Vencidos",
            (
                f"{calcular_porcentaje(vencidos_n, total_registros):.1f}%"
            ),
            f"{vencidos_n:,} operativos",
            "red",
        )

    with kpi_4:

        tarjeta_kpi(
            "Próximos a vencer",
            f"{proximos_n:,}",
            "operativos",
            "amber",
        )

    st.markdown(
        '<div class="dashboard-space"></div>',
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # FILA 2
    # --------------------------------------------------------

    (
        margen_izq,
        kpi_5,
        sep_1,
        kpi_6,
        sep_2,
        kpi_7,
        margen_der,
    ) = st.columns(
        [
            0.75,
            1,
            0.10,
            1,
            0.10,
            1,
            0.75,
        ]
    )

    with kpi_5:

        tarjeta_kpi(
            "Ejecutables",
            (
                f"{calcular_porcentaje(ejecutables_n, total_registros):.1f}%"
            ),
            f"{ejecutables_n:,} casos",
            "teal",
        )

    with kpi_6:

        tarjeta_kpi(
            "No ejecutables",
            (
                f"{calcular_porcentaje(no_ejecutables_n, total_registros):.1f}%"
            ),
            f"{no_ejecutables_n:,} casos",
            "red",
        )

    with kpi_7:

        tarjeta_kpi(
            "Escalados a proyectos",
            f"{escalados_n:,}",
            "fuera de alerta operativa",
            "amber",
        )

    # ========================================================
    # ALERTAS
    # ========================================================

    st.markdown(
        '<div class="dashboard-space"></div>',
        unsafe_allow_html=True,
    )

    alertas = st.columns(
        3,
        gap="medium",
    )

    with alertas[0].container(
        border=True,
    ):

        if vencidos_n > 0:

            st.markdown(
                ":red-badge[Crítico]"
            )

            st.write(
                f"**{vencidos_n:,}** "
                f"casos vencidos operativos requieren "
                f"priorización."
            )

        else:

            st.markdown(
                ":green-badge[Controlado]"
            )

            st.write(
                "No existen casos vencidos "
                "operativos."
            )

    with alertas[1].container(
        border=True,
    ):

        if proximos_n > 0:

            st.markdown(
                ":orange-badge[Atención]"
            )

            st.write(
                f"**{proximos_n:,}** "
                f"casos operativos están próximos a vencer."
            )

        else:

            st.markdown(
                ":green-badge[Controlado]"
            )

            st.write(
                "No existen casos próximos a vencer "
                "operativos."
            )

    with alertas[2].container(
        border=True,
    ):

        st.markdown(
            ":blue-badge[Proyectos]"
        )

        st.write(
            f"**{escalados_n:,}** "
            f"casos escalados a proyectos "
            f"se gestionan fuera de la alerta operativa."
        )

    # ========================================================
    # ESTADO PQRSD + VIABILIDAD
    # ========================================================

    izquierda, derecha = st.columns(
        2,
        gap="medium",
    )

    # --------------------------------------------------------
    # ESTADO PQRSD
    # --------------------------------------------------------

    with izquierda.container(
        border=True,
    ):

        st.subheader(
            "Estado de PQRSD",
            icon=":material/donut_small:",
        )

        if "ESTADO PQRS CON PDR" in df.columns:

            fig = px.pie(
                df,
                names="ESTADO PQRS CON PDR",
                hole=0.62,
                color="ESTADO PQRS CON PDR",
                color_discrete_map={
                    "VIGENTE": PALETTE["teal"],
                    "PRÓXIMO A VENCER": PALETTE["amber"],
                    "PROXIMO A VENCER": PALETTE["amber"],
                    "VENCIDO": PALETTE["red"],
                },
            )

            st.plotly_chart(
                figura_base(fig),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna de estado PQRSD."
            )

    # --------------------------------------------------------
    # VIABILIDAD
    # --------------------------------------------------------

    with derecha.container(
        border=True,
    ):

        st.subheader(
            "Viabilidad de mantenimiento",
            icon=":material/build:",
        )

        if "VIABILIDAD DE MTTO" in df.columns:

            fig = px.pie(
                df,
                names="VIABILIDAD DE MTTO",
                hole=0.62,
                color="VIABILIDAD DE MTTO",
                color_discrete_map={
                    "EJECUTABLE": PALETTE["teal"],
                    "NO EJECUTABLE": PALETTE["red"],
                },
            )

            st.plotly_chart(
                figura_base(fig),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna de viabilidad."
            )

    # ========================================================
    # ACCIÓN + VISITAS
    # ========================================================

    izquierda, derecha = st.columns(
        2,
        gap="medium",
    )

    # --------------------------------------------------------
    # ACCIÓN REQUERIDA
    # --------------------------------------------------------

    with izquierda.container(
        border=True,
    ):

        st.subheader(
            "Acción requerida",
            icon=":material/task_alt:",
        )

        col_accion = buscar_columna(
            df,
            [
                "ACCIÓN",
                "ACCION",
                "ACCIÓN REQUERIDA",
                "ACCION REQUERIDA",
            ],
        )

        if col_accion:

            acciones = limpiar_texto(
                df[col_accion]
            )

            resumen_accion = (
                acciones
                .value_counts()
                .sort_values()
                .reset_index()
            )

            resumen_accion.columns = [
                "Acción",
                "Casos",
            ]

            fig = px.bar(
                resumen_accion,
                x="Casos",
                y="Acción",
                orientation="h",
                text="Casos",
                color_discrete_sequence=[
                    PALETTE["blue"]
                ],
            )

            fig.update_traces(
                textposition="outside",
                cliponaxis=False,
            )

            st.plotly_chart(
                figura_base(
                    fig,
                    max(
                        290,
                        len(resumen_accion) * 42,
                    ),
                ),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna ACCIÓN."
            )

    # --------------------------------------------------------
    # PROGRAMACIÓN DE VISITAS
    # --------------------------------------------------------

    with derecha.container(
        border=True,
    ):

        st.subheader(
            "Programación de visitas",
            icon=":material/event_available:",
        )

        visitas = (
            df["_ESTADO_VISITA"]
            .value_counts()
            .sort_values()
            .reset_index()
        )

        visitas.columns = [
            "Estado",
            "Casos",
        ]

        fig = px.bar(
            visitas,
            x="Casos",
            y="Estado",
            orientation="h",
            text="Casos",
            color="Estado",
            color_discrete_map={
                "Visita vencida": PALETTE["red"],
                "Visita hoy": PALETTE["amber"],
                "Visita programada": PALETTE["teal"],
                "Visita programada - caso vencido": PALETTE["red"],
                "Sin fecha": PALETTE["slate"],
            },
        )

        fig.update_traces(
            textposition="outside",
            cliponaxis=False,
        )

        st.plotly_chart(
            figura_base(
                fig,
                max(
                    290,
                    len(visitas) * 45,
                ),
            ),
            width="stretch",
        )

    # ========================================================
    # DEPARTAMENTO + BLOQUEOS
    # ========================================================

    izquierda, derecha = st.columns(
        2,
        gap="medium",
    )

    # --------------------------------------------------------
    # DEPARTAMENTO
    # --------------------------------------------------------

    with izquierda.container(
        border=True,
    ):

        st.subheader(
            "Casos por departamento",
            icon=":material/location_on:",
        )

        if "DEPARTAMENTO" in df.columns:

            resumen = (
                df["DEPARTAMENTO"]
                .fillna("Sin departamento")
                .astype(str)
                .str.strip()
                .value_counts()
                .head(10)
                .sort_values()
                .reset_index()
            )

            resumen.columns = [
                "Departamento",
                "Casos",
            ]

            fig = px.bar(
                resumen,
                x="Casos",
                y="Departamento",
                orientation="h",
                text="Casos",
                color_discrete_sequence=[
                    PALETTE["blue"]
                ],
            )

            fig.update_traces(
                textposition="outside",
                cliponaxis=False,
            )

            st.plotly_chart(
                figura_base(fig),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna de departamento."
            )

    # --------------------------------------------------------
    # BLOQUEOS
    # --------------------------------------------------------

    with derecha.container(
        border=True,
    ):

        st.subheader(
            "Principales bloqueos",
            icon=":material/block:",
        )

        if {
            "VIABILIDAD DE MTTO",
            "NOVEDAD",
        }.issubset(df.columns):

            viabilidad_bloqueos = (
                df["VIABILIDAD DE MTTO"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            bloqueos = df.loc[
                viabilidad_bloqueos.eq(
                    "NO EJECUTABLE"
                ),
                "NOVEDAD",
            ]

            resumen = (
                bloqueos
                .dropna()
                .astype(str)
                .str.strip()
                .replace(
                    "",
                    "Sin información",
                )
                .value_counts()
                .head(10)
                .sort_values()
                .reset_index()
            )

            resumen.columns = [
                "Novedad",
                "Casos",
            ]

            if resumen.empty:

                st.success(
                    "No hay bloqueos en los filtros seleccionados."
                )

            else:

                fig = px.bar(
                    resumen,
                    x="Casos",
                    y="Novedad",
                    orientation="h",
                    text="Casos",
                    color_discrete_sequence=[
                        PALETTE["red"]
                    ],
                )

                fig.update_traces(
                    textposition="outside",
                    cliponaxis=False,
                )

                st.plotly_chart(
                    figura_base(fig),
                    width="stretch",
                )

        else:

            st.info(
                "No se encontraron las columnas "
                "de viabilidad y novedad."
            )

    # ========================================================
    # TABLA DETALLADA
    # ========================================================

    with st.container(
        border=True,
    ):

        st.subheader(
            "Dataframe",
            icon=":material/table_chart:",
        )

        buscar = st.text_input(
            "Buscar en la tabla",
            placeholder=(
                "Municipio, técnico, cédula o novedad..."
            ),
        )

        columnas = [
            "ALIADOS",
            "ID BENEFICIARIO",
            "DEPARTAMENTO",
            "MUNICIPIO",
            "PRIORIDAD",
            "ESTADO PQRS CON PDR",
            "ESTADO",
            "ACCIÓN",
            "VIABILIDAD DE MTTO",
            "FECHA VISITA",
            "TÉCNICO",
            "NOVEDAD",
        ]

        columnas_disponibles = [
            c
            for c in columnas
            if c in df.columns
        ]

        tabla = df[
            columnas_disponibles
        ].copy()

        if buscar:

            coincidencia = (
                tabla
                .astype(str)
                .apply(
                    lambda serie:
                    serie.str.contains(
                        buscar,
                        case=False,
                        na=False,
                        regex=False,
                    )
                )
                .any(axis=1)
            )

            tabla = tabla[
                coincidencia
            ]

        st.dataframe(
            tabla,
            height=420,
            width="stretch",
            hide_index=True,
        )


# ============================================================
# EJECUCIÓN
# ============================================================

st.sidebar.caption(
    "Dashboard ejecutivo · RoadMap"
)

datos = cargar_roadmap()

if datos.empty:

    st.title(
        "Dashboard ejecutivo",
        icon=":material/monitoring:",
    )

    if ARCHIVO_GUARDADO.exists():

        st.warning(
            "No hay registros para la combinación "
            "de filtros seleccionada."
        )

    else:

        st.info(
            "Carga un archivo Excel con la hoja "
            "**RoadMap** para comenzar.",
            icon=":material/upload:",
        )

else:

    dashboard(
        datos
    )
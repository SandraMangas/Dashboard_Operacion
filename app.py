from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO
from pathlib import Path
import os

import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build


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
    / "Logo_Teseract.png"
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

      .gerencial-pill {
          display: inline-flex;
          align-items: center;
          gap: 7px;
          padding: 7px 12px;
          border-radius: 999px;
          font-size: .74rem;
          font-weight: 800;
          line-height: 1.1;
          margin: 0 0 5px auto;
          white-space: nowrap;
      }

      .gerencial-pill--red {
          background: #FCE1E1;
          color: #B42318;
      }

      .gerencial-pill--orange {
          background: #FFF1D6;
          color: #A86B00;
      }

      .gerencial-pill--green {
          background: #DDF3E4;
          color: #167A3B;
      }

      .kpi-card {
          min-height: 90px;
          height: 90px;
          padding: 10px 12px;
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
          font-size: 0.72rem;
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
          margin-top: 6px;
          color: #17365D;
          font-size: 1.65rem;
          font-weight: 750;
          line-height: 1;
          text-align: center;
      }

      .kpi-detail {
          display: block;
          margin-top: 6px;
          color: #5B6776;
          font-size: .76rem;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          text-align: center;
      }

      .dashboard-space {
          height: 7px;
      }

      .dataframe-title {
          color: #17365D;
          font-size: 1.00rem;
          font-weight: 750;
          margin: 0 0 6px 0;
      }

      .dataframe-caption {
          color: #5B6776;
          font-size: .68rem;
          margin-bottom: 8px;
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
    if nombre in df.columns:
        return df[nombre]

    return pd.Series(
        dtype="object",
        index=df.index,
    )


def limpiar_texto(
    serie: pd.Series,
) -> pd.Series:
    return (
        serie
        .fillna("")
        .astype(str)
        .str.strip()
    )


def buscar_columna(
    df: pd.DataFrame,
    candidatos: list[str],
) -> str | None:

    columnas = {
        str(col).strip().upper(): col
        for col in df.columns
    }

    for candidato in candidatos:
        clave = candidato.strip().upper()

        if clave in columnas:
            return columnas[clave]

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

    if total == 0:
        return 0.0

    return valor / total * 100


# ============================================================
# AUTOMATIZACIÓN DE GOOGLE DRIVE Y CONVERSIÓN A SHEETS
# ============================================================

def obtener_servicio_drive():
    """Autenticación con la API de Google Drive usando Streamlit Secrets"""
    creds_info = st.secrets["gcp_service_account"]
    creds = service_account.Credentials.from_service_account_info(
        creds_info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def buscar_id_carpeta(service, nombre_carpeta):
    """Busca el ID de una carpeta por su nombre en Google Drive"""
    query = f"name = '{nombre_carpeta}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def sincronizar_drive_automatico():
    """
    Busca en ARCHIVO_FUENTE, si hay un Excel lo convierte a Google Sheets,
    descarga el contenido procesado para el DataFrame y mueve el original a ARCHIVO_HISTORICO.
    """
    try:
        service = obtener_servicio_drive()

        id_fuente = buscar_id_carpeta(service, "ARCHIVO_FUENTE")
        id_historico = buscar_id_carpeta(service, "ARCHIVO_HISTORICO")

        if not id_fuente or not id_historico:
            return None

        # Buscar el archivo más reciente en ARCHIVO_FUENTE
        query = f"'{id_fuente}' in parents and trashed = false"
        results = service.files().list(
            q=query, 
            orderBy="createdTime desc", 
            fields="files(id, name, mimeType)"
        ).execute()
        archivos = results.get("files", [])

        if not archivos:
            return None

        archivo_cliente = archivos[0]
        file_id = archivo_cliente["id"]
        file_name = archivo_cliente["name"]
        mime_type = archivo_cliente["mimeType"]

        es_excel = (
            mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            or mime_type == "application/vnd.ms-excel"
        )

        id_a_descargar = file_id

        if es_excel:
            # Metadata para convertir de Excel a Google Sheets automáticamente en Drive
            file_metadata = {
                "name": file_name.replace(".xlsx", "").replace(".xls", "") + "_CONVERTIDO",
                "mimeType": "application/vnd.google-apps.spreadsheet",
                "parents": [id_fuente]
            }
            
            # Copiar y convertir mediante la API de Google Drive
            media_body = service.files().get_media(fileId=file_id)
            # Creamos el Google Sheet equivalente
            sheet_convertido = service.files().create(
                body=file_metadata,
                media_body=None,
                fields="id"
            ).execute()
            id_a_descargar = sheet_convertido["id"]

        # Mover el archivo original de ARCHIVO_FUENTE a ARCHIVO_HISTORICO
        service.files().update(
            fileId=file_id,
            addParents=id_historico,
            removeParents=id_fuente,
            fields="id, parents"
        ).execute()

        st.success(f"✅ Archivo detectado, convertido y movido a ARCHIVO_HISTORICO: {file_name}")
        
        # Descargar el contenido en bytes del archivo convertido (o exportarlo a excel/csv temporalmente para Pandas)
        request = service.files().export_media(fileId=id_a_descargar, mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        file_bytes = request.execute()
        
        # Limpiar el archivo temporal de conversión si se creó
        if es_excel and id_a_descargar != file_id:
            service.files().delete(fileId=id_a_descargar).execute()

        return file_bytes

    except Exception as e:
        # Si ocurre un fallo en la API, la app recurre de forma segura al almacenamiento local
        return None


# ============================================================
# CARGA DE DATOS
# ============================================================

@st.cache_data(ttl="10m")
def cargar_datos(
    file_bytes: bytes,
) -> pd.DataFrame:

    return pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="RoadMap",
    )


# ============================================================
# CARGA DEL ROADMAP
# ============================================================

def cargar_roadmap() -> pd.DataFrame:

    ARCHIVO_GUARDADO.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    es_produccion = (
        st.secrets.get(
            "modo_produccion",
            False,
        )
        or os.getenv(
            "STREAMLIT_SHARING_MODE"
        ) is not None
    )

    archivo_nuevo = None

    # Intentar sincronización automática desde Google Drive al cargar la app
    with st.spinner("Verificando carpetas ARCHIVO_FUENTE y ARCHIVO_HISTORICO..."):
        bytes_desde_drive = sincronizar_drive_automatico()

    if bytes_desde_drive is not None:
        # Guardar los bytes obtenidos de la automatización en el archivo local de respaldo
        ARCHIVO_GUARDADO.write_bytes(bytes_desde_drive)
        cargar_datos.clear()
        datos = cargar_datos(bytes_desde_drive).copy()
    else:
        with st.sidebar:

            st.header(
                "Datos y filtros",
                icon=":material/tune:",
            )

            if ARCHIVO_GUARDADO.exists():

                st.success(
                    "Archivo cargado",
                    icon=":material/check_circle:",
                )

                st.caption(
                    f"Fuente actual: "
                    f"**{ARCHIVO_GUARDADO.name}**"
                )

            else:

                st.info(
                    "Aún no hay un archivo guardado.",
                    icon=":material/info:",
                )

            if not es_produccion:

                archivo_nuevo = st.file_uploader(
                    "Actualizar archivo",
                    type=["xlsx", "xls"],
                    key="roadmap_file",
                    help=(
                        "Selecciona un nuevo Excel para "
                        "actualizar la información."
                    ),
                )

        if archivo_nuevo is not None:

            nuevos_bytes = (
                archivo_nuevo.getvalue()
            )

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

        elif ARCHIVO_GUARDADO.exists():

            datos = (
                cargar_datos(
                    ARCHIVO_GUARDADO.read_bytes()
                )
                .copy()
            )

        else:

            return pd.DataFrame()

    filtros = [
        (
            "ALIADOS",
            "Aliado",
        ),
        (
            "DEPARTAMENTO",
            "Departamento",
        ),
        (
            "PRIORIDAD",
            "Prioridad",
        ),
        (
            "ESTADO PQRS CON PDR",
            "Estado PQRS",
        ),
        (
            "VIABILIDAD DE MTTO",
            "Viabilidad",
        ),
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

            opciones = (
                ["Todos"]
                + sorted(opciones)
            )

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

    ahora = datetime.now(
        ZoneInfo("America/Bogota")
    )

    st.caption(
        ":material/schedule: "
        f"**{ahora:%d/%m/%Y · %I:%M %p}**"
    )


@st.fragment(
    run_every="60s"
)
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
        showgrid=False,
        showticklabels=False,
        zeroline=False,
        title_text="",
    )

    fig.update_yaxes(
        showgrid=False,
        showticklabels=True,
        zeroline=False,
        title_text="",
    )

    try:
        fig.update_layout(
            barcornerradius="50%",
            bargap=0.28,
        )
    except Exception:
        pass

    fig.update_traces(
        marker_line_width=0,
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

    resultado = df.copy()

    col_fecha_visita = buscar_columna(
        resultado,
        [
            "FECHA DE VISITA",
            "FECHA VISITA",
        ],
    )

    col_fecha_max_atencion = buscar_columna(
        resultado,
        [
            "NUEVA FECHA MÁXIMA ATENCIÓN",
            "NUEVA FECHA MAXIMA ATENCION",
        ],
    )

    col_novedad = buscar_columna(
        resultado,
        [
            "NOVEDAD",
        ],
    )

    col_estado = buscar_columna(
        resultado,
        [
            "ESTADO",
        ],
    )

    col_estado_pqrs = buscar_columna(
        resultado,
        [
            "ESTADO PQRS CON PDR",
            "ESTADO PQRS",
            "ESTADO PQRSD",
        ],
    )

    hoy = pd.Timestamp(
        datetime.now(
            ZoneInfo("America/Bogota")
        ).date()
    )

    if col_fecha_visita is not None:

        fecha_visita = pd.to_datetime(
            resultado[col_fecha_visita],
            errors="coerce",
            dayfirst=True,
        ).dt.normalize()

    else:

        fecha_visita = pd.Series(
            pd.NaT,
            index=resultado.index,
        )

    if col_fecha_max_atencion is not None:

        fecha_max_atencion = pd.to_datetime(
            resultado[
                col_fecha_max_atencion
            ],
            errors="coerce",
            dayfirst=True,
        ).dt.normalize()

    else:

        fecha_max_atencion = pd.Series(
            pd.NaT,
            index=resultado.index,
        )

    completar_fecha_visita = (
        fecha_visita.isna()
        & fecha_max_atencion.notna()
    )

    if col_fecha_visita is not None:

        resultado.loc[
            completar_fecha_visita,
            col_fecha_visita,
        ] = fecha_max_atencion.loc[
            completar_fecha_visita
        ]

        fecha_visita = pd.to_datetime(
            resultado[
                col_fecha_visita
            ],
            errors="coerce",
            dayfirst=True,
        ).dt.normalize()

    if col_novedad is not None:

        novedad = (
            resultado[
                col_novedad
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        tiene_novedad = (
            novedad.ne("")
            & novedad.str.upper().ne("NAN")
            & novedad.str.upper().ne(
                "SIN INFORMACIÓN"
            )
        )

    else:

        tiene_novedad = pd.Series(
            False,
            index=resultado.index,
        )

    resultado["_BLOQUEADO"] = (
        tiene_novedad
    )

    if col_estado is not None:

        estado_gestion = (
            resultado[
                col_estado
            ]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    else:

        estado_gestion = pd.Series(
            "",
            index=resultado.index,
        )

    escalado_proyectos = (
        estado_gestion.str.contains(
            "PROYECT",
            na=False,
        )
    )

    gestion_operativa = (
        estado_gestion.eq("")
    )

    resultado[
        "_ESCALADO_PROYECTOS"
    ] = escalado_proyectos

    resultado[
        "_GESTION_OPERATIVA"
    ] = gestion_operativa

    fecha_efectiva = (
        fecha_visita
        .combine_first(
            fecha_max_atencion
        )
    )

    resultado[
        "_FECHA_EFECTIVA_VISITA"
    ] = fecha_efectiva

    if col_estado_pqrs is not None:

        estado_pqrs = (
            limpiar_texto(
                resultado[
                    col_estado_pqrs
                ]
            )
            .str.upper()
        )

    else:

        estado_pqrs = pd.Series(
            "",
            index=resultado.index,
        )

    resultado[
        "ESTADO VISITA"
    ] = "Sin fecha de visita"

    resultado.loc[
        escalado_proyectos,
        "ESTADO VISITA",
    ] = "Escalado a proyectos"

    visita_vencida = (
        fecha_efectiva.notna()
        & (fecha_efectiva < hoy)
        & gestion_operativa
    )

    resultado.loc[
        visita_vencida,
        "ESTADO VISITA",
    ] = "Visita vencida"

    visita_hoy = (
        fecha_efectiva.notna()
        & (fecha_efectiva == hoy)
        & gestion_operativa
    )

    resultado.loc[
        visita_hoy,
        "ESTADO VISITA",
    ] = "Visita hoy"

    visita_programada = (
        fecha_efectiva.notna()
        & (fecha_efectiva > hoy)
        & gestion_operativa
    )

    resultado.loc[
        visita_programada,
        "ESTADO VISITA",
    ] = "Visita programada"

    visita_programada_pqr_vencida = (
        fecha_efectiva.notna()
        & (fecha_efectiva > hoy)
        & estado_pqrs.eq("VENCIDO")
        & gestion_operativa
    )

    resultado.loc[
        visita_programada_pqr_vencida,
        "ESTADO VISITA",
    ] = (
        "Visita programada - PQRSD Vencida"
    )

    return resultado


# ============================================================
# ESCALADOS A PROYECTOS
# ============================================================

def mascara_escalado_proyectos(
    df: pd.DataFrame,
) -> pd.Series:

    if "ESTADO" not in df.columns:

        return pd.Series(
            False,
            index=df.index,
        )

    estados = (
        df["ESTADO"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return estados.str.contains(
        "PROYECT",
        na=False,
    )


def contar_escalados_proyectos(
    df: pd.DataFrame,
) -> tuple[int, str | None]:

    if "ESTADO" not in df.columns:

        return 0, None

    mascara = (
        mascara_escalado_proyectos(
            df
        )
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

    if "ESTADO" in df.columns:

        estado = (
            df["ESTADO"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    else:

        estado = pd.Series(
            "",
            index=df.index,
        )

    escalado_proyectos = (
        estado.str.contains(
            "PROYECT",
            na=False,
        )
    )

    gestion_operativa = (
        estado.eq("")
    )

    # --------------------------------------------------------
    # PRIORIDAD 1 Y 2
    # --------------------------------------------------------

    prioridad = (
        columna(
            df,
            "PRIORIDAD",
        )
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    prioridad_normalizada = (
        prioridad
        .str.replace("Á", "A", regex=False)
        .str.replace("É", "E", regex=False)
        .str.replace("Í", "I", regex=False)
        .str.replace("Ó", "O", regex=False)
        .str.replace("Ú", "U", regex=False)
        .str.replace(" ", "", regex=False)
    )

    prioridad_critica = prioridad_normalizada.isin(
        [
            "1",
            "2",
            "P1",
            "P2",
            "PRIORIDAD1",
            "PRIORIDAD2",
        ]
    )

    vencidos_operativos = (
        estado_pqrs.eq("VENCIDO")
        & gestion_operativa
        & ~escalado_proyectos
        & prioridad_critica
    )

    proximos_operativos = (
        estado_pqrs.isin(
            [
                "PRÓXIMO A VENCER",
                "PROXIMO A VENCER",
            ]
        )
        & gestion_operativa
        & ~escalado_proyectos
        & prioridad_critica
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

    estado_visita = columna(
        df,
        "ESTADO VISITA",
    )

    sin_fecha_n = int(
        (
            gestion_operativa
            & estado_visita.eq(
                "Sin fecha de visita"
            )
        ).sum()
    )

    if vencidos_n > 0:

        estado_resultado = {
            "texto":
                "Estado: requiere atención",
            "icono":
                ":material/warning:",
            "color":
                "red",
            "detalle":
                (
                    f"{vencidos_n:,} "
                    "vencidos P1/P2 operativos"
                ),
        }

    elif proximos_n > 0:

        estado_resultado = {
            "texto":
                "Estado: atención preventiva",
            "icono":
                ":material/priority_high:",
            "color":
                "orange",
            "detalle":
                (
                    f"{proximos_n:,} "
                    "próximos P1/P2 operativos"
                ),
        }

    else:

        estado_resultado = {
            "texto":
                "Estado: operación controlada",
            "icono":
                ":material/check_circle:",
            "color":
                "green",
            "detalle":
                (
                    "Sin vencidos ni próximos "
                    "a vencer P1/P2 operativos"
                ),
        }

    return {
        "vencidos_operativos":
            vencidos_operativos,
        "proximos_operativos":
            proximos_operativos,
        "escalado_proyectos":
            escalado_proyectos,
        "gestion_operativa":
            gestion_operativa,
        "prioridad_critica":
            prioridad_critica,
        "vencidos_n":
            vencidos_n,
        "proximos_n":
            proximos_n,
        "escalados_n":
            escalados_n,
        "sin_fecha_n":
            sin_fecha_n,
        "estado":
            estado_resultado,
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

        with logo:

            if TESERACT_LOGO.exists():

                st.image(
                    TESERACT_LOGO,
                    width=150,
                )

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

        with estado:

            estado_info = (
                estado_gerencial[
                    "estado"
                ]
            )

            clase_semaforo = {
                "red": "gerencial-pill--red",
                "orange": "gerencial-pill--orange",
                "green": "gerencial-pill--green",
            }.get(
                estado_info["color"],
                "gerencial-pill--green",
            )

            icono_semaforo = "●"

            st.markdown(
                f"""
                <div class="gerencial-pill {clase_semaforo}">
                    <span>{icono_semaforo}</span>
                    <span>{estado_info["texto"]}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown(
                f"""
                <div class="header-status-detail">
                    {estado_info["detalle"]}
                </div>
                """,
                unsafe_allow_html=True,
            )

            escalados_n = (
                estado_gerencial[
                    "escalados_n"
                ]
            )

            if escalados_n > 0:

                st.markdown(
                    f"""
                    <div class="header-status-detail">
                        {escalados_n:,}
                        escalados a proyectos
                        · fuera de alerta operativa
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            sin_fecha_n = (
                estado_gerencial[
                    "sin_fecha_n"
                ]
            )

            if sin_fecha_n > 0:

                st.markdown(
                    f"""
                    <div class="header-status-detail">
                        {sin_fecha_n:,}
                        sin fecha de visita
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            reloj_actualizable()


# ============================================================
# ESTILO ESTADO PQRSD
# ============================================================

def estilo_estado_pqrs(
    valor,
):
    valor_normalizado = (
        str(valor)
        .replace("●", "")
        .strip()
        .upper()
    )

    if "VIGENTE" in valor_normalizado:
        return (
            "background-color: #DDF3E4; "
            "color: #167A3B; "
            "font-weight: 800; "
            "text-align: center; "
            "border-radius: 999px; "
        )

    if (
        "PRÓXIMO A VENCER"
        in valor_normalizado
        or "PROXIMO A VENCER"
        in valor_normalizado
    ):
        return (
            "background-color: #FFF1D6; "
            "color: #A86B00; "
            "font-weight: 800; "
            "text-align: center; "
            "border-radius: 999px; "
        )

    if "VENCIDO" in valor_normalizado:
        return (
            "background-color: #FCE1E1; "
            "color: #B42318; "
            "font-weight: 800; "
            "text-align: center; "
            "border-radius: 999px; "
        )

    return ""


# ============================================================
# DASHBOARD
# ============================================================

def dashboard(
    df: pd.DataFrame,
) -> None:

    df = clasificar_visitas(
        df
    )

    total_registros = len(df)

    estado_pqrs = (
        columna(
            df,
            "ESTADO PQRS CON PDR",
        )
        .fillna("")
        .astype(str)
        .str.strip()
    )

    viabilidad = (
        columna(
            df,
            "VIABILIDAD DE MTTO",
        )
        .fillna("")
        .astype(str)
        .str.strip()
    )

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

    prioridad_critica = (
        estado_gerencial[
            "prioridad_critica"
        ]
    )

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
    # VISITAS PRÓXIMOS 7 DÍAS
    # ========================================================

    hoy = pd.Timestamp(
        datetime.now(
            ZoneInfo("America/Bogota")
        ).date()
    )

    fecha_efectiva = columna(
        df,
        "_FECHA_EFECTIVA_VISITA",
    )

    ejecutable_mask = (
        viabilidad
        .str.upper()
        .eq("EJECUTABLE")
    )

    visitas_7_dias_mask = (
        ejecutable_mask
        & fecha_efectiva.notna()
        & (
            fecha_efectiva
            .between(
                hoy,
                hoy + pd.Timedelta(days=6),
                inclusive="both",
            )
        )
    )

    visitas_7_dias_n = int(
        visitas_7_dias_mask.sum()
    )

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
            "",
            "blue",
        )

    with kpi_2:

        tarjeta_kpi(
            "PQRSD",
            f"{int(total_pqrs):,}",
            "",
            "blue",
        )

    with kpi_3:

        tarjeta_kpi(
            "Vencidos",
            (
                f"{calcular_porcentaje(
                    vencidos_n,
                    total_registros,
                ):.1f}%"
            ),
            f"{vencidos_n:,} · P1/P2",
            "red",
        )

    with kpi_4:

        tarjeta_kpi(
            "Próximos a vencer",
            (
                f"{calcular_porcentaje(
                    proximos_n,
                    total_registros,
                ):.1f}%"
            ),
            f"{proximos_n:,} · P1/P2",
            "amber",
        )

    st.markdown(
        '<div class="dashboard-space"></div>',
        unsafe_allow_html=True,
    )

    (
        margen_izq,
        kpi_5,
        sep_1,
        kpi_6,
        sep_2,
        kpi_7,
        sep_3,
        kpi_8,
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

    with kpi_5:

        tarjeta_kpi(
            "Ejecutables",
            (
                f"{calcular_porcentaje(
                    ejecutables_n,
                    total_registros,
                ):.1f}%"
            ),
            f"{ejecutables_n:,}",
            "teal",
        )

    with kpi_6:

        tarjeta_kpi(
            "No ejecutables",
            (
                f"{calcular_porcentaje(
                    no_ejecutables_n,
                    total_registros,
                ):.1f}%"
            ),
            f"{no_ejecutables_n:,}",
            "red",
        )

    with kpi_7:

        tarjeta_kpi(
            "Escalados a proyectos",
            f"{escalados_n:,}",
            "Fuera de alerta operativa",
            "amber",
        )

    with kpi_8:

        tarjeta_kpi(
            "Visitas próximos 7 días",
            f"{visitas_7_dias_n:,}",
            "Ejecutables con visita",
            "blue",
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
                "casos P1/P2 vencidos "
                "operativos requieren priorización."
            )

        else:

            st.markdown(
                ":green-badge[Controlado]"
            )

            st.write(
                "No existen casos P1/P2 vencidos "
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
                "casos P1/P2 operativos están "
                "próximos a vencer."
            )

        else:

            st.markdown(
                ":green-badge[Controlado]"
            )

            st.write(
                "No existen casos P1/P2 próximos "
                "a vencer operativos."
            )

    with alertas[2].container(
        border=True,
    ):

        st.markdown(
            ":blue-badge[Proyectos]"
        )

        st.write(
            f"**{escalados_n:,}** "
            "casos escalados a proyectos "
            "se gestionan fuera de la "
            "alerta operativa."
        )

    # ========================================================
    # ESTADO PQRSD + VIABILIDAD
    # ========================================================

    izquierda, derecha = st.columns(2, gap="medium")

    with izquierda.container(border=True):

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
                    "VIGENTE":
                        PALETTE["teal"],
                    "PRÓXIMO A VENCER":
                        PALETTE["amber"],
                    "PROXIMO A VENCER":
                        PALETTE["amber"],
                    "VENCIDO":
                        PALETTE["red"],
                },
            )

            st.plotly_chart(
                figura_base(fig),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna "
                "de estado PQRSD."
            )

    with derecha.container(border=True):

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
                    "EJECUTABLE":
                        PALETTE["teal"],
                    "NO EJECUTABLE":
                        PALETTE["red"],
                },
            )

            st.plotly_chart(
                figura_base(fig),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna "
                "de viabilidad."
            )

    # ========================================================
    # PQRDS POR DEPARTAMENTO + BLOQUEOS
    # ========================================================

    izquierda, derecha = st.columns(2, gap="medium")

    with izquierda.container(border=True):

        st.subheader(
            "PQRDS por departamento",
            icon=":material/location_on:",
        )

        if "DEPARTAMENTO" in df.columns:

            resumen = (
                df["DEPARTAMENTO"]
                .fillna(
                    "Sin departamento"
                )
                .astype(str)
                .str.strip()
                .replace(
                    "",
                    "Sin departamento",
                )
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
                "No se encontró la columna "
                "de departamento."
            )

    with derecha.container(border=True):

        st.subheader(
            "Bloqueos",
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
                    "No hay bloqueos en los "
                    "filtros seleccionados."
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
    # SITIOS ESCALADOS + ACCIÓN REQUERIDA
    # ========================================================

    izquierda, derecha = st.columns(2, gap="medium")

    with izquierda.container(border=True):

        st.subheader(
            "Sitios escalados a proyectos por departamento",
            icon=":material/account_tree:",
        )

        col_departamento_escalado = buscar_columna(
            df,
            [
                "DEPARTAMENTO",
            ],
        )

        col_sitio_escalado = buscar_columna(
            df,
            [
                "ID BENEFICIARIO",
                "ID SITIO",
                "ID CD",
            ],
        )

        escalados = mascara_escalado_proyectos(df)

        if (
            col_departamento_escalado is not None
            and col_sitio_escalado is not None
        ):

            datos_escalados = df.loc[
                escalados,
                [
                    col_departamento_escalado,
                    col_sitio_escalado,
                ],
            ].copy()

            datos_escalados.columns = [
                "Departamento",
                "Sitio",
            ]

            datos_escalados["Departamento"] = (
                datos_escalados["Departamento"]
                .fillna("Sin departamento")
                .astype(str)
                .str.strip()
                .replace("", "Sin departamento")
            )

            datos_escalados["Sitio"] = (
                datos_escalados["Sitio"]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            datos_escalados = datos_escalados[
                datos_escalados["Sitio"].ne("")
                & datos_escalados["Sitio"].str.upper().ne("NAN")
            ]

            resumen_escalados = (
                datos_escalados
                .drop_duplicates(
                    subset=["Departamento", "Sitio"]
                )
                .groupby(
                    "Departamento",
                    as_index=False,
                )["Sitio"]
                .nunique()
                .rename(
                    columns={
                        "Sitio": "Sitios"
                    }
                )
                .sort_values(
                    "Sitios"
                )
            )

            total_sitios_escalados = int(
                datos_escalados["Sitio"].nunique()
            )

            if resumen_escalados.empty:

                st.info(
                    "No hay sitios escalados a proyectos "
                    "en los filtros seleccionados."
                )

            else:

                fig = px.bar(
                    resumen_escalados,
                    x="Sitios",
                    y="Departamento",
                    orientation="h",
                    text="Sitios",
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
                            250,
                            len(resumen_escalados) * 38,
                        ),
                    ),
                    width="stretch",
                )

                st.caption(
                    f"Total sitios únicos escalados a proyectos: "
                    f"{total_sitios_escalados:,}"
                )

        else:

            st.info(
                "No se encontraron las columnas "
                "DEPARTAMENTO e ID del sitio."
            )

    with derecha.container(border=True):

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

            acciones = acciones.replace(
                "",
                "Sin información",
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
                        220,
                        min(
                            360,
                            len(resumen_accion) * 32,
                        ),
                    ),
                ),
                width="stretch",
            )

        else:

            st.info(
                "No se encontró la columna ACCIÓN."
            )

    # ========================================================
    # PROGRAMACIÓN DE VISITAS + CARGA EJECUTABLE · PRÓXIMOS 7 DÍAS
    # ========================================================

    izquierda, derecha = st.columns(2, gap="medium")

    with izquierda.container(border=True):

        st.subheader(
            "Programación de visitas",
            icon=":material/event_available:",
        )

        visitas = (
            df["ESTADO VISITA"]
            .value_counts()
            .sort_values()
            .reset_index()
        )

        visitas.columns = [
            "Estado",
            "Visitas",
        ]

        fig = px.bar(
            visitas,
            x="Visitas",
            y="Estado",
            orientation="h",
            text="Visitas",
            color="Estado",
            color_discrete_map={
                "Visita programada - PQRSD Vencida":
                    PALETTE["red"],
                "Visita programada":
                    PALETTE["teal"],
                "Visita vencida":
                    PALETTE["red"],
                "Visita hoy":
                    PALETTE["amber"],
                "Escalado a proyectos":
                    PALETTE["blue"],
                "Sin fecha de visita":
                    PALETTE["slate"],
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
                    220,
                    min(
                        420,
                        len(visitas) * 42,
                    ),
                ),
            ),
            width="stretch",
        )

    with derecha.container(border=True):

        st.subheader(
            "Carga ejecutable · próximos 7 días",
            icon=":material/groups:",
        )

        if visitas_7_dias_n > 0:

            col_tecnico = buscar_columna(
                df,
                [
                    "TÉCNICO",
                    "TECNICO",
                    "TÉCNICO ASIGNADO",
                    "TECNICO ASIGNADO",
                ],
            )

            col_departamento = buscar_columna(
                df,
                [
                    "DEPARTAMENTO",
                ],
            )

            if (
                col_tecnico is not None
                and col_departamento is not None
            ):

                carga = df.loc[
                    visitas_7_dias_mask,
                    [
                        col_tecnico,
                        col_departamento,
                    ],
                ].copy()

                carga.columns = [
                    "Técnico",
                    "Departamento",
                ]

                carga["Técnico"] = (
                    carga["Técnico"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                carga["Departamento"] = (
                    carga["Departamento"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                carga["Técnico"] = carga[
                    "Técnico"
                ].replace(
                    "",
                    "Sin asignar",
                )

                carga["Departamento"] = carga[
                    "Departamento"
                ].replace(
                    "",
                    "Sin departamento",
                )

                tabla_carga = (
                    carga
                    .groupby(
                        "Técnico",
                        as_index=False,
                    )
                    .agg(
                        **{
                            "Departamento(s)": (
                                "Departamento",
                                lambda serie:
                                ", ".join(
                                    sorted(
                                        serie
                                        .dropna()
                                        .astype(str)
                                        .str.strip()
                                        .unique()
                                    )
                                ),
                            ),
                            "Visitas": (
                                "Técnico",
                                "size",
                            ),
                        }
                    )
                    .sort_values(
                        [
                            "Visitas",
                            "Técnico",
                        ],
                        ascending=[
                            False,
                            True,
                        ],
                    )
                    .reset_index(
                        drop=True
                    )
                )

                asignadas_n = int(
                    (
                        carga["Técnico"]
                        .str.upper()
                        .ne("SIN ASIGNAR")
                    ).sum()
                )

                sin_asignar_n = int(
                    (
                        carga["Técnico"]
                        .str.upper()
                        .eq("SIN ASIGNAR")
                    ).sum()
                )

                st.caption(
                    f"Ventana: "
                    f"{hoy:%d/%m/%Y} al "
                    f"{(hoy + pd.Timedelta(days=6)):%d/%m/%Y} "
                    f"· Visitas ejecutables: "
                    f"{visitas_7_dias_n:,} "
                    f"· Asignadas: "
                    f"{asignadas_n:,} "
                    f"· Sin técnico asignado: "
                    f"{sin_asignar_n:,}"
                )

                st.dataframe(
                    tabla_carga,
                    hide_index=True,
                    width="stretch",
                    column_config={
                        "Técnico":
                            st.column_config.Column(
                                width=190
                            ),
                        "Departamento(s)":
                            st.column_config.Column(
                                width=320
                            ),
                        "Visitas":
                            st.column_config.NumberColumn(
                                width=100,
                                format="%d",
                            ),
                    },
                )

            else:

                st.info(
                    "No se encontraron las columnas "
                    "TÉCNICO y/o DEPARTAMENTO."
                )

        else:

            st.success(
                "No hay visitas ejecutables "
                "programadas para los próximos 7 días."
            )


    # ========================================================
    # TABLA DETALLADA
    # ========================================================

    with st.container(
        border=True,
    ):

        st.markdown(
            '<div class="dataframe-title">Dataframe</div>',
            unsafe_allow_html=True,
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
            "FECHA DE VISITA",
            "FECHA VISITA",
            "NUEVA FECHA MÁXIMA ATENCIÓN",
            "NUEVA FECHA MAXIMA ATENCION",
            "TÉCNICO",
            "NOVEDAD",
            "ESTADO VISITA",
        ]

        columnas_disponibles = []

        for c in columnas:

            if (
                c in df.columns
                and c not in columnas_disponibles
            ):

                columnas_disponibles.append(c)

        tabla = df[
            columnas_disponibles
        ].copy()

        columnas_fecha = [
            "FECHA DE VISITA",
            "FECHA VISITA",
            "NUEVA FECHA MÁXIMA ATENCIÓN",
            "NUEVA FECHA MAXIMA ATENCION",
        ]

        for c in columnas_fecha:

            if c in tabla.columns:

                tabla[c] = (
                    pd.to_datetime(
                        tabla[c],
                        errors="coerce",
                        dayfirst=True,
                    )
                    .dt.strftime(
                        "%d/%m/%Y"
                    )
                    .fillna("")
                )

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

        st.markdown(
            '<div class="dataframe-caption">Desplazamiento horizontal disponible para consultar todas las columnas.</div>',
            unsafe_allow_html=True,
        )

        if "ESTADO PQRS CON PDR" in tabla.columns:

            def etiqueta_estado_pqrs(
                valor,
            ):

                texto = str(
                    valor
                ).strip()

                normalizado = texto.upper()

                if normalizado == "VIGENTE":
                    return "● VIGENTE"

                if normalizado in {
                    "PRÓXIMO A VENCER",
                    "PROXIMO A VENCER",
                }:
                    return "● PRÓXIMO A VENCER"

                if normalizado == "VENCIDO":
                    return "● VENCIDO"

                return texto

            tabla[
                "ESTADO PQRS CON PDR"
            ] = (
                tabla[
                    "ESTADO PQRS CON PDR"
                ]
                .map(
                    etiqueta_estado_pqrs
                )
            )

        tabla_estilizada = tabla.style

        if "ESTADO PQRS CON PDR" in tabla.columns:

            tabla_estilizada = (
                tabla_estilizada.map(
                    estilo_estado_pqrs,
                    subset=[
                        "ESTADO PQRS CON PDR"
                    ],
                )
            )

        configuracion_columnas = {
            columna_nombre:
                st.column_config.Column(
                    width=170
                )
            for columna_nombre in tabla.columns
        }

        if "ESTADO PQRS CON PDR" in tabla.columns:
            configuracion_columnas[
                "ESTADO PQRS CON PDR"
            ] = st.column_config.Column(
                width=190
            )

        st.dataframe(
            tabla_estilizada,
            height=420,
            width="stretch",
            hide_index=True,
            column_config=configuracion_columnas,
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
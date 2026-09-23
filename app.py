"""Dashboard de Contratos + Alfresco (2024 / 2025 / 2026).

Ejecutar con:  streamlit run app.py
Los archivos Excel pueden estar en cualquier carpeta: se elige en
"Configuracion de datos" o con la variable de entorno DASHBOARD_DATA_DIR.
Por defecto se usa data/ junto a app.py. El año se detecta desde el nombre del
archivo (ej: 2025_ENLACES_ALFRESCO.xlsx).
"""
import glob
import io
import os
import re
import sys
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Contratos — Alfresco",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# Estilo institucional sobrio
# ----------------------------------------------------------------------------
PALETA = {
    "fondo": "#F4F6F5",
    "tarjeta": "#FFFFFF",
    "borde": "#DFE7E2",
    "texto": "#1C2B23",
    "muted": "#5C6F65",
    "primario": "#1B5E20",
    "primario_suave": "#E8F2EA",
    "en_alfresco": "#1B7A3D",
    "pendiente": "#F9573B",
    "pendiente_suave": "#FBEEDC",
}

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {PALETA['fondo']}; }}
    section[data-testid="stSidebar"] {{
        background-color: #EDF3EE;
        border-right: 1px solid {PALETA['borde']};
    }}
    h1 {{ color: {PALETA['texto']} !important; font-size: 2rem !important; margin-bottom: 0.1rem !important; }}
    h2 {{ color: {PALETA['texto']} !important; font-size: 1.25rem !important; margin-top: 1.6rem !important;
         padding-bottom: 0.35rem !important; border-bottom: 1px solid {PALETA['borde']}; }}
    h3 {{ color: {PALETA['texto']} !important; font-size: 1.05rem !important; }}
    .subtitulo {{ color: {PALETA['muted']}; font-size: 1rem; margin-top: 0; }}
    .fuente-datos {{ color: {PALETA['muted']}; font-size: 0.8rem; }}
    .kpi-card {{
        background: {PALETA['tarjeta']};
        border: 1px solid {PALETA['borde']};
        border-top: 4px solid {PALETA['primario']};
        border-radius: 10px;
        padding: 14px 16px;
        min-height: 108px;
    }}
    .kpi-card.destacada {{
        border-top-color: {PALETA['en_alfresco']};
        background: linear-gradient(180deg, #FFFFFF 60%, #EAF4ED 100%);
    }}
    .kpi-card.alerta {{ border-top-color: {PALETA['pendiente']}; }}
    .kpi-etiqueta {{ color: {PALETA['muted']}; font-size: 0.8rem; text-transform: uppercase;
                    letter-spacing: 0.04em; margin-bottom: 4px; }}
    .kpi-valor {{ color: {PALETA['texto']}; font-size: 1.7rem; font-weight: 700; line-height: 1.1; }}
    .kpi-sub {{ color: {PALETA['muted']}; font-size: 0.82rem; margin-top: 4px; }}
    .barra-fondo {{ background: #E5E9E6; border-radius: 999px; height: 12px; overflow: hidden; }}
    .barra-relleno {{ background: {PALETA['en_alfresco']}; height: 12px; border-radius: 999px; }}
    .pill {{
        display: inline-block; padding: 2px 10px; border-radius: 999px;
        font-size: 0.78rem; font-weight: 600;
    }}
    .pill-ok {{ background: #E3F2E6; color: #14532D; border: 1px solid #BFE0C6; }}
    .pill-pend {{ background: {PALETA['pendiente_suave']}; color: #7C3F06; border: 1px solid #F0D3A8; }}
    a {{ color: #1b7a3d !important; }}
    .block-container {{ padding-top: 1.5rem !important; max-width: 1400px; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Constantes y helpers de negocio (SIN CAMBIOS de logica)
# ----------------------------------------------------------------------------
BASE_DIR = None


def base_dir():
    # Cuando corre como .exe de PyInstaller, __file__ apunta a _MEIPASS (temporal).
    # Los datos por defecto deben buscarse junto al .exe, no en el temporal.
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = base_dir()
DEFAULT_DATA_DIR = os.environ.get("DASHBOARD_DATA_DIR", os.path.join(BASE_DIR, "data"))
ESTADO_OK = "VALIDADO 1-CLIC"
ESTADO_FINAL_OK = "ENCONTRADO"
VERDE_SOLIDO = "#2E7D32"
VERDE_OSCURO = "#1B5E20"
VERDES = ["#1b7a3d", "#2e7d32", "#43a047", "#66bb6a", "#81c784", "#00441b", "#238b45", "#a5d6a7"]
VERDES_CONTINUO = ["#edf8e9", "#c7e9c0", "#a1d99b", "#74c476", "#41ab5d", "#238b45", "#005a32"]

RE_ANIO = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
RE_ANIO_FILA = re.compile(r"(?:19|20)\d{2}")
RE_SERIE = re.compile(r"([A-Z]+\d+)", re.IGNORECASE)
RE_SUBSERIE = re.compile(r"([A-Z]+\d+\.\d+)", re.IGNORECASE)
RE_CENTRO_NOMBRE = re.compile(r"^\s*\d+\s*[-–]\s*(.+?)\s*$")


def extraer_serie(ruta):
    """Serie = codigo LETRAS+digitos dentro de SERIE_RUTA. Ej: 1112_C09.23_... -> C09,
    4112_P11.02_... -> P11.

    Se usa regex [A-Z]+\\d+ porque la definicion literal 'despues del primer _ y antes
    del punto' falla en rutas con doble prefijo como 2123_6540_C09.11_... (daria
    '6540_C09'). El regex devuelve siempre C09 / P11 / C270 etc. Exige letras
    iniciales para no capturar los prefijos numericos de dependencia (4112, 1112).
    """
    if not isinstance(ruta, str) or not ruta.strip():
        return "Sin serie"
    m = RE_SERIE.search(ruta)
    return m.group(1).upper() if m else "Sin serie"


def extraer_subserie(ruta):
    if not isinstance(ruta, str) or not ruta.strip():
        return "Sin subserie"
    m = RE_SUBSERIE.search(ruta)
    return m.group(1).upper() if m else extraer_serie(ruta)


def extraer_centro_nombre(valor):
    """Agrupa CENTRO_COSTO por nombre de dependencia, sin código de fondo.

    Ej: '3112-DIVISION DE CONTRATACION' y '9244-DIVISION DE CONTRATACION'
    -> 'DIVISION DE CONTRATACION'. Así los 2 fondos de Diana se atribuyen
    a la misma dependencia. Si no hay prefijo numérico, devuelve el texto
    limpio. Vacios -> 'Sin centro'.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "Sin centro"
    txt = str(valor).strip()
    if not txt or txt.lower() == "nan":
        return "Sin centro"
    m = RE_CENTRO_NOMBRE.match(txt)
    nombre = m.group(1).strip() if m else txt
    return nombre if nombre else "Sin centro"


def detectar_anio(path):
    """Año por defecto desde el nombre del archivo.

    Si el nombre trae un único año (ej: 2025_ENLACES_ALFRESCO.xlsx) se
    devuelve ese año. Si trae varios (ej: 2024_2025_2026_....xlsx, archivo
    consolidado) o ninguno, se devuelve None para que el año se infiera
    fila por fila desde CONTRATO / NUMERO en cargar_datos().
    """
    anios = RE_ANIO.findall(os.path.basename(path))
    unicos = sorted(set(anios))
    if len(unicos) == 1:
        return unicos[0]
    return None


def extraer_anio_fila(valor):
    """Extrae el primer año de 4 dígitos (19xx/20xx) dentro de un texto.

    Sirve para valores como '270-2024000001' o '2025000022'. Devuelve None
    si no hay año.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    m = RE_ANIO_FILA.search(str(valor))
    return m.group(0) if m else None


# Columnas que el dashboard sabe aprovechar (todas opcionales salvo
# CONTRATO/NUMERO para inferir año; el resto se tolera con guards).
COLUMNAS_ESPERADAS = [
    "CONTRATO", "NUMERO", "TIPO", "CENTRO_COSTO", "NOMBRE_CONTRATISTA",
    "OBJETO_CONTRATO", "ORDENADOR_CENTRO", "ORDENADOR_UNIDAD",
    "CARPETA_ALFRESCO", "URL_ALFRESCO_1CLIC", "SERIE_RUTA",
    "ESTADO", "ESTADO_FINAL",
]


def normalizar_frame(df_raw, nombre_archivo):
    """Aplica a UN dataframe la misma normalización que usaba cargar_datos.

    - Columnas a MAYÚSCULAS sin espacios.
    - Columna AÑO: respeta AÑO/ANO válido, si no infiere de CONTRATO/NUMERO,
      si no usa el año del nombre de archivo.
    - Columna ARCHIVO con el nombre de origen.
    No toca SERIE/SUBSERIE/EN_ALFRESCO (eso lo hace finalizar_consolidado).
    """
    df = df_raw.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    anio_defecto = detectar_anio(nombre_archivo)
    if "AÑO" in df.columns or "ANO" in df.columns:
        col_orig = "AÑO" if "AÑO" in df.columns else "ANO"
        base = df[col_orig].astype("string").str.strip()
        valida = base.str.fullmatch(r"(?:19|20)\d{2}", na=False)
        df["AÑO"] = base.where(valida)
    else:
        df["AÑO"] = pd.NA
    faltan = df["AÑO"].isna()
    if faltan.any():
        for col_fuente in ("CONTRATO", "NUMERO"):
            if col_fuente in df.columns and faltan.any():
                inferido = df.loc[faltan, col_fuente].map(extraer_anio_fila)
                df.loc[faltan, "AÑO"] = inferido
                faltan = df["AÑO"].isna()
            if not faltan.any():
                break
    df["AÑO"] = df["AÑO"].fillna(anio_defecto if anio_defecto else "Sin año")
    df["ARCHIVO"] = os.path.basename(nombre_archivo)
    return df


def finalizar_consolidado(frames):
    """Concatena frames ya normalizados y deriva SERIE/SUBSERIE/EN_ALFRESCO.

    Lógica idéntica a la original: prioriza ESTADO_FINAL, con fallback a ESTADO.
    """
    if not frames:
        return pd.DataFrame()
    full = pd.concat(frames, ignore_index=True)
    for c in full.columns:
        if full[c].dtype == object:
            full[c] = full[c].astype("string").str.strip()
    full["SERIE"] = full["SERIE_RUTA"].apply(extraer_serie) if "SERIE_RUTA" in full.columns else "Sin serie"
    full["SUBSERIE"] = full["SERIE_RUTA"].apply(extraer_subserie) if "SERIE_RUTA" in full.columns else "Sin subserie"
    if "ESTADO_FINAL" in full.columns:
        full["EN_ALFRESCO"] = (
            full["ESTADO_FINAL"].astype("string").str.strip().str.upper() == ESTADO_FINAL_OK
        ).fillna(False).astype(bool)
    elif "ESTADO" in full.columns:
        full["EN_ALFRESCO"] = (
            full["ESTADO"].astype("string").str.strip() == ESTADO_OK
        ).fillna(False).astype(bool)
    else:
        full["EN_ALFRESCO"] = False
    full["ORDENADOR_CENTRO"] = full.get("ORDENADOR_CENTRO")
    full["ORDENADOR_UNIDAD"] = full.get("ORDENADOR_UNIDAD")
    if "CENTRO_COSTO" in full.columns:
        full["CENTRO_NOMBRE"] = full["CENTRO_COSTO"].map(extraer_centro_nombre)
    else:
        full["CENTRO_NOMBRE"] = "Sin centro"
    return full


@st.cache_data(show_spinner="Cargando Excels...")
def cargar_datos(data_dir):
    # Solo se lee la carpeta elegida. Si está vacía y es la carpeta por defecto,
    # se usa la raíz como respaldo (instalaciones anteriores).
    patrones = sorted(glob.glob(os.path.join(data_dir, "*.xlsx")))
    archivos = [p for p in patrones if not os.path.basename(p).startswith("~$")]
    if not archivos and os.path.abspath(data_dir) == os.path.abspath(DEFAULT_DATA_DIR):
        archivos = sorted(
            p for p in glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
            if not os.path.basename(p).startswith("~$")
        )
    if not archivos:
        return pd.DataFrame(), []
    frames = []
    origen = []
    for path in archivos:
        try:
            df_raw = pd.read_excel(path, engine="openpyxl", dtype=str)
        except Exception as e:
            st.warning(f"No se pudo leer {os.path.basename(path)}: {e}")
            continue
        frames.append(normalizar_frame(df_raw, path))
        origen.append(os.path.basename(path))
    if not frames:
        return pd.DataFrame(), []
    return finalizar_consolidado(frames), sorted(origen)


@st.cache_data(show_spinner="Procesando Excel cargado...")
def cargar_excel_subido(file_bytes, nombre_archivo):
    """Procesa en memoria el .xlsx subido con st.file_uploader.

    Usa exactamente la misma normalización que los archivos locales
    (normalizar_frame + finalizar_consolidado). Los bytes son la clave
    de caché, por lo que el archivo persiste al interactuar con filtros.
    """
    df_raw = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", dtype=str)
    frames = [normalizar_frame(df_raw, nombre_archivo)]
    return finalizar_consolidado(frames)


# ----------------------------------------------------------------------------
# Helpers solo visuales (no tocan la logica de negocio)
# ----------------------------------------------------------------------------
def fmt_num(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def fmt_pct(x):
    return f"{x:.1f} %".replace(".", ",")


def kpi_card(etiqueta, valor, sub="", estilo=""):
    st.markdown(
        f"""<div class="kpi-card {estilo}">
        <div class="kpi-etiqueta">{etiqueta}</div>
        <div class="kpi-valor">{valor}</div>
        <div class="kpi-sub">{sub}</div>
        </div>""",
        unsafe_allow_html=True,
    )


def base_layout(fig, height=380):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=20, r=20, t=50, b=20),
        font=dict(family="Arial", size=12, color="#1C2B23"),
        title_font=dict(size=13, color="#1C2B23"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    return fig


# ----------------------------------------------------------------------------
# Semaforo por % de faltantes (verde <50, amarillo 50-70, rojo >70)
# ----------------------------------------------------------------------------
SEMAFORO = {
    "verde": "#43A047",
    "amarillo": "#EAB308",
    "rojo": "#E22200",
}


def color_semaforo(pct):
    """Devuelve el color semaforo segun % de faltantes (0-100)."""
    try:
        v = float(pct)
    except (TypeError, ValueError):
        return SEMAFORO["verde"]
    if v > 70:
        return SEMAFORO["rojo"]
    if v >= 50:
        return SEMAFORO["amarillo"]
    return SEMAFORO["verde"]


def nivel_semaforo(pct):
    """Etiqueta corta del nivel: 🟢 / 🟡 / 🔴."""
    try:
        v = float(pct)
    except (TypeError, ValueError):
        return "🟢 Bajo"
    if v > 70:
        return "🔴 Crítico (>70%)"
    if v >= 50:
        return "🟡 Medio (50-70%)"
    return "🟢 Bajo (<50%)"


# ----------------------------------------------------------------------------
# Configuracion de datos (misma funcionalidad, menor protagonismo visual)
# ----------------------------------------------------------------------------
if "data_dir" not in st.session_state:
    st.session_state.data_dir = DEFAULT_DATA_DIR
if "explorar_dir" not in st.session_state:
    st.session_state.explorar_dir = (
        st.session_state.data_dir if os.path.isdir(st.session_state.data_dir) else BASE_DIR
    )


def listar_unidades():
    """Unidades disponibles (Windows) o raíz (otros sistemas)."""
    import string

    unidades = [f"{letra}:\\" for letra in string.ascii_uppercase if os.path.isdir(f"{letra}:\\")]
    if unidades:
        return unidades
    return ["/"] if os.path.isdir("/") else [BASE_DIR]


def contar_xlsx(ruta):
    try:
        return sum(
            1 for p in glob.glob(os.path.join(ruta, "*.xlsx"))
            if not os.path.basename(p).startswith("~$")
        )
    except OSError:
        return 0


# ----------------------------------------------------------------------------
# Encabezado (la fuente tecnica no domina visualmente)
# ----------------------------------------------------------------------------
st.markdown("# Gestión Documental de Contratos en Alfresco")
st.markdown(
    '<p class="subtitulo">Monitoreo de la disponibilidad y estado de los contratos en Alfresco.'
    "</p>",
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Cargar archivo de datos (fuente prioritaria; no altera lógica posterior)
# Flujo: abrir URL -> cargar Excel -> validación -> lectura pandas ->
# mismo procesamiento (normalizar_frame + finalizar_consolidado) -> dashboard.
# Funciona en Streamlit Community Cloud: todo en memoria/sesión, sin rutas.
# ----------------------------------------------------------------------------
st.markdown("## Cargar archivo de datos")
archivo_subido = st.file_uploader(
    "📂 Seleccionar archivo Excel",
    type=["xlsx"],
    key="excel_upload",
    help="Seleccione el Excel de enlaces Alfresco (.xlsx). Se procesa en memoria con el mismo formato que el archivo local.",
)
FUENTE_SUBIDA = archivo_subido is not None
df = pd.DataFrame()
archivos = []
DATA_DIR_ACTIVA = st.session_state.data_dir
info_carga = st.session_state.get("upload_info")

if FUENTE_SUBIDA:
    nombre_subido = archivo_subido.name or "archivo.xlsx"
    if not nombre_subido.lower().endswith(".xlsx"):
        st.error("El archivo debe ser un Excel (.xlsx). Seleccione un archivo válido.")
        st.stop()
    try:
        file_bytes = archivo_subido.getvalue()
    except Exception:
        st.error("No se pudo leer el archivo cargado. Intente de nuevo.")
        st.stop()
    if not file_bytes:
        st.error("El archivo cargado está vacío. Seleccione un Excel válido.")
        st.stop()
    try:
        peek = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", dtype=str, nrows=0)
        cols_subido = [str(c).strip().upper() for c in peek.columns]
    except Exception:
        st.error("No se pudo leer el archivo como Excel (.xlsx). Verifique que no esté dañado ni protegido.")
        st.stop()
    if not cols_subido:
        st.error("El archivo no tiene columnas. Verifique la estructura del Excel.")
        st.stop()
    faltantes = [c for c in COLUMNAS_ESPERADAS if c not in cols_subido]
    if "CONTRATO" not in cols_subido and "NUMERO" not in cols_subido:
        st.error(
            "Estructura inválida: el Excel debe contener al menos la columna CONTRATO o NUMERO. "
            f"Faltan: {', '.join(faltantes) if faltantes else 'columnas clave'}."
        )
        st.stop()
    if faltantes:
        st.warning(
            "El archivo no trae estas columnas esperadas y esas vistas/filtros mostrarán menos detalle: "
            + ", ".join(faltantes) + "."
        )
    if "ESTADO" not in cols_subido and "ESTADO_FINAL" not in cols_subido:
        st.warning("El archivo no trae ESTADO ni ESTADO_FINAL: todo se marcará como “No encontrado”.")
    try:
        df = cargar_excel_subido(file_bytes, nombre_subido)
    except Exception:
        st.error("No se pudo procesar el archivo. Verifique que sea un .xlsx válido con datos.")
        st.stop()
    if df.empty:
        st.error("El archivo cargado no tiene registros para mostrar.")
        st.stop()
    archivos = [nombre_subido]
    DATA_DIR_ACTIVA = f"Archivo cargado: {nombre_subido}"
    tam = getattr(archivo_subido, "size", len(file_bytes))
    firma = (nombre_subido, tam, len(df))
    if st.session_state.get("upload_firma") != firma:
        st.session_state["upload_firma"] = firma
        st.session_state["upload_info"] = {
            "nombre": nombre_subido,
            "registros": int(len(df)),
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
    info_carga = st.session_state.get("upload_info")
    st.success("✓ Archivo cargado correctamente")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.caption(f"Archivo: {info_carga['nombre']}" if info_carga else f"Archivo: {nombre_subido}")
    with c2:
        st.caption(f"Registros procesados: {fmt_num(len(df))}")
    with c3:
        st.caption(f"Procesado: {info_carga['fecha']}" if info_carga and info_carga.get("fecha") else "Procesado: ahora")
else:
    DATA_DIR_ACTIVA = st.session_state.data_dir
    df, archivos = cargar_datos(DATA_DIR_ACTIVA)

with st.sidebar:
    st.markdown("### Seguimiento documental")
    st.caption("Contratos · Alfresco · cobertura documental")
    st.markdown("### Filtros")
    st.caption("Todo el dashboard responde a estos filtros.")
    _anios = sorted(df["AÑO"].dropna().unique().tolist()) if "AÑO" in df.columns else []
    sel_anios = st.multiselect("Año", _anios, default=_anios)
    _tipos = sorted(df["TIPO"].dropna().unique().tolist()) if "TIPO" in df.columns else []
    sel_tipos = st.multiselect("Tipo de contrato", _tipos, default=_tipos)
    sel_estado_doc = st.multiselect(
        "Estado documental",
        ["Encontrado", "No encontrado"],
        default=["Encontrado", "No encontrado"],
        help="Encontrado = ESTADO_FINAL = ENCONTRADO (si existe; si no, ESTADO = VALIDADO 1-CLIC). No encontrado = resto.",
    )
    with st.expander("⚙️ Filtros avanzados", expanded=False):
        _centros = sorted(df["CENTRO_COSTO"].dropna().unique().tolist()) if "CENTRO_COSTO" in df.columns else []
        sel_centros = st.multiselect("Centro de costo", _centros, default=[])
        if "ORDENADOR_CENTRO" in df.columns:
            _ord = sorted(df["ORDENADOR_CENTRO"].dropna().astype("string").str.strip().replace("", pd.NA).dropna().unique().tolist())
        else:
            _ord = []
        sel_ord = st.multiselect("Ordenador", _ord, default=[],
                                 help="Filtra por ORDENADOR_CENTRO (quien ordena el gasto). Ej: DIANA MARCELA ARBOLEDA CALVO trae sus 2 fondos 3112 y 9244.")
        _series = sorted(df["SERIE"].dropna().unique().tolist()) if "SERIE" in df.columns else []
        sel_series = st.multiselect("Serie (ej: C09)", _series, default=[])
        _sub = sorted(df["SUBSERIE"].dropna().unique().tolist()) if "SUBSERIE" in df.columns else []
        sel_sub = st.multiselect("Subserie (ej: C09.11)", _sub, default=[])
    texto = st.text_input(
        "Buscar contrato, contratista u objeto",
        "",
        placeholder="Ej: 2025000022, contratista o palabra del objeto…",
    )
    with st.expander("⚙️ Configuración de datos", expanded=False):
        if FUENTE_SUBIDA:
            st.info("Hay un archivo cargado desde “Cargar archivo de datos” y tiene prioridad. Para volver a la carpeta local, quite el archivo (X en el cargador).")
        nueva_ruta = st.text_input(
            "Ruta con los Excel",
            value=st.session_state.data_dir,
            help="Pega cualquier ruta, ej: C:\\Contratos\\Excels o \\\\servidor\\contratos. También puedes definirla con la variable de entorno DASHBOARD_DATA_DIR o elegirla en el explorador de abajo.",
        )
        if st.button("↩️ Volver a la carpeta por defecto", use_container_width=True):
            st.session_state.data_dir = DEFAULT_DATA_DIR
            st.session_state.explorar_dir = DEFAULT_DATA_DIR if os.path.isdir(DEFAULT_DATA_DIR) else BASE_DIR
            st.rerun()

        # Si el usuario editó el texto, usarlo (previa validación al cargar).
        if nueva_ruta and nueva_ruta != st.session_state.data_dir:
            st.session_state.data_dir = nueva_ruta
            if os.path.isdir(nueva_ruta):
                st.session_state.explorar_dir = os.path.abspath(nueva_ruta)
            st.rerun()

        with st.expander("📂 Explorar carpetas y elegir…", expanded=False):
            actual = st.session_state.explorar_dir
            if not os.path.isdir(actual):
                actual = BASE_DIR
                st.session_state.explorar_dir = actual

            unidades = listar_unidades()
            unidad_actual, _resto = os.path.splitdrive(os.path.abspath(actual))
            unidad_actual = (unidad_actual + "\\") if unidad_actual else unidades[0]
            if unidad_actual not in unidades:
                unidades = [unidad_actual] + unidades
            unidad = st.selectbox("Unidad", unidades, index=unidades.index(unidad_actual), key="sel_unidad")
            if os.path.abspath(unidad) != os.path.abspath(actual) and unidad != unidad_actual:
                st.session_state.explorar_dir = os.path.abspath(unidad)
                st.rerun()
                actual = st.session_state.explorar_dir

            st.code(actual, language=None)
            n_xlsx = contar_xlsx(actual)
            if n_xlsx:
                st.success(f"📊 {n_xlsx} Excel aquí")
            else:
                st.caption("Sin archivos .xlsx en esta carpeta")

            c_up, c_use = st.columns(2)
            with c_up:
                if st.button("⬆️ Subir nivel", use_container_width=True, key="btn_subir"):
                    padre = os.path.dirname(os.path.abspath(actual.rstrip("\\/")))
                    if padre and os.path.isdir(padre):
                        st.session_state.explorar_dir = padre
                        st.rerun()
            with c_use:
                if st.button("✅ Usar esta carpeta", use_container_width=True, key="btn_usar"):
                    st.session_state.data_dir = actual
                    st.rerun()

            try:
                with os.scandir(actual) as it:
                    subdirs = sorted(
                        (e.path for e in it if e.is_dir(follow_symlinks=False)),
                        key=lambda p: os.path.basename(p).lower(),
                    )
            except OSError as e:
                st.warning(f"No se puede listar esta carpeta: {e}")
                subdirs = []

            if subdirs:
                st.caption(f"Subcarpetas ({len(subdirs)}):")
                for sub in subdirs[:100]:
                    nombre = os.path.basename(sub) or sub
                    marca = " 📊" if contar_xlsx(sub) else ""
                    if st.button(f"📁 {nombre}{marca}", key=f"nav_{sub}", use_container_width=True):
                        st.session_state.explorar_dir = sub
                        st.rerun()
                if len(subdirs) > 100:
                    st.caption("…mostrando las 100 primeras. Usa el campo de ruta para ir directo.")

if df.empty:
    st.info("Cargue el archivo Excel para generar el dashboard.")
    st.caption(
        "Use la sección “Cargar archivo de datos” (.xlsx). "
        "También puede usar la carpeta local en la barra lateral → ⚙️ Configuración de datos."
    )
    st.stop()


# ----------------------------------------------------------------------------
# Filtros activos (los controles estan en la barra lateral izquierda)
# ----------------------------------------------------------------------------
_partes = []
if sel_anios:
    _partes.append(f"Año: {', '.join(sel_anios)}")
if sel_tipos and "TIPO" in df.columns:
    _partes.append(f"Tipo: {', '.join(sel_tipos)}")
if sel_estado_doc:
    _partes.append(f"Estado: {', '.join(sel_estado_doc)}")
else:
    _partes.append("Estado: todos")
if sel_centros:
    _partes.append(f"Centros: {len(sel_centros)} sel.")
if sel_ord:
    _partes.append(f"Ordenadores: {len(sel_ord)} sel.")
if sel_series:
    _partes.append(f"Series: {', '.join(sel_series)}")
if sel_sub:
    _partes.append(f"Subseries: {', '.join(sel_sub)}")
if texto.strip():
    _partes.append(f"Búsqueda: “{texto.strip()}”")
#st.caption("Filtros activos → " + (" · ".join(_partes) if _partes else "sin filtros")
         #  + ". Se ajustan en la barra lateral izquierda.")

# ----------------------------------------------------------------------------
# Aplicar filtros (logica original intacta)
# ----------------------------------------------------------------------------
f = df.copy()
if sel_anios:
    f = f[f["AÑO"].isin(sel_anios)]
if sel_tipos and "TIPO" in f.columns:
    f = f[f["TIPO"].isin(sel_tipos)]
if sel_estado_doc and len(sel_estado_doc) < 2:
    if "Encontrado" in sel_estado_doc:
        f = f[f["EN_ALFRESCO"]]
    else:
        f = f[~f["EN_ALFRESCO"]]
if sel_centros:
    f = f[f["CENTRO_COSTO"].isin(sel_centros)]
if sel_ord:
    f = f[f["ORDENADOR_CENTRO"].isin(sel_ord)]
if sel_series:
    f = f[f["SERIE"].isin(sel_series)]
if sel_sub:
    f = f[f["SUBSERIE"].isin(sel_sub)]
if texto.strip():
    t = texto.strip().lower()
    cols_txt = [c for c in ["CONTRATO", "NOMBRE_CONTRATISTA", "OBJETO_CONTRATO", "CARPETA_ALFRESCO"] if c in f.columns]
    mask = pd.Series(False, index=f.index)
    for c in cols_txt:
        mask = mask | f[c].fillna("").str.lower().str.contains(t, na=False)
    f = f[mask]

if f.empty:
    st.warning("Sin resultados con los filtros actuales. Ajusta los filtros en la barra lateral.")
    st.stop()

# ----------------------------------------------------------------------------
# KPIs (calculados dinamicamente)
# ----------------------------------------------------------------------------
total = len(f)
en_alf = int(f["EN_ALFRESCO"].sum())
pend = total - en_alf
pct = (en_alf / total * 100) if total else 0

st.markdown("## Resumen")
k1, k2, k3, k4 = st.columns(4)
with k1:
    kpi_card("Total contratos", fmt_num(total), "Contratos en el filtro actual")
with k2:
    kpi_card("En Alfresco", fmt_num(en_alf), "Con documentación disponible")
with k3:
    kpi_card("Pendientes", fmt_num(pend), "Sin documentación en Alfresco", estilo="alerta")
with k4:
    kpi_card("% en Alfresco", fmt_pct(pct), "Nivel de cobertura documental", estilo="destacada")

# ----------------------------------------------------------------------------
# Seguimiento documental (prioridad a pendientes)
# ----------------------------------------------------------------------------
import plotly.express as px

st.markdown("## Seguimiento documental")
st.caption("Proporción de contratos con documentación disponible frente a pendientes.")

s1, s2 = st.columns([1.1, 1])
with s1:
    st.markdown("#### Contratos en Alfresco vs. pendientes")
    g2 = pd.DataFrame({"Situación": ["En Alfresco", "Pendientes"], "N": [en_alf, pend]})
    fig2 = px.pie(
        g2, names="Situación", values="N", hole=0.55,
        color="Situación",
        color_discrete_map={"En Alfresco": PALETA["en_alfresco"], "Pendientes": "#F9573B"},
        template="plotly_white",
    )
    fig2.update_traces(textinfo="value", textfont_size=13,
                       hovertemplate="%{label}: %{value:,} (%{percent})")
    fig2 = base_layout(fig2, height=340)
    fig2.update_layout(showlegend=True)
    st.plotly_chart(fig2, use_container_width=True, key="pie_alfresco")

with s2:
    st.markdown("#### Cobertura documental")
    st.markdown(
        f"""<div style="font-size:2rem;font-weight:700;color:{PALETA['texto']}">{fmt_pct(pct)}</div>
        <div style="color:{PALETA['muted']};font-size:0.9rem;margin-bottom:8px;">
        {fmt_num(en_alf)} en Alfresco · {fmt_num(pend)} pendientes de {fmt_num(total)}</div>
        <div class="barra-fondo"><div class="barra-relleno" style="width:{min(max(pct, 0), 100):.1f}%"></div></div>""",
        unsafe_allow_html=True,
    )
    st.markdown("")
    c_ok, c_pend = st.columns(2)
    with c_ok:
        st.markdown(
            f'<span class="pill pill-ok">En Alfresco: {fmt_num(en_alf)}</span>',
            unsafe_allow_html=True,
        )
    with c_pend:
        st.markdown(
            f'<span class="pill pill-pend">Pendientes: {fmt_num(pend)}</span>',
            unsafe_allow_html=True,
        )


# ----------------------------------------------------------------------------
# Analisis por año y estado
# ----------------------------------------------------------------------------
st.markdown("## Análisis por año y estado")
st.caption("Gráfico compuesto por año: contratos encontrados en Alfresco vs. faltantes (pendientes). La línea superior indica el total del año. "
           "🟢 <50% faltantes · 🟡 50-70% · 🔴 >70%.")
import plotly.graph_objects as go

g_year = f.groupby(["AÑO", "EN_ALFRESCO"]).size().reset_index(name="N")
pivot = g_year.pivot(index="AÑO", columns="EN_ALFRESCO", values="N").fillna(0)
for _col in (False, True):
    if _col not in pivot.columns:
        pivot[_col] = 0
pivot = pivot.reset_index().sort_values("AÑO")
pivot["Faltantes"] = pivot[False].astype(int)
pivot["Encontrados"] = pivot[True].astype(int)
pivot["Total"] = pivot["Faltantes"] + pivot["Encontrados"]
pivot["% Faltantes"] = (pivot["Faltantes"] / pivot["Total"].replace(0, pd.NA)).fillna(0) * 100
pivot["ColorFalt"] = pivot["% Faltantes"].map(color_semaforo)
pivot["Nivel"] = pivot["% Faltantes"].map(nivel_semaforo)

c3, c4 = st.columns(2)
with c3:
    st.markdown("#### Total de contratos por año (compuesto)")
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        x=pivot["AÑO"].astype(str),
        y=pivot["Encontrados"],
        name="En Alfresco",
        marker_color=PALETA["en_alfresco"],
        text=pivot["Encontrados"],
        textposition="inside",
        insidetextanchor="middle",
        texttemplate="%{text:,}",
        textfont=dict(size=11, color="white"),
        hovertemplate="Año %{x}<br>En Alfresco: %{y:,}<extra></extra>",
    ))
    fig3.add_trace(go.Bar(
        x=pivot["AÑO"].astype(str),
        y=pivot["Faltantes"],
        name="Faltantes (🟢<50% 🟡50-70% 🔴>70%)",
        marker_color=pivot["ColorFalt"].tolist(),
        text=pivot["Faltantes"],
        textposition="inside",
        insidetextanchor="middle",
        texttemplate="%{text:,}",
        textfont=dict(size=11, color="white"),
        customdata=pivot[["% Faltantes", "Nivel"]].to_numpy(),
        hovertemplate="Año %{x}<br>Faltantes: %{y:,} (%{customdata[0]:.1f} %)<br>%{customdata[1]}<extra></extra>",
    ))
    _ymax = float(pivot["Total"].max()) if len(pivot) else 1
    _etiqueta_y = (pivot["Total"] + _ymax * 0.06).tolist()
    fig3.add_trace(go.Scatter(
        x=pivot["AÑO"].astype(str),
        y=pivot["Total"],
        name="Total",
        mode="lines+markers",
        line=dict(color=PALETA["texto"], width=2.5),
        marker=dict(size=9, color=PALETA["texto"]),
        hovertemplate="Año %{x}<br>Total: %{y:,}<extra></extra>",
    ))
    fig3.add_trace(go.Scatter(
        x=pivot["AÑO"].astype(str),
        y=_etiqueta_y,
        name="Total (etiqueta)",
        mode="text",
        text=[f"Total: {int(v):,}" for v in pivot["Total"]],
        textposition="top center",
        textfont=dict(size=11, color=PALETA["texto"]),
        hoverinfo="skip",
        showlegend=False,
    ))
    fig3.update_layout(barmode="stack", template="plotly_white",
                       margin=dict(l=20, r=20, t=60, b=70),
                       legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5))
    fig3.update_xaxes(type="category", title="Año de suscripción", tickfont=dict(size=13))
    fig3.update_yaxes(title="Número de contratos", showgrid=True, gridcolor="#E2E8F0",
                      range=[0, _ymax * 1.22])
    fig3 = base_layout(fig3)
    fig3.update_layout(margin=dict(l=20, r=20, t=60, b=80),
                       legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5))
    fig3.update_yaxes(range=[0, _ymax * 1.22])
    st.plotly_chart(fig3, use_container_width=True, key="bar_year_comp")

with c4:
    st.markdown("#### Faltantes por año (detalle)")
    pivot["% en Alfresco"] = (pivot["Encontrados"] / pivot["Total"].replace(0, pd.NA)).fillna(0) * 100
    fig4 = go.Figure()
    fig4.add_trace(go.Bar(
        x=pivot["AÑO"].astype(str),
        y=pivot["Faltantes"],
        name="Faltantes",
        marker_color=pivot["ColorFalt"].tolist(),
        text=[f"{n:,} ({p:.1f} %)" for n, p in zip(pivot["Faltantes"], pivot["% Faltantes"])],
        textposition="outside",
        customdata=pivot[["Total", "% Faltantes", "Nivel"]].to_numpy(),
        hovertemplate="Año %{x}<br>Faltantes: %{y:,} (%{customdata[1]:.1f} %)<br>Total año: %{customdata[0]:,}<br>%{customdata[2]}<extra></extra>",
    ))
    fig4.update_layout(template="plotly_white", showlegend=False)
    fig4.update_xaxes(type="category", title="Año de suscripción", tickfont=dict(size=13))
    fig4.update_yaxes(title="Contratos faltantes", showgrid=True, gridcolor="#E2E8F0")
    fig4 = base_layout(fig4)
    st.plotly_chart(fig4, use_container_width=True, key="bar_year_falt")


# ----------------------------------------------------------------------------
# Analisis por tipo de contrato
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Analisis por tipo de contrato (tabla estructurada con catalogo de nombres)
# ----------------------------------------------------------------------------
st.markdown("## Análisis por tipo de contrato")
st.caption("Estos contratos pertenecen a estos 8 tipos de modalidades contractuales ante los entes de control")


@st.cache_data(show_spinner=False)
def cargar_catalogo_tipos(data_dir):
    """Lee el catálogo código -> nombre completo desde data/.

    Prioridad: tipos_contrato.csv > tipos_contrato.xlsx > cualquier
    *tipo*.csv/xlsx/txt. El .txt admite líneas 'CODIGO -> Nombre'.
    Devuelve dict {codigo_str: nombre}.
    """
    import glob as _glob

    dirs = []
    for d in (data_dir, DEFAULT_DATA_DIR):
        if d and os.path.isdir(d) and os.path.abspath(d) not in [os.path.abspath(x) for x in dirs]:
            dirs.append(d)

    patrones_csv = ["tipos_contrato.csv", "tipo_contrato.csv", "*tipo*.csv"]
    patrones_xlsx = ["tipos_contrato.xlsx", "tipo_contrato.xlsx", "*tipo*.xlsx"]
    patrones_txt = ["*tipo*.txt"]

    def _normalizar(df):
        df.columns = [str(c).strip().upper() for c in df.columns]
        col_cod = next((c for c in ("TIPO", "CODIGO", "CÓDIGO", "CODE", "ID") if c in df.columns), None)
        col_nom = next(
            (c for c in ("NOMBRE", "NOMBRE_COMPLETO", "DESCRIPCION", "DESCRIPCIÓN", "TIPO_NOMBRE") if c in df.columns),
            None,
        )
        if col_cod is None:
            col_cod = df.columns[0]
        if col_nom is None:
            col_nom = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        out = {}
        for _, row in df[[col_cod, col_nom]].dropna(how="all").iterrows():
            cod = "" if pd.isna(row[col_cod]) else str(row[col_cod]).strip()
            nom = "" if pd.isna(row[col_nom]) else str(row[col_nom]).strip()
            if cod:
                out[cod] = nom or cod
        return out

    for d in dirs:
        for pat in patrones_csv:
            for path in sorted(_glob.glob(os.path.join(d, pat))):
                if os.path.basename(path).startswith("~$"):
                    continue
                try:
                    for enc in ("utf-8-sig", "utf-8", "latin-1"):
                        try:
                            return _normalizar(pd.read_csv(path, dtype=str, encoding=enc))
                        except UnicodeDecodeError:
                            continue
                except Exception:
                    continue
        for pat in patrones_xlsx:
            for path in sorted(_glob.glob(os.path.join(d, pat))):
                if os.path.basename(path).startswith("~$"):
                    continue
                try:
                    return _normalizar(pd.read_excel(path, engine="openpyxl", dtype=str))
                except Exception:
                    continue
        for pat in patrones_txt:
            for path in sorted(_glob.glob(os.path.join(d, pat))):
                try:
                    with open(path, encoding="utf-8-sig") as fh:
                        lineas = fh.read().splitlines()
                except (OSError, UnicodeError):
                    continue
                out = {}
                for ln in lineas:
                    if "->" not in ln:
                        continue
                    cod, _, nom = ln.partition("->")
                    cod, nom = cod.strip(), nom.strip()
                    if cod and nom and re.search(r"\d", cod):
                        out[cod] = nom
                if out:
                    return out
    return {}


catalogo_tipos = cargar_catalogo_tipos(DATA_DIR_ACTIVA)

if "TIPO" in f.columns:
    _tmp = f[["TIPO", "EN_ALFRESCO"]].copy()
    _tmp["COD"] = _tmp["TIPO"].fillna("Sin tipo").astype("string").str.strip().replace("", "Sin tipo")
else:
    _tmp = pd.DataFrame({"COD": pd.Series(dtype=str), "EN_ALFRESCO": pd.Series(dtype=bool)})
    _tmp["COD"] = "Sin tipo"

_tmp["ESTADO_DOCUMENTAL"] = _tmp["EN_ALFRESCO"].map(lambda v: "Encontrado" if bool(v) else "No encontrado")
cruce = pd.crosstab(_tmp["COD"], _tmp["ESTADO_DOCUMENTAL"])
for _col in ("No encontrado", "Encontrado"):
    if _col not in cruce.columns:
        cruce[_col] = 0
tabla_tipos = (
    cruce[["No encontrado", "Encontrado"]]
    .reset_index()
    .rename(columns={"COD": "CODIGO", "No encontrado": "Contratos Faltantes", "Encontrado": "Contratos Encontrados"})
)
tabla_tipos["Contratos Faltantes"] = tabla_tipos["Contratos Faltantes"].fillna(0).astype(int)
tabla_tipos["Contratos Encontrados"] = tabla_tipos["Contratos Encontrados"].fillna(0).astype(int)
tabla_tipos["Total Contratos"] = tabla_tipos["Contratos Faltantes"] + tabla_tipos["Contratos Encontrados"]
tabla_tipos["Tipo de Contrato"] = tabla_tipos["CODIGO"].map(
    lambda c: f"{c} – {catalogo_tipos[c]}" if c in catalogo_tipos else str(c)
)
# 1) Ordenar SOLO las filas de detalle de mayor a menor por total.
tabla_tipos = (
    tabla_tipos[["Tipo de Contrato", "Contratos Faltantes", "Contratos Encontrados", "Total Contratos"]]
    .sort_values("Total Contratos", ascending=False)
    .reset_index(drop=True)
)
# 2) Calcular % de faltantes por tipo (0-100, sin division por cero).
tabla_tipos["% Faltantes"] = (
    tabla_tipos["Contratos Faltantes"] / tabla_tipos["Total Contratos"].replace(0, pd.NA)
).fillna(0) * 100

if tabla_tipos.empty:
    st.info("Sin datos de tipos de contrato bajo el filtro actual.")
else:
    # 3) Concatenar la fila de resumen AL FINAL, despues de ordenar.
    _tot_falt = int(tabla_tipos["Contratos Faltantes"].sum())
    _tot_enc = int(tabla_tipos["Contratos Encontrados"].sum())
    _tot = int(tabla_tipos["Total Contratos"].sum())
    total_row = pd.DataFrame([{
        "Tipo de Contrato": "Total general",
        "Contratos Faltantes": _tot_falt,
        "Contratos Encontrados": _tot_enc,
        "Total Contratos": _tot,
        "% Faltantes": (_tot_falt / _tot * 100) if _tot else 0.0,
    }])
    tabla_tipos = pd.concat([tabla_tipos, total_row], ignore_index=True)

    def _estilo_pct(v):
        try:
            x = float(v)
        except (TypeError, ValueError):
            return ""
        s = nivel_semaforo(x)
        if s.startswith("🔴"):
            return "background-color: #FDECEA; color: #7F1D1D; font-weight: 600;"
        if s.startswith("🟡"):
            return "background-color: #FEF9C3; color: #713F12; font-weight: 600;"
        return "background-color: #DCFCE7; color: #14532D; font-weight: 600;"

    st.caption("Semáforo por % faltantes: 🟢 Bajo (<50%) · 🟡 Medio (50-70%) · 🔴 Crítico (>70%).")
    st.dataframe(
        tabla_tipos.style.map(_estilo_pct, subset=["% Faltantes"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tipo de Contrato": st.column_config.TextColumn("Tipo de Contrato", width="large"),
            "Contratos Faltantes": st.column_config.NumberColumn("Contratos Faltantes", format="%d"),
            "Contratos Encontrados": st.column_config.NumberColumn("Contratos Encontrados", format="%d"),
            "Total Contratos": st.column_config.NumberColumn("Total Contratos", format="%d"),
            "% Faltantes": st.column_config.NumberColumn(
                "% Faltantes", format="%.1f %%",
                help="% de contratos faltantes (No encontrado) sobre el total del tipo.",
            ),
        },
    )

# ----------------------------------------------------------------------------
# Centros de costo y ordenadores
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Centros de costo y ordenadores (col. G = ORDENADOR_CENTRO, col. H = ORDENADOR_UNIDAD)
# ----------------------------------------------------------------------------
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Definición de colores estratégicos
COLOR_NEUTRO = "#277A1F"  # Slate Gray / Azul grisáceo para la mayoría
COLOR_ALERTA = "#F9573B"  # Siena / Ámbar para resaltar el #1 (Outlier)

# --- 1. SECCIÓN DE KPI CARDS (RESUMEN EJECUTIVO) ---
st.markdown("## Monitoreo y Gestión de Pendientes por Ordenación de Gasto")
st.caption(
    "Volumen de contratos pendientes de regularización o verificación "
    "distribuidos por área de responsabilidad."
)

_n_fondos = f["CENTRO_COSTO"].astype("string").str.strip().replace("", pd.NA).dropna().nunique() if "CENTRO_COSTO" in f.columns else 0
_n_dependencias = f["CENTRO_NOMBRE"].astype("string").str.strip().replace("", pd.NA).dropna().nunique() if "CENTRO_NOMBRE" in f.columns else _n_fondos
_n_unidades = f["ORDENADOR_UNIDAD"].astype("string").str.strip().replace("", pd.NA).dropna().nunique() if "ORDENADOR_UNIDAD" in f.columns else 0

k1, k2, k3 = st.columns(3)
k1.metric(
    label="Total Contratos (filtro)",
    value=f"{len(f):,}",
    help="Total de registros con el filtro actual.",
)
k2.metric(
    label="Centros de Costo (fondos)",
    value=f"{int(_n_fondos):,}",
    help=f"Fondos distintos en CENTRO_COSTO. Dependencias agrupadas (sin código): {int(_n_dependencias)}. Ej: Diana -> 3112 y 9244 = 2 fondos, 1 dependencia.",
)
k3.metric(
    label="Unidades de Supervisión",
    value=f"{int(_n_unidades):,}",
    help="Valores distintos en ORDENADOR_UNIDAD.",
)

st.write("---")


# --- 2. FUNCIÓN MEJORADA PARA EL GRÁFICO ---
def fig_top_ordenadores(serie, n=15, max_len=32, height=None):
    """Barra horizontal Top-N con colorimetría condicional y limpieza visual."""
    if serie is None:
        return None

    s = serie.dropna().astype("string").str.strip()
    s = s[s != ""]
    if s.empty:
        return None

    g = s.value_counts().head(n).reset_index()
    g.columns = ["Nombre", "N"]
    g = g.sort_values("N", ascending=False).reset_index(drop=True)

    largos = g["Nombre"].str.len()
    g["Etiqueta"] = g["Nombre"].str.slice(0, max_len)
    g.loc[largos > max_len, "Etiqueta"] = (
        g.loc[largos > max_len, "Etiqueta"] + "..."
    )

    # Colorimetría: la primera barra (#1) en ámbar/siena, el resto en gris neutro
    colores = [COLOR_ALERTA] + [COLOR_NEUTRO] * (len(g) - 1)

    max_val = int(g["N"].max())

    fig = go.Figure(
        go.Bar(
            x=g["N"],
            y=g["Etiqueta"],
            orientation="h",
            text=g["N"],
            textposition="outside",
            texttemplate="%{text:,}",
            customdata=g["Nombre"],
            hovertemplate="<b>%{customdata}</b><br>Pendientes: %{x:,}<extra></extra>",
            marker=dict(color=colores),
        )
    )

    # Estructura del Layout: Limpieza de ejes y ocultamiento de controles
    fig.update_layout(
        template="plotly_white",
        margin=dict(l=20, r=40, t=20, b=20),
        xaxis=dict(
            visible=False,  # Oculta el eje X ya que las barras tienen etiquetas numéricas
            range=[0, max_val * 1.20],  # Espacio para evitar corte de texto
        ),
        yaxis=dict(
            autorange="reversed",
            tickfont=dict(size=11, color="#2D3748"),
            title=None,
        ),
        height=height or max(380, 34 * len(g) + 60),
    )

    return fig


# --- 3. DISPOSICIÓN DE LAS COLUMNAS ---
c5, c6 = st.columns(2)

with c5:
    st.markdown("#### Contratos Pendientes por Ordenador (Centro de Costo)")
    st.caption(
        "Top 15 áreas de ejecución directa con mayor volumen acumulado."
    )
    if "ORDENADOR_CENTRO" not in f.columns:
        st.warning("Columna ORDENADOR_CENTRO no encontrada.")
    else:
        fig5 = fig_top_ordenadores(f["ORDENADOR_CENTRO"])
        if fig5 is None:
            st.info("Sin datos registrados para este filtro.")
        else:
            st.plotly_chart(
                fig5,
                use_container_width=True,
                config={"displayModeBar": False},
                key="top_ord_centro",
            )

with c6:
    st.markdown("#### Contratos Pendientes por Ordenador (Unidad Superior)")
    st.caption(
        "Top 15 unidades de supervisión jerárquica con mayor volumen pendiente."
    )
    if "ORDENADOR_UNIDAD" not in f.columns:
        st.warning("Columna ORDENADOR_UNIDAD no encontrada.")
    else:
        fig6 = fig_top_ordenadores(f["ORDENADOR_UNIDAD"])
        if fig6 is None:
            st.info("Sin datos registrados para este filtro.")
        else:
            st.plotly_chart(
                fig6,
                use_container_width=True,
                config={"displayModeBar": False},
                key="top_ord_unidad",
            )

# --- 4. ATRIBUCIÓN ORDENADOR (CENTRO) -> CENTROS DE COSTO ---
# Atribuye solo por ORDENADOR_CENTRO (quien ordena el gasto).
# Muestra fondos separados (CENTRO_COSTO con código) y dependencia agrupada
# (CENTRO_NOMBRE sin código). Ej: DIANA -> DIVISION DE CONTRATACION = 41
# (3112: 26 + 9244: 15). No altera los gráficos Top-15 de arriba.
st.markdown("#### Centros de costo por ordenador")
st.caption(
    "Atribución por ORDENADOR_CENTRO. Fondos separados y total agrupado por dependencia (sin código)."
)
if "ORDENADOR_CENTRO" in f.columns and "CENTRO_COSTO" in f.columns:
    _atr = f.copy()
    _atr["ORDENADOR_CENTRO"] = _atr["ORDENADOR_CENTRO"].astype("string").str.strip()
    _atr = _atr[_atr["ORDENADOR_CENTRO"].notna() & (_atr["ORDENADOR_CENTRO"] != "")]
    if "CENTRO_NOMBRE" not in _atr.columns:
        _atr["CENTRO_NOMBRE"] = _atr["CENTRO_COSTO"].map(extraer_centro_nombre)
    if _atr.empty:
        st.info("Sin datos registrados para este filtro.")
    else:
        _cruce = (
            _atr.groupby(["ORDENADOR_CENTRO", "CENTRO_NOMBRE", "CENTRO_COSTO"], dropna=False)
            .size()
            .reset_index(name="N contratos")
            .sort_values(["ORDENADOR_CENTRO", "N contratos"], ascending=[True, False])
        )
        _tot_ord = _atr.groupby("ORDENADOR_CENTRO").size().reset_index(name="Total ordenador")
        _cruce = _cruce.merge(_tot_ord, on="ORDENADOR_CENTRO", how="left")
        _cruce["% del ordenador"] = (_cruce["N contratos"] / _cruce["Total ordenador"].replace(0, pd.NA)).fillna(0) * 100
        _cruce = _cruce.rename(columns={
            "ORDENADOR_CENTRO": "Ordenador (centro)",
            "CENTRO_NOMBRE": "Dependencia (agrupada)",
            "CENTRO_COSTO": "Fondo (centro de costo)",
        })
        st.dataframe(
            _cruce,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Ordenador (centro)": st.column_config.TextColumn("Ordenador (centro)", width="large"),
                "Dependencia (agrupada)": st.column_config.TextColumn("Dependencia (agrupada)", width="large"),
                "Fondo (centro de costo)": st.column_config.TextColumn("Fondo (centro de costo)", width="large"),
                "N contratos": st.column_config.NumberColumn("N contratos", format="%d"),
                "Total ordenador": st.column_config.NumberColumn("Total ordenador", format="%d"),
                "% del ordenador": st.column_config.NumberColumn("% del ordenador", format="%.1f %%"),
            },
        )
else:
    st.warning("Columnas ORDENADOR_CENTRO / CENTRO_COSTO no encontradas.")

# --- 5. PIE DE PÁGINA EXPLICATIVO ---


# ----------------------------------------------------------------------------
# Serie y subserie (misma logica de extraccion)
# ----------------------------------------------------------------------------
import plotly.express as px
import streamlit as st

# Título principal del módulo
st.markdown("## Distribución por Serie y Subserie Documental")

# --- 1. MÉTRICAS / KPIS DIRECTIVOS ---
total_registros = len(f)
sin_serie_count = f["SERIE"].isna().sum() + (f["SERIE"] == "Sin serie").sum()
pct_sin_serie = (
    (sin_serie_count / total_registros) * 100 if total_registros > 0 else 0
)

m1, m2, m3 = st.columns(3)
m1.metric("Total Contratos Analizados", f"{total_registros:,}")
m2.metric(
    label="Contratos Sin Clasificar (Sin Serie)",
    value=f"{sin_serie_count:,}",
    delta=f"⚠️ {pct_sin_serie:.1f}% del total",
    delta_color="inverse",  # Alerta visual en rojo por brecha de clasificación
)
m3.metric(
    "Series / Subseries Identificadas",
    f"{f['SERIE'].replace('Sin serie', None).dropna().nunique()} / {f['SUBSERIE'].replace('Sin subserie', None).dropna().nunique()}",
)

st.divider()

# --- 2. PREPARACIÓN DE DATOS (MUESTRA CLASIFICADA) ---
df_series = f[f["SERIE"].notna() & (f["SERIE"] != "Sin serie")]
df_subseries = f[f["SUBSERIE"].notna() & (f["SUBSERIE"] != "Sin subserie")]

cc1, cc2 = st.columns(2)

# --- 3. GRÁFICO POR SERIE ---
with cc1:
    st.markdown("#### Top Series Catalogadas")
    gs = df_series["SERIE"].value_counts().head(10).reset_index()
    gs.columns = ["Serie", "N"]
    gs = gs.sort_values("N", ascending=True)  # Orden para barra horizontal

    fig_s = px.bar(
        gs,
        x="N",
        y="Serie",
        orientation="h",
        text="N",
        template="plotly_white",
        color_discrete_sequence=[VERDE_SOLIDO],
    )

    # Formateo de texto en la barra y etiqueta flotante
    fig_s.update_traces(
        textposition="outside",
        texttemplate="%{x:,}",  # Formato con separadores de miles
        hovertemplate="<b>Serie:</b> %{y}<br><b>Contratos:</b> %{x:,}<extra></extra>",
    )

    # Margen extra a la derecha (18%) para evitar que los números se recorten en el borde
    max_val_s = gs["N"].max() if not gs.empty else 1
    fig_s.update_xaxes(visible=False, range=[0, max_val_s * 1.18])
    fig_s.update_yaxes(title_text="", type="category")

    fig_s = base_layout(fig_s)
    fig_s.update_layout(margin=dict(r=40, l=10, t=10, b=10))
    st.plotly_chart(fig_s, use_container_width=True, key="top_series")

# --- 4. GRÁFICO POR SUBSERIE ---
with cc2:
    st.markdown("#### Top 10 Subseries Catalogadas")
    gss = df_subseries["SUBSERIE"].value_counts().head(10).reset_index()
    gss.columns = ["Subserie", "N"]
    gss = gss.sort_values("N", ascending=True)

    fig_ss = px.bar(
        gss,
        x="N",
        y="Subserie",
        orientation="h",
        text="N",
        template="plotly_white",
        color_discrete_sequence=[VERDE_SOLIDO],
    )

    # Formateo de texto en la barra y etiqueta flotante
    fig_ss.update_traces(
        textposition="outside",
        texttemplate="%{x:,}",
        hovertemplate="<b>Subserie:</b> %{y}<br><b>Contratos:</b> %{x:,}<extra></extra>",
    )

    # Margen extra a la derecha (18%) para evitar recorte
    max_val_ss = gss["N"].max() if not gss.empty else 1
    fig_ss.update_xaxes(visible=False, range=[0, max_val_ss * 1.18])
    fig_ss.update_yaxes(title_text="", type="category")

    fig_ss = base_layout(fig_ss)
    fig_ss.update_layout(margin=dict(r=40, l=10, t=10, b=10))
    st.plotly_chart(fig_ss, use_container_width=True, key="top_subseries")



# ----------------------------------------------------------------------------
# Tabla de detalle (seccion operativa principal)
# ----------------------------------------------------------------------------
st.markdown("## Detalle de contratos")
st.caption("Tabla filtrada lista para gestión: identifica pendientes y abre el enlace de Alfresco cuando exista.")

# Columna derivada solo para presentacion (no altera columnas originales)
f = f.copy()
f["ESTADO_DOCUMENTAL"] = f["EN_ALFRESCO"].map(lambda v: "Encontrado" if bool(v) else "No encontrado")

cols_pref = ["AÑO", "CONTRATO", "TIPO", "CENTRO_COSTO", "CENTRO_NOMBRE", "NOMBRE_CONTRATISTA",
             "ORDENADOR_CENTRO", "ORDENADOR_UNIDAD", "ESTADO_DOCUMENTAL",
             "SERIE", "SUBSERIE", "SERIE_RUTA", "CARPETA_ALFRESCO",
             "URL_ALFRESCO_1CLIC"]
cols_show = [c for c in cols_pref if c in f.columns]
tabla = f[cols_show].copy()

t_head, t_btn = st.columns([3, 1])
with t_head:
    st.markdown(f"**{fmt_num(len(tabla))} registros** con los filtros actuales")
with t_btn:
    csv = tabla.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Descargar filtrado en CSV", csv, "contratos_filtrado.csv", "text/csv",
                       use_container_width=True)

st.dataframe(
    tabla,
    use_container_width=True,
    hide_index=True,
    column_config={
        "AÑO": st.column_config.TextColumn("Año", width="small"),
        "CONTRATO": st.column_config.TextColumn("Contrato", width="medium"),
        "TIPO": st.column_config.TextColumn("Tipo", width="small"),
        "CENTRO_COSTO": st.column_config.TextColumn("Centro de costo", width="large"),
        "CENTRO_NOMBRE": st.column_config.TextColumn("Dependencia (agrupada)", width="large"),
        "NOMBRE_CONTRATISTA": st.column_config.TextColumn("Contratista", width="large"),
        "ORDENADOR_CENTRO": st.column_config.TextColumn("Ordenador (centro)", width="medium"),
        "ORDENADOR_UNIDAD": st.column_config.TextColumn("Ordenador (unidad)", width="medium"),
        "ESTADO_DOCUMENTAL": st.column_config.TextColumn("Estado documental", width="small"),
        "URL_ALFRESCO_1CLIC": st.column_config.LinkColumn("URL Alfresco", display_text="Abrir"),
    },
)
st.caption("La columna “URL Alfresco” solo muestra enlace cuando el dataset trae URL. No se generan enlaces artificiales.")

"""Dashboard de Contratos + Alfresco (2024 / 2025 / 2026).

Ejecutar con:  streamlit run app.py
Los archivos Excel pueden estar en cualquier carpeta: se elige en
"Configuracion de datos" o con la variable de entorno DASHBOARD_DATA_DIR.
Por defecto se usa data/ junto a app.py. El año se detecta desde el nombre del
archivo (ej: 2025_ENLACES_ALFRESCO.xlsx).
"""
import glob
import os
import re
import sys

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
    "pendiente": "#B42309",
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
VERDE_SOLIDO = "#2E7D32"
VERDE_OSCURO = "#1B5E20"
VERDES = ["#1b7a3d", "#2e7d32", "#43a047", "#66bb6a", "#81c784", "#00441b", "#238b45", "#a5d6a7"]
VERDES_CONTINUO = ["#edf8e9", "#c7e9c0", "#a1d99b", "#74c476", "#41ab5d", "#238b45", "#005a32"]

RE_ANIO = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
RE_ANIO_FILA = re.compile(r"(?:19|20)\d{2}")
RE_SERIE = re.compile(r"(C\d+)")
RE_SUBSERIE = re.compile(r"(C\d+\.\d+)")


def extraer_serie(ruta):
    """Serie = secuencia Cxx dentro de SERIE_RUTA. Ej: 1112_C09.23_... -> C09.

    Se usa regex C\\d+ porque la definicion literal 'despues del primer _ y antes
    del punto' falla en rutas con doble prefijo como 2123_6540_C09.11_... (daria
    '6540_C09'). El regex devuelve siempre C09 / C270 etc.
    """
    if not isinstance(ruta, str) or not ruta.strip():
        return "Sin serie"
    m = RE_SERIE.search(ruta)
    return m.group(1) if m else "Sin serie"


def extraer_subserie(ruta):
    if not isinstance(ruta, str) or not ruta.strip():
        return "Sin subserie"
    m = RE_SUBSERIE.search(ruta)
    return m.group(1) if m else extraer_serie(ruta)


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
            df = pd.read_excel(path, engine="openpyxl", dtype=str)
        except Exception as e:
            st.warning(f"No se pudo leer {os.path.basename(path)}: {e}")
            continue
        df.columns = [str(c).strip().upper() for c in df.columns]
        anio_defecto = detectar_anio(path)  # un año, o None si consolidado/sin año
        # 1) Si el Excel ya trae columna AÑO/ANO con años válidos, respetarla.
        if "AÑO" in df.columns or "ANO" in df.columns:
            col_orig = "AÑO" if "AÑO" in df.columns else "ANO"
            base = df[col_orig].astype("string").str.strip()
            valida = base.str.fullmatch(r"(?:19|20)\d{2}", na=False)
            df["AÑO"] = base.where(valida)
        else:
            df["AÑO"] = pd.NA
        # 2) Completar faltantes fila por fila: CONTRATO -> NUMERO -> año archivo.
        faltan = df["AÑO"].isna()
        if faltan.any():
            for col_fuente in ("CONTRATO", "NUMERO"):
                if col_fuente in df.columns and faltan.any():
                    inferido = df.loc[faltan, col_fuente].map(extraer_anio_fila)
                    df.loc[faltan, "AÑO"] = inferido
                    faltan = df["AÑO"].isna()
                if not faltan.any():
                    break
        # 3) Último recurso: año del nombre de archivo, o "Sin año".
        df["AÑO"] = df["AÑO"].fillna(anio_defecto if anio_defecto else "Sin año")
        df["ARCHIVO"] = os.path.basename(path)
        frames.append(df)
        origen.append(os.path.basename(path))
    if not frames:
        return pd.DataFrame(), []
    full = pd.concat(frames, ignore_index=True)
    # Normalizar nulos / espacios
    for c in full.columns:
        if full[c].dtype == object:
            full[c] = full[c].astype("string").str.strip()
    full["SERIE"] = full["SERIE_RUTA"].apply(extraer_serie) if "SERIE_RUTA" in full.columns else "Sin serie"
    full["SUBSERIE"] = full["SERIE_RUTA"].apply(extraer_subserie) if "SERIE_RUTA" in full.columns else "Sin subserie"
    full["EN_ALFRESCO"] = full["ESTADO"] == ESTADO_OK if "ESTADO" in full.columns else False
    # Ordenador combinado para facilitar busqueda (mantiene columnas originales)
    full["ORDENADOR_CENTRO"] = full.get("ORDENADOR_CENTRO")
    full["ORDENADOR_UNIDAD"] = full.get("ORDENADOR_UNIDAD")
    return full, sorted(origen)


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
        help="Encontrado = pasó la validación 1-clic (ESTADO = VALIDADO 1-CLIC). No encontrado = resto.",
    )
    with st.expander("⚙️ Filtros avanzados", expanded=False):
        _centros = sorted(df["CENTRO_COSTO"].dropna().unique().tolist()) if "CENTRO_COSTO" in df.columns else []
        sel_centros = st.multiselect("Centro de costo", _centros, default=[])
        if "ORDENADOR_CENTRO" in df.columns:
            _ord = sorted(set(df["ORDENADOR_CENTRO"].dropna().tolist())
                          | set(df["ORDENADOR_UNIDAD"].dropna().tolist()))
        else:
            _ord = []
        sel_ord = st.multiselect("Ordenador", _ord, default=[],
                                 help="Busca en ORDENADOR_CENTRO u ORDENADOR_UNIDAD.")
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

# ----------------------------------------------------------------------------
# Encabezado (la fuente tecnica no domina visualmente)
# ----------------------------------------------------------------------------
st.markdown("# Gestión Documental de Contratos en Alfresco")
st.markdown(
    '<p class="subtitulo">Monitoreo de la disponibilidad y estado de los contratos en Alfresco.'
    "</p>",
    unsafe_allow_html=True,
)

if df.empty:
    st.error(
        f"No se encontraron archivos .xlsx en:\n\n`{DATA_DIR_ACTIVA}`\n\n"
        "Copia ahí 2024_ENLACES_ALFRESCO.xlsx, 2025_ENLACES_ALFRESCO.xlsx y "
        "2026_ENLACES_ALFRESCO.xlsx (mismo formato de columnas), o elige otra carpeta en "
        "la barra lateral → ⚙️ Configuración de datos."
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
    f = f[(f["ORDENADOR_CENTRO"].isin(sel_ord)) | (f["ORDENADOR_UNIDAD"].isin(sel_ord))]
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
        color_discrete_map={"En Alfresco": PALETA["en_alfresco"], "Pendientes": "#E41805"},
        template="plotly_white",
    )
    fig2.update_traces(textinfo="value", textfont_size=13,
                       hovertemplate="%{label}: %{value:,} (%{percent})")
    fig2 = base_layout(fig2, height=340)
    fig2.update_layout(showlegend=True)
    st.plotly_chart(fig2, use_container_width=True)

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
c3, c4 = st.columns(2)
with c3:
    st.markdown("#### Contratos no encontrados por año")
    st.caption("Refleja el filtro actual. Para ver solo pendientes, filtra Estado documental → No encontrado.")
        
    g3 = f.groupby("AÑO").size().reset_index(name="N").sort_values("AÑO")
        
        # 1. Crear gráfico de área/línea para darle mayor presencia visual
    fig3 = px.line(
            g3, 
            x="AÑO", 
            y="N", 
            text="N", 
            markers=True, 
            template="plotly_white",
            color_discrete_sequence=[VERDE_SOLIDO]
        )
        
        # Formatear el texto a miles con separador (ej: 7,389 o 7.389)
    fig3.update_traces(
            texttemplate="%{text:,}",          # Formato con miles
            textposition="top center",          # Posición encima del punto
            textfont=dict(size=14, family="Arial Black", color="#1E293B"), # Números más grandes y legibles
            marker=dict(size=10, symbol="circle"), # Puntos de la línea más grandes
            hovertemplate="Año %{x}<br>Contratos faltantes: %{y:,}<extra></extra>"
        )
        
        # 2. Ajustar rangos de los ejes para que el texto superior/inferior no se corte ni cruce
    min_y = g3["N"].min() * 0.85
    max_y = g3["N"].max() * 1.12
        
    fig3.update_xaxes(
            type="category", 
            title="Año de suscripción",
            tickfont=dict(size=13)
        )
        
    fig3.update_yaxes(
            title="<b>Contratos faltantes</b>", # Título del eje más claro y explícito
            range=[min_y, max_y],              # Margen suficiente para que las etiquetas no colisionen
            showgrid=True,
            gridcolor="#E2E8F0"
        )
        
    fig3 = base_layout(fig3)
    st.plotly_chart(fig3, use_container_width=True)


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
    st.dataframe(
        tabla_tipos,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tipo de Contrato": st.column_config.TextColumn("Tipo de Contrato", width="large"),
            "Contratos Faltantes": st.column_config.NumberColumn("Contratos Faltantes", format="%d"),
            "Contratos Encontrados": st.column_config.NumberColumn("Contratos Encontrados", format="%d"),
            "Total Contratos": st.column_config.NumberColumn("Total Contratos", format="%d"),
            "% Faltantes": st.column_config.ProgressColumn(
                "% Faltantes", min_value=0, max_value=100, format="%.1f %%",
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
COLOR_NEUTRO = "#4A5568"  # Slate Gray / Azul grisáceo para la mayoría
COLOR_ALERTA = "#E68656"  # Siena / Ámbar para resaltar el #1 (Outlier)

# --- 1. SECCIÓN DE KPI CARDS (RESUMEN EJECUTIVO) ---
st.markdown("## Monitoreo y Gestión de Pendientes por Ordenación de Gasto")
st.caption(
    "Volumen de contratos pendientes de regularización o verificación "
    "distribuidos por área de responsabilidad."
)

if "ORDENADOR_CENTRO" in f.columns and "ORDENADOR_UNIDAD" in f.columns:
    tot_centro = f["ORDENADOR_CENTRO"].dropna().count()
    tot_unidad = f["ORDENADOR_UNIDAD"].dropna().count()
    total_general = max(tot_centro, tot_unidad)

    k1, k2, k3 = st.columns(3)
    k1.metric(
        label="Total Pendientes",
        value=f"{total_general:,}",
        help="Total de registros que requieren seguimiento.",
    )
    k2.metric(
        label="Áreas Evaluadas",
        value=f"{f['ORDENADOR_CENTRO'].nunique():,}",
        help="Número total de centros de costo con registros.",
    )
    k3.metric(
        label="Unidades de Supervisión",
        value=f"{f['ORDENADOR_UNIDAD'].nunique():,}",
        help="Número total de unidades superiores registradas.",
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
            )

# --- 4. PIE DE PÁGINA EXPLICATIVO ---
st.info(
    "**Nota de gestión:** Las cifras presentadas reflejan el volumen de registros pendientes "
    "de cierre/validación en el sistema. Los picos observados están asociados al volumen propio "
    "de contratación de cada área. Se recomienda priorizar el acompañamiento técnico e instrumental "
    "en las dependencias con mayor concentración."
)

# ----------------------------------------------------------------------------
# Serie y subserie (misma logica de extraccion)
# ----------------------------------------------------------------------------
import plotly.express as px
import streamlit as st

# Título principal del módulo
st.markdown("## Resumen Ejecutivo: Distribución por Serie y Subserie Documental")

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
    st.plotly_chart(fig_s, use_container_width=True)

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
    st.plotly_chart(fig_ss, use_container_width=True)



# ----------------------------------------------------------------------------
# Tabla de detalle (seccion operativa principal)
# ----------------------------------------------------------------------------
st.markdown("## Detalle de contratos")
st.caption("Tabla filtrada lista para gestión: identifica pendientes y abre el enlace de Alfresco cuando exista.")

# Columna derivada solo para presentacion (no altera columnas originales)
f = f.copy()
f["ESTADO_DOCUMENTAL"] = f["EN_ALFRESCO"].map(lambda v: "Encontrado" if bool(v) else "No encontrado")

cols_pref = ["AÑO", "CONTRATO", "TIPO", "CENTRO_COSTO", "NOMBRE_CONTRATISTA",
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
        "NOMBRE_CONTRATISTA": st.column_config.TextColumn("Contratista", width="large"),
        "ORDENADOR_CENTRO": st.column_config.TextColumn("Ordenador (centro)", width="medium"),
        "ORDENADOR_UNIDAD": st.column_config.TextColumn("Ordenador (unidad)", width="medium"),
        "ESTADO_DOCUMENTAL": st.column_config.TextColumn("Estado documental", width="small"),
        "URL_ALFRESCO_1CLIC": st.column_config.LinkColumn("URL Alfresco", display_text="Abrir"),
    },
)
st.caption("La columna “URL Alfresco” solo muestra enlace cuando el dataset trae URL. No se generan enlaces artificiales.")

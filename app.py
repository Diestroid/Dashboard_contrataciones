"""Dashboard de Contratos + Alfresco (2024 / 2025 / 2026).
Ejecutar con:  streamlit run app.py
Los archivos Excel pueden estar en cualquier carpeta: se elige desde la barra
lateral ("Carpeta de datos") o con la variable de entorno DASHBOARD_DATA_DIR.
Por defecto se usa data/ junto a app.py. El año se detecta desde el nombre del
archivo (ej: 2025_ENLACES_ALFRESCO.xlsx).
"""
import glob
import os
import re
import sys

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard Contratos Alfresco", layout="wide")

# Estilo: fondo claro + verdes
st.markdown(
    """
    <style>
    .stApp { background-color: #f4faf5; }
    section[data-testid="stSidebar"] { background-color: #e8f5e9; }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #c8e6c9;
        border-left: 6px solid #2e7d32;
        border-radius: 12px;
        padding: 12px;
    }
    div[data-testid="stMetric"] label { color: #1b5e20 !important; }
    h1, h2, h3 { color: #1b5e20 !important; }
    .stProgress > div > div > div > div { background-color: #2e7d32; }
    a { color: #1b7a3d !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

def base_dir():
    # Cuando corre como .exe de PyInstaller, __file__ apunta a _MEIPASS (temporal).
    # Los datos por defecto deben buscarse junto al .exe, no en el temporal.
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = base_dir()
DEFAULT_DATA_DIR = os.environ.get("DASHBOARD_DATA_DIR", os.path.join(BASE_DIR, "data"))
ESTADO_OK = "VALIDADO 1-CLIC"
VERDES = ["#1b7a3d", "#2e7d32", "#43a047", "#66bb6a", "#81c784", "#00441b", "#238b45", "#a5d6a7"]
VERDES_CONTINUO = ["#edf8e9", "#c7e9c0", "#a1d99b", "#74c476", "#41ab5d", "#238b45", "#005a32"]

# ---------------------------------------------------------------- helpers
RE_ANIO = re.compile(r"(19|20)\d{2}")
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
    m = RE_ANIO.search(os.path.basename(path))
    return m.group(0) if m else "Sin año"


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
        df["AÑO"] = detectar_anio(path)
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


# ---------------------------------------------------------------- carpeta de datos
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


st.sidebar.header("📁 Carpeta de datos")
nueva_ruta = st.sidebar.text_input(
    "Ruta con los Excel",
    value=st.session_state.data_dir,
    help="Pega cualquier ruta, ej: C:\\Contratos\\Excels o \\\\servidor\\contratos. También puedes definirla con la variable de entorno DASHBOARD_DATA_DIR o elegirla en el explorador de abajo.",
)
if st.sidebar.button("↩️ Volver a la carpeta por defecto", use_container_width=True):
    st.session_state.data_dir = DEFAULT_DATA_DIR
    st.session_state.explorar_dir = DEFAULT_DATA_DIR if os.path.isdir(DEFAULT_DATA_DIR) else BASE_DIR
    st.rerun()

# Si el usuario editó el texto, usarlo (previa validación al cargar).
if nueva_ruta and nueva_ruta != st.session_state.data_dir:
    st.session_state.data_dir = nueva_ruta
    if os.path.isdir(nueva_ruta):
        st.session_state.explorar_dir = os.path.abspath(nueva_ruta)

with st.sidebar.expander("📂 Explorar carpetas y elegir…", expanded=False):
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

DATA_DIR_ACTIVA = st.session_state.data_dir

df, archivos = cargar_datos(DATA_DIR_ACTIVA)

st.title("📊 Dashboard de Contratos — Alfresco")
if df.empty:
    st.error(
        f"No se encontraron archivos .xlsx en:\n\n`{DATA_DIR_ACTIVA}`\n\n"
        "Copia ahí 2024_ENLACES_ALFRESCO.xlsx, 2025_ENLACES_ALFRESCO.xlsx y "
        "2026_ENLACES_ALFRESCO.xlsx (mismo formato de columnas), o elige otra carpeta arriba."
    )
    st.stop()

st.caption(f"📁 Carpeta: `{DATA_DIR_ACTIVA}`  •  Archivos: {', '.join(archivos)}  •  Total filas: **{len(df):,}**".replace(",", "."))

# ---------------------------------------------------------------- sidebar
st.sidebar.header("Filtros")

anios = sorted(df["AÑO"].dropna().unique().tolist())
sel_anios = st.sidebar.multiselect("Año(s)", anios, default=anios)

tipos = sorted(df["TIPO"].dropna().unique().tolist()) if "TIPO" in df.columns else []
sel_tipos = st.sidebar.multiselect("Tipo contrato", tipos, default=tipos)

estados = sorted(df["ESTADO"].dropna().unique().tolist()) if "ESTADO" in df.columns else []
sel_estados = st.sidebar.multiselect("Estado", estados, default=estados)

centros = sorted(df["CENTRO_COSTO"].dropna().unique().tolist()) if "CENTRO_COSTO" in df.columns else []
sel_centros = st.sidebar.multiselect("Centro de costo", centros, default=[])

ordenadores_union = sorted(
    set(df["ORDENADOR_CENTRO"].dropna().tolist()) | set(df["ORDENADOR_UNIDAD"].dropna().tolist())
) if "ORDENADOR_CENTRO" in df.columns else []
sel_ord = st.sidebar.multiselect("Ordenador (centro o unidad)", ordenadores_union, default=[])

series = sorted(df["SERIE"].dropna().unique().tolist())
sel_series = st.sidebar.multiselect("Serie (ej: C09)", series, default=[])

subseries = sorted(df["SUBSERIE"].dropna().unique().tolist())
sel_sub = st.sidebar.multiselect("Subserie (ej: C09.11, C09.23)", subseries, default=[])

alfresco_opt = st.sidebar.radio("Ubicación Alfresco", ["Todos", "En Alfresco ✅", "Pendientes ⏳"], index=0)
texto = st.sidebar.text_input("🔎 Buscar (contrato, contratista, objeto)", "")

# ---------------------------------------------------------------- aplicar filtros
f = df.copy()
if sel_anios:
    f = f[f["AÑO"].isin(sel_anios)]
if sel_tipos and "TIPO" in f.columns:
    f = f[f["TIPO"].isin(sel_tipos)]
if sel_estados and "ESTADO" in f.columns:
    f = f[f["ESTADO"].isin(sel_estados)]
if sel_centros:
    f = f[f["CENTRO_COSTO"].isin(sel_centros)]
if sel_ord:
    f = f[(f["ORDENADOR_CENTRO"].isin(sel_ord)) | (f["ORDENADOR_UNIDAD"].isin(sel_ord))]
if sel_series:
    f = f[f["SERIE"].isin(sel_series)]
if sel_sub:
    f = f[f["SUBSERIE"].isin(sel_sub)]
if alfresco_opt == "En Alfresco ✅":
    f = f[f["EN_ALFRESCO"]]
elif alfresco_opt == "Pendientes ⏳":
    f = f[~f["EN_ALFRESCO"]]
if texto.strip():
    t = texto.strip().lower()
    cols_txt = [c for c in ["CONTRATO", "NOMBRE_CONTRATISTA", "OBJETO_CONTRATO", "CARPETA_ALFRESCO"] if c in f.columns]
    mask = pd.Series(False, index=f.index)
    for c in cols_txt:
        mask = mask | f[c].fillna("").str.lower().str.contains(t, na=False)
    f = f[mask]

if f.empty:
    st.warning("Sin resultados con los filtros actuales. Ajusta los filtros.")
    st.stop()

# ---------------------------------------------------------------- KPIs
total = len(f)
en_alf = int(f["EN_ALFRESCO"].sum())
pend = total - en_alf
pct = (en_alf / total * 100) if total else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("N.º contratos", f"{total:,}".replace(",", "."))
k2.metric("En Alfresco (VALIDADO 1-CLIC)", f"{en_alf:,}".replace(",", "."))
k3.metric("Pendientes", f"{pend:,}".replace(",", "."))
k4.metric("% en Alfresco", f"{pct:.1f}%")
st.progress(min(max(pct / 100, 0.0), 1.0))

# ---------------------------------------------------------------- gráficos
import plotly.express as px

c1, c2 = st.columns(2)
with c1:
    st.subheader("Por estado")
    g = f["ESTADO"].fillna("Sin estado").value_counts().reset_index()
    g.columns = ["Estado", "N"]
    fig = px.bar(g, x="Estado", y="N", text="N", color="Estado",
                 color_discrete_sequence=VERDES, template="plotly_white")
    fig.update_layout(xaxis_tickangle=-25, height=380, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)
with c2:
    st.subheader("Alfresco vs pendientes")
    g2 = pd.DataFrame({"Situación": ["En Alfresco", "Pendientes"], "N": [en_alf, pend]})
    fig2 = px.pie(g2, names="Situación", values="N", hole=0.45,
                  color="Situación",
                  color_discrete_map={"En Alfresco": "#2e7d32", "Pendientes": "#a5d6a7"},
                  template="plotly_white")
    fig2.update_layout(height=380)
    st.plotly_chart(fig2, use_container_width=True)

c3, c4 = st.columns(2)
with c3:
    st.subheader("Por año")
    g3 = f.groupby("AÑO").size().reset_index(name="N").sort_values("AÑO")
    fig3 = px.bar(g3, x="AÑO", y="N", text="N", color="AÑO",
                  color_discrete_sequence=VERDES, template="plotly_white")
    fig3.update_layout(height=380, showlegend=False)
    st.plotly_chart(fig3, use_container_width=True)
with c4:
    st.subheader("Por tipo")
    g4 = f["TIPO"].fillna("Sin tipo").value_counts().reset_index()
    g4.columns = ["Tipo", "N"]
    fig4 = px.bar(g4, x="Tipo", y="N", text="N", color="Tipo",
                  color_discrete_sequence=VERDES, template="plotly_white")
    fig4.update_layout(height=380, showlegend=False)
    st.plotly_chart(fig4, use_container_width=True)

c5, c6 = st.columns(2)
with c5:
    st.subheader("Top 15 centros de costo")
    g5 = f["CENTRO_COSTO"].fillna("Sin centro").value_counts().head(15).reset_index()
    g5.columns = ["Centro", "N"]
    fig5 = px.bar(g5, x="N", y="Centro", orientation="h", text="N", color="N",
                  color_continuous_scale=VERDES_CONTINUO, template="plotly_white")
    fig5.update_layout(height=520, yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig5, use_container_width=True)
with c6:
    st.subheader("Top 15 ordenadores")
    ord_serie = pd.concat([f["ORDENADOR_CENTRO"], f["ORDENADOR_UNIDAD"]]).dropna()
    ord_serie = ord_serie[ord_serie.str.strip() != ""]
    g6 = ord_serie.value_counts().head(15).reset_index()
    g6.columns = ["Ordenador", "N"]
    fig6 = px.bar(g6, x="N", y="Ordenador", orientation="h", text="N", color="N",
                  color_continuous_scale=VERDES_CONTINUO, template="plotly_white")
    fig6.update_layout(height=520, yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig6, use_container_width=True)

st.subheader("Por serie / subserie")
cc1, cc2 = st.columns(2)
with cc1:
    gs = f["SERIE"].fillna("Sin serie").value_counts().head(20).reset_index()
    gs.columns = ["Serie", "N"]
    st.plotly_chart(px.bar(gs, x="Serie", y="N", text="N", color="Serie",
                           color_discrete_sequence=VERDES, template="plotly_white"), use_container_width=True)
with cc2:
    gss = f["SUBSERIE"].fillna("Sin subserie").value_counts().head(20).reset_index()
    gss.columns = ["Subserie", "N"]
    st.plotly_chart(px.bar(gss, x="Subserie", y="N", text="N", color="Subserie",
                           color_discrete_sequence=VERDES, template="plotly_white"), use_container_width=True)
st.caption("Serie se extrae con regex `C\\d+` sobre SERIE_RUTA (ej: `1112_C09.23_...` → **C09**). Subserie con `C\\d+.\\d+` (→ **C09.23**). Así funciona también con rutas de doble prefijo como `2123_6540_C09.11_...`.")

# ---------------------------------------------------------------- tabla
st.subheader(f"Detalle ({len(f):,} registros)".replace(",", "."))
cols_pref = ["AÑO", "CONTRATO", "TIPO", "CENTRO_COSTO", "NOMBRE_CONTRATISTA",
             "ORDENADOR_CENTRO", "ORDENADOR_UNIDAD", "SERIE", "SUBSERIE",
             "SERIE_RUTA", "ESTADO", "CARPETA_ALFRESCO", "URL_ALFRESCO_1CLIC", "VALIDACION"]
cols_show = [c for c in cols_pref if c in f.columns]
tabla = f[cols_show].copy()

st.dataframe(
    tabla,
    use_container_width=True,
    hide_index=True,
    column_config={
        "URL_ALFRESCO_1CLIC": st.column_config.LinkColumn("URL Alfresco", display_text="Abrir 📂"),
    },
)

csv = tabla.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Descargar filtrado en CSV", csv, "contratos_filtrado.csv", "text/csv")

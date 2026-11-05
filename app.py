"""
Taco 2026 - Análisis operativo de logística (Streamlit)

Ejecución:
    pip install streamlit pandas plotly openpyxl numpy
    streamlit run app.py

Fuente de datos:
    - Link de Google Sheets (la planilla debe estar compartida como "Cualquier persona con el enlace
      puede ver"), o bien
    - Archivo local CSV / Excel.
    Opcional: guardar el link en .streamlit/secrets.toml  ->  SHEET_URL = "https://docs.google.com/spreadsheets/d/XXXX/edit"

Criterio de "viaje":
    La planilla suele tener una fila por cliente/entrega. Un VIAJE (salida) se arma agrupando las filas
    por Fecha + Camión + Chofer + Tipo de reparto (o por una columna de ID de viaje, si existe y se elige
    en la barra lateral). La cantidad de clientes de un viaje es la cantidad de clientes distintos.
"""

import io
import re
import unicodedata

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------------------------------
st.set_page_config(page_title="Taco 2026 - Logística", page_icon="🚚", layout="wide")

HOJA = "TACO"
SHEET_URL_DEFAULT = ""  # Pegá acá el link de la planilla si no querés usar secrets

# Valores que se consideran "vacíos"
NULOS = {"", "NAN", "NONE", "NULL", "-", "--", "S/D", "N/A", "#N/A", "NA", "NAT"}

# Zonas especiales (comparación sin tildes y en mayúsculas)
INTERIOR = ["SAN JUSTO", "SAN JAVIER"]

# Nombres posibles de cada columna + posición de respaldo (índice base 0 de la letra de Excel)
# W=22, X=23, Y=24, Z=25, AA=26, AH=33, AI=34
CANDIDATOS = {
    "fecha": ["FECHA"],
    "mes": ["MES"],
    "semana": ["SEMANA"],
    "tipo": ["TIPO DE REPARTO", "TIPO REPARTO"],
    "camion": ["CAMION", "PATENTE"],
    "chofer": ["CHOFER REAL", "CHOFER"],
    "ayud1": ["AYUDANTE 1", "AYUDANTE1"],
    "ayud2": ["AYUDANTE 2", "AYUDANTE2"],
    "zona": ["ZONA"],
    "cliente": ["CLIENTE"],
}
POSICION = {"tipo": 22, "camion": 23, "chofer": 24, "ayud1": 25, "ayud2": 26, "zona": 33, "cliente": 34}
ETIQUETAS = {
    "fecha": "Fecha", "mes": "Mes", "semana": "Semana", "tipo": "Tipo de reparto (W)",
    "camion": "Camión (X)", "chofer": "Chofer Real (Y)", "ayud1": "Ayudante 1 (Z)",
    "ayud2": "Ayudante 2 (AA)", "zona": "Zona (AH)", "cliente": "Cliente (AI)",
}


# ----------------------------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------------------------------------------------
def mostrar(fig):
    """Muestra un gráfico Plotly ocupando todo el ancho (compatible con distintas versiones)."""
    try:
        st.plotly_chart(fig, width="stretch")
    except Exception:
        st.plotly_chart(fig, use_container_width=True)


def tabla(df):
    """Muestra un DataFrame ocupando todo el ancho (compatible con distintas versiones)."""
    try:
        st.dataframe(df, width="stretch", hide_index=True)
    except Exception:
        st.dataframe(df, use_container_width=True, hide_index=True)


def norm_txt(x):
    """Limpia un texto: sin tildes, mayúsculas, sin espacios repetidos. Devuelve NaN si queda vacío."""
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return np.nan
    s = unicodedata.normalize("NFKD", str(x))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s).strip().upper()
    return np.nan if s in NULOS else s


def detectar_columna(raw, clave):
    """Busca la columna por nombre de encabezado; si no, usa la posición (letra de Excel)."""
    nombres = {c: re.sub(r"\.\d+$", "", str(norm_txt(c))) for c in raw.columns}
    coincidencias = [c for c, n in nombres.items() if n in CANDIDATOS[clave]]
    pos = POSICION.get(clave)
    if coincidencias:
        # Si hay varias columnas con el mismo nombre (ej. ZONA repetida), se prefiere la de la posición esperada
        if pos is not None and pos < len(raw.columns) and raw.columns[pos] in coincidencias:
            return raw.columns[pos]
        return coincidencias[0]
    if pos is not None and pos < len(raw.columns):
        return raw.columns[pos]
    return None


def url_a_csv(url):
    """Convierte el link de Google Sheets (o su ID) en un link de exportación CSV de la hoja TACO."""
    url = url.strip()
    if "output=csv" in url or "format=csv" in url or "tqx=out:csv" in url:
        return url
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    sheet_id = m.group(1) if m else url
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={HOJA}"


@st.cache_data(ttl=600, show_spinner="Cargando datos desde Google Sheets…")
def cargar_url(url, fila_enc):
    return pd.read_csv(url_a_csv(url), header=fila_enc - 1, dtype=str)


@st.cache_data(show_spinner="Leyendo archivo…")
def cargar_archivo(contenido, nombre, fila_enc):
    if nombre.lower().endswith((".xlsx", ".xlsm", ".xls")):
        xl = pd.ExcelFile(io.BytesIO(contenido))
        hoja = next((h for h in xl.sheet_names if h.strip().upper() == HOJA), xl.sheet_names[0])
        return xl.parse(hoja, header=fila_enc - 1, dtype=str)
    return pd.read_csv(io.BytesIO(contenido), header=fila_enc - 1, dtype=str)


def parsear_fechas(s):
    """Convierte texto a fecha (formato día/mes/año o ISO)."""
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce", format="mixed")
    except (TypeError, ValueError):
        return pd.to_datetime(s, dayfirst=True, errors="coerce")


def moda(s):
    """Valor más frecuente de una serie (ignorando nulos)."""
    s = s.dropna()
    return s.mode().iloc[0] if not s.empty else np.nan


def preparar_filas(raw, cols, col_id):
    """Limpia el dataset original y devuelve una fila por cliente/entrega con columnas estándar."""
    df = pd.DataFrame(index=raw.index)
    for clave in CANDIDATOS:
        origen = cols.get(clave)
        df[clave] = raw[origen] if origen else np.nan

    # Texto limpio para campos categóricos
    for c in ["tipo", "camion", "chofer", "ayud1", "ayud2", "zona", "cliente"]:
        df[c] = df[c].map(norm_txt)

    # Fecha, Mes y Semana (se derivan de la fecha; si falta, se usa lo que traiga la planilla)
    df["Fecha"] = parsear_fechas(df["fecha"])
    mes = df["Fecha"].dt.strftime("%Y-%m")
    iso = df["Fecha"].dt.isocalendar()
    semana = iso["year"].astype("Int64").astype(str) + "-S" + iso["week"].astype("Int64").astype(str).str.zfill(2)
    semana = semana.where(df["Fecha"].notna(), np.nan)
    df["Mes"] = mes.fillna(df["mes"].map(norm_txt))
    df["Semana"] = semana.fillna(df["semana"].map(norm_txt))

    df = df.rename(columns={"tipo": "Tipo", "camion": "Camion", "chofer": "Chofer", "ayud1": "Ayudante1",
                            "ayud2": "Ayudante2", "zona": "Zona", "cliente": "Cliente"})
    df = df[["Fecha", "Mes", "Semana", "Tipo", "Camion", "Chofer", "Ayudante1", "Ayudante2", "Zona", "Cliente"]]

    # Se descartan filas totalmente vacías en lo operativo
    df = df.dropna(how="all", subset=["Camion", "Chofer", "Zona", "Cliente"])

    # Identificador de viaje
    if col_id:
        clave = raw.loc[df.index, col_id].astype(str).str.strip()
        clave = clave.where(~clave.str.upper().isin(NULOS), "FILA-" + df.index.astype(str))
    else:
        clave = (df["Fecha"].dt.strftime("%Y-%m-%d").fillna("SF") + "|" + df["Camion"].fillna("SC") + "|"
                 + df["Chofer"].fillna("SCH") + "|" + df["Tipo"].fillna("ST"))
    df["ViajeID"] = pd.factorize(clave)[0] + 1
    return df.reset_index(drop=True)


def armar_viajes(filas):
    """Agrupa las filas (clientes) en viajes/salidas."""
    viajes = filas.groupby("ViajeID", sort=False).agg(
        Fecha=("Fecha", "min"), Mes=("Mes", moda), Semana=("Semana", moda), Tipo=("Tipo", moda),
        Camion=("Camion", moda), Chofer=("Chofer", moda), Ayudante1=("Ayudante1", moda),
        Ayudante2=("Ayudante2", moda), Zona=("Zona", moda),
        Zonas=("Zona", lambda s: ", ".join(sorted(s.dropna().unique()))),
        Clientes=("Cliente", "nunique"), Paradas=("Cliente", "size"),
    ).reset_index()
    # Si la planilla no trae clientes, se usa la cantidad de filas
    viajes["Clientes"] = viajes["Clientes"].where(viajes["Clientes"] > 0, viajes["Paradas"])
    return viajes


def contiene(serie, texto):
    """True si la zona contiene el texto indicado."""
    return serie.fillna("").str.contains(texto, regex=False)


@st.cache_data(show_spinner=False)
def a_excel(viajes, detalle):
    """Genera el archivo Excel (hojas Viajes y Detalle) en memoria."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        viajes.to_excel(w, sheet_name="Viajes", index=False)
        detalle.to_excel(w, sheet_name="Detalle", index=False)
    return buf.getvalue()


def ranking(df, col, top):
    """Ranking de repartos livianos por persona (col) con total y % de livianos."""
    t = (df.dropna(subset=[col]).groupby(col)
         .agg(Total=("Categoria", "size"), Livianos=("Categoria", lambda s: int((s == "Liviano").sum())))
         .reset_index())
    t["% Liviano"] = (t["Livianos"] / t["Total"] * 100).round(1)
    return t[t["Livianos"] > 0].sort_values("Livianos", ascending=False).head(top)


# ----------------------------------------------------------------------------------------------
# BARRA LATERAL: ORIGEN DE DATOS
# ----------------------------------------------------------------------------------------------
st.title("🚚 Taco 2026 · Análisis operativo de distribución")

st.sidebar.header("📂 Datos")
origen = st.sidebar.radio("Origen", ["Google Sheets (URL)", "Archivo local (CSV/Excel)"])
fila_enc = st.sidebar.number_input("Fila del encabezado", min_value=1, value=1, step=1,
                                   help="Número de fila donde están los nombres de las columnas.")

raw = None
try:
    if origen.startswith("Google"):
        try:
            url_def = st.secrets["SHEET_URL"]
        except Exception:
            url_def = SHEET_URL_DEFAULT
        url = st.sidebar.text_input("Link o ID de la planilla", value=url_def)
        if url:
            raw = cargar_url(url, int(fila_enc))
    else:
        archivo = st.sidebar.file_uploader("Subí el archivo", type=["csv", "xlsx", "xlsm", "xls"])
        if archivo is not None:
            raw = cargar_archivo(archivo.getvalue(), archivo.name, int(fila_enc))
except Exception as e:
    st.error(f"No se pudo cargar la planilla: {e}")
    st.info("Verificá que esté compartida con 'Cualquier persona con el enlace' y que exista la pestaña TACO.")
    st.stop()

if raw is None or raw.empty:
    st.info("👈 Cargá el link de Google Sheets o un archivo local para comenzar.")
    st.stop()

# Mapeo de columnas (autodetectado, editable)
with st.sidebar.expander("🧩 Mapeo de columnas"):
    opciones = ["(ninguna)"] + [str(c) for c in raw.columns]
    cols = {}
    for clave, etiqueta in ETIQUETAS.items():
        det = detectar_columna(raw, clave)
        idx = opciones.index(str(det)) if det is not None and str(det) in opciones else 0
        sel = st.selectbox(etiqueta, opciones, index=idx, key=f"map_{clave}")
        cols[clave] = None if sel == "(ninguna)" else sel
    sel_id = st.selectbox("ID de viaje (opcional)", ["Automático"] + [str(c) for c in raw.columns],
                          help="Automático = Fecha + Camión + Chofer + Tipo de reparto.")
    col_id = None if sel_id == "Automático" else sel_id

if not cols["chofer"] or not cols["zona"]:
    st.error("Hace falta al menos mapear las columnas Chofer Real y Zona.")
    st.stop()

filas = preparar_filas(raw, cols, col_id)
viajes_all = armar_viajes(filas)

# Parámetros de análisis
with st.sidebar.expander("⚙️ Parámetros"):
    umbral_liviano = st.slider("Reparto liviano: clientes ≤", 1, 30, 9)
    tope_fijo = st.slider("Chofer 'Fijo': zonas distintas ≤", 1, 5, 2)
viajes_all["Categoria"] = np.where(viajes_all["Clientes"] <= umbral_liviano, "Liviano", "Pesado")

# ----------------------------------------------------------------------------------------------
# FILTROS GLOBALES
# ----------------------------------------------------------------------------------------------
st.sidebar.header("🔎 Filtros")
mascara = pd.Series(True, index=viajes_all.index)

if viajes_all["Fecha"].notna().any():
    fmin, fmax = viajes_all["Fecha"].min().date(), viajes_all["Fecha"].max().date()
    rango = st.sidebar.date_input("Rango de fechas", value=(fmin, fmax), min_value=fmin, max_value=fmax)
    if isinstance(rango, (tuple, list)) and len(rango) == 2 and (rango[0], rango[1]) != (fmin, fmax):
        mascara &= viajes_all["Fecha"].between(pd.Timestamp(rango[0]), pd.Timestamp(rango[1]) + pd.Timedelta(days=1))

for etiqueta, campo in [("Mes", "Mes"), ("Semana", "Semana"), ("Chofer Real", "Chofer"),
                        ("Tipo de reparto", "Tipo"), ("Camión", "Camion")]:
    opciones_f = sorted(viajes_all[campo].dropna().unique())
    elegidos = st.sidebar.multiselect(etiqueta, opciones_f, placeholder="Todos")
    if elegidos:
        mascara &= viajes_all[campo].isin(elegidos)

v = viajes_all[mascara].copy()
if v.empty:
    st.warning("No hay viajes con los filtros seleccionados.")
    st.stop()

detalle = filas[filas["ViajeID"].isin(v["ViajeID"])]

# Tabla larga de ayudantes (una fila por viaje y ayudante)
largo = (v.melt(id_vars=["ViajeID", "Chofer", "Categoria", "Mes", "Semana"],
                value_vars=["Ayudante1", "Ayudante2"], value_name="Ayudante")
         .dropna(subset=["Ayudante"]).drop(columns="variable"))

# ----------------------------------------------------------------------------------------------
# TARJETAS KPI
# ----------------------------------------------------------------------------------------------
es_interior = v["Zona"].isin(INTERIOR)
n_interior = int(es_interior.sum())
n_linea = int((v["Zona"].notna() & ~es_interior).sum())
es_raf = contiene(v["Zona"], "RAFAELA")
es_sun = contiene(v["Zona"], "SUNCHALES")
es_parana = contiene(v["Zona"], "PARANA")
n_liv = int((v["Categoria"] == "Liviano").sum())

st.caption(f"{len(v):,} viajes (salidas) · {len(detalle):,} entregas · "
           f"{v['Chofer'].nunique()} choferes · {v['Camion'].nunique()} camiones".replace(",", "."))

k = st.columns(5)
k[0].metric("Zonas de línea", n_linea, help="Viajes a todas las zonas excepto SAN JUSTO y SAN JAVIER.")
k[1].metric("Interior", n_interior, help="Viajes a SAN JUSTO y SAN JAVIER.")
k[2].metric("Rafaela + Sunchales", int(es_raf.sum() + es_sun.sum()),
            delta=f"Rafaela {int(es_raf.sum())} · Sunchales {int(es_sun.sum())}", delta_color="off")
k[3].metric("Cruce a Paraná", int(es_parana.sum()))
k[4].metric(f"Repartos livianos (≤{umbral_liviano})", f"{n_liv / len(v) * 100:.1f}%",
            delta=f"{n_liv} de {len(v)} viajes", delta_color="off")

st.divider()

# ----------------------------------------------------------------------------------------------
# MÓDULOS DE VISUALIZACIÓN
# ----------------------------------------------------------------------------------------------
tab_a, tab_b, tab_c, tab_d = st.tabs(["A · Intercambiabilidad", "B · Equipos", "C · Repartos livianos", "D · Operativos"])

# ---------- A) Matriz de intercambiabilidad de choferes ----------
with tab_a:
    st.subheader("Choferes Fijos vs. Multizona")
    vz = v.dropna(subset=["Chofer", "Zona"])
    zc = (vz.groupby("Chofer").agg(Zonas=("Zona", "nunique"), Viajes=("ViajeID", "size")).reset_index())
    zc["Perfil"] = np.where(zc["Zonas"] <= tope_fijo, f"Fijo (≤{tope_fijo} zonas)", "Multizona")
    mapa_color = {f"Fijo (≤{tope_fijo} zonas)": "#2E86AB", "Multizona": "#F18F01"}

    c1, c2 = st.columns([2, 1])
    with c1:
        fig = px.bar(zc.sort_values("Zonas"), x="Zonas", y="Chofer", color="Perfil", orientation="h",
                     text="Zonas", hover_data=["Viajes"], color_discrete_map=mapa_color,
                     title="Zonas distintas atendidas por chofer")
        fig.update_layout(height=max(400, 24 * len(zc)), yaxis_title=None)
        mostrar(fig)
    with c2:
        resumen = zc["Perfil"].value_counts().reset_index()
        resumen.columns = ["Perfil", "Choferes"]
        mostrar(px.pie(resumen, names="Perfil", values="Choferes", color="Perfil",
                       color_discrete_map=mapa_color, hole=0.5, title="Distribución de perfiles"))

    st.markdown("**Matriz Chofer × Zona (cantidad de viajes)**")
    ct = pd.crosstab(vz["Chofer"], vz["Zona"])
    fig = px.imshow(ct, text_auto=True, aspect="auto", color_continuous_scale="Blues",
                    labels=dict(x="Zona", y="Chofer", color="Viajes"))
    fig.update_layout(height=max(400, 24 * len(ct)))
    mostrar(fig)

# ---------- B) Composición de equipos ----------
with tab_b:
    st.subheader("Duplas y tríos habituales")
    ambos = sorted(set(v["Chofer"].dropna()) & set(largo["Ayudante"]))

    duplas = (largo.dropna(subset=["Chofer"]).groupby(["Chofer", "Ayudante"]).size()
              .reset_index(name="Viajes").sort_values("Viajes", ascending=False))
    duplas["Ayudante también es chofer"] = np.where(duplas["Ayudante"].isin(ambos), "Sí", "No")

    top_n = st.slider("Cantidad de duplas a mostrar", 5, 40, 15, key="top_duplas")
    d = duplas.head(top_n).copy()
    d["Dupla"] = d["Chofer"] + "  +  " + d["Ayudante"]
    fig = px.bar(d.sort_values("Viajes"), x="Viajes", y="Dupla", orientation="h", text="Viajes",
                 color="Ayudante también es chofer", title="Duplas Chofer + Ayudante más frecuentes",
                 color_discrete_map={"Sí": "#E4572E", "No": "#2E86AB"})
    fig.update_layout(height=max(400, 26 * len(d)), yaxis_title=None)
    mostrar(fig)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Ayudante habitual por chofer**")
        viajes_chofer = v.groupby("Chofer").size().rename("Viajes del chofer")
        hab = duplas.drop_duplicates("Chofer").set_index("Chofer").join(viajes_chofer)
        hab["% de sus viajes"] = (hab["Viajes"] / hab["Viajes del chofer"] * 100).round(1)
        tabla(hab.reset_index().rename(columns={"Ayudante": "Ayudante habitual", "Viajes": "Viajes juntos"})
              [["Chofer", "Ayudante habitual", "Viajes juntos", "Viajes del chofer", "% de sus viajes"]])
    with c2:
        st.markdown("**Tríos (Chofer + Ayudante 1 + Ayudante 2)**")
        trios = (v.dropna(subset=["Chofer"]).fillna({"Ayudante1": "—", "Ayudante2": "—"})
                 .groupby(["Chofer", "Ayudante1", "Ayudante2"]).size().reset_index(name="Viajes")
                 .sort_values("Viajes", ascending=False))
        tabla(trios)

    st.subheader("Ayudantes que también operan como choferes")
    if ambos:
        como_chofer = v.groupby("Chofer").size()
        como_ayud = largo.groupby("Ayudante").size()
        dual = pd.DataFrame({"Persona": ambos,
                             "Viajes como chofer": [int(como_chofer.get(p, 0)) for p in ambos],
                             "Viajes como ayudante": [int(como_ayud.get(p, 0)) for p in ambos]})
        dual["% como chofer"] = (dual["Viajes como chofer"] /
                                 (dual["Viajes como chofer"] + dual["Viajes como ayudante"]) * 100).round(1)
        largo_dual = dual.melt(id_vars="Persona", value_vars=["Viajes como chofer", "Viajes como ayudante"],
                               var_name="Rol", value_name="Viajes")
        fig = px.bar(largo_dual, x="Persona", y="Viajes", color="Rol", barmode="group",
                     color_discrete_sequence=["#2E86AB", "#F18F01"])
        mostrar(fig)
        tabla(dual.sort_values("Viajes como ayudante", ascending=False))
    else:
        st.info("No hay ayudantes que figuren como choferes en el período filtrado.")

# ---------- C) Repartos livianos ----------
with tab_c:
    st.subheader(f"Repartos livianos (≤ {umbral_liviano} clientes) vs. pesados")
    gran = st.radio("Agrupar por", ["Mes", "Semana"], horizontal=True)
    per = v.dropna(subset=[gran]).groupby([gran, "Categoria"]).size().unstack(fill_value=0)
    for cat in ["Liviano", "Pesado"]:
        if cat not in per:
            per[cat] = 0
    per["% Liviano"] = (per["Liviano"] / (per["Liviano"] + per["Pesado"]) * 100).round(1)
    per = per.reset_index().sort_values(gran)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(per.melt(id_vars=gran, value_vars=["Liviano", "Pesado"], var_name="Categoría",
                              value_name="Viajes"),
                     x=gran, y="Viajes", color="Categoría", barmode="stack", text_auto=True,
                     color_discrete_map={"Liviano": "#59A14F", "Pesado": "#B07AA1"},
                     title="Viajes livianos y pesados")
        fig.update_xaxes(type="category")
        mostrar(fig)
    with c2:
        fig = px.line(per, x=gran, y="% Liviano", markers=True, title="% de repartos livianos")
        fig.update_xaxes(type="category")
        fig.update_yaxes(range=[0, 100])
        mostrar(fig)

    fig = px.histogram(v, x="Clientes", nbins=max(10, int(v["Clientes"].max())),
                       title="Distribución de clientes por viaje")
    fig.add_vline(x=umbral_liviano + 0.5, line_dash="dash", line_color="red",
                  annotation_text=f"Límite liviano ({umbral_liviano})")
    mostrar(fig)

    st.subheader("Ranking de repartos livianos")
    top_r = st.slider("Cantidad de personas por ranking", 5, 30, 10, key="top_rank")
    c1, c2 = st.columns(2)
    for col_ui, df_src, campo, titulo in [(c1, v, "Chofer", "Choferes"), (c2, largo, "Ayudante", "Ayudantes")]:
        with col_ui:
            r = ranking(df_src, campo, top_r)
            if r.empty:
                st.info(f"Sin repartos livianos para {titulo.lower()}.")
                continue
            fig = px.bar(r.sort_values("Livianos"), x="Livianos", y=campo, orientation="h", text="Livianos",
                         hover_data=["Total", "% Liviano"], title=f"{titulo} con más repartos livianos",
                         color_discrete_sequence=["#59A14F"])
            fig.update_layout(yaxis_title=None)
            mostrar(fig)

# ---------- D) Gráficos operativos ----------
with tab_d:
    st.subheader("Frecuencia de choferes por zona")
    zonas_orden = v["Zona"].value_counts().index.tolist()
    zona_sel = st.selectbox("Zona", ["(Todas las zonas)"] + zonas_orden, index=1 if zonas_orden else 0)
    base = v if zona_sel.startswith("(") else v[v["Zona"] == zona_sel]
    freq = base.dropna(subset=["Chofer"]).groupby("Chofer").size().reset_index(name="Viajes").sort_values("Viajes")
    fig = px.bar(freq, x="Viajes", y="Chofer", orientation="h", text="Viajes",
                 title=f"Viajes por chofer · {zona_sel}")
    fig.update_layout(height=max(350, 24 * len(freq)), yaxis_title=None)
    mostrar(fig)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Viajes a Paraná por camión y mes")
        par = v[es_parana].dropna(subset=["Mes"])
        if par.empty:
            st.info("Sin viajes a Paraná en el período.")
        else:
            pc = par.groupby(["Mes", "Camion"], dropna=False).size().reset_index(name="Viajes").fillna("S/D")
            fig = px.bar(pc.sort_values("Mes"), x="Mes", y="Viajes", color="Camion", barmode="group",
                         text_auto=True)
            fig.update_xaxes(type="category")
            mostrar(fig)
            with st.expander("Ver tabla"):
                tabla(pc.pivot_table(index="Camion", columns="Mes", values="Viajes", fill_value=0).reset_index())
    with c2:
        st.subheader("Salidas mensuales a Rafaela y Sunchales")
        rs = v[es_raf | es_sun].dropna(subset=["Mes"]).copy()
        if rs.empty:
            st.info("Sin salidas a Rafaela o Sunchales en el período.")
        else:
            rs["Destino"] = np.where(contiene(rs["Zona"], "RAFAELA"), "RAFAELA", "SUNCHALES")
            rc = rs.groupby(["Mes", "Destino"]).size().reset_index(name="Salidas")
            fig = px.bar(rc.sort_values("Mes"), x="Mes", y="Salidas", color="Destino", barmode="group",
                         text_auto=True, color_discrete_map={"RAFAELA": "#2E86AB", "SUNCHALES": "#F18F01"})
            fig.update_xaxes(type="category")
            mostrar(fig)

# ----------------------------------------------------------------------------------------------
# EXPORTACIÓN Y DATOS LIMPIOS
# ----------------------------------------------------------------------------------------------
st.divider()
st.subheader("📥 Exportar datos filtrados")

v_exp = v.drop(columns=["Categoria"]).assign(Categoria=v["Categoria"]).copy()
v_exp["Fecha"] = v_exp["Fecha"].dt.date
det_exp = detalle.copy()
det_exp["Fecha"] = det_exp["Fecha"].dt.date

b1, b2, b3 = st.columns(3)
b1.download_button("⬇️ Viajes (CSV)", v_exp.to_csv(index=False).encode("utf-8-sig"),
                   file_name="taco_viajes_filtrados.csv", mime="text/csv")
b2.download_button("⬇️ Detalle de entregas (CSV)", det_exp.to_csv(index=False).encode("utf-8-sig"),
                   file_name="taco_detalle_filtrado.csv", mime="text/csv")
b3.download_button("⬇️ Excel (Viajes + Detalle)", a_excel(v_exp, det_exp), file_name="taco_filtrado.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("🔍 Tabla de datos limpios (viajes)"):
    tabla(v_exp)
with st.expander("🔍 Tabla de datos limpios (detalle de entregas)"):
    tabla(det_exp)

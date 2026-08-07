import re
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Tablero de Control - Instrumentos PRODUCE",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# PALETA DE COLORES (base app_mod)
# ============================================================
NAVY = "#1a2634"
DARK = "#2c3e50"
TEAL = "#16a085"
SLATE = "#34495e"
RED = "#e74c3c"
AMBER = "#f1c40f"
BLUE = "#3498db"
GREEN = "#2ecc71"
INK_FAINT = "#8a97a8"
LINE = "#dde3ea"

STATUS_COLOR = {"crit": RED, "bajo": AMBER, "medio": BLUE, "alto": GREEN}
STATUS_LABEL = {
    "crit": "Crítico (<25%)",
    "bajo": "Bajo (25–50%)",
    "medio": "Medio (50–75%)",
    "alto": "Alto (>75%)",
}

# ============================================================
# ESTILOS GLOBALES ADICIONALES
# ============================================================
st.markdown(f"""
<style>
.topbar {{
    background: linear-gradient(135deg, {NAVY} 0%, #122a44 55%, #1a3a5c 100%);
    color:#fff; padding:34px 40px; border-radius:16px; margin-bottom:18px;
    position:relative; overflow:hidden;
}}
.topbar::before {{
    content:""; position:absolute; inset:0;
    background-image:
        radial-gradient(circle at 85% -10%, rgba(31,174,143,.35), transparent 45%),
        repeating-linear-gradient(90deg, rgba(255,255,255,.035) 0 1px, transparent 1px 64px);
    pointer-events:none;
}}
.topbar-inner {{ position:relative; }}
.eyebrow {{
    display:inline-flex; align-items:center; gap:8px; font-size:12.5px; letter-spacing:.09em;
    text-transform:uppercase; color:#9fd9c9; font-weight:600; margin-bottom:12px; font-family:Arial;
}}
.eyebrow .dot {{ width:7px; height:7px; border-radius:50%; background:{TEAL}; box-shadow:0 0 0 4px rgba(31,174,143,.25); }}
.topbar h1 {{ font-size:32px; font-weight:800; margin:0 0 8px; letter-spacing:-.01em; font-family:Arial; }}
.topbar p {{ margin:0; color:#b9c6d8; font-size:14.5px; max-width:640px; line-height:1.5; font-family:Arial; }}
.topbar-meta {{
    display:flex; gap:22px; margin-top:18px; flex-wrap:wrap;
    font-size:12.5px; color:#8fa3bb; font-family:Arial;
}}
.topbar-meta b {{ color:#e7edf5; font-weight:600; }}
.note-box {{
    background:#e8e8e8; border-radius:12px; border-left:4px solid {AMBER};
    padding:14px 20px; font-family:Arial; font-size:12.5px; color:#333; line-height:1.55;
}}
.chip {{
    display:inline-block; font-size:11.5px; font-weight:600; color:{DARK}; background:#eef3f8;
    border:1px solid #d9e3ec; padding:4px 10px; border-radius:20px; margin:2px 4px 2px 0;
}}
.section-title {{ font-size:19px; font-weight:700; color:{DARK}; margin: 10px 0 2px; font-family:Arial; }}
.section-hint {{ font-size:12.5px; color:{INK_FAINT}; margin-bottom:10px; font-family:Arial; }}
.kpi-mini {{
    background:#fff; border:1px solid {LINE}; border-radius:12px; padding:14px 16px;
    box-shadow: 1px 1px 6px rgba(0,0,0,0.06);
}}
.kpi-mini-label {{ font-size:11px; text-transform:uppercase; letter-spacing:.05em; color:{INK_FAINT}; font-weight:700; font-family:Arial; }}
.kpi-mini-value {{ font-size:15px; font-weight:700; margin-top:4px; color:{DARK}; font-family:Arial; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# AUTENTICACIÓN CON CONTRASEÑA
# ============================================================
#def verificar_password():
#    def password_ingresada():
#        if st.session_state["password"] == st.secrets["password"]:
#            st.session_state["password_correcta"] = True
#            del st.session_state["password"]
#        else:
#            st.session_state["password_correcta"] = False
#
#    if "password_correcta" not in st.session_state:
#        st.text_input("Contraseña", type="password", on_change=password_ingresada, key="password")
#        return False
#    elif not st.session_state["password_correcta"]:
#        st.text_input("Contraseña", type="password", on_change=password_ingresada, key="password")
#        st.error("😕 Contraseña incorrecta")
#        return False
#    else:
#        return True
#
#if not verificar_password():
#    st.stop()

# ============================================================
# 1. CARGA DE DATOS (directa, sin carga por el usuario)
# ============================================================
NOMBRE_ARCHIVO = "Instrumentos bi.xlsx"  # nombre exacto tal como está en el repositorio

COLMAP = {
    "cod": "cod",
    "Entidad que reporta": "entidad",
    "Tipo Instrumento": "tipo",
    "Nombre del Instrumento": "instrumento",
    "Norma aprobatoria": "norma",
    "Objetivo Estrategico": "objetivo",
    "Linea de accion": "linea",
    "Accion Estrategica / Iniciativa": "accion",
    "Responsable de realizar el seguimiento y evaluacion de indicadores": "responsable_seg",
    "Responsable del indicador": "responsable_ind",
    "Observacion": "observacion",
}

INVALID_DEP = {"", "-", "no se encontro", "no encontrado"}


def _try_float(x):
    try:
        if x is None:
            return None
        return float(x)
    except (TypeError, ValueError):
        return None


@st.cache_data(show_spinner="Procesando Excel…")
def load_data(path) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name="BASE")
    df = pd.DataFrame()
    for src, dst in COLMAP.items():
        df[dst] = raw[src] if src in raw.columns else ""

    # Indicador: preferir "Indicador(es)"; si viene vacío, usar "Indicador mod"
    ind_a = raw.get("Indicador(es)")
    ind_b = raw.get("Indicador mod")
    if ind_a is not None and ind_b is not None:
        df["indicador"] = ind_a.fillna(ind_b)
    elif ind_a is not None:
        df["indicador"] = ind_a
    elif ind_b is not None:
        df["indicador"] = ind_b
    else:
        df["indicador"] = ""

    df["meta"] = raw["Meta global"].apply(_try_float)
    df["avance"] = raw["Avance acum mod"].apply(_try_float)
    df["medible"] = df["meta"].notna() & df["avance"].notna()
    df["pct"] = df.apply(
        lambda r: round(r["avance"] / r["meta"] * 100, 2)
        if r["medible"] and r["meta"] not in (0, None)
        else None,
        axis=1,
    )
    df["pct_capped"] = df["pct"].apply(lambda v: min(v, 100) if v is not None else None)

    for col in ["entidad", "tipo", "instrumento", "norma", "objetivo", "responsable_ind", "indicador"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["objetivo_norm"] = df["objetivo"].apply(
        lambda v: v if v not in ("", "-") else "Sin objetivo estratégico asociado"
    )
    return df


try:
    DATA = load_data(NOMBRE_ARCHIVO)
except FileNotFoundError:
    st.error(f"No se encontró el archivo **{NOMBRE_ARCHIVO}**. Verifica que esté en el mismo repositorio.")
    st.stop()

col_instrumento = "instrumento"

ENTIDADES = sorted(e for e in DATA["entidad"].unique() if e)
TIPOS = sorted(t for t in DATA["tipo"].unique() if t)
INSTRUMENTOS = sorted(i for i in DATA["instrumento"].unique() if i)

# ============================================================
# ENCABEZADO
# ============================================================
st.markdown(f"""
<div class="topbar">
  <div class="topbar-inner">
    <div class="eyebrow"><span class="dot"></span>PRODUCE · Seguimiento de indicadores</div>
    <h1>📊 Tablero de Control — Instrumentos PRODUCE</h1>
    <p>Procesado con información copilada por la DGPAR, obteniendo {DATA[col_instrumento].nunique()} instrumentos
    de gestión y {len(DATA)} indicadores asociados.</p>
    <div class="topbar-meta">
      <div>Fuente: <b>{NOMBRE_ARCHIVO}</b></div>
      <div>Entidades responsables: <b>{len(ENTIDADES)}</b></div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================
def split_dependencias(text: str):
    if not text:
        return []
    parts = re.split(r"/|,|\sy\s", text, flags=re.IGNORECASE)
    out = []
    for p in parts:
        p = p.strip()
        if p and p.lower() not in INVALID_DEP:
            out.append(p)
    return out


def status_of(pct):
    if pct is None or pd.isna(pct):
        return None
    if pct < 25:
        return "crit"
    if pct < 50:
        return "bajo"
    if pct < 75:
        return "medio"
    return "alto"


def avg_pct(items: pd.DataFrame):
    vals = items.loc[items["medible"] & items["pct"].notna(), "pct_capped"]
    return float(vals.mean()) if len(vals) else None


def tarjeta_kpi(titulo, valor, color=DARK):
    st.markdown(f"""
    <div style="background-color:{color}; color:white; padding:18px; border-radius:12px;
                text-align:center; font-family:Arial;
                box-shadow: 2px 2px 8px rgba(0,0,0,0.2);">
        <div style="font-size:13px; opacity:0.8;">{titulo}</div>
        <div style="font-size:32px; font-weight:bold;">{valor}</div>
    </div>
    """, unsafe_allow_html=True)


def gauge_figure(value, height=220):
    v = 0 if value is None or pd.isna(value) else value
    fig = go.Figure(go.Indicator(
        mode="number+gauge",
        value=v,
        title={"text": "% Avance Promedio"},
        number={'suffix': "%", 'font': {'size': 44}},
        gauge={
            'axis': {'range': [0, 100]},
            'bar': {'color': DARK},
            'steps': [
                {'range': [0, 25], 'color': RED},
                {'range': [25, 50], 'color': AMBER},
                {'range': [50, 75], 'color': BLUE},
                {'range': [75, 100], 'color': GREEN},
            ],
            'threshold': {'line': {'color': "black", 'width': 4}, 'thickness': 0.75, 'value': v},
        }
    ))
    fig.update_layout(height=height, margin=dict(t=40, b=10, l=20, r=20))
    return fig


# ============================================================
# 2. KPIs GENERALES
# ============================================================
total_general = DATA[col_instrumento].nunique()
indicadores_general = len(DATA)
medibles_df = DATA[DATA["medible"]]
indicadores_medibles = len(medibles_df)
avance_general = avg_pct(DATA)

k1, k2, k3, k4 = st.columns(4)
with k1:
    tarjeta_kpi("Total Instrumentos", total_general)
with k2:
    tarjeta_kpi("Total Indicadores", indicadores_general)
with k3:
    tarjeta_kpi("Indicadores Medibles", indicadores_medibles, color=TEAL)
with k4:
    tarjeta_kpi("% Avance General", f"{avance_general:.1f}%" if avance_general is not None else "—", color=SLATE)

st.markdown("<br>", unsafe_allow_html=True)

tab_resumen, tab_subtotales = st.tabs(["📊 Resumen general", "🗂️ Subtotales por instrumento y dependencia"])

# ============================================================
# TAB 1 — RESUMEN GENERAL
# ============================================================
with tab_resumen:

    # ---- 3. GRÁFICO DE BARRAS + DONA DE ESTADOS ----
    c1, c2 = st.columns([1.35, 1])
    with c1:
        conteo_entidad = medibles_df["entidad"].value_counts().reset_index()
        conteo_entidad.columns = ["entidad", "count"]
        total_indicadores_grafico = len(medibles_df)

        fig_barras = px.bar(conteo_entidad,
                             x='entidad', y='count',
                             title='Cantidad de Indicadores por Área Responsable',
                             labels={'count': 'Cantidad', 'entidad': 'Área Responsable'},
                             color='entidad', text='count')
        fig_barras.add_annotation(
            text=f"Total: {total_indicadores_grafico}",
            xref="paper", yref="paper", x=1, y=-0.15, showarrow=False,
            font=dict(size=13, color="gray"), align="right"
        )
        fig_barras.update_layout(margin=dict(b=100), height=400, showlegend=False)
        st.plotly_chart(fig_barras, use_container_width=True)

    with c2:
        buckets = {k: 0 for k in STATUS_COLOR}
        for p in medibles_df["pct"]:
            s = status_of(p)
            if s:
                buckets[s] += 1
        fig_donut = go.Figure(go.Pie(
            labels=[STATUS_LABEL[k] for k in buckets],
            values=list(buckets.values()),
            marker=dict(colors=[STATUS_COLOR[k] for k in buckets]),
            hole=0.62,
        ))
        fig_donut.update_layout(height=400, margin=dict(l=10, r=10, t=30, b=10),
                                 title="Indicadores por estado de avance", showlegend=True,
                                 legend=dict(orientation="h", yanchor="bottom", y=-0.3))
        st.plotly_chart(fig_donut, use_container_width=True)

    # --- Nota: instrumentos sin indicador/meta definida ---
    instrumentos_sin_indicador = sorted(
        name for name, g in DATA.groupby(col_instrumento) if not g["medible"].any() and name
    )
    lista_html = "".join([f"<li>{nombre}</li>" for nombre in instrumentos_sin_indicador]) or "<li>Ninguno.</li>"
    st.markdown(f"""
    <div class="note-box">
        <b>Nota:</b> Los siguientes instrumentos de gestión presentan acciones a seguir.
        Sin embargo, los indicadores y/o metas no se encuentran definidas o en su defecto los indicadores no han sido reportados.
        <ul style="margin:8px 0 0 0; padding-left:20px;">
            {lista_html}
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ---- 4. CRUCE NORMA x ENTIDAD ----
    st.markdown('<div class="section-title">🔗 Cumplimiento cruzado: Norma × Entidad</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-hint">Elige un instrumento para ver qué entidades van cumpliendo, '
        'o una entidad para ver qué instrumentos va cumpliendo</div>',
        unsafe_allow_html=True,
    )
    cc1, cc2 = st.columns(2)

    instr_by_count = DATA[col_instrumento].value_counts().index.tolist()
    with cc1:
        st.markdown("**Por instrumento / norma → avance por entidad**")
        sel_instr = st.selectbox("Instrumento (norma)", instr_by_count, key="cross_instr")
        rows = DATA[DATA[col_instrumento] == sel_instr]
        norma_txt = rows["norma"].iloc[0] if len(rows) else "—"
        by_ent = (
            rows.groupby("entidad")
            .apply(lambda g: pd.Series({"n": len(g), "pct": avg_pct(g)}))
            .reset_index()
            .sort_values("pct", ascending=True, na_position="first")
        )
        st.caption(f"Norma: **{norma_txt}** · Entidades involucradas: **{len(by_ent)}** · Indicadores: **{len(rows)}**")
        colors = [STATUS_COLOR[status_of(p)] if pd.notna(p) else "#c9d2dd" for p in by_ent["pct"]]
        fig_ci = go.Figure(go.Bar(
            x=by_ent["pct"].fillna(0), y=by_ent["entidad"], orientation="h",
            marker_color=colors,
            text=[f"{p:.1f}%" if pd.notna(p) else "N/D" for p in by_ent["pct"]],
            textposition="outside",
        ))
        fig_ci.update_layout(height=320, margin=dict(l=10, r=30, t=10, b=10), xaxis=dict(range=[0, 105]))
        st.plotly_chart(fig_ci, use_container_width=True)

    ent_by_count = DATA["entidad"].value_counts().index.tolist()
    with cc2:
        st.markdown("**Por entidad → avance por instrumento / norma**")
        sel_ent = st.selectbox("Entidad responsable", ent_by_count, key="cross_ent")
        rows2 = DATA[DATA["entidad"] == sel_ent]
        by_instr = (
            rows2.groupby(col_instrumento)
            .apply(lambda g: pd.Series({"n": len(g), "pct": avg_pct(g)}))
            .reset_index()
            .sort_values("pct", ascending=True, na_position="first")
        )
        st.caption(f"Instrumentos que reporta: **{len(by_instr)}** · Indicadores: **{len(rows2)}**")
        labels = [i if len(i) <= 40 else i[:40] + "…" for i in by_instr[col_instrumento]]
        colors2 = [STATUS_COLOR[status_of(p)] if pd.notna(p) else "#c9d2dd" for p in by_instr["pct"]]
        fig_ce = go.Figure(go.Bar(
            x=by_instr["pct"].fillna(0), y=labels, orientation="h",
            marker_color=colors2,
            text=[f"{p:.1f}%" if pd.notna(p) else "N/D" for p in by_instr["pct"]],
            textposition="outside",
            hovertext=by_instr[col_instrumento],
        ))
        fig_ce.update_layout(height=320, margin=dict(l=10, r=30, t=10, b=10), xaxis=dict(range=[0, 105]))
        st.plotly_chart(fig_ce, use_container_width=True)

    st.markdown("---")

    # ---- 5. MATRIZ CRUZADA ENTIDAD x TIPO ----
    st.markdown('<div class="section-title">🧮 Matriz cruzada: Entidad × Tipo de instrumento</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-hint">N° de indicadores que reporta cada entidad, por tipo de instrumento de gestión</div>',
        unsafe_allow_html=True,
    )
    matrix = pd.crosstab(DATA["entidad"], DATA["tipo"])
    matrix = matrix.reindex(index=ENTIDADES, columns=TIPOS, fill_value=0)
    matrix["Total"] = matrix.sum(axis=1)
    total_row = matrix.sum(axis=0)
    total_row.name = "Total"
    matrix_display = pd.concat([matrix, total_row.to_frame().T])

    fig_hm = px.imshow(
        matrix.drop(columns="Total"),
        text_auto=True,
        color_continuous_scale=[[0, "#f7f9fb"], [1, TEAL]],
        aspect="auto",
    )
    fig_hm.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig_hm, use_container_width=True)
    with st.expander("Ver tabla con totales"):
        st.dataframe(matrix_display, use_container_width=True)

    st.markdown("---")

    # ============================================================
    # 6. FILTRO Y DETALLE POR INSTRUMENTO
    # ============================================================
    st.markdown('<div class="section-title">🔍 Detalle por Instrumento</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-hint">Filtra para inspeccionar el avance de un recorte específico</div>',
                unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    f_ent = f1.selectbox("Entidad responsable", ["Todas"] + ENTIDADES, key="f_ent")
    f_ins = f2.selectbox("Instrumento", ["Todos"] + INSTRUMENTOS, key="f_ins")
    f_tip = f3.selectbox("Tipo de instrumento", ["Todos"] + TIPOS, key="f_tip")
    f_q = f4.text_input("Buscar indicador", placeholder="ej. economía circular, MYPE, digital…", key="f_q")

    data = medibles_df.copy()
    if f_ent != "Todas":
        data = data[data["entidad"] == f_ent]
    if f_ins != "Todos":
        data = data[data[col_instrumento] == f_ins]
    if f_tip != "Todos":
        data = data[data["tipo"] == f_tip]
    if f_q:
        ql = f_q.lower()
        data = data[
            data["indicador"].str.lower().str.contains(ql)
            | data[col_instrumento].str.lower().str.contains(ql)
            | data["norma"].str.lower().str.contains(ql)
        ]

    cantidad_indicadores = len(data)
    avance_prom_filtrado = avg_pct(data)
    responsables = sorted(set(data["responsable_seg"]) - {"", "-"})
    texto_responsable = ", ".join(responsables) if responsables else "Sin dato"

    col_tarjeta, col_responsable, col_gauge = st.columns([1, 1, 2])

    with col_tarjeta:
        st.markdown(f"""
        <div style="background-color:{DARK}; color:white; padding:20px; border-radius:12px;
                    height:220px; text-align:center; font-family:Arial;
                    box-shadow: 2px 2px 8px rgba(0,0,0,0.2); display:flex; flex-direction:column;
                    justify-content:center;">
            <div style="font-size:14px; opacity:0.8;">N° de Indicadores</div>
            <div style="font-size:36px; font-weight:bold;">{cantidad_indicadores}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_responsable:
        st.markdown(f"""
        <div style="background-color:{DARK}; color:white; padding:20px; border-radius:12px;
                    height:220px; text-align:center; font-family:Arial;
                    box-shadow: 2px 2px 8px rgba(0,0,0,0.2); display:flex; flex-direction:column;
                    justify-content:center; overflow:hidden;">
            <div style="font-size:14px; opacity:0.8;">Responsable del seguimiento</div>
            <div style="font-size:16px; font-weight:bold; margin-top:8px;">{texto_responsable}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_gauge:
        if avance_prom_filtrado is None:
            st.info("Sin datos de avance disponibles para este filtro.")
        else:
            st.plotly_chart(gauge_figure(avance_prom_filtrado, height=220), use_container_width=True)

    st.markdown("---")

    # ---- 7. TABLA DE INDICADORES ----
    st.markdown('<div class="section-title">📋 Tabla de Indicadores</div>', unsafe_allow_html=True)
    st.caption(f"{len(data)} filas")
    tabla = data[["entidad", col_instrumento, "norma", "indicador", "meta", "avance", "pct"]].rename(columns={
        "entidad": "Entidad", col_instrumento: "Instrumento", "norma": "Norma aprobatoria",
        "indicador": "Indicador", "meta": "Meta", "avance": "Avance", "pct": "% Avance",
    })
    st.dataframe(
        tabla,
        use_container_width=True,
        height=480,
        hide_index=True,
        column_config={
            "% Avance": st.column_config.ProgressColumn("% Avance", format="%.1f%%", min_value=0, max_value=100),
            "Meta": st.column_config.NumberColumn("Meta", format="%.2f"),
            "Avance": st.column_config.NumberColumn("Avance", format="%.2f"),
        },
    )

# ============================================================
# TAB 2 — SUBTOTALES POR INSTRUMENTO / DEPENDENCIA / OBJETIVO
# ============================================================
with tab_subtotales:
    st.markdown('<div class="section-title">Subtotales</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-hint">Elige el corte: por instrumento de gestión, por dependencia participante '
        'o por objetivo estratégico</div>',
        unsafe_allow_html=True,
    )

    sub_tab_instr, sub_tab_dep, sub_tab_obj = st.tabs(
        ["Por instrumento de gestión", "Por dependencia", "Por objetivo estratégico"]
    )

    def render_master_detail(build_fn, name_col, key_prefix, subtitle_fn, chips_label, chips_col,
                              extra_col_label, note_html=None):
        rows = build_fn()
        if note_html:
            st.markdown(note_html, unsafe_allow_html=True)
        search = st.text_input("Buscar…", key=f"{key_prefix}_search")
        options = [r[name_col] for r in rows]
        if search:
            sl = search.lower()
            options = [o for o in options if sl in o.lower()]
        if not options:
            st.info("Sin resultados.")
            return
        labels = {o: f"{o}  ·  {next(r for r in rows if r[name_col] == o)['n']} indicador(es)" for o in options}
        sel = st.selectbox("Selecciona:", options, format_func=lambda o: labels[o], key=f"{key_prefix}_sel")
        r = next(x for x in rows if x[name_col] == sel)

        st.markdown(f"### {r[name_col]}")
        st.caption(subtitle_fn(r))
        m1, m2, m3 = st.columns(3)
        m1.metric("N° Indicadores", r["n"])
        m2.metric("N° Medibles", r["n_medibles"])
        m3.metric(extra_col_label, r["n_extra"])

        gc, cc = st.columns([1, 1.4])
        with gc:
            st.plotly_chart(gauge_figure(r["pct_prom"], height=220), use_container_width=True)
        with cc:
            st.markdown(f"**{chips_label}**")
            chips_html = "".join(f'<span class="chip">{c}</span>' for c in r[chips_col]) or "<i>Ninguna identificada</i>"
            st.markdown(chips_html, unsafe_allow_html=True)

        etiqueta = "este instrumento" if key_prefix == "instr" else ("esta dependencia" if key_prefix == "dep" else "este objetivo")
        st.markdown(f"**Indicadores de {etiqueta}**")
        items_df = r["items"][["indicador", "meta", "avance", "pct"]].rename(
            columns={"indicador": "Indicador", "meta": "Meta", "avance": "Avance", "pct": "% Avance"}
        )
        st.dataframe(
            items_df, use_container_width=True, hide_index=True, height=340,
            column_config={
                "% Avance": st.column_config.ProgressColumn("% Avance", format="%.1f%%", min_value=0, max_value=100),
            },
        )

    # ---- Por instrumento ----
    def build_instr_subtotals():
        rows = []
        for instrumento, g in DATA.groupby(col_instrumento):
            if not instrumento:
                continue
            deps = set()
            g["responsable_ind"].apply(lambda t: deps.update(split_dependencias(t)))
            rows.append({
                "instrumento": instrumento,
                "entidad": g["entidad"].iloc[0],
                "tipo": g["tipo"].iloc[0],
                "items": g,
                "n": len(g),
                "n_medibles": int(g["medible"].sum()),
                "n_extra": len(deps),
                "pct_prom": avg_pct(g),
                "dependencias": sorted(deps),
            })
        return sorted(rows, key=lambda r: r["n"], reverse=True)

    with sub_tab_instr:
        render_master_detail(
            build_instr_subtotals, "instrumento", "instr",
            subtitle_fn=lambda r: f"{r['entidad']} · {r['tipo']}",
            chips_label="Dependencias que participan", chips_col="dependencias",
            extra_col_label="N° Dependencias",
        )

    # ---- Por dependencia ----
    def build_dep_subtotals():
        rows_map = {}
        for idx, row in DATA.iterrows():
            for dep in split_dependencias(row["responsable_ind"]):
                rows_map.setdefault(dep, {"idx": [], "instrumentos": set()})
                rows_map[dep]["idx"].append(idx)
                rows_map[dep]["instrumentos"].add(row[col_instrumento])
        rows = []
        for dep, info in rows_map.items():
            g = DATA.loc[info["idx"]]
            rows.append({
                "dependencia": dep,
                "items": g,
                "n": len(g),
                "n_medibles": int(g["medible"].sum()),
                "n_extra": len(info["instrumentos"]),
                "pct_prom": avg_pct(g),
                "instrumentos": sorted(info["instrumentos"]),
            })
        return sorted(rows, key=lambda r: r["n"], reverse=True)

    with sub_tab_dep:
        st.markdown(
            """<div class="note-box"><b>Nota:</b> la "dependencia" es la unidad concreta responsable de cada
            indicador (columna "Responsable del indicador" del Excel). Un indicador puede tener más de una
            dependencia corresponsable (ej. "DGPAR/PNDP"), en cuyo caso se cuenta para cada una.</div>""",
            unsafe_allow_html=True,
        )
        render_master_detail(
            build_dep_subtotals, "dependencia", "dep",
            subtitle_fn=lambda r: "Dependencia / unidad responsable",
            chips_label="Instrumentos en los que participa", chips_col="instrumentos",
            extra_col_label="N° Instrumentos",
        )

    # ---- Por objetivo estratégico ----
    def build_obj_subtotals():
        rows = []
        for objetivo, g in DATA.groupby("objetivo_norm"):
            rows.append({
                "objetivo": objetivo,
                "items": g,
                "n": len(g),
                "n_medibles": int(g["medible"].sum()),
                "n_extra": g[col_instrumento].nunique(),
                "pct_prom": avg_pct(g),
                "instrumentos": sorted(i for i in g[col_instrumento].unique() if i),
                "entidades": sorted(e for e in g["entidad"].unique() if e),
            })
        sin = "Sin objetivo estratégico asociado"
        return sorted(rows, key=lambda r: (r["objetivo"] == sin, -r["n"]))

    with sub_tab_obj:
        st.markdown(
            """<div class="note-box"><b>Nota:</b> se agrupa por el texto del "Objetivo Estratégico" declarado en
            cada instrumento. Los indicadores cuyo instrumento no define un objetivo estratégico explícito se
            agrupan en "Sin objetivo estratégico asociado".</div>""",
            unsafe_allow_html=True,
        )
        render_master_detail(
            build_obj_subtotals, "objetivo", "obj",
            subtitle_fn=lambda r: ", ".join(r["entidades"]),
            chips_label="Instrumentos que aportan a este objetivo", chips_col="instrumentos",
            extra_col_label="N° Instrumentos",
        )

st.caption("Tablero generado a partir de la base de datos del repositorio.")

# ============================================================
# 8. DESCARGAR BASE ACTUALIZADA (deshabilitado)
# ============================================================
#import io
#buffer = io.BytesIO()
#DATA.to_excel(buffer, index=False)
#st.download_button(
#    label="⬇️ Descargar Excel actualizado",
#    data=buffer.getvalue(),
#    file_name="Instrumentos_bi_actualizado.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)#)

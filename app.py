import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# ============================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Tablero de Control - Instrumentos PRODUCE",
    page_icon="📊",
    layout="wide"
)
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
 
#    if "password_correcta" not in st.session_state:
#        st.text_input("Contraseña", type="password", on_change=password_ingresada, key="password")
#        return False
#    elif not st.session_state["password_correcta"]:
#        st.text_input("Contraseña", type="password", on_change=password_ingresada, key="password")
#        st.error("😕 Contraseña incorrecta")
#        return False
#    else:
#        return True
 
#if not verificar_password():
#    st.stop()
 
# ============================================================
# ENCABEZADO
# ============================================================
st.markdown("""
<div style="background-color:#1a2634; padding:20px; border-radius:12px; margin-bottom:15px;">
    <h1 style="color:white; font-family:Arial; margin:0;">📊 Tablero de Control - Instrumentos PRODUCE</h1>
    <p style="color:#bdc3c7; font-family:Arial; margin:5px 0 0 0;">Seguimiento de indicadores por instrumento</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# 1. CARGA DE DATOS
# ============================================================
# El archivo debe estar en el mismo repositorio de GitHub que app.py
NOMBRE_ARCHIVO = "Instrumentos bi.xlsx"  # nombre exacto tal como está en el repositorio

df = pd.read_excel(NOMBRE_ARCHIVO, sheet_name="BASE")

col_instrumento = 'Nombre del Instrumento'  # ajusta el nombre exacto si difiere

# ============================================================
# 2. LIMPIEZA DE DATOS
# ============================================================
df['Meta mod'] = pd.to_numeric(df['Meta global'], errors='coerce')
df['Avance acum mod'] = pd.to_numeric(df['Avance acum mod'], errors='coerce')

df['pct_cumplimiento'] = (df['Avance acum mod'] / df['Meta mod']) * 100
df['pct_cumplimiento'] = df['pct_cumplimiento'].clip(upper=150)

df['Indicador med'] = np.where(df['pct_cumplimiento'].isna(), 'No', 'Si')

df_filtrado = df[df['Indicador med'] == 'Si']

# ============================================================
# 3. KPIs GENERALES
# ============================================================
total_general = df[col_instrumento].nunique()
indicadores_general = df['Indicador mod'].dropna().shape[0]
indicadores_medibles = (df_filtrado['Indicador med'] == 'Si').sum()
avance_general = df_filtrado['pct_cumplimiento'].mean()


def tarjeta_kpi(titulo, valor, color="#2c3e50"):
    st.markdown(f"""
    <div style="background-color:{color}; color:white; padding:18px; border-radius:12px;
                text-align:center; font-family:Arial;
                box-shadow: 2px 2px 8px rgba(0,0,0,0.2);">
        <div style="font-size:13px; opacity:0.8;">{titulo}</div>
        <div style="font-size:32px; font-weight:bold;">{valor}</div>
    </div>
    """, unsafe_allow_html=True)


k1, k2, k3, k4 = st.columns(4)
with k1:
    tarjeta_kpi("Total Instrumentos", total_general)
with k2:
    tarjeta_kpi("Total Indicadores", indicadores_general)
with k3:
    tarjeta_kpi("Indicadores Medibles", indicadores_medibles, color="#16a085")
with k4:
    tarjeta_kpi("% Avance General", f"{avance_general:.1f}%", color="#34495e")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 4. GRÁFICO DE BARRAS
# ============================================================
df_grafico = df_filtrado[df_filtrado['Indicador med'] == 'Si']
conteo_entidad = df_grafico['Entidad que reporta'].value_counts().reset_index()
total_indicadores_grafico = len(df_grafico)

fig_barras = px.bar(conteo_entidad,
                     x='Entidad que reporta', y='count',
                     title='Cantidad de Indicadores por Área Responsable',
                     labels={'count': 'Cantidad', 'Entidad que reporta': 'Área Responsable'},
                     color='Entidad que reporta', text='count')

fig_barras.add_annotation(
    text=f"Total: {total_indicadores_grafico}",
    xref="paper", yref="paper",
    x=1, y=-0.15,
    showarrow=False,
    font=dict(size=13, color="gray"),
    align="right"
)
fig_barras.update_layout(margin=dict(b=100), height=400)
st.plotly_chart(fig_barras, use_container_width=True)

# --- Nota: instrumentos sin indicador/meta definida ---
instrumentos_sin_indicador = sorted(
    df[(df['Indicador med'] == 'No') & (df['unico'] == 1)][col_instrumento].dropna().unique().tolist()
)
 
lista_html = "".join([f"<li>{nombre}</li>" for nombre in instrumentos_sin_indicador])
 
st.markdown(f"""
<div style="background-color:#e8e8e8; border-radius:12px; padding:16px 20px; font-family:Arial; font-size:11px; color:#333;">
    <b>Nota:</b> Los siguientes instrumentos de gestión presentan acciones a seguir.
    Sin embargo, los indicadores y/o metas no se encuentran definidas o en su defecto los indicadores no han sido reportados.
    <ul style="margin:8px 0 0 0; padding-left:20px;">
        {lista_html}
    </ul>
</div>
""", unsafe_allow_html=True)
 
st.markdown("---")

# ============================================================
# 5. FILTRO POR INSTRUMENTO
# ============================================================
st.markdown("### 🔍 Detalle por Instrumento")
 
instrumentos = sorted(df_filtrado[col_instrumento].dropna().unique().tolist())
instrumentos.insert(0, 'Todos')
 
seleccion = st.selectbox("Instrumento:", instrumentos)
 
if seleccion == 'Todos':
    data = df_filtrado
else:
    data = df_filtrado[df_filtrado[col_instrumento] == seleccion]
 
cantidad_indicadores = data['Indicador mod'].dropna().shape[0]
avance_prom_filtrado = data['pct_cumplimiento'].mean()
 
if seleccion == 'Todos':
    texto_responsable = "-"
else:
    responsables = sorted(
        data['Responsable de realizar el seguimiento y evaluacion de indicadores'].dropna().unique().tolist()
    )
    texto_responsable = ", ".join(responsables) if responsables else "Sin dato"
 
col_tarjeta, col_responsable, col_gauge = st.columns([1, 1, 2])
 
with col_tarjeta:
    st.markdown(f"""
    <div style="background-color:#2c3e50; color:white; padding:20px; border-radius:12px;
                height:220px; text-align:center; font-family:Arial;
                box-shadow: 2px 2px 8px rgba(0,0,0,0.2); display:flex; flex-direction:column;
                justify-content:center;">
        <div style="font-size:14px; opacity:0.8;">N° de Indicadores</div>
        <div style="font-size:36px; font-weight:bold;">{cantidad_indicadores}</div>
    </div>
    """, unsafe_allow_html=True)
 
with col_responsable:
    st.markdown(f"""
    <div style="background-color:#2c3e50; color:white; padding:20px; border-radius:12px;
                height:220px; text-align:center; font-family:Arial;
                box-shadow: 2px 2px 8px rgba(0,0,0,0.2); display:flex; flex-direction:column;
                justify-content:center; overflow:hidden;">
        <div style="font-size:14px; opacity:0.8;">Responsable del seguimiento</div>
        <div style="font-size:16px; font-weight:bold; margin-top:8px;">{texto_responsable}</div>
    </div>
    """, unsafe_allow_html=True)
 
with col_gauge:
    if pd.isna(avance_prom_filtrado):
        st.info("Sin datos de avance disponibles para este instrumento.")
    else:
        fig_gauge = go.Figure(go.Indicator(
            mode="number+gauge",
            value=avance_prom_filtrado,
            title={"text": "% Avance Promedio"},
            number={'font': {'size': 50}},
            gauge={
                'axis': {'range': [0, 100]},
                'bar': {'color': "#2c3e50"},
                'steps': [
                    {'range': [0, 40], 'color': "#e74c3c"},
                    {'range': [40, 70], 'color': "#f1c40f"},
                    {'range': [70, 100], 'color': "#2ecc71"}
                ],
                'threshold': {
                    'line': {'color': "black", 'width': 4},
                    'thickness': 0.75,
                    'value': avance_prom_filtrado
                }
            }
        ))
        fig_gauge.update_layout(height=220, margin=dict(t=40, b=10, l=20, r=20))
        st.plotly_chart(fig_gauge, use_container_width=True)
 
st.markdown("---")
 
# ============================================================
# 6. TABLA DE INDICADORES
# ============================================================
st.markdown("### 📋 Tabla de Indicadores")
 
tabla = data[[
    'Entidad que reporta',
    'Indicador mod',
    'Meta global',
    'Avance acum mod',
    'pct_cumplimiento'
]].rename(columns={
    'Indicador mod': 'Indicador',
    'Avance acum mod': 'Avance',
    'pct_cumplimiento': '% Avance'
})
 
columnas_centradas = ['Meta global', 'Avance', '% Avance']
 
filas_html = ""
for _, fila in tabla.iterrows():
    filas_html += "<tr>"
    for col in tabla.columns:
        align = "center" if col in columnas_centradas else "left"
        valor = fila[col]
        if pd.isna(valor):
            valor = ""
        elif isinstance(valor, float):
            valor = f"{valor:.2f}"
        filas_html += f"<td style='text-align:{align}; padding:8px; border-bottom:1px solid #eee;'>{valor}</td>"
    filas_html += "</tr>"
 
encabezado_html = "".join([
    f"<th style='text-align:{'center' if col in columnas_centradas else 'left'}; padding:8px; background-color:#2c3e50; color:white;'>{col}</th>"
    for col in tabla.columns
])
 
st.markdown(f"""
<div style="overflow-x:auto;">
<table style="width:100%; border-collapse:collapse; font-family:Arial; font-size:13px;">
    <thead><tr>{encabezado_html}</tr></thead>
    <tbody>{filas_html}</tbody>
</table>
</div>
""", unsafe_allow_html=True)
# ============================================================
# 7. DESCARGAR BASE ACTUALIZADA
# ============================================================
#import io
#buffer = io.BytesIO()
#df.to_excel(buffer, index=False)
#st.download_button(
#    label="⬇️ Descargar Excel actualizado",
#    data=buffer.getvalue(),
#    file_name="Instrumentos_bi_actualizado.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)

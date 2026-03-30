# -*- coding: utf-8 -*-
"""
app.py — App Streamlit para analisis de ABC de Productos (Grupo Petri).
Soporta un solo PDF (analisis individual) o multiples PDFs (comparacion de sucursales).
"""
import io
import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from extractor import extract_pdf
from analyzer import analizar
from excel_report import generar_excel, generar_excel_comparacion
from pdf_report import generar_pdf, generar_pdf_comparacion

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title='Reportes Petri',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded',
)

C_AZUL   = '#1B3A5C'
C_DORADO = '#C8A84E'
C_VERDE  = '#27AE60'
C_ROJO   = '#E74C3C'

COLORES_SUCURSALES = [
    '#1B3A5C', '#C8A84E', '#27AE60', '#E74C3C',
    '#8E44AD', '#2E86C1', '#D35400', '#1ABC9C',
]

# ---------------------------------------------------------------------------
# Estilos globales
# ---------------------------------------------------------------------------
st.markdown("""
<style>
  [data-testid="stMetricValue"] { font-size: 1.35rem; font-weight: 700; }
  .section-title {
      background: #1B3A5C; color: white;
      padding: 8px 16px; border-radius: 6px;
      margin: 20px 0 12px 0; font-size: 1.05rem; font-weight: 700;
  }
  .comp-card {
      background: #F8F9FA; border: 1px solid #DDE; border-radius: 8px;
      padding: 14px 16px; margin-bottom: 8px;
  }
  div[data-testid="stDataFrame"] { font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Helpers de formato
# ---------------------------------------------------------------------------

def fmt_m(v):
    return f"$ {v:,.2f}"

def fmt_pct(v):
    return f"{v:.1f}%"

def _seccion(titulo):
    st.markdown(f'<div class="section-title">📌 {titulo}</div>', unsafe_allow_html=True)

def _nombre_sucursal(uploaded_file, meta):
    """Nombre de la sucursal: del nombre de archivo, luego del metadato."""
    nombre = os.path.splitext(uploaded_file.name)[0].strip()
    return nombre if nombre else meta.get('sucursal', 'Sucursal')


# ---------------------------------------------------------------------------
# Carga y cache
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def _cargar_pdf(pdf_bytes):
    import traceback
    buf = io.BytesIO(pdf_bytes)
    try:
        df, meta, totals = extract_pdf(buf, verbose=False)
    except Exception as e:
        raise RuntimeError(f"Error al leer el PDF: {type(e).__name__}: {e}\n{traceback.format_exc()}") from e
    return df, meta, totals


# ---------------------------------------------------------------------------
# Graficos comunes (modo individual)
# ---------------------------------------------------------------------------

def _grafico_barras_h(df, x_col, y_col, titulo, x_label, color=C_AZUL, n=15):
    df_plot = df.head(n).copy()
    df_plot['desc_corta'] = df_plot[y_col].str[:45]
    fig = px.bar(
        df_plot.iloc[::-1],
        x=x_col, y='desc_corta', orientation='h',
        title=titulo,
        labels={x_col: x_label, 'desc_corta': ''},
        color_discrete_sequence=[color],
    )
    fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white',
        font_family='Arial', title_font_size=14, title_font_color=C_AZUL,
        xaxis=dict(gridcolor='#EEEEEE', tickformat=',.0f'),
        yaxis=dict(tickfont=dict(size=10)),
        margin=dict(l=10, r=20, t=40, b=20), height=420,
    )
    return fig


def _grafico_pareto(pareto_df, n=50):
    df_plot = pareto_df.head(n).copy()
    df_plot['desc_corta'] = df_plot['descripcion'].str[:30]
    fig = go.Figure()
    fig.add_bar(x=df_plot['desc_corta'], y=df_plot['precio'],
                name='Precio venta', marker_color=C_AZUL, opacity=0.85)
    fig.add_scatter(x=df_plot['desc_corta'], y=df_plot['participacion_acum'],
                    name='Acumulado %', yaxis='y2',
                    line=dict(color=C_DORADO, width=2),
                    mode='lines+markers', marker=dict(size=4))
    fig.update_layout(
        title='Analisis Pareto — Top %d Productos' % n,
        title_font_color=C_AZUL,
        xaxis=dict(tickangle=-45, tickfont=dict(size=8)),
        yaxis=dict(title='Precio Venta ($)', tickformat='$,.0f', gridcolor='#EEEEEE'),
        yaxis2=dict(title='Acumulado %', overlaying='y', side='right',
                    range=[0, 105], tickformat='.0f', ticksuffix='%'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        plot_bgcolor='white', paper_bgcolor='white', font_family='Arial',
        height=420, margin=dict(l=10, r=10, t=50, b=80),
    )
    fig.add_hline(y=80, line_dash='dash', line_color=C_ROJO,
                  yref='y2', annotation_text='80%', annotation_position='right')
    return fig


def _grafico_dona(distribucion):
    labels = [d['rango'] for d in distribucion]
    values = [d['cantidad'] for d in distribucion]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.45,
        marker_colors=[C_ROJO, C_DORADO, C_VERDE, C_AZUL, '#8E44AD'],
        textinfo='percent+label', textfont_size=11,
    ))
    fig.update_layout(
        title='Distribucion por Rango de Margen', title_font_color=C_AZUL,
        showlegend=True, legend=dict(orientation='v', font=dict(size=10)),
        paper_bgcolor='white', font_family='Arial',
        height=380, margin=dict(l=10, r=10, t=50, b=10),
    )
    return fig


def _tabla_top(df):
    display = df.copy()
    display['descripcion'] = display['descripcion'].str[:50]
    fmt = {}
    for c in ['costo', 'precio', 'rentabilidad']:
        if c in display.columns: fmt[c] = fmt_m
    for c in ['margen', 'participacion']:
        if c in display.columns: fmt[c] = fmt_pct
    st.dataframe(display.style.format(fmt), use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# MODO INDIVIDUAL — reporte de una sola sucursal
# ---------------------------------------------------------------------------

def _modo_individual(uploaded, sucursal_nombre, df, meta, totals):

    if df.empty:
        st.error('No se pudieron extraer datos. Verificar que sea un ABC de Productos valido.')
        return

    analisis = analizar(df, meta, totals)
    res = analisis['resumen']

    fecha_d = meta.get('fecha_desde', '')
    fecha_h = meta.get('fecha_hasta', '')

    st.title(f'📊 Reporte de Ventas — {sucursal_nombre}')
    st.caption(f'Período: {fecha_d} al {fecha_h}  |  {res["cantidad_productos"]} productos')

    # --- Sección 1: Resumen ---
    _seccion('Resumen Ejecutivo')
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Total Ventas', fmt_m(res['total_ventas']))
    c2.metric('Total Costos', fmt_m(res['total_costo']))
    c3.metric('Rentabilidad', fmt_m(res['total_rentabilidad']))
    c4.metric('Margen Global', fmt_pct(res['margen_global']))

    c5, c6, c7 = st.columns(3)
    if res.get('mas_vendido'):
        mv = res['mas_vendido']
        c5.metric('Más vendido', mv['descripcion'][:35],
                  delta=f"{mv['unidades']:,.0f} un.", delta_color='off')
    if res.get('mas_rentable'):
        mr = res['mas_rentable']
        c6.metric('Más rentable', mr['descripcion'][:35],
                  delta=fmt_m(mr['rentabilidad']), delta_color='normal')
    if res.get('mas_facturado'):
        mf = res['mas_facturado']
        c7.metric('Mayor facturación', mf['descripcion'][:35],
                  delta=fmt_m(mf['precio']), delta_color='normal')

    # --- Sección 2: Top Unidades ---
    _seccion('Top 15 — Productos más Vendidos (unidades)')
    ct, cg = st.columns([1, 1.4])
    with ct: _tabla_top(analisis['top_unidades'])
    with cg:
        st.plotly_chart(_grafico_barras_h(
            analisis['top_unidades'], 'unidades', 'descripcion',
            'Top 15 por Unidades', 'Unidades', C_AZUL), use_container_width=True)

    # --- Sección 3: Top Rentabilidad ---
    _seccion('Top 15 — Productos por Rentabilidad ($)')
    ct, cg = st.columns([1, 1.4])
    with ct: _tabla_top(analisis['top_rentabilidad'])
    with cg:
        st.plotly_chart(_grafico_barras_h(
            analisis['top_rentabilidad'], 'rentabilidad', 'descripcion',
            'Top 15 por Rentabilidad', 'Rentabilidad ($)', C_VERDE), use_container_width=True)

    # --- Sección 4: Top Precio ---
    _seccion('Top 15 — Productos por Precio de Venta Total')
    ct, cg = st.columns([1, 1.4])
    with ct: _tabla_top(analisis['top_precio'])
    with cg:
        st.plotly_chart(_grafico_barras_h(
            analisis['top_precio'], 'precio', 'descripcion',
            'Top 15 por Facturación', 'Precio ($)', C_DORADO), use_container_width=True)

    # --- Sección 5: Pareto ---
    _seccion('Análisis Pareto — Regla 80/20')
    ps = analisis['pareto_stats']
    n80, pct80, tot = ps.get('n_productos_80pct', 0), ps.get('pct_productos_para_80', 0), ps.get('total_productos', 0)
    st.info(f'**{n80} productos** ({pct80:.1f}% del total) generan el **80% de la facturación**. '
            f'Los {tot - n80} restantes ({100 - pct80:.1f}%) generan el 20% restante.')
    st.plotly_chart(_grafico_pareto(analisis['pareto_df'], n=min(60, tot)), use_container_width=True)

    # --- Sección 6: Distribución margen ---
    _seccion('Distribución por Rango de Margen')
    cg, ct = st.columns([1.3, 1])
    with cg: st.plotly_chart(_grafico_dona(analisis['distribucion_margen']), use_container_width=True)
    with ct:
        df_dist = pd.DataFrame(analisis['distribucion_margen'])[
            ['rango', 'cantidad', 'pct_cantidad', 'precio_total', 'pct_ventas']]
        df_dist.columns = ['Rango', 'Productos', '% Prod.', 'Ventas ($)', '% Ventas']
        st.dataframe(df_dist.style.format(
            {'% Prod.': '{:.1f}%', 'Ventas ($)': '$ {:,.2f}', '% Ventas': '{:.1f}%'}),
            use_container_width=True, hide_index=True)

    # --- Sección 7: Tabla completa ---
    _seccion('Tabla Completa de Productos')
    df_full = analisis['df_completo'].copy()
    c_busq, c_filt = st.columns([2, 1])
    with c_busq:
        busqueda = st.text_input('🔍 Buscar (descripción o código)', key='busq_ind')
    with c_filt:
        filtro = st.selectbox('Filtrar por margen', ['Todos', 'Solo positivos', '> 100%'], key='filt_ind')
    if busqueda:
        mask = (df_full['descripcion'].str.contains(busqueda, case=False, na=False) |
                df_full['codigo'].str.contains(busqueda, case=False, na=False))
        df_full = df_full[mask]
    if filtro == 'Solo positivos': df_full = df_full[df_full['margen'] >= 0]
    elif filtro == '> 100%': df_full = df_full[df_full['margen'] > 100]

    st.caption(f'Mostrando {len(df_full)} de {len(analisis["df_completo"])} productos')
    st.dataframe(df_full.style.format({
        'unidades': '{:,.2f}', 'costo': '$ {:,.2f}', 'precio': '$ {:,.2f}',
        'rentabilidad': '$ {:,.2f}', 'margen': '{:.2f}%', 'participacion': '{:.2f}%',
    }), use_container_width=True, hide_index=True, height=420)

    # --- Sección 8: Exportar ---
    _seccion('📥 Exportar Reporte')
    nombre_base = f'reporte_{sucursal_nombre.replace(" ", "_")}_{fecha_d.replace("/", "-")}_{fecha_h.replace("/", "-")}'

    col_xls, col_pdf = st.columns(2)

    with col_xls:
        st.write('**Excel** — 5 hojas: Resumen, Datos Completos, Top Vendidos, Top Rentabilidad, Análisis Pareto.')
        excel_bytes = generar_excel(analisis, meta)
        st.download_button('⬇️ Descargar Excel', data=excel_bytes,
                           file_name=nombre_base + '.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           use_container_width=True, type='primary')

    with col_pdf:
        st.write('**PDF** — Carátula, índice, KPIs, gráficos en hoja completa y tabla de datos.')
        with st.spinner('Generando PDF...'):
            pdf_bytes = generar_pdf(analisis, meta)
        st.download_button('⬇️ Descargar PDF', data=pdf_bytes,
                           file_name=nombre_base + '.pdf',
                           mime='application/pdf',
                           use_container_width=True, type='primary')

    with st.expander('🔍 Validación de extracción'):
        st.json({'productos': int(len(df)),
                 'sum_costo': float(round(df['costo'].sum(), 2)),
                 'sum_precio': float(round(df['precio'].sum(), 2)),
                 'totals_pdf': {k: float(v) for k, v in totals.items()}})


# ---------------------------------------------------------------------------
# MODO COMPARACION — multiples sucursales
# ---------------------------------------------------------------------------

def _graf_comp_barras(df_comp, y_col, titulo, y_label, fmt_tick='$,.0f'):
    """Barra horizontal comparando sucursales para una metrica."""
    fig = px.bar(
        df_comp.sort_values(y_col),
        x=y_col, y='sucursal', orientation='h',
        title=titulo,
        labels={y_col: y_label, 'sucursal': ''},
        color='sucursal',
        color_discrete_sequence=COLORES_SUCURSALES,
        text=y_col,
    )
    fig.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
    fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white', font_family='Arial',
        title_font_color=C_AZUL, showlegend=False,
        xaxis=dict(tickformat=fmt_tick, gridcolor='#EEEEEE'),
        margin=dict(l=10, r=40, t=40, b=20), height=max(280, len(df_comp) * 55 + 80),
    )
    return fig


def _graf_comp_agrupado(df_comp, cols, titulo):
    """Barras agrupadas: ventas, costo, rentabilidad por sucursal."""
    fig = go.Figure()
    labels = {'total_ventas': 'Ventas', 'total_costo': 'Costo', 'total_rentabilidad': 'Rentabilidad'}
    colors = [C_AZUL, C_ROJO, C_VERDE]
    for col, color in zip(cols, colors):
        fig.add_bar(name=labels.get(col, col), x=df_comp['sucursal'],
                    y=df_comp[col], marker_color=color,
                    text=df_comp[col].apply(lambda v: f'${v/1e6:.1f}M'),
                    textposition='outside')
    fig.update_layout(
        barmode='group',
        plot_bgcolor='white', paper_bgcolor='white', font_family='Arial',
        yaxis=dict(tickformat='$,.0f', gridcolor='#EEEEEE'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0.0, xanchor='left'),
        margin=dict(l=10, r=10, t=60, b=20),
        height=420,
    )
    return fig


def _graf_comp_margen(df_comp):
    """Barras de margen global por sucursal."""
    fig = px.bar(
        df_comp.sort_values('margen_global'),
        x='margen_global', y='sucursal', orientation='h',
        title='Margen Global por Sucursal (%)',
        labels={'margen_global': 'Margen Global (%)', 'sucursal': ''},
        color='sucursal', color_discrete_sequence=COLORES_SUCURSALES,
        text='margen_global',
    )
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white', font_family='Arial',
        title_font_color=C_AZUL, showlegend=False,
        xaxis=dict(ticksuffix='%', gridcolor='#EEEEEE'),
        margin=dict(l=10, r=60, t=40, b=20),
        height=max(280, len(df_comp) * 55 + 80),
    )
    return fig


def _graf_top_productos_comp(datos, metrica, label, n=10):
    """
    Grafico de top productos comparando sucursales.
    Muestra los N productos mas relevantes del total combinado, con barra por sucursal.
    """
    # Unir todos los DataFrames con columna 'sucursal'
    dfs = []
    for suc, an in datos.items():
        top = an['df_completo'].nlargest(50, metrica)[['descripcion', metrica]].copy()
        top['sucursal'] = suc
        dfs.append(top)
    df_all = pd.concat(dfs, ignore_index=True)

    # Top N productos por suma total
    top_productos = (df_all.groupby('descripcion')[metrica].sum()
                     .nlargest(n).index.tolist())
    df_plot = df_all[df_all['descripcion'].isin(top_productos)].copy()
    df_plot['desc_corta'] = df_plot['descripcion'].str[:35]

    fig = px.bar(
        df_plot, x=metrica, y='desc_corta', color='sucursal',
        orientation='h', barmode='group',
        labels={metrica: label, 'desc_corta': '', 'sucursal': 'Sucursal'},
        color_discrete_sequence=COLORES_SUCURSALES,
    )
    fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white', font_family='Arial',
        xaxis=dict(tickformat=',.0f', gridcolor='#EEEEEE'),
        yaxis=dict(tickfont=dict(size=9)),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0.0, xanchor='left'),
        margin=dict(l=10, r=20, t=60, b=20),
        height=max(400, n * 45 + 100),
    )
    return fig


def _graf_dist_margen_comp(datos):
    """Barras apiladas de distribucion de margen por sucursal."""
    rangos = ['< 0% (perdida)', '0% - 30%', '30% - 60%', '60% - 100%', '> 100%']
    colors = [C_ROJO, C_DORADO, '#F0B429', C_VERDE, C_AZUL]
    fig = go.Figure()
    for rango, color in zip(rangos, colors):
        y_vals = []
        for suc, an in datos.items():
            dist = {d['rango']: d['pct_cantidad'] for d in an['distribucion_margen']}
            y_vals.append(dist.get(rango, 0))
        fig.add_bar(name=rango, x=list(datos.keys()), y=y_vals,
                    marker_color=color, text=[f'{v:.0f}%' for v in y_vals],
                    textposition='inside')
    fig.update_layout(
        barmode='stack',
        plot_bgcolor='white', paper_bgcolor='white', font_family='Arial',
        yaxis=dict(ticksuffix='%', gridcolor='#EEEEEE'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0.0, xanchor='left'),
        margin=dict(l=10, r=10, t=80, b=20), height=420,
    )
    return fig


def _modo_comparacion(archivos_datos):
    """
    archivos_datos: lista de (uploaded_file, df, meta, totals)
    """
    # Construir analisis por sucursal
    datos = {}   # {nombre_sucursal: analisis_dict}
    metas = {}
    for uploaded, df, meta, totals in archivos_datos:
        nombre = _nombre_sucursal(uploaded, meta)
        datos[nombre] = analizar(df, meta, totals)
        metas[nombre] = meta

    sucursales = list(datos.keys())
    periodo = metas[sucursales[0]].get('fecha_desde', '') + ' al ' + metas[sucursales[0]].get('fecha_hasta', '')

    # Tabla resumen para graficos
    filas_comp = []
    for suc, an in datos.items():
        r = an['resumen']
        filas_comp.append({
            'sucursal': suc,
            'total_ventas': r['total_ventas'],
            'total_costo': r['total_costo'],
            'total_rentabilidad': r['total_rentabilidad'],
            'margen_global': r['margen_global'],
            'cantidad_productos': r['cantidad_productos'],
            'cantidad_activos': r['cantidad_activos'],
            'pareto_n80': an['pareto_stats'].get('n_productos_80pct', 0),
            'pareto_pct80': an['pareto_stats'].get('pct_productos_para_80', 0),
        })
    df_comp = pd.DataFrame(filas_comp)

    # ================================================================
    # TABS: Comparacion + una por sucursal
    # ================================================================
    tab_names = ['🏢 Comparación General'] + [f'📍 {s}' for s in sucursales]
    tabs = st.tabs(tab_names)

    # ----------------------------------------------------------------
    # TAB 0 — Comparacion General
    # ----------------------------------------------------------------
    with tabs[0]:
        st.title(f'📊 Comparación de Sucursales — Grupo Petri')
        st.caption(f'Período: {periodo}  |  {len(sucursales)} sucursales cargadas')

        # -- Métricas resumen --
        _seccion('Resumen por Sucursal')
        tabla_res = df_comp[['sucursal', 'total_ventas', 'total_costo',
                              'total_rentabilidad', 'margen_global', 'cantidad_activos']].copy()
        tabla_res.columns = ['Sucursal', 'Ventas ($)', 'Costo ($)',
                              'Rentabilidad ($)', 'Margen Global', 'Productos activos']

        # Totales
        totales_row = {
            'Sucursal': '⚡ TOTAL',
            'Ventas ($)': df_comp['total_ventas'].sum(),
            'Costo ($)': df_comp['total_costo'].sum(),
            'Rentabilidad ($)': df_comp['total_rentabilidad'].sum(),
            'Margen Global': (df_comp['total_rentabilidad'].sum() /
                              abs(df_comp['total_costo'].sum()) * 100
                              if df_comp['total_costo'].sum() != 0 else 0),
            'Productos activos': df_comp['cantidad_activos'].sum(),
        }
        tabla_res = pd.concat([tabla_res, pd.DataFrame([totales_row])], ignore_index=True)

        st.dataframe(
            tabla_res.style
              .format({'Ventas ($)': '$ {:,.2f}', 'Costo ($)': '$ {:,.2f}',
                       'Rentabilidad ($)': '$ {:,.2f}', 'Margen Global': '{:.2f}%',
                       'Productos activos': '{:,.0f}'})
              .apply(lambda row: ['font-weight:bold; background:#EAF4FB'] * len(row)
                     if row['Sucursal'] == '⚡ TOTAL' else [''] * len(row), axis=1),
            use_container_width=True, hide_index=True,
        )

        # -- Gráficos de métricas principales --
        _seccion('Ventas, Costo y Rentabilidad por Sucursal')
        st.plotly_chart(
            _graf_comp_agrupado(df_comp, ['total_ventas', 'total_costo', 'total_rentabilidad'],
                                'Comparación: Ventas / Costo / Rentabilidad'),
            use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(_graf_comp_barras(df_comp, 'total_ventas',
                'Ventas Totales por Sucursal ($)', 'Ventas ($)'),
                use_container_width=True)
        with c2:
            st.plotly_chart(_graf_comp_margen(df_comp), use_container_width=True)

        # -- Top productos comparados --
        _seccion('Top 10 Productos por Facturación — Comparación entre Sucursales')
        st.plotly_chart(_graf_top_productos_comp(datos, 'precio', 'Precio Venta ($)'),
                        use_container_width=True)

        _seccion('Top 10 Productos por Rentabilidad — Comparación entre Sucursales')
        st.plotly_chart(_graf_top_productos_comp(datos, 'rentabilidad', 'Rentabilidad ($)'),
                        use_container_width=True)

        # -- Distribución de márgenes --
        _seccion('Distribución de Márgenes por Sucursal')
        c1, c2 = st.columns([1.6, 1])
        with c1:
            st.plotly_chart(_graf_dist_margen_comp(datos), use_container_width=True)
        with c2:
            # Tabla Pareto comparativa
            pareto_comp = df_comp[['sucursal', 'pareto_n80', 'pareto_pct80']].copy()
            pareto_comp.columns = ['Sucursal', 'Prods. para 80% ventas', '% del catálogo']
            st.caption('**Análisis Pareto por Sucursal**')
            st.dataframe(pareto_comp.style.format({'% del catálogo': '{:.1f}%'}),
                         use_container_width=True, hide_index=True)

        # -- Exportar --
        _seccion('📥 Exportar Reporte Comparativo')
        fecha_d0 = metas[sucursales[0]].get('fecha_desde', '').replace('/', '-')
        fecha_h0 = metas[sucursales[0]].get('fecha_hasta', '').replace('/', '-')
        nombre_base = f'comparacion_sucursales_{fecha_d0}_{fecha_h0}'

        col_xls, col_pdf = st.columns(2)

        with col_xls:
            st.write('**Excel** — Hoja de comparación general + resumen individual por sucursal.')
            with st.spinner('Generando Excel...'):
                xls_comp = generar_excel_comparacion(datos, metas, df_comp)
            st.download_button(
                '⬇️ Descargar Excel Comparativo',
                data=xls_comp,
                file_name=nombre_base + '.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True, type='primary',
            )

        with col_pdf:
            st.write('**PDF** — Carátula, índice, comparación general y sección detallada por sucursal.')
            with st.spinner('Generando PDF (puede tardar unos segundos)...'):
                pdf_comp = generar_pdf_comparacion(datos, metas, df_comp)
            st.download_button(
                '⬇️ Descargar PDF Comparativo',
                data=pdf_comp,
                file_name=nombre_base + '.pdf',
                mime='application/pdf',
                use_container_width=True, type='primary',
            )

    # ----------------------------------------------------------------
    # TABS individuales por sucursal
    # ----------------------------------------------------------------
    for i, (suc, tab) in enumerate(zip(sucursales, tabs[1:])):
        an = datos[suc]
        meta = metas[suc]
        res = an['resumen']
        color_suc = COLORES_SUCURSALES[i % len(COLORES_SUCURSALES)]

        with tab:
            st.subheader(f'📍 {suc}')
            st.caption(f"Período: {meta.get('fecha_desde','')} al {meta.get('fecha_hasta','')} "
                       f"| {res['cantidad_productos']} productos")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric('Ventas', fmt_m(res['total_ventas']))
            c2.metric('Costos', fmt_m(res['total_costo']))
            c3.metric('Rentabilidad', fmt_m(res['total_rentabilidad']))
            c4.metric('Margen', fmt_pct(res['margen_global']))

            _seccion('Top 15 por Unidades')
            ct, cg = st.columns([1, 1.4])
            with ct: _tabla_top(an['top_unidades'])
            with cg:
                st.plotly_chart(_grafico_barras_h(
                    an['top_unidades'], 'unidades', 'descripcion',
                    f'Top 15 por Unidades — {suc}', 'Unidades', color_suc),
                    use_container_width=True)

            _seccion('Top 15 por Rentabilidad ($)')
            ct, cg = st.columns([1, 1.4])
            with ct: _tabla_top(an['top_rentabilidad'])
            with cg:
                st.plotly_chart(_grafico_barras_h(
                    an['top_rentabilidad'], 'rentabilidad', 'descripcion',
                    f'Top 15 por Rentabilidad — {suc}', 'Rentabilidad ($)', color_suc),
                    use_container_width=True)

            _seccion('Top 15 por Facturación ($)')
            ct, cg = st.columns([1, 1.4])
            with ct: _tabla_top(an['top_precio'])
            with cg:
                st.plotly_chart(_grafico_barras_h(
                    an['top_precio'], 'precio', 'descripcion',
                    f'Top 15 por Facturación — {suc}', 'Precio ($)', color_suc),
                    use_container_width=True)

            _seccion('Análisis Pareto')
            ps = an['pareto_stats']
            n80, pct80, tot = ps.get('n_productos_80pct',0), ps.get('pct_productos_para_80',0), ps.get('total_productos',0)
            st.info(f'**{n80} productos** ({pct80:.1f}%) generan el 80% de la facturación.')
            st.plotly_chart(_grafico_pareto(an['pareto_df'], n=min(60, tot)),
                            use_container_width=True)

            _seccion('Distribución por Rango de Margen')
            cg, ct = st.columns([1.3, 1])
            with cg: st.plotly_chart(_grafico_dona(an['distribucion_margen']), use_container_width=True)
            with ct:
                df_dist = pd.DataFrame(an['distribucion_margen'])[
                    ['rango', 'cantidad', 'pct_cantidad', 'precio_total', 'pct_ventas']]
                df_dist.columns = ['Rango', 'Prods.', '% Prods.', 'Ventas ($)', '% Ventas']
                st.dataframe(df_dist.style.format(
                    {'% Prods.': '{:.1f}%', 'Ventas ($)': '$ {:,.2f}', '% Ventas': '{:.1f}%'}),
                    use_container_width=True, hide_index=True)

            _seccion('Tabla Completa')
            df_full = an['df_completo'].copy()
            c_b, c_f = st.columns([2, 1])
            with c_b:
                busq = st.text_input('🔍 Buscar', key=f'busq_{suc}')
            with c_f:
                filt = st.selectbox('Filtrar margen', ['Todos', 'Solo positivos', '> 100%'], key=f'filt_{suc}')
            if busq:
                mask = (df_full['descripcion'].str.contains(busq, case=False, na=False) |
                        df_full['codigo'].str.contains(busq, case=False, na=False))
                df_full = df_full[mask]
            if filt == 'Solo positivos': df_full = df_full[df_full['margen'] >= 0]
            elif filt == '> 100%': df_full = df_full[df_full['margen'] > 100]
            st.caption(f'{len(df_full)} de {len(an["df_completo"])} productos')
            st.dataframe(df_full.style.format({
                'unidades': '{:,.2f}', 'costo': '$ {:,.2f}', 'precio': '$ {:,.2f}',
                'rentabilidad': '$ {:,.2f}', 'margen': '{:.2f}%', 'participacion': '{:.2f}%',
            }), use_container_width=True, hide_index=True, height=380)

            _seccion(f'📥 Exportar {suc}')
            excel_ind = generar_excel(an, meta)
            st.download_button(
                f'⬇️ Descargar Excel — {suc}',
                data=excel_ind,
                file_name=f'reporte_{suc.replace(" ","_")}_{meta.get("fecha_desde","").replace("/","-")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True, key=f'dl_{suc}',
            )


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

def _sidebar():
    with st.sidebar:
        st.markdown(
            f'<div style="background:{C_AZUL};color:white;padding:14px;'
            f'border-radius:8px;text-align:center;font-weight:700;'
            f'font-size:1.1rem;letter-spacing:1px;">GRUPO PETRI</div>',
            unsafe_allow_html=True,
        )
        st.markdown('---')
        st.subheader('📂 Cargar PDF(s)')
        st.caption('Un PDF → reporte individual. Varios PDFs → comparación de sucursales.')
        archivos = st.file_uploader(
            'ABC de Productos (.pdf)',
            type=['pdf'],
            accept_multiple_files=True,
            help='El nombre del archivo se usa como nombre de sucursal.',
        )
        if archivos:
            st.markdown('**Archivos cargados:**')
            for f in archivos:
                nombre = os.path.splitext(f.name)[0]
                st.markdown(f'- 📄 {nombre}')
        st.markdown('---')
        st.caption('Generador de Reportes v1.1')
    return archivos


# ---------------------------------------------------------------------------
# Pantalla de bienvenida
# ---------------------------------------------------------------------------

def _pantalla_inicio():
    st.markdown(f"""
    <div style="text-align:center; padding: 60px 20px;">
      <h1 style="color:{C_AZUL}; font-size:2.2rem;">📊 Generador de Reportes</h1>
      <h3 style="color:#555; font-weight:400;">Grupo Petri — ABC de Productos</h3>
      <hr style="border-color:{C_DORADO}; width:200px; margin:20px auto;">
      <p style="color:#666; font-size:1rem;">
        <b>Un PDF</b> → Reporte individual completo de la sucursal.<br><br>
        <b>Varios PDFs</b> → Comparación automática entre sucursales<br>
        con tabs individuales por cada una.<br><br>
        El nombre del archivo se usa como nombre de la sucursal.
      </p>
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    archivos = _sidebar()

    if not archivos:
        _pantalla_inicio()
        return

    # Cargar todos los PDFs
    archivos_datos = []
    errores = []
    progress = st.progress(0, text='Procesando PDFs...')
    for i, f in enumerate(archivos):
        progress.progress((i + 1) / len(archivos), text=f'Procesando {f.name}...')
        df, meta, totals = _cargar_pdf(f.read())
        if df.empty:
            errores.append(f.name)
        else:
            archivos_datos.append((f, df, meta, totals))
    progress.empty()

    if errores:
        st.warning(f'No se pudo procesar: {", ".join(errores)}')

    if not archivos_datos:
        st.error('No se pudo extraer datos de ningún PDF.')
        return

    if len(archivos_datos) == 1:
        uploaded, df, meta, totals = archivos_datos[0]
        sucursal = _nombre_sucursal(uploaded, meta)
        _modo_individual(uploaded, sucursal, df, meta, totals)
    else:
        _modo_comparacion(archivos_datos)


if __name__ == '__main__':
    main()

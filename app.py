# -*- coding: utf-8 -*-
"""
app.py — App Streamlit para analisis de ABC de Productos (Grupo Petri).
"""
import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from extractor import extract_pdf
from analyzer import analizar
from excel_report import generar_excel

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title='Reportes Petri',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded',
)

# Paleta
C_AZUL    = '#1B3A5C'
C_DORADO  = '#C8A84E'
C_VERDE   = '#27AE60'
C_ROJO    = '#E74C3C'
C_FONDO   = '#F8F9FA'

# ---------------------------------------------------------------------------
# Estilos globales
# ---------------------------------------------------------------------------
st.markdown("""
<style>
  [data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: 700; }
  .section-title {
      background: #1B3A5C; color: white;
      padding: 8px 16px; border-radius: 6px;
      margin: 20px 0 12px 0; font-size: 1.05rem; font-weight: 700;
  }
  .negative-alert {
      background: #FDECEA; border-left: 4px solid #E74C3C;
      padding: 10px 16px; border-radius: 4px; margin: 8px 0;
  }
  .ok-banner {
      background: #D5F5E3; border-left: 4px solid #27AE60;
      padding: 10px 16px; border-radius: 4px;
  }
  div[data-testid="stDataFrame"] { font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Helpers de formato
# ---------------------------------------------------------------------------

def fmt_moneda(v):
    return f"$ {v:,.2f}"


def fmt_pct(v):
    return f"{v:.1f}%"


def _seccion(titulo):
    st.markdown(f'<div class="section-title">📌 {titulo}</div>', unsafe_allow_html=True)


def _tabla_top(df, value_col, value_fmt):
    """Muestra un DataFrame de top-N con formateo."""
    display = df.copy()
    display['descripcion'] = display['descripcion'].str[:50]
    fmt_cols = {}
    for c in ['costo', 'precio', 'rentabilidad']:
        if c in display.columns:
            fmt_cols[c] = fmt_moneda
    for c in ['margen', 'participacion']:
        if c in display.columns:
            fmt_cols[c] = fmt_pct
    if value_col in fmt_cols:
        pass  # ya incluido
    st.dataframe(
        display.style.format(fmt_cols),
        use_container_width=True,
        hide_index=True,
    )


# ---------------------------------------------------------------------------
# Graficos Plotly
# ---------------------------------------------------------------------------

def _grafico_barras_h(df, x_col, y_col, titulo, x_label, color=C_AZUL, n=15):
    """Grafico de barras horizontal con los top N."""
    df_plot = df.head(n).copy()
    df_plot['desc_corta'] = df_plot[y_col].str[:45]
    fig = px.bar(
        df_plot.iloc[::-1],  # invertir para que el mayor quede arriba
        x=x_col,
        y='desc_corta',
        orientation='h',
        title=titulo,
        labels={x_col: x_label, 'desc_corta': ''},
        color_discrete_sequence=[color],
    )
    fig.update_layout(
        plot_bgcolor='white',
        paper_bgcolor='white',
        font_family='Arial',
        title_font_size=14,
        title_font_color=C_AZUL,
        xaxis=dict(gridcolor='#EEEEEE', tickformat=',.0f'),
        yaxis=dict(tickfont=dict(size=10)),
        margin=dict(l=10, r=20, t=40, b=20),
        height=420,
    )
    return fig


def _grafico_pareto(pareto_df, n=50):
    """Grafico de barras + linea acumulativa (Pareto)."""
    df_plot = pareto_df.head(n).copy()
    df_plot['desc_corta'] = df_plot['descripcion'].str[:30]

    fig = go.Figure()
    fig.add_bar(
        x=df_plot['desc_corta'],
        y=df_plot['precio'],
        name='Precio venta',
        marker_color=C_AZUL,
        opacity=0.85,
    )
    fig.add_scatter(
        x=df_plot['desc_corta'],
        y=df_plot['participacion_acum'],
        name='Acumulado %',
        yaxis='y2',
        line=dict(color=C_DORADO, width=2),
        mode='lines+markers',
        marker=dict(size=4),
    )
    fig.update_layout(
        title='Analisis Pareto — Top %d Productos' % n,
        title_font_color=C_AZUL,
        xaxis=dict(tickangle=-45, tickfont=dict(size=8)),
        yaxis=dict(title='Precio Venta ($)', tickformat='$,.0f', gridcolor='#EEEEEE'),
        yaxis2=dict(title='Acumulado %', overlaying='y', side='right',
                    range=[0, 105], tickformat='.0f', ticksuffix='%'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        plot_bgcolor='white',
        paper_bgcolor='white',
        font_family='Arial',
        height=420,
        margin=dict(l=10, r=10, t=50, b=80),
    )
    # Linea de referencia 80%
    fig.add_hline(y=80, line_dash='dash', line_color=C_ROJO,
                  yref='y2', annotation_text='80%',
                  annotation_position='right')
    return fig


def _grafico_dona(distribucion):
    """Grafico de dona por rangos de margen."""
    labels = [d['rango'] for d in distribucion]
    values = [d['cantidad'] for d in distribucion]
    colors = [C_ROJO, C_DORADO, C_VERDE, C_AZUL, '#8E44AD']
    fig = go.Figure(go.Pie(
        labels=labels,
        values=values,
        hole=0.45,
        marker_colors=colors,
        textinfo='percent+label',
        textfont_size=11,
    ))
    fig.update_layout(
        title='Distribucion por Rango de Margen',
        title_font_color=C_AZUL,
        showlegend=True,
        legend=dict(orientation='v', font=dict(size=10)),
        paper_bgcolor='white',
        font_family='Arial',
        height=380,
        margin=dict(l=10, r=10, t=50, b=10),
    )
    return fig


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
        st.subheader('📂 Cargar PDF')
        uploaded = st.file_uploader(
            'Subir ABC de Productos (.pdf)',
            type=['pdf'],
            help='PDF generado por el sistema de Grupo Petri',
        )
        st.markdown('---')
        st.caption('Generador de Reportes v1.0')
    return uploaded


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
        Subí el PDF de <b>ABC de Productos</b> desde el panel lateral<br>
        para generar el reporte completo con análisis visual y exportación a Excel.
      </p>
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# App principal
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner='Extrayendo datos del PDF...')
def _cargar_pdf(pdf_bytes):
    buf = io.BytesIO(pdf_bytes)
    df, meta, totals = extract_pdf(buf, verbose=False)
    return df, meta, totals


def main():
    uploaded = _sidebar()

    if uploaded is None:
        _pantalla_inicio()
        return

    # ------- Carga y extraccion -------
    with st.spinner('Procesando PDF...'):
        df, meta, totals = _cargar_pdf(uploaded.read())

    if df.empty:
        st.error('No se pudieron extraer datos del PDF. Verificar que sea un ABC de Productos valido.')
        return

    analisis = analizar(df, meta, totals)
    res = analisis['resumen']

    # ------- Encabezado del reporte -------
    sucursal = meta.get('sucursal', 'Sucursal')
    fecha_d  = meta.get('fecha_desde', '')
    fecha_h  = meta.get('fecha_hasta', '')
    st.title(f'📊 Reporte de Ventas — {sucursal}')
    st.caption(f'Período: {fecha_d} al {fecha_h}  |  {res["cantidad_productos"]} productos en el sistema')

    # ================================================================
    # SECCION 1 — Resumen Ejecutivo
    # ================================================================
    _seccion('Resumen Ejecutivo')
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Total Ventas', fmt_moneda(res['total_ventas']))
    c2.metric('Total Costos', fmt_moneda(res['total_costo']))
    c3.metric('Rentabilidad Total', fmt_moneda(res['total_rentabilidad']))
    c4.metric('Margen Global', fmt_pct(res['margen_global']))

    c5, c6, c7 = st.columns(3)
    if res.get('mas_vendido'):
        mv = res['mas_vendido']
        c5.metric('Producto mas vendido', mv['descripcion'][:35],
                  delta=f"{mv['unidades']:,.0f} unidades",
                  delta_color='off')
    if res.get('mas_rentable'):
        mr = res['mas_rentable']
        c6.metric('Producto mas rentable', mr['descripcion'][:35],
                  delta=fmt_moneda(mr['rentabilidad']),
                  delta_color='normal')
    if res.get('mas_facturado'):
        mf = res['mas_facturado']
        c7.metric('Mayor facturacion', mf['descripcion'][:35],
                  delta=fmt_moneda(mf['precio']),
                  delta_color='normal')

    # ================================================================
    # SECCION 2 — Top 15 por Unidades
    # ================================================================
    _seccion('Top 15 — Productos mas Vendidos (por unidades)')
    top_u = analisis['top_unidades']
    col_t, col_g = st.columns([1, 1.4])
    with col_t:
        _tabla_top(top_u, 'unidades', fmt_moneda)
    with col_g:
        fig = _grafico_barras_h(top_u, 'unidades', 'descripcion',
                                'Top 15 Productos por Unidades Vendidas',
                                'Unidades Vendidas', color=C_AZUL)
        st.plotly_chart(fig, use_container_width=True)

    # ================================================================
    # SECCION 3 — Top 15 por Rentabilidad
    # ================================================================
    _seccion('Top 15 — Productos por Rentabilidad ($)')
    top_r = analisis['top_rentabilidad']
    col_t, col_g = st.columns([1, 1.4])
    with col_t:
        _tabla_top(top_r, 'rentabilidad', fmt_moneda)
    with col_g:
        fig = _grafico_barras_h(top_r, 'rentabilidad', 'descripcion',
                                'Top 15 Productos por Rentabilidad',
                                'Rentabilidad ($)', color=C_VERDE)
        st.plotly_chart(fig, use_container_width=True)

    # ================================================================
    # SECCION 4 — Top 15 por Precio de Venta
    # ================================================================
    _seccion('Top 15 — Productos por Precio de Venta Total')
    top_p = analisis['top_precio']
    col_t, col_g = st.columns([1, 1.4])
    with col_t:
        _tabla_top(top_p, 'precio', fmt_moneda)
    with col_g:
        fig = _grafico_barras_h(top_p, 'precio', 'descripcion',
                                'Top 15 Productos por Facturacion',
                                'Precio de Venta ($)', color=C_DORADO)
        st.plotly_chart(fig, use_container_width=True)

    # ================================================================
    # SECCION 5 — Pareto
    # ================================================================
    _seccion('Analisis Pareto — Regla 80/20')
    pstats = analisis['pareto_stats']
    n80 = pstats.get('n_productos_80pct', 0)
    pct80 = pstats.get('pct_productos_para_80', 0)
    total_p = pstats.get('total_productos', 0)

    st.info(
        f'**{n80} productos** ({pct80:.1f}% del total) generan el **80% de la facturación**. '
        f'Los {total_p - n80} restantes ({100 - pct80:.1f}%) generan el 20% restante.'
    )
    fig_pareto = _grafico_pareto(analisis['pareto_df'], n=min(60, total_p))
    st.plotly_chart(fig_pareto, use_container_width=True)

    # ================================================================
    # SECCION 6 — Margen Negativo
    # ================================================================
    _seccion('⚠️ Productos con Margen Negativo')
    neg = analisis['margen_negativo']
    if neg.empty:
        st.markdown('<div class="ok-banner">✅ No hay productos con margen negativo en este período.</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown(
            f'<div class="negative-alert">🚨 Se detectaron <b>{len(neg)} productos</b> vendidos por debajo del costo.</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(
            neg.style
               .format({'costo': fmt_moneda, 'precio': fmt_moneda,
                        'rentabilidad': fmt_moneda, 'margen': fmt_pct,
                        'participacion': fmt_pct})
               .map(lambda v: f'color:{C_ROJO}; font-weight:bold' if isinstance(v, float) and v < 0 else '',
                    subset=['margen', 'rentabilidad']),
            use_container_width=True,
            hide_index=True,
        )

    # ================================================================
    # SECCION 7 — Distribucion por Rango de Margen
    # ================================================================
    _seccion('Distribucion por Rango de Margen')
    dist = analisis['distribucion_margen']
    col_g, col_t = st.columns([1.3, 1])
    with col_g:
        st.plotly_chart(_grafico_dona(dist), use_container_width=True)
    with col_t:
        df_dist = pd.DataFrame(dist)
        df_dist = df_dist[['rango', 'cantidad', 'pct_cantidad', 'precio_total', 'pct_ventas']]
        df_dist.columns = ['Rango', 'Productos', '% Productos', 'Ventas ($)', '% Ventas']
        st.dataframe(
            df_dist.style.format({
                '% Productos': '{:.1f}%',
                'Ventas ($)': '$ {:,.2f}',
                '% Ventas': '{:.1f}%',
            }),
            use_container_width=True,
            hide_index=True,
        )

    # ================================================================
    # SECCION 8 — Tabla Completa
    # ================================================================
    _seccion('Tabla Completa de Productos')
    df_full = analisis['df_completo'].copy()

    # Buscador
    col_search, col_filter = st.columns([2, 1])
    with col_search:
        busqueda = st.text_input('🔍 Buscar producto (descripcion o codigo)', '')
    with col_filter:
        filtro_margen = st.selectbox('Filtrar por margen', [
            'Todos', 'Solo negativos', 'Solo positivos', '> 100%'
        ])

    if busqueda:
        mask = (df_full['descripcion'].str.contains(busqueda, case=False, na=False) |
                df_full['codigo'].str.contains(busqueda, case=False, na=False))
        df_full = df_full[mask]

    if filtro_margen == 'Solo negativos':
        df_full = df_full[df_full['margen'] < 0]
    elif filtro_margen == 'Solo positivos':
        df_full = df_full[df_full['margen'] >= 0]
    elif filtro_margen == '> 100%':
        df_full = df_full[df_full['margen'] > 100]

    st.caption(f'Mostrando {len(df_full)} de {len(analisis["df_completo"])} productos')
    st.dataframe(
        df_full.style.format({
            'unidades': '{:,.2f}',
            'costo': '$ {:,.2f}',
            'precio': '$ {:,.2f}',
            'rentabilidad': '$ {:,.2f}',
            'margen': '{:.2f}%',
            'participacion': '{:.2f}%',
        }),
        use_container_width=True,
        hide_index=True,
        height=420,
    )

    # ================================================================
    # SECCION 9 — Exportar Excel
    # ================================================================
    _seccion('📥 Exportar a Excel')
    st.write('Descargá el reporte completo en formato Excel con 6 hojas: Resumen, Datos Completos, '
             'Top Vendidos, Top Rentabilidad, Margen Negativo y Análisis Pareto.')

    excel_bytes = generar_excel(analisis, meta)
    nombre_archivo = 'reporte_%s_%s_%s.xlsx' % (
        sucursal.replace(' ', '_'),
        fecha_d.replace('/', '-'),
        fecha_h.replace('/', '-'),
    )
    st.download_button(
        label='⬇️ Descargar Reporte Excel',
        data=excel_bytes,
        file_name=nombre_archivo,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
        type='primary',
    )

    # Validacion discreta
    with st.expander('🔍 Validacion de extraccion'):
        st.json({
            'productos_extraidos': int(len(df)),
            'sum_costo': float(round(df['costo'].sum(), 2)),
            'sum_precio': float(round(df['precio'].sum(), 2)),
            'sum_rentabilidad': float(round(df['rentabilidad'].sum(), 2)),
            'totals_pdf': {k: float(v) for k, v in totals.items()},
        })


if __name__ == '__main__':
    main()

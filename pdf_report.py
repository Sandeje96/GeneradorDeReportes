# -*- coding: utf-8 -*-
"""
pdf_report.py — Genera reportes PDF profesionales.
Requiere: reportlab >= 4.0, kaleido == 0.2.1
"""
import io
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape as rl_landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.colors import HexColor, Color
from reportlab.platypus import (
    Paragraph, Spacer, Table, TableStyle, PageBreak,
    Image, NextPageTemplate, HRFlowable, KeepTogether,
)
from reportlab.platypus.doctemplate import BaseDocTemplate, PageTemplate
from reportlab.platypus.frames import Frame

# ─── Paleta ──────────────────────────────────────────────────────────────────
C_AZUL   = HexColor('#1B3A5C')
C_DORADO = HexColor('#C8A84E')
C_VERDE  = HexColor('#27AE60')
C_ROJO   = HexColor('#E74C3C')
C_GRIS   = HexColor('#F8F9FA')
C_BORDE  = HexColor('#CCCCCC')

PC_AZUL  = '#1B3A5C'
PC_DORA  = '#C8A84E'
PC_VERD  = '#27AE60'
PC_ROJO  = '#E74C3C'
COLORES_S = [PC_AZUL, PC_DORA, PC_VERD, PC_ROJO, '#9B59B6', '#1ABC9C', '#E67E22']

# ─── Páginas y márgenes ───────────────────────────────────────────────────────
PAGE_P = A4                   # 595.3 × 841.9 pt  (portrait)
PAGE_L = rl_landscape(A4)    # 841.9 × 595.3 pt  (landscape)
MG     = 1.5 * cm
MG_T   = 1.8 * cm
MG_B   = 1.2 * cm
FT_H   = 0.9 * cm            # alto del footer

def _avail(page):
    w, h = page
    return w - 2 * MG, h - MG_T - MG_B - FT_H

AVAIL_P = _avail(PAGE_P)   # ~510 × 693 pt
AVAIL_L = _avail(PAGE_L)   # ~756 × 447 pt

# ─── Estilos tipográficos ─────────────────────────────────────────────────────
def _sty(name, **kw):
    base = dict(fontName='Helvetica', fontSize=10,
                textColor=HexColor('#1A1A1A'), spaceAfter=3,
                spaceBefore=2, leading=13)
    base.update(kw)
    return ParagraphStyle(name, **base)

S = {
    'sec':   _sty('sec',  fontName='Helvetica-Bold', fontSize=14, textColor=C_AZUL,
                  spaceBefore=8, spaceAfter=5),
    'sub':   _sty('sub',  fontName='Helvetica-Bold', fontSize=11, textColor=C_AZUL,
                  spaceBefore=5, spaceAfter=3),
    'norm':  _sty('norm', fontSize=10),
    'small': _sty('sm',   fontSize=8, textColor=HexColor('#555555')),
    'idx_h': _sty('idh',  fontName='Helvetica-Bold', fontSize=12, textColor=C_AZUL,
                  spaceBefore=10, spaceAfter=2),
    'idx_i': _sty('idi',  fontSize=11, leftIndent=20, spaceBefore=2, spaceAfter=2),
    'idx_s': _sty('ids',  fontSize=10, leftIndent=36, textColor=HexColor('#555555'),
                  spaceBefore=1, spaceAfter=1),
    'cen':   _sty('cen',  alignment=TA_CENTER),
    'pie':   _sty('pie',  fontSize=7, textColor=colors.grey, alignment=TA_CENTER),
    'div':   _sty('div',  fontName='Helvetica-Bold', fontSize=22, textColor=colors.white,
                  alignment=TA_CENTER),
    'div_s': _sty('divs', fontSize=14, textColor=Color(1, 1, 1, 0.85),
                  alignment=TA_CENTER),
}


# ─── Clase documento (portrait + landscape) ───────────────────────────────────
class _PetriDoc(BaseDocTemplate):

    def __init__(self, buf, cover, footer_txt=''):
        self._cover = cover
        self._footer = footer_txt
        BaseDocTemplate.__init__(
            self, buf,
            leftMargin=MG, rightMargin=MG,
            topMargin=MG_T, bottomMargin=MG_B + FT_H,
            title=cover.get('titulo', ''),
            author='Grupo Petri',
        )
        wp, hp = AVAIL_P
        wl, hl = AVAIL_L
        fp = Frame(MG, MG_B + FT_H, wp, hp, id='p')
        fl = Frame(MG, MG_B + FT_H, wl, hl, id='l')
        fc = Frame(0, 0, PAGE_P[0], PAGE_P[1], id='c')  # full-page cover

        self.addPageTemplates([
            PageTemplate('cover',     pagesize=PAGE_P, frames=[fc], onPage=self._pg_cover),
            PageTemplate('portrait',  pagesize=PAGE_P, frames=[fp], onPage=self._pg_footer_p),
            PageTemplate('landscape', pagesize=PAGE_L, frames=[fl], onPage=self._pg_footer_l),
        ])

    # ── Callbacks ──────────────────────────────────────────────────────────────

    def _pg_cover(self, canvas, doc):
        canvas.saveState()
        cd = self._cover
        w, h = PAGE_P

        # Fondo azul oscuro
        canvas.setFillColor(C_AZUL)
        canvas.rect(0, 0, w, h, fill=1, stroke=0)

        # Franja dorada decorativa
        canvas.setFillColor(C_DORADO)
        canvas.rect(0, h * 0.40, w, 4.5, fill=1, stroke=0)
        canvas.rect(0, h * 0.40 - 11, w, 2, fill=1, stroke=0)

        # Bloque superior con nombre de empresa
        canvas.setFont('Helvetica', 10)
        canvas.setFillColor(Color(1, 1, 1, 0.55))
        canvas.drawCentredString(w / 2, h * 0.88, 'GRUPO PETRI')
        # Líneas decorativas flanqueando el nombre
        canvas.setStrokeColor(Color(1, 1, 1, 0.25))
        canvas.setLineWidth(0.5)
        canvas.line(w * 0.15, h * 0.893, w * 0.38, h * 0.893)
        canvas.line(w * 0.62, h * 0.893, w * 0.85, h * 0.893)

        # Título principal
        canvas.setFont('Helvetica-Bold', 34)
        canvas.setFillColor(C_DORADO)
        canvas.drawCentredString(w / 2, h * 0.65, cd.get('titulo', 'Reporte de Ventas'))

        # Tipo de reporte
        canvas.setFont('Helvetica', 16)
        canvas.setFillColor(Color(1, 1, 1, 0.88))
        canvas.drawCentredString(w / 2, h * 0.58, cd.get('subtitulo', ''))

        # Sucursal(es)
        canvas.setFont('Helvetica-Bold', 20)
        canvas.setFillColor(colors.white)
        canvas.drawCentredString(w / 2, h * 0.49, cd.get('sucursal', ''))

        # Período
        canvas.setFont('Helvetica', 13)
        canvas.setFillColor(Color(1, 1, 1, 0.78))
        canvas.drawCentredString(w / 2, h * 0.42, cd.get('periodo', ''))

        # Pie de carátula
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(Color(1, 1, 1, 0.45))
        canvas.drawCentredString(
            w / 2, h * 0.05,
            f'Generado el {date.today().strftime("%d/%m/%Y")}  ·  Generador de Reportes — Grupo Petri'
        )
        canvas.restoreState()

    def _pg_footer_p(self, canvas, doc):
        self._draw_footer(canvas, doc, PAGE_P)

    def _pg_footer_l(self, canvas, doc):
        self._draw_footer(canvas, doc, PAGE_L)

    def _draw_footer(self, canvas, doc, page):
        canvas.saveState()
        w = page[0]
        y = (MG_B + FT_H) * 0.32
        canvas.setFont('Helvetica', 7)
        canvas.setFillColor(colors.grey)
        canvas.drawString(MG, y, self._footer)
        canvas.drawRightString(w - MG, y, f'Página {doc.page}')
        canvas.setStrokeColor(HexColor('#E0E0E0'))
        canvas.setLineWidth(0.5)
        canvas.line(MG, (MG_B + FT_H) * 0.65, w - MG, (MG_B + FT_H) * 0.65)
        canvas.restoreState()


# ─── Conversión gráfico → Image flowable ─────────────────────────────────────

def _fig_img(fig, avail_w, avail_h, w_px=1500, h_px=800):
    """Exporta un gráfico plotly a PNG y devuelve un Image flowable."""
    png = fig.to_image(format='png', width=w_px, height=h_px, scale=1.5)
    buf = io.BytesIO(png)
    ar = h_px / w_px
    draw_w = min(avail_w, avail_h / ar)
    draw_h = draw_w * ar
    img = Image(buf, width=draw_w, height=draw_h)
    img.hAlign = 'CENTER'
    return img


# ─── Gráficos plotly para el PDF ──────────────────────────────────────────────

def _graf_barras_h(df, col_val, titulo, xlabel, color=PC_AZUL, n=15, fmt_money=True):
    df_top = df.nlargest(n, col_val)[['descripcion', col_val]].copy()
    df_top = df_top.sort_values(col_val)
    df_top['_lbl'] = df_top['descripcion'].str[:48]
    txt = df_top[col_val].apply(lambda v: f'$ {v:,.0f}' if fmt_money else f'{v:,.2f}')
    fig = go.Figure(go.Bar(
        x=df_top[col_val], y=df_top['_lbl'],
        orientation='h', marker_color=color,
        text=txt, textposition='outside', textfont=dict(size=10),
    ))
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=17, color=PC_AZUL, family='Arial')),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=11),
        xaxis=dict(tickformat=',.0f', gridcolor='#EEEEEE', title=xlabel),
        yaxis=dict(tickfont=dict(size=10), automargin=True),
        margin=dict(l=400, r=120, t=60, b=40),
        height=700, width=1500,
    )
    return fig


def _graf_pareto(pareto_df, pareto_stats):
    if pareto_df is None or pareto_df.empty:
        return go.Figure()
    n = len(pareto_df)
    n80 = pareto_stats.get('n_productos_80pct', 0)
    pct80 = pareto_stats.get('pct_productos_para_80', 0)
    x = list(range(1, n + 1))
    fig = go.Figure()
    fig.add_bar(x=x, y=pareto_df['precio'].values, name='Precio Venta',
                marker_color=PC_AZUL, yaxis='y1')
    fig.add_scatter(x=x, y=pareto_df['participacion_acum'].values,
                    name='Acumulado %', mode='lines+markers',
                    line=dict(color=PC_DORA, width=2.5),
                    marker=dict(size=5), yaxis='y2')
    fig.update_layout(
        title=dict(
            text=f'Análisis Pareto — {n80} productos ({pct80:.1f}%) concentran el 80% de la facturación',
            font=dict(size=15, color=PC_AZUL, family='Arial'),
        ),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=11),
        yaxis=dict(title='Precio Venta ($)', tickformat='$,.0f', gridcolor='#EEEEEE'),
        yaxis2=dict(title='Acumulado %', ticksuffix='%', overlaying='y',
                    side='right', range=[0, 110]),
        xaxis=dict(title='Productos (orden descendente de facturación)', gridcolor='#EEEEEE'),
        legend=dict(orientation='h', y=1.02, x=0.0, xanchor='left', yanchor='bottom'),
        margin=dict(l=20, r=70, t=70, b=40),
        height=640, width=1400,
        shapes=[
            dict(type='line', x0=n80, x1=n80, y0=0, y1=1, yref='paper',
                 line=dict(color=PC_VERD, dash='dash', width=2)),
            dict(type='line', x0=0, x1=1, y0=80, y1=80, xref='paper', yref='y2',
                 line=dict(color=PC_ROJO, dash='dash', width=2)),
        ],
        annotations=[
            dict(x=n80, y=1.02, yref='paper', text=f' {n80} prods.',
                 showarrow=False, font=dict(color=PC_VERD, size=10)),
        ],
    )
    return fig


def _graf_donut(distribucion_margen, titulo='Distribución por Rango de Margen'):
    COLS = {'< 0% (perdida)': PC_ROJO, '0% - 30%': PC_DORA,
            '30% - 60%': '#F0B429', '60% - 100%': PC_VERD, '> 100%': PC_AZUL}
    labels = [d['rango'] for d in distribucion_margen if d['cantidad'] > 0]
    values = [d['cantidad'] for d in distribucion_margen if d['cantidad'] > 0]
    clrs   = [COLS.get(l, '#95A5A6') for l in labels]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.45,
        marker_colors=clrs, textinfo='label+percent', textfont_size=12,
    ))
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=16, color=PC_AZUL, family='Arial')),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=12),
        legend=dict(orientation='v', x=1.01, y=0.5),
        margin=dict(l=20, r=20, t=60, b=20),
        height=560, width=900,
    )
    return fig


def _graf_comp_grouped(df_comp):
    fig = go.Figure()
    for col, name, color in [('total_ventas', 'Ventas', PC_AZUL),
                               ('total_costo', 'Costo', PC_ROJO),
                               ('total_rentabilidad', 'Rentabilidad', PC_VERD)]:
        fig.add_bar(name=name, x=df_comp['sucursal'], y=df_comp[col],
                    marker_color=color,
                    text=df_comp[col].apply(lambda v: f'${v/1e6:.1f}M'),
                    textposition='outside', textfont=dict(size=11))
    fig.update_layout(
        barmode='group',
        title=dict(text='Ventas, Costo y Rentabilidad por Sucursal',
                   font=dict(size=17, color=PC_AZUL, family='Arial')),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=12),
        yaxis=dict(tickformat='$,.0f', gridcolor='#EEEEEE', title='$ Pesos AR'),
        legend=dict(orientation='h', y=1.02, x=0.0, xanchor='left', yanchor='bottom'),
        margin=dict(l=20, r=20, t=80, b=40),
        height=600, width=1400,
    )
    return fig


def _graf_comp_top(datos, col_val, titulo, xlabel, n=10):
    dfs = []
    for suc, an in datos.items():
        top = an['df_completo'].nlargest(50, col_val)[['descripcion', col_val]].copy()
        top['sucursal'] = suc
        dfs.append(top)
    df_all = pd.concat(dfs, ignore_index=True)
    top_prods = (df_all.groupby('descripcion')[col_val].sum()
                 .nlargest(n).index.tolist())
    df_plot = df_all[df_all['descripcion'].isin(top_prods)].copy()
    df_plot['_lbl'] = df_plot['descripcion'].str[:42]
    fig = px.bar(df_plot, x=col_val, y='_lbl', color='sucursal',
                 orientation='h', barmode='group',
                 labels={col_val: xlabel, '_lbl': '', 'sucursal': 'Sucursal'},
                 color_discrete_sequence=COLORES_S)
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=16, color=PC_AZUL, family='Arial')),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=11),
        xaxis=dict(tickformat=',.0f', gridcolor='#EEEEEE'),
        yaxis=dict(tickfont=dict(size=10), automargin=True),
        legend=dict(orientation='h', y=1.02, x=0.0, xanchor='left', yanchor='bottom'),
        margin=dict(l=400, r=20, t=80, b=40),
        height=max(500, n * 55 + 130), width=1500,
    )
    return fig


def _graf_comp_margen_dist(datos):
    rangos = ['< 0% (perdida)', '0% - 30%', '30% - 60%', '60% - 100%', '> 100%']
    cols   = [PC_ROJO, PC_DORA, '#F0B429', PC_VERD, PC_AZUL]
    fig = go.Figure()
    for rango, color in zip(rangos, cols):
        y_vals = []
        for suc, an in datos.items():
            dist = {d['rango']: d['pct_cantidad'] for d in an['distribucion_margen']}
            y_vals.append(dist.get(rango, 0))
        fig.add_bar(name=rango, x=list(datos.keys()), y=y_vals,
                    marker_color=color,
                    text=[f'{v:.0f}%' for v in y_vals], textposition='inside',
                    textfont=dict(size=11))
    fig.update_layout(
        barmode='stack',
        title=dict(text='Distribución de Márgenes por Sucursal (% de productos)',
                   font=dict(size=16, color=PC_AZUL, family='Arial')),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=12),
        yaxis=dict(ticksuffix='%', gridcolor='#EEEEEE', title='% de productos'),
        legend=dict(orientation='h', y=1.02, x=0.0, xanchor='left', yanchor='bottom'),
        margin=dict(l=20, r=20, t=80, b=40),
        height=550, width=1200,
    )
    return fig


# ─── Helpers de flowables ────────────────────────────────────────────────────

def _titulo_seccion(texto):
    return [
        HRFlowable(width='100%', thickness=2, color=C_AZUL, spaceAfter=4),
        Paragraph(texto, S['sec']),
        Spacer(1, 4),
    ]


def _kpi_table(kpis, avail_w):
    """kpis: list[(label, valor_str)]"""
    col_lbl = avail_w * 0.55
    col_val = avail_w * 0.43
    data = []
    for lbl, val in kpis:
        data.append([
            Paragraph(f'<b>{lbl}</b>',
                      ParagraphStyle('kl', fontName='Helvetica-Bold', fontSize=10, textColor=C_AZUL)),
            Paragraph(val,
                      ParagraphStyle('kv', fontName='Helvetica', fontSize=10,
                                     textColor=HexColor('#1A1A1A'), alignment=TA_RIGHT)),
        ])
    cmds = [
        ('GRID',         (0, 0), (-1, -1), 0.5, C_BORDE),
        ('TOPPADDING',   (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
        ('LEFTPADDING',  (0, 0), (0, -1),  10),
        ('RIGHTPADDING', (1, 0), (1, -1),  10),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ]
    for i in range(len(data)):
        bg = colors.white if i % 2 == 0 else C_GRIS
        cmds.append(('BACKGROUND', (0, i), (-1, i), bg))
    t = Table(data, colWidths=[col_lbl, col_val])
    t.setStyle(TableStyle(cmds))
    return t


def _comp_table(df_comp, avail_w):
    """Tabla comparativa de sucursales."""
    headers = ['Sucursal', 'Ventas ($)', 'Costo ($)', 'Rentabilidad ($)',
               'Margen Global', 'Cant. Prods.', 'Con Ventas', 'Prods. 80%']
    ncols = len(headers)
    col_w = avail_w / ncols

    def fmt(v, tipo):
        if tipo == 'm': return f'$ {v:,.0f}'
        if tipo == 'p': return f'{v:.1f}%'
        if tipo == 'i': return f'{int(v):,}'
        return str(v)

    rows = [[Paragraph(f'<b>{h}</b>',
                        ParagraphStyle('ch', fontName='Helvetica-Bold', fontSize=8,
                                       textColor=colors.white, alignment=TA_CENTER))
             for h in headers]]

    for _, r in df_comp.iterrows():
        rows.append([
            Paragraph(str(r['sucursal']),
                      ParagraphStyle('cd', fontName='Helvetica-Bold', fontSize=9)),
            Paragraph(fmt(r['total_ventas'], 'm'),
                      ParagraphStyle('cv', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
            Paragraph(fmt(r['total_costo'], 'm'),
                      ParagraphStyle('cv2', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
            Paragraph(fmt(r['total_rentabilidad'], 'm'),
                      ParagraphStyle('cv3', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
            Paragraph(fmt(r['margen_global'], 'p'),
                      ParagraphStyle('cv4', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
            Paragraph(fmt(r['cantidad_productos'], 'i'),
                      ParagraphStyle('cv5', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
            Paragraph(fmt(r['cantidad_activos'], 'i'),
                      ParagraphStyle('cv6', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
            Paragraph(fmt(r['pareto_n80'], 'i'),
                      ParagraphStyle('cv7', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
        ])

    cmds = [
        ('BACKGROUND',   (0, 0), (-1, 0),  C_AZUL),
        ('GRID',         (0, 0), (-1, -1), 0.5, C_BORDE),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
        ('LEFTPADDING',  (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ]
    for i in range(1, len(rows)):
        bg = colors.white if i % 2 == 1 else C_GRIS
        cmds.append(('BACKGROUND', (0, i), (-1, i), bg))

    # Fila de totales
    tv = df_comp['total_ventas'].sum()
    tc = df_comp['total_costo'].sum()
    tr = df_comp['total_rentabilidad'].sum()
    mg = (tr / tv * 100) if tv else 0
    total_row = [
        Paragraph('<b>TOTAL</b>', ParagraphStyle('tt', fontName='Helvetica-Bold',
                                                  fontSize=9, textColor=colors.white)),
        Paragraph(fmt(tv, 'm'), ParagraphStyle('tv', fontName='Helvetica-Bold', fontSize=9,
                                                textColor=colors.white, alignment=TA_RIGHT)),
        Paragraph(fmt(tc, 'm'), ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=9,
                                                textColor=colors.white, alignment=TA_RIGHT)),
        Paragraph(fmt(tr, 'm'), ParagraphStyle('tr', fontName='Helvetica-Bold', fontSize=9,
                                                textColor=colors.white, alignment=TA_RIGHT)),
        Paragraph(fmt(mg, 'p'),  ParagraphStyle('tm', fontName='Helvetica-Bold', fontSize=9,
                                                textColor=colors.white, alignment=TA_RIGHT)),
        Paragraph('', S['small']), Paragraph('', S['small']), Paragraph('', S['small']),
    ]
    rows.append(total_row)
    cmds.append(('BACKGROUND', (0, len(rows) - 1), (-1, len(rows) - 1), C_AZUL))

    t = Table(rows, colWidths=[col_w] * ncols)
    t.setStyle(TableStyle(cmds))
    return t


def _data_table(df, avail_w, max_rows=None):
    """Tabla completa de productos. Si max_rows se limita el total."""
    if max_rows:
        df = df.head(max_rows)

    headers = ['Cód.', 'Descripción', 'Unidades', 'Costo ($)', 'Precio ($)',
               'Rentabilidad ($)', 'Margen (%)']
    # Anchos proporcionales
    cws = [avail_w * p for p in [0.07, 0.30, 0.10, 0.13, 0.13, 0.14, 0.10]]

    def _p(txt, bold=False, align=TA_LEFT, size=8, color=HexColor('#1A1A1A')):
        fn = 'Helvetica-Bold' if bold else 'Helvetica'
        return Paragraph(str(txt), ParagraphStyle('_', fontName=fn, fontSize=size,
                                                   textColor=color, alignment=align))

    data = [[_p(h, bold=True, align=TA_CENTER, color=colors.white) for h in headers]]

    for _, r in df.iterrows():
        data.append([
            _p(r['codigo'],    align=TA_CENTER),
            _p(r['descripcion'][:45]),
            _p(f"{r['unidades']:,.2f}", align=TA_RIGHT),
            _p(f"$ {r['costo']:,.2f}",        align=TA_RIGHT),
            _p(f"$ {r['precio']:,.2f}",        align=TA_RIGHT),
            _p(f"$ {r['rentabilidad']:,.2f}",  align=TA_RIGHT),
            _p(f"{r['margen']:.2f}%",          align=TA_RIGHT),
        ])

    cmds = [
        ('BACKGROUND',   (0, 0), (-1, 0),  C_AZUL),
        ('GRID',         (0, 0), (-1, -1), 0.3, C_BORDE),
        ('TOPPADDING',   (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
        ('LEFTPADDING',  (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ]
    for i in range(1, len(data)):
        bg = colors.white if i % 2 == 1 else C_GRIS
        cmds.append(('BACKGROUND', (0, i), (-1, i), bg))

    t = Table(data, colWidths=cws, repeatRows=1)
    t.setStyle(TableStyle(cmds))
    return t


def _divider_page(nombre, subtitulo=''):
    """Página separadora para cada sucursal en el reporte comparativo."""
    # Se implementa como tabla de celda única con fondo azul que llena el frame
    content = [
        Paragraph(nombre, S['div']),
        Spacer(1, 12),
        Paragraph(subtitulo, S['div_s']),
    ]
    return content


# ─── Bloque de sección individual (una sucursal) ─────────────────────────────

def _story_sucursal(nombre, an, meta, avail_p, avail_l, include_data=True, max_data=None):
    """Genera los flowables para una sucursal completa."""
    story = []
    res = an['resumen']
    df  = an['df_completo']
    fd  = meta.get('fecha_desde', '')
    fh  = meta.get('fecha_hasta', '')
    wp, hp = avail_p
    wl, hl = avail_l

    # ── KPIs ──────────────────────────────────────────────────────────────────
    story += [NextPageTemplate('portrait'), PageBreak()]
    story += _titulo_seccion(f'Resumen Ejecutivo — {nombre}')
    story.append(Paragraph(f'Período: {fd} al {fh}', S['small']))
    story.append(Spacer(1, 8))

    kpis = [
        ('Total Ventas (Precio)',    f"$ {res['total_ventas']:,.2f}"),
        ('Total Costos',             f"$ {res['total_costo']:,.2f}"),
        ('Rentabilidad Total',       f"$ {res['total_rentabilidad']:,.2f}"),
        ('Margen Global',            f"{res['margen_global']:.2f}%"),
        ('Cantidad de Productos',    f"{int(res['cantidad_productos']):,}"),
        ('Productos con ventas > 0', f"{int(res['cantidad_activos']):,}"),
    ]
    if res.get('mas_vendido'):
        mv = res['mas_vendido']
        kpis.append(('Prod. más vendido (unid.)',
                      f"{mv['descripcion'][:40]} — {mv['unidades']:,.2f} un."))
    if res.get('mas_rentable'):
        mr = res['mas_rentable']
        kpis.append(('Prod. más rentable ($)',
                      f"{mr['descripcion'][:40]} — $ {mr['rentabilidad']:,.2f}"))

    story.append(_kpi_table(kpis, wp))

    # ── Top 15 por Unidades ───────────────────────────────────────────────────
    fig = _graf_barras_h(df, 'unidades',
                          f'Top 15 Productos Más Vendidos — {nombre}',
                          'Unidades Vendidas', color=PC_AZUL, fmt_money=False)
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(fig, wl, hl))

    # ── Top 15 por Rentabilidad ────────────────────────────────────────────────
    fig = _graf_barras_h(df, 'rentabilidad',
                          f'Top 15 Productos por Rentabilidad — {nombre}',
                          '$ Rentabilidad', color=PC_VERD)
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(fig, wl, hl))

    # ── Top 15 por Precio ─────────────────────────────────────────────────────
    fig = _graf_barras_h(df, 'precio',
                          f'Top 15 Productos por Facturación — {nombre}',
                          '$ Precio Venta', color=PC_DORA)
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(fig, wl, hl))

    # ── Pareto ────────────────────────────────────────────────────────────────
    fig = _graf_pareto(an.get('pareto_df'), an.get('pareto_stats', {}))
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(fig, wl, hl))

    # ── Distribución de márgenes ──────────────────────────────────────────────
    fig = _graf_donut(an.get('distribucion_margen', []),
                      f'Distribución de Márgenes — {nombre}')
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(fig, wl, min(hl, wl * 0.6 / 0.9)))

    # ── Tabla de datos ────────────────────────────────────────────────────────
    if include_data:
        story += [NextPageTemplate('portrait'), PageBreak()]
        story += _titulo_seccion(f'Datos Completos — {nombre}')
        n_total = len(df)
        if max_data and n_total > max_data:
            story.append(Paragraph(
                f'Mostrando top {max_data} productos por facturación de {n_total:,} totales.',
                S['small']))
            story.append(Spacer(1, 4))
        story.append(_data_table(df.sort_values('precio', ascending=False), wp, max_rows=max_data))

    return story


# ─── API pública ──────────────────────────────────────────────────────────────

def generar_pdf(analisis, metadata):
    """
    Genera PDF para una sola sucursal.
    Devuelve bytes.
    """
    suc = metadata.get('sucursal', 'Sucursal')
    fd  = metadata.get('fecha_desde', '')
    fh  = metadata.get('fecha_hasta', '')

    cover = dict(
        titulo='Reporte de Ventas',
        subtitulo='Análisis Individual de Sucursal',
        sucursal=suc,
        periodo=f'Período: {fd} al {fh}',
    )
    footer = f'Grupo Petri — {suc} — {fd} al {fh}'

    buf = io.BytesIO()
    doc = _PetriDoc(buf, cover, footer)
    wp, hp = AVAIL_P
    wl, hl = AVAIL_L

    story = []

    # ── Carátula (página 1) ────────────────────────────────────────────────────
    story.append(Spacer(1, 1))        # placeholder — el fondo lo dibuja onPage
    story.append(NextPageTemplate('portrait'))
    story.append(PageBreak())

    # ── Índice (página 2) ──────────────────────────────────────────────────────
    story += _titulo_seccion('Índice')
    indice = [
        ('1', 'Resumen Ejecutivo',              'Métricas clave: ventas, costos, rentabilidad, margen global'),
        ('2', 'Top 15 Productos Más Vendidos',  'Productos con mayor volumen de unidades vendidas'),
        ('3', 'Top 15 por Rentabilidad',         'Productos que generaron mayor ganancia en $'),
        ('4', 'Top 15 por Facturación',          'Productos con mayor precio de venta total'),
        ('5', 'Análisis Pareto (80/20)',         'Qué porcentaje del catálogo genera el 80% de la facturación'),
        ('6', 'Distribución de Márgenes',        'Clasificación de productos por rango de margen porcentual'),
        ('7', 'Datos Completos',                 f'Tabla detallada de todos los productos ({len(analisis["df_completo"]):,} registros)'),
    ]
    for num, titulo, desc in indice:
        story.append(Paragraph(f'<b>{num}.</b>  <b>{titulo}</b>', S['idx_h']))
        story.append(Paragraph(desc, S['idx_s']))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        f'Generado el {date.today().strftime("%d/%m/%Y")}  ·  Sucursal: {suc}  ·  Período: {fd} al {fh}',
        S['small']))
    # Sin PageBreak extra — _story_sucursal ya inicia con su propio PageBreak

    # ── Secciones de contenido ─────────────────────────────────────────────────
    story += _story_sucursal(suc, analisis, metadata, AVAIL_P, AVAIL_L,
                              include_data=True, max_data=None)

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def generar_pdf_comparacion(datos, metas, df_comp):
    """
    Genera PDF comparativo multi-sucursal.
    datos: {nombre: analisis_dict}
    metas: {nombre: meta_dict}
    df_comp: DataFrame resumen comparativo
    Devuelve bytes.
    """
    sucursales = list(datos.keys())
    m0 = metas.get(sucursales[0], {}) if sucursales else {}
    fd = m0.get('fecha_desde', '')
    fh = m0.get('fecha_hasta', '')
    lista_suc = ', '.join(sucursales)

    cover = dict(
        titulo='Reporte de Ventas',
        subtitulo='Análisis Comparativo de Sucursales',
        sucursal=lista_suc if len(lista_suc) <= 60 else f'{len(sucursales)} Sucursales',
        periodo=f'Período: {fd} al {fh}',
    )
    footer = f'Grupo Petri — Comparación de Sucursales — {fd} al {fh}'

    buf = io.BytesIO()
    doc = _PetriDoc(buf, cover, footer)
    wp, hp = AVAIL_P
    wl, hl = AVAIL_L

    story = []

    # ── Carátula ───────────────────────────────────────────────────────────────
    story.append(Spacer(1, 1))
    story.append(NextPageTemplate('portrait'))
    story.append(PageBreak())

    # ── Índice ─────────────────────────────────────────────────────────────────
    story += _titulo_seccion('Índice')
    story.append(Paragraph('<b>SECCIÓN A — Comparación General</b>', S['idx_h']))
    for num, tit, desc in [
        ('A1', 'Resumen Comparativo',              'Tabla con los indicadores clave de cada sucursal lado a lado'),
        ('A2', 'Ventas, Costo y Rentabilidad',     'Gráfico comparativo de los tres indicadores financieros clave'),
        ('A3', 'Top 10 por Rentabilidad',           'Los productos más rentables en $ comparando sucursales'),
        ('A4', 'Top 10 por Facturación',            'Los productos con mayor precio de venta comparando sucursales'),
        ('A5', 'Distribución de Márgenes',          'Cómo se distribuyen los márgenes en cada sucursal'),
    ]:
        story.append(Paragraph(f'<b>{num}.</b>  <b>{tit}</b>', S['idx_h']))
        story.append(Paragraph(desc, S['idx_s']))

    story.append(Spacer(1, 8))
    story.append(Paragraph('<b>SECCIÓN B — Detalle por Sucursal</b>', S['idx_h']))
    for i, suc in enumerate(sucursales, 1):
        story.append(Paragraph(
            f'<b>B{i}.</b>  <b>{suc}</b> — KPIs + Top 15 Vendidos + Top 15 Rentabilidad + '
            f'Top 15 Facturación + Pareto + Distribución de Márgenes + Top 50 Productos',
            S['idx_i']))

    story.append(Spacer(1, 6))
    story.append(Paragraph(
        f'Generado el {date.today().strftime("%d/%m/%Y")}  ·  Período: {fd} al {fh}',
        S['small']))
    story.append(PageBreak())

    # ── SECCIÓN A: Comparación General ────────────────────────────────────────

    # A1 — Tabla resumen
    story += _titulo_seccion('A1 · Resumen Comparativo de Sucursales')
    story.append(Spacer(1, 4))
    story.append(_comp_table(df_comp, wp))
    story.append(PageBreak())

    # A2 — Ventas/Costo/Rentabilidad
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(_graf_comp_grouped(df_comp), wl, hl))

    # A3 — Top 10 Rentabilidad comparativo
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(
        _graf_comp_top(datos, 'rentabilidad',
                       'Top 10 Productos por Rentabilidad — Comparación entre Sucursales',
                       '$ Rentabilidad'),
        wl, hl))

    # A4 — Top 10 Precio comparativo
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(
        _graf_comp_top(datos, 'precio',
                       'Top 10 Productos por Facturación — Comparación entre Sucursales',
                       '$ Precio Venta'),
        wl, hl))

    # A5 — Distribución márgenes comparativo
    story += [NextPageTemplate('landscape'), PageBreak()]
    story.append(_fig_img(_graf_comp_margen_dist(datos), wl, hl))

    # ── SECCIÓN B: Una sección por sucursal ───────────────────────────────────
    for i, suc in enumerate(sucursales, 1):
        an   = datos[suc]
        meta = metas.get(suc, {})

        # Página separadora con fondo azul
        story += [NextPageTemplate('portrait'), PageBreak()]
        # Tabla de una celda que llena el frame con fondo azul
        div_data = [[
            Paragraph(f'<br/><br/><br/><br/>'
                      f'<b>B{i}</b><br/><br/>'
                      f'<font size="22" color="white"><b>{suc}</b></font><br/><br/>'
                      f'<font size="13" color="#C8A84E">Detalle Individual de Sucursal</font>',
                      ParagraphStyle('dc', fontName='Helvetica-Bold', fontSize=14,
                                     textColor=colors.white, alignment=TA_CENTER,
                                     leading=28)),
        ]]
        div_t = Table(div_data, colWidths=[wp], rowHeights=[hp * 0.65])
        div_t.setStyle(TableStyle([
            ('BACKGROUND',   (0, 0), (-1, -1), C_AZUL),
            ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
            ('ROUNDEDCORNERS', [8]),
        ]))
        story.append(div_t)
        story.append(Spacer(1, 12))
        story.append(Paragraph(
            f'Período: {meta.get("fecha_desde", "")} al {meta.get("fecha_hasta", "")}',
            S['cen']))

        # Contenido de la sucursal (top 50 en tabla de datos)
        story += _story_sucursal(suc, an, meta, AVAIL_P, AVAIL_L,
                                  include_data=True, max_data=50)

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

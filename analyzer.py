# -*- coding: utf-8 -*-
"""
analyzer.py — Analisis del DataFrame extraido del PDF ABC de Productos.
Todas las funciones reciben un DataFrame con columnas:
  codigo, descripcion, unidades, costo, precio, rentabilidad, margen, participacion
"""
import pandas as pd


# ---------------------------------------------------------------------------
# Filtros base
# ---------------------------------------------------------------------------

def _activos(df):
    """Productos con al menos 1 unidad vendida (positiva)."""
    return df[df['unidades'] > 0].copy()


# ---------------------------------------------------------------------------
# Resumen ejecutivo
# ---------------------------------------------------------------------------

def get_resumen(df, metadata=None, totals=None):
    """
    Retorna dict con metricas generales del reporte.
    metadata: dict con fecha_desde, fecha_hasta, sucursal
    totals:   dict con costos, precios, rentabilidad del PDF
    """
    activos = _activos(df)

    total_ventas = df['precio'].sum()
    total_costo = df['costo'].sum()
    total_rentab = df['rentabilidad'].sum()
    margen_global = (total_rentab / abs(total_costo) * 100) if total_costo != 0 else 0.0

    # Producto mas vendido (por unidades)
    mas_vendido = None
    if not activos.empty:
        idx = activos['unidades'].idxmax()
        mas_vendido = activos.loc[idx, ['codigo', 'descripcion', 'unidades']].to_dict()

    # Producto mas rentable (por rentabilidad $)
    mas_rentable = None
    if not activos.empty:
        idx = activos['rentabilidad'].idxmax()
        mas_rentable = activos.loc[idx, ['codigo', 'descripcion', 'rentabilidad']].to_dict()

    # Producto de mayor precio total
    mas_facturado = None
    if not activos.empty:
        idx = activos['precio'].idxmax()
        mas_facturado = activos.loc[idx, ['codigo', 'descripcion', 'precio']].to_dict()

    return {
        'total_ventas': total_ventas,
        'total_costo': total_costo,
        'total_rentabilidad': total_rentab,
        'margen_global': margen_global,
        'cantidad_productos': len(df),
        'cantidad_activos': len(activos),
        'mas_vendido': mas_vendido,
        'mas_rentable': mas_rentable,
        'mas_facturado': mas_facturado,
        'fecha_desde': (metadata or {}).get('fecha_desde', ''),
        'fecha_hasta': (metadata or {}).get('fecha_hasta', ''),
        'sucursal': (metadata or {}).get('sucursal', ''),
        'totals_pdf': totals or {},
    }


# ---------------------------------------------------------------------------
# Rankings
# ---------------------------------------------------------------------------

def get_top_unidades(df, n=15):
    """Top N productos por unidades vendidas (solo positivos)."""
    activos = _activos(df)
    cols = ['codigo', 'descripcion', 'unidades', 'costo', 'precio', 'rentabilidad', 'margen']
    return activos.nlargest(n, 'unidades')[cols].reset_index(drop=True)


def get_top_rentabilidad(df, n=15):
    """Top N productos por rentabilidad en pesos."""
    activos = _activos(df)
    cols = ['codigo', 'descripcion', 'unidades', 'costo', 'precio', 'rentabilidad', 'margen']
    return activos.nlargest(n, 'rentabilidad')[cols].reset_index(drop=True)


def get_top_precio(df, n=15):
    """Top N productos por precio de venta total."""
    activos = _activos(df)
    cols = ['codigo', 'descripcion', 'unidades', 'costo', 'precio', 'rentabilidad', 'margen']
    return activos.nlargest(n, 'precio')[cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Margen negativo
# ---------------------------------------------------------------------------

def get_margen_negativo(df):
    """Productos con margen < 0%, ordenados de menor a mayor."""
    neg = df[df['margen'] < 0].copy()
    cols = ['codigo', 'descripcion', 'unidades', 'costo', 'precio', 'rentabilidad', 'margen']
    return neg.sort_values('margen')[cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Analisis Pareto
# ---------------------------------------------------------------------------

def get_pareto(df):
    """
    Analisis 80/20: cuantos productos representan el 80% de la facturacion.
    Retorna DataFrame con columnas:
      codigo, descripcion, precio, precio_acumulado, participacion_acum_pct
    y dict con estadisticas del pareto.
    """
    activos = _activos(df)
    if activos.empty:
        return pd.DataFrame(), {}

    total = activos['precio'].sum()
    pareto_df = (
        activos[['codigo', 'descripcion', 'precio', 'unidades', 'rentabilidad']]
        .sort_values('precio', ascending=False)
        .reset_index(drop=True)
    )
    pareto_df['precio_acum'] = pareto_df['precio'].cumsum()
    pareto_df['participacion_acum'] = pareto_df['precio_acum'] / total * 100

    # Cuantos productos llegan al 80%
    n_80 = int((pareto_df['participacion_acum'] <= 80).sum()) + 1
    n_80 = min(n_80, len(pareto_df))
    pct_productos_80 = n_80 / len(pareto_df) * 100

    stats = {
        'total_productos': len(pareto_df),
        'n_productos_80pct': n_80,
        'pct_productos_para_80': pct_productos_80,
    }
    return pareto_df, stats


# ---------------------------------------------------------------------------
# Distribucion por rangos de margen
# ---------------------------------------------------------------------------

RANGOS_MARGEN = [
    ('< 0% (perdida)',  None,  0),
    ('0% - 30%',           0, 30),
    ('30% - 60%',         30, 60),
    ('60% - 100%',        60, 100),
    ('> 100%',           100, None),
]


def get_distribucion_margen(df):
    """
    Distribucion de productos por rangos de margen.
    Retorna lista de dicts: {rango, cantidad, pct_cantidad, precio_total, pct_ventas}
    """
    total_productos = len(df)
    total_ventas = df['precio'].sum()

    resultado = []
    for label, lo, hi in RANGOS_MARGEN:
        if lo is None:
            mask = df['margen'] < hi
        elif hi is None:
            mask = df['margen'] >= lo
        else:
            mask = (df['margen'] >= lo) & (df['margen'] < hi)

        sub = df[mask]
        cantidad = len(sub)
        precio_total = sub['precio'].sum()

        resultado.append({
            'rango': label,
            'cantidad': cantidad,
            'pct_cantidad': cantidad / total_productos * 100 if total_productos > 0 else 0,
            'precio_total': precio_total,
            'pct_ventas': precio_total / total_ventas * 100 if total_ventas > 0 else 0,
        })

    return resultado


# ---------------------------------------------------------------------------
# Funcion principal: devuelve todo de una vez
# ---------------------------------------------------------------------------

def analizar(df, metadata=None, totals=None):
    """Ejecuta todos los analisis y retorna un dict completo."""
    pareto_df, pareto_stats = get_pareto(df)
    return {
        'resumen': get_resumen(df, metadata, totals),
        'top_unidades': get_top_unidades(df),
        'top_rentabilidad': get_top_rentabilidad(df),
        'top_precio': get_top_precio(df),
        'margen_negativo': get_margen_negativo(df),
        'pareto_df': pareto_df,
        'pareto_stats': pareto_stats,
        'distribucion_margen': get_distribucion_margen(df),
        'df_completo': df,
    }

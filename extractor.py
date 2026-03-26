# -*- coding: utf-8 -*-
"""
extractor.py — Extrae datos del PDF "Abc de Productos" de Grupo Petri.
Usa extract_text() pagina por pagina y parsea linea por linea con regex.
"""
import re
import pdfplumber
import pandas as pd

# ---------------------------------------------------------------------------
# Helpers de conversion numerica
# ---------------------------------------------------------------------------

def parse_number(s):
    """Convierte string numerico a float. Parens = negativo. Devuelve None si falla."""
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    negative = s.startswith('(') and s.endswith(')')
    s = s.strip('()')
    s = s.replace(',', '')
    try:
        val = float(s)
        return -val if negative else val
    except ValueError:
        return None


def parse_pct(s):
    """Convierte '52,524.43%' o '(0.01)%' a float (52524.43 o -0.01)."""
    if s is None:
        return None
    s = s.strip()
    negative = s.startswith('(')
    s = s.strip('()').rstrip('%').replace(',', '')
    try:
        val = float(s)
        return -val if negative else val
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Extraccion de metadatos
# ---------------------------------------------------------------------------

def extract_metadata(first_page_text):
    """Extrae Fecha desde, Fecha hasta y Sucursal del encabezado."""
    lines = first_page_text.split('\n')
    metadata = {'fecha_desde': '', 'fecha_hasta': '', 'sucursal': ''}
    for i, line in enumerate(lines):
        if 'Fecha desde:' in line and 'Sucursal:' in line:
            # La siguiente linea tiene los valores
            if i + 1 < len(lines):
                vals = lines[i + 1].strip().split()
                if len(vals) >= 3:
                    metadata['fecha_desde'] = vals[0]
                    metadata['fecha_hasta'] = vals[1]
                    metadata['sucursal'] = ' '.join(vals[2:])
            break
    return metadata


def extract_totals(last_page_text):
    """Extrae totales de la ultima pagina."""
    totals = {}
    patterns = {
        'costos': r'Costos:\s*\$\s*([\d,]+\.\d{2})',
        'precios': r'Precios:\s*\$\s*([\d,]+\.\d{2})',
        'rentabilidad': r'Rentabilidad:\s*\$\s*([\d,]+\.\d{2})',
        'margenes': r'Margenes:\s*([\d,]+\.\d{2})\s*%',
    }
    for key, pat in patterns.items():
        m = re.search(pat, last_page_text)
        if m:
            totals[key] = float(m.group(1).replace(',', ''))
    return totals


# ---------------------------------------------------------------------------
# Parser de linea de producto
# ---------------------------------------------------------------------------

# Numero con parens opcionales (cualquier cantidad de digitos y comas)
_NUM = r'\(?[\d,]+\.\d{2}\)?'
# Porcentaje con parens opcionales
_PCT = r'\(?[\d,]+\.\d{2}\)?%'
# Porcentaje con miles bien formateados (para separar del rentab pegado)
_PCT_PROPER = r'\(?(?:\d{1,3}(?:,\d{3})*)\.\d{2}\)?%'


def _split_rentab_marg(token):
    """
    Separa un token que puede ser:
      - 'marg%'              -> rentab=None, marg='marg%'
      - '-1.29%'             -> rentab=None, marg='-1.29%'  (negativo con -)
      - '(1.29)%'            -> rentab=None, marg='(1.29)%' (negativo con parens)
      - 'rentab_valmarg%'   -> rentab='rentab_val', marg='marg%'
    """
    if not token.endswith('%'):
        return token, None

    # Si empieza con - o (, es un marg negativo puro (sin rentab pegado)
    if token.startswith('-') or token.startswith('('):
        return None, token

    # Buscar el marg con formato miles correcto al FINAL del token
    m = re.search(r'(\(?(?:\d{1,3}(?:,\d{3})*)\.\d{2}\)?%)$', token)
    if m:
        marg = m.group(1)
        rentab = token[:m.start()].strip()
        return rentab if rentab else None, marg
    # Fallback: todo el token es el marg
    return None, token


def parse_product_line(line):
    """
    Parsea una linea de producto del PDF.
    Estrategia de derecha a izquierda.
    Retorna dict con campos o None si no parsea.
    """
    line = line.strip()
    if not re.match(r'^\d+\s', line):
        return None

    original = line

    # --- 1. Extraer part% del final ---
    m = re.search(r'\s(' + _PCT + r')\s*$', line)
    if not m:
        return None
    part_str = m.group(1)
    line = line[:m.start()].rstrip()

    # --- 2. Extraer el token que contiene marg (y posiblemente rentab pegado) ---
    # Buscar el ultimo % en la linea restante
    pct_pos = line.rfind('%')
    if pct_pos < 0:
        return None
    # Escanear hacia atras hasta encontrar espacio (inicio del token)
    start = pct_pos
    while start > 0 and line[start - 1] not in (' ', '\t'):
        start -= 1
    combined = line[start:pct_pos + 1].strip()
    line = line[:start].rstrip()

    rentab_str, marg_str = _split_rentab_marg(combined)

    # --- 3. Si rentab NO estaba pegado al marg, extraerlo ahora por separado ---
    if rentab_str is None or rentab_str == '':
        m = re.search(r'\s(' + _NUM + r')\s*$', line)
        if not m:
            m = re.search(r'(' + _NUM + r')\s*$', line)
        if not m:
            return None
        rentab_str = m.group(1)
        line = line[:m.start()].rstrip()

    # --- 4. Extraer precio del final de la linea restante ---
    m = re.search(r'\s(' + _NUM + r')\s*$', line)
    if not m:
        m = re.search(r'(' + _NUM + r')\s*$', line)
    if not m:
        return None
    precio_str = m.group(1)
    line = line[:m.start()].rstrip()

    # --- 4. Extraer costo ---
    m = re.search(r'\s(' + _NUM + r')\s*$', line)
    if not m:
        m = re.search(r'(' + _NUM + r')\s*$', line)
    if not m:
        return None
    costo_str = m.group(1)
    line = line[:m.start()].rstrip()

    # --- 5. Extraer unidades ---
    m = re.search(r'\s(' + _NUM + r')\s*$', line)
    if not m:
        m = re.search(r'(' + _NUM + r')\s*$', line)
    if not m:
        return None
    unidades_str = m.group(1)
    line = line[:m.start()].rstrip()

    # --- 6. Extraer codigo y descripcion ---
    m = re.match(r'^(\d+)\s+(.+)$', line)
    if not m:
        return None

    return {
        'codigo': m.group(1),
        'descripcion': m.group(2).strip(),
        'unidades_str': unidades_str,
        'costo_str': costo_str,
        'precio_str': precio_str,
        'rentab_str': rentab_str,
        'marg_str': marg_str,
        'part_str': part_str,
    }


# ---------------------------------------------------------------------------
# Lineas a ignorar
# ---------------------------------------------------------------------------

HEADER_STARTS = (
    'Abc de Productos',
    'Fecha desde:',
    'Rubro y SubRubro:',
    'Uid:',
    'Turno:',
    '1900/01/01',
    'Todos',
    'Todas',
    'No Aplicable',
    'Productor:',
    'Costos:',
    'Precios:',
    'Rentabilidad:',
    'Margenes:',
    u'C\u00f3digo',   # "Código Descripción..."
    'Codigo',
)

TIME_RE = re.compile(r'^\d{2}:\d{2}\s*$')
DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}')


def is_skip_line(line):
    stripped = line.strip()
    if not stripped:
        return True
    if TIME_RE.match(stripped):
        return True
    if DATE_RE.match(stripped):
        return True
    for h in HEADER_STARTS:
        if stripped.startswith(h):
            return True
    return False


def is_continuation_line(line):
    """Linea de continuacion de descripcion: no empieza con codigo."""
    stripped = line.strip()
    return bool(stripped) and not re.match(r'^\d+\s', stripped)


# ---------------------------------------------------------------------------
# Extraccion principal
# ---------------------------------------------------------------------------

def extract_pdf(pdf_path, verbose=True):
    """
    Extrae todos los productos del PDF.
    Retorna (DataFrame, metadata_dict, totals_dict).
    """
    rows = []
    metadata = {}
    totals = {}

    with pdfplumber.open(pdf_path) as pdf:
        n_pages = len(pdf.pages)
        if verbose:
            print("Total paginas: %d" % n_pages)

        for page_idx, page in enumerate(pdf.pages):
            page_num = page_idx + 1
            text = page.extract_text()
            if not text:
                if verbose:
                    print("  [WARN] Pagina %d sin texto" % page_num)
                continue

            # Metadatos solo de la primera pagina
            if page_idx == 0:
                metadata = extract_metadata(text)

            # Totales solo de la ultima pagina
            if page_idx == n_pages - 1:
                totals = extract_totals(text)

            lines = text.split('\n')
            pending_row = None  # fila parseada esperando posible continuacion

            for line in lines:
                if is_skip_line(line):
                    # Si hay fila pendiente, commitearla
                    if pending_row is not None:
                        rows.append(pending_row)
                        pending_row = None
                    continue

                if is_continuation_line(line):
                    # Agregar a la descripcion de la fila pendiente
                    if pending_row is not None:
                        pending_row['descripcion'] = (
                            pending_row['descripcion'] + ' ' + line.strip()
                        )
                    # Si no hay fila pendiente, ignorar
                    continue

                # Intentar parsear como fila de producto
                parsed = parse_product_line(line)
                if parsed is not None:
                    if pending_row is not None:
                        rows.append(pending_row)
                    pending_row = parsed
                else:
                    # No es producto ni continuacion ni header conocido
                    # Puede ser una linea con formato raro — ignorar con warning
                    if pending_row is not None:
                        rows.append(pending_row)
                        pending_row = None
                    if verbose and line.strip():
                        print("  [SKIP] Pag %d: %r" % (page_num, line.strip()[:80]))

            # Al final de la pagina, commitear fila pendiente
            if pending_row is not None:
                rows.append(pending_row)
                pending_row = None

    # --- Construir DataFrame ---
    records = []
    parse_errors = 0
    for r in rows:
        unidades = parse_number(r['unidades_str'])
        costo = parse_number(r['costo_str'])
        precio = parse_number(r['precio_str'])
        rentab_raw = parse_number(r['rentab_str'])
        marg = parse_pct(r['marg_str'])
        part = parse_pct(r['part_str'])

        if unidades is None or costo is None or precio is None:
            parse_errors += 1
            if verbose:
                print("  [ERR] No se pudo parsear fila: %s %s" % (
                    r.get('codigo', '?'), r.get('descripcion', '?')))
            continue

        # Rentabilidad: usar precio-costo si el valor extraido es inconsistente
        rentab_computed = precio - costo
        if rentab_raw is None:
            rentab = rentab_computed
        else:
            if precio != 0 and abs(rentab_raw - rentab_computed) / max(abs(precio), 1) > 0.01:
                rentab = rentab_computed  # fallback
            else:
                rentab = rentab_raw

        # Margen: recalcular si costo > 0, sino usar extraido
        if costo != 0 and marg is not None:
            marg_computed = (rentab / abs(costo)) * 100
            # Si el marg extraido difiere mucho del calculado, usar el calculado
            if abs(marg_computed - marg) / max(abs(marg_computed), 1) > 0.05:
                marg = marg_computed
        elif costo != 0:
            marg = (rentab / abs(costo)) * 100
        # si costo == 0, marg queda como fue extraido (puede ser 0.00%)

        records.append({
            'codigo': r['codigo'],
            'descripcion': r['descripcion'],
            'unidades': unidades,
            'costo': costo,
            'precio': precio,
            'rentabilidad': rentab,
            'margen': marg,
            'participacion': part,
        })

    df = pd.DataFrame(records)
    if not df.empty:
        df['codigo'] = df['codigo'].astype(str)

    if verbose and parse_errors:
        print("  [WARN] %d filas con error de parseo descartadas" % parse_errors)

    return df, metadata, totals


# ---------------------------------------------------------------------------
# Validacion
# ---------------------------------------------------------------------------

def validate(df, totals, verbose=True):
    """Valida las sumas contra los totales del PDF. Retorna True si pasa."""
    expected = {
        'costos': totals.get('costos', 0),
        'precios': totals.get('precios', 0),
        'rentabilidad': totals.get('rentabilidad', 0),
    }

    actual = {
        'costos': df['costo'].sum(),
        'precios': df['precio'].sum(),
        'rentabilidad': df['rentabilidad'].sum(),
    }

    ok = True
    for key in ['costos', 'precios', 'rentabilidad']:
        exp = expected[key]
        act = actual[key]
        diff_pct = abs(act - exp) / max(abs(exp), 1) * 100 if exp != 0 else float('inf')
        status = 'OK' if diff_pct <= 1.0 else 'FAIL'
        if status == 'FAIL':
            ok = False
        if verbose:
            print("  %-14s esperado: %15.2f  extraido: %15.2f  diff: %.4f%%  [%s]" % (
                key.upper(), exp, act, diff_pct, status))

    return ok


# ---------------------------------------------------------------------------
# Main — prueba con sample/central.pdf
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    import sys
    import os

    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    PDF = os.path.join(os.path.dirname(__file__), 'sample', 'central.pdf')
    print("=" * 65)
    print("EXTRAYENDO: %s" % PDF)
    print("=" * 65)

    df, meta, totals = extract_pdf(PDF, verbose=True)

    print("\n--- METADATOS ---")
    print("  Fecha desde : %s" % meta.get('fecha_desde', '?'))
    print("  Fecha hasta : %s" % meta.get('fecha_hasta', '?'))
    print("  Sucursal    : %s" % meta.get('sucursal', '?'))

    print("\n--- TOTALES PDF ---")
    for k, v in totals.items():
        print("  %-15s $ %.2f" % (k, v))

    print("\n--- RESULTADO DE EXTRACCION ---")
    print("  Productos extraidos: %d" % len(df))

    print("\n--- VALIDACION DE SUMAS ---")
    ok = validate(df, totals, verbose=True)

    print("\n--- PRODUCTOS CON MARGEN NEGATIVO ---")
    neg = df[df['margen'] < 0].sort_values('margen')
    if neg.empty:
        print("  Ninguno.")
    else:
        for _, row in neg.iterrows():
            print("  [%s] %-45s  margen: %.2f%%" % (
                row['codigo'], row['descripcion'][:45], row['margen']))

    print("\n--- PRIMERAS 5 FILAS ---")
    cols = ['codigo', 'descripcion', 'unidades', 'costo', 'precio', 'margen']
    print(df[cols].head(5).to_string(index=False))

    print("\n--- ULTIMAS 5 FILAS ---")
    print(df[cols].tail(5).to_string(index=False))

    print("\n--- SUMAS RAPIDAS ---")
    print("  sum(costo)        : $ %.2f" % df['costo'].sum())
    print("  sum(precio)       : $ %.2f" % df['precio'].sum())
    print("  sum(rentabilidad) : $ %.2f" % df['rentabilidad'].sum())

    print("\n%s" % ("EXTRACCION OK" if ok else "EXTRACCION FALLO — revisar warnings"))

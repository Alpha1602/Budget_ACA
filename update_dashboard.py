#!/usr/bin/env python3
"""
update_dashboard.py
===================
Extrae datos de un Excel FCST (hoja "Datos HTML" + "Ventas 2025")
y actualiza todas las variables JS del dashboard index.html.

Uso:
    python3 update_dashboard.py <excel.xlsx> <template.html> [output.html]

Si no se especifica output.html, sobreescribe el template.
"""

import sys
import json
import re
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("ERROR: Instala openpyxl primero:  pip install openpyxl")


# ── Constantes ────────────────────────────────────────────────────────────────
RETAIL_MAP_DH = {
    'ATT': 'AT&T', 'AT&T': 'AT&T', 'Prime': 'AT&T',
    'MacStore': 'MacStore', 'iShop': 'iShop',
    'Liverpool': 'Liverpool', 'Coppel': 'Coppel',
    'Mercado Libre': 'MercadoLibre', 'MercadoLibre': 'MercadoLibre',
}
RETAIL_MAP_25 = {
    'ATT': 'AT&T', 'AT&T': 'AT&T', 'Prime': 'AT&T',
    'MacStore': 'MacStore', 'iShop': 'iShop',
    'Liverpool': 'Liverpool', 'Coppel': 'Coppel',
    'Mercado Libre': 'MercadoLibre', 'MercadoLibre': 'MercadoLibre',
}
ALL_RETAILS = ['AT&T', 'MacStore', 'iShop', 'Liverpool', 'Coppel', 'MercadoLibre']


# ── Helpers ───────────────────────────────────────────────────────────────────
def fl(v):
    """Convierte a float seguro."""
    try:
        return float(v) if v not in (None, '') else 0.0
    except (TypeError, ValueError):
        return 0.0


def quarterly(monthly_dict):
    """Suma mensual → trimestral [Q1, Q2, Q3, Q4]."""
    result = {}
    for r in ALL_RETAILS:
        m = monthly_dict.get(r, [0] * 12)
        result[r] = [round(sum(m[0:3])), round(sum(m[3:6])),
                     round(sum(m[6:9])), round(sum(m[9:12]))]
    return result


def js_retail_obj(d):
    """Genera objeto JS con todos los retails."""
    parts = []
    for r in ALL_RETAILS:
        v = [round(x) for x in d.get(r, [0] * 12)]
        parts.append(f'"{r}":{json.dumps(v)}')
    return '{' + ','.join(parts) + '}'


def js_quarter_obj(q):
    parts = []
    for r in ALL_RETAILS:
        v = q.get(r, [0, 0, 0, 0])
        parts.append(f'"{r}":{json.dumps(v)}')
    return '{' + ','.join(parts) + '}'


# ── Extracción: Datos HTML (2026 FCST) ───────────────────────────────────────
def extract_datos_html(ws):
    """
    Columnas (1-indexed):
      9  = Retail
      6  = Categoría
      1  = Marca
      15-26 = FCST piezas Jan-Dec
      27-38 = Valor MXN Jan-Dec
    """
    fcst_val  = defaultdict(lambda: [0.0] * 12)
    fcst_pcs  = defaultdict(lambda: [0.0] * 12)
    marca_raw = defaultdict(lambda: defaultdict(lambda: defaultdict(
        lambda: {'val': [0.0] * 12, 'pcs': [0.0] * 12}
    )))
    cat_data  = {r: {} for r in ALL_RETAILS}
    all_brands = set()

    for row in ws.iter_rows(min_row=3, values_only=True):
        retail_raw = str(row[8]).strip() if row[8] else ''
        retail = RETAIL_MAP_DH.get(retail_raw)
        if not retail:
            continue

        brand = str(row[0]).strip() if row[0] else ''
        if not brand:
            continue
        if brand.upper() == 'TECH21':
            brand = 'Tech21'
        cat = str(row[5]).strip() if row[5] else 'Other'
        all_brands.add(brand)

        for i in range(12):
            fp = fl(row[16 + i])   # FCST piezas (col 16=Jan..col 27=Dec)
            fv = fl(row[28 + i])   # Valor MXN (col 28=Jan..col 39=Dec)

            fcst_pcs[retail][i] += fp
            fcst_val[retail][i] += fv

            marca_raw[brand][retail][cat]['val'][i] += fv
            marca_raw[brand][retail][cat]['pcs'][i] += fp

            if cat not in cat_data[retail]:
                cat_data[retail][cat] = {'val': [0.0] * 12, 'pcs': [0.0] * 12}
            cat_data[retail][cat]['val'][i] += fv
            cat_data[retail][cat]['pcs'][i] += fp

    return fcst_val, fcst_pcs, marca_raw, cat_data, sorted(all_brands)


# ── Extracción: Ventas 2025 ───────────────────────────────────────────────────
def extract_ventas_2025(ws):
    """
    Columnas (1-indexed):
      1  = Customer (retail)
      9-20  = piezas Jan-Dec
      23-34 = valor MXN Jan-Dec
    """
    actual_val = defaultdict(lambda: [0.0] * 12)
    actual_pcs = defaultdict(lambda: [0.0] * 12)

    for row in ws.iter_rows(min_row=3, values_only=True):
        retail_raw = str(row[0]).strip() if row[0] else ''
        retail = RETAIL_MAP_25.get(retail_raw)
        if not retail:
            continue
        for i in range(12):
            pv = fl(row[8 + i])    # piezas
            vv = fl(row[22 + i])   # valor MXN
            actual_pcs[retail][i] += pv
            actual_val[retail][i] += vv

    return actual_val, actual_pcs


# ── Extracción: Datos HTML → sinMovData (SKUs sin movimiento) ────────────────
def extract_sin_movimiento_consenso(ws):
    """
    Hoja Datos HTML — columnas (0-indexed):
      0  = Marca,  1  = SKU Retail,  4  = Descripción
      5  = Categoría,  6  = Sub Categoría,  7  = Clase
      8  = Retail  ('ATT', 'MacStore', 'iShop', 'Liverpool', 'Coppel')
      11 = Cost unitario,  12 = Status ACA,  13 = Status Retail
      14 = OH Retail (unidades),  15 = OH ACA
      16-27 = FCST piezas Jan-Dec
      28-39 = Valor MXN Jan-Dec

    Condición sin-movimiento: OH Retail > 0  AND  suma(FCST Ago-Dic) == 0
    MOI: OH Retail / avg(FCST Jun col21, FCST Jul col22)
    """
    RETAIL_DH_INV = {'ATT', 'MacStore', 'iShop', 'Liverpool', 'Coppel'}
    records = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 28:
            continue

        retail = str(row[8]).strip() if row[8] else ''
        if retail not in RETAIL_DH_INV:
            continue

        oh_retail = fl(row[14])
        if oh_retail <= 0:
            continue

        # Sin movimiento = sin FCST para meses FUTUROS (Ago-Dic, cols 23-27)
        fcst_future = sum(fl(row[23 + i]) for i in range(5))  # Aug Sep Oct Nov Dec
        if fcst_future > 0:
            continue  # tiene plan de venta futuro → no es sin movimiento

        # FCST 12 meses completos (cols 16-27)
        fcst_months = [fl(row[16 + i]) for i in range(12)]

        marca   = str(row[0]).strip()  if row[0]  else ''
        sku     = str(row[1]).strip()  if row[1]  else ''
        desc    = str(row[4]).strip()  if row[4]  else ''
        cat     = str(row[5]).strip()  if row[5]  else ''
        subcat  = str(row[6]).strip()  if row[6]  else ''
        clase   = str(row[7]).strip()  if row[7]  else ''
        status  = str(row[12]).strip() if row[12] else ''
        oh_aca  = fl(row[15])

        # MOI: avg FCST Jul (col22) y Aug (col23)
        venta_jul = fl(row[22])
        venta_ago = fl(row[23])
        avg_fcst  = round((venta_jul + venta_ago) / 2, 2) if (venta_jul + venta_ago) > 0 else 0.0

        moi = round(oh_retail / avg_fcst, 1) if avg_fcst > 0 else None

        cost_unit    = fl(row[11])
        costo_retail = round(cost_unit * oh_retail, 2)
        val_riesgo   = costo_retail

        # FCST Jul-Dec para columnas de la tabla (cols 22-27)
        fcst_jul = int(fl(row[22])); fcst_aug = int(fl(row[23])); fcst_sep = int(fl(row[24]))
        fcst_oct = int(fl(row[25])); fcst_nov = int(fl(row[26])); fcst_dec = int(fl(row[27]))

        records.append({
            'marca':        marca,
            'sku':          sku,
            'desc':         desc,
            'cat':          cat,
            'subcat':       subcat,
            'clase':        clase,
            'cliente':      retail,
            'cost':         round(fl(row[11]), 2),
            'status_aca':   status,
            'oh_retail':    int(oh_retail),
            'oh_aca':       int(oh_aca),
            'fcst_jul':     fcst_jul,
            'fcst_aug':     fcst_aug,
            'fcst_sep':     fcst_sep,
            'fcst_oct':     fcst_oct,
            'fcst_nov':     fcst_nov,
            'fcst_dec':     fcst_dec,
            'ttl_fcst':     int(sum(fcst_months)),
            'moi':          moi,
            'sin_mov':      avg_fcst,
            'val_riesgo':   round(val_riesgo, 2),
            'costo_retail': costo_retail,
            'costo_landed': costo_retail,
            'costo_usd':    0.0,
        })

    return records




# ── Extracción: Budget 2026 (hoja "Budget 120M") ─────────────────────────────
def extract_budget_2026(ws):
    """
    Hoja 'Budget 120M' (SKU-level, min_row=4):
      col0  = Marca
      col9  = Retail  ('ATT','MacStore','iShop','Liverpool','Coppel')
      col23-34 = Budget piezas Jan-Dec
      col36-47 = Budget $ MXN Jan-Dec

    Devuelve:
      budget_val  {retail: [12 meses MXN]}
      budget_pcs  {retail: [12 meses pzas]}
      budget_marca {marca: {retail: {val:[12], pcs:[12]}}}
    """
    RMAP = {'ATT':'AT&T','iShop':'iShop','Liverpool':'Liverpool',
            'MacStore':'MacStore','Coppel':'Coppel','MercadoLibre':'MercadoLibre'}
    BRAND_NORM = {'TECH21':'Tech21','HYPERGEAR':'Hypergear'}

    def _new():
        return {'val':[0.0]*12,'pcs':[0.0]*12}

    bv = defaultdict(lambda: [0.0]*12)
    bp = defaultdict(lambda: [0.0]*12)
    # marca → retail → {val, pcs}
    bm   = defaultdict(lambda: defaultdict(_new))
    # cat → retail → {val, pcs}
    bc   = defaultdict(lambda: defaultdict(_new))
    # marca → cat → retail → {val, pcs}
    bmc  = defaultdict(lambda: defaultdict(lambda: defaultdict(_new)))

    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row or len(row) < 48 or row[9] is None:
            continue
        retail_key = str(row[9]).strip()
        retail = RMAP.get(retail_key)
        if not retail:
            continue
        marca_raw = str(row[0]).strip() if row[0] else ''
        marca = BRAND_NORM.get(marca_raw.upper(), marca_raw)
        cat   = str(row[5]).strip() if row[5] else 'Others'
        for i in range(12):
            p = fl(row[23 + i]); v = fl(row[36 + i])
            bp[retail][i] += p;           bv[retail][i] += v
            bm[marca][retail]['pcs'][i]         += p
            bm[marca][retail]['val'][i]         += v
            bc[cat][retail]['pcs'][i]            += p
            bc[cat][retail]['val'][i]            += v
            bmc[marca][cat][retail]['pcs'][i]    += p
            bmc[marca][cat][retail]['val'][i]    += v

    def _round(d):
        return {r: {'val':[round(x) for x in e['val']],'pcs':[int(x) for x in e['pcs']]}
                for r, e in d.items()}

    budget_val      = {r: [round(v) for v in vals] for r, vals in bv.items()}
    budget_pcs      = {r: [int(v)   for v in vals] for r, vals in bp.items()}
    budget_marca    = {m: _round(rd)  for m, rd  in bm.items()}
    budget_cat      = {c: _round(rd)  for c, rd  in bc.items()}
    budget_marca_cat = {m: {c: _round(rd) for c, rd in cats.items()}
                        for m, cats in bmc.items()}
    return budget_val, budget_pcs, budget_marca, budget_cat, budget_marca_cat


def build_budget_js(budget_val, budget_pcs, budget_marca, budget_cat, budget_marca_cat):
    import json as _json
    line1 = 'const budgetVal = '         + _json.dumps(budget_val,       ensure_ascii=False, separators=(',',':')) + ';'
    line2 = 'const budgetPcs = '         + _json.dumps(budget_pcs,       ensure_ascii=False, separators=(',',':')) + ';'
    line3 = 'const budgetMarcaRaw = '    + _json.dumps(budget_marca,     ensure_ascii=False, separators=(',',':')) + ';'
    line4 = 'const budgetCatRaw = '      + _json.dumps(budget_cat,       ensure_ascii=False, separators=(',',':')) + ';'
    line5 = 'const budgetMarcaCatRaw = ' + _json.dumps(budget_marca_cat, ensure_ascii=False, separators=(',',':')) + ';'
    return '\n'.join([line1, line2, line3, line4, line5])


# ── Extracción: Todos los SKUs con OH Retail (para tabla MOI completa) ────────
def extract_all_skus_with_oh(ws):
    """
    Todos los SKUs con OH Retail > 0 — sin filtro de FCST.
    Hoja Datos HTML (0-indexed): ver extract_sin_movimiento_consenso.
    MOI: avg FCST Jun (col21) y FCST Jul (col22).
    """
    RETAIL_DH_INV = {'ATT', 'MacStore', 'iShop', 'Liverpool', 'Coppel'}
    records = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 28:
            continue
        retail = str(row[8]).strip() if row[8] else ''
        if retail not in RETAIL_DH_INV:
            continue
        oh_retail = fl(row[14])
        if oh_retail <= 0:
            continue

        marca  = str(row[0]).strip() if row[0] else ''
        sku    = str(row[1]).strip() if row[1] else ''
        desc   = str(row[4]).strip() if row[4] else ''
        cat    = str(row[5]).strip() if row[5] else ''
        oh_aca = fl(row[15])

        # MOI: avg FCST Jul (col22) y Aug (col23)
        venta_jul = fl(row[22])
        venta_ago = fl(row[23])
        avg_pcs   = round((venta_jul + venta_ago) / 2, 1) if (venta_jul + venta_ago) > 0 else 0.0
        moi       = round(oh_retail / avg_pcs, 1) if avg_pcs > 0 else None

        costo_retail = round(fl(row[11]) * oh_retail, 2)

        records.append({
            'cliente':      retail,
            'cat':          cat,
            'marca':        marca,
            'sku':          sku,
            'desc':         desc,
            'oh_retail':    int(oh_retail),
            'oh_aca':       int(oh_aca),
            'avg_pcs':      avg_pcs,
            'moi':          moi,
            'costo_retail': costo_retail,
        })
    return records


def build_all_sku_js(records):
    return 'const allSkuData = ' + json.dumps(records, ensure_ascii=False, separators=(',', ':')) + ';'

def build_sin_mov_js(records):
    """Genera la línea const sinMovData = [...]; lista de registros."""
    return 'const sinMovData = ' + json.dumps(records, ensure_ascii=False, separators=(',', ':')) + ';'


# ── Extracción: moiDataByClient (OH Retail por retail+categoría) ──────────────
RETAILS_CONSENSO = ['ATT', 'MacStore', 'iShop', 'Liverpool', 'Coppel']

def extract_moi_data_by_client(ws_datos_html):
    """
    Consolida OH Retail y FCST mensual por retail+categoría para moiDataByClient.
    Fuente única: Datos HTML.
      Col 8  = Retail, Col 5 = Categoría, Col 0 = Marca
      Col 14 = OH Retail (pzas), Col 15 = OH ACA (pzas), Col 11 = Cost unitario
      Cols 16-27 = FCST pzas Jan-Dec, Cols 28-39 = FCST valor Jan-Dec

    monthly_pcs/val = promedio Jul-Dic FCST (índices 6-11).
    """
    RETAIL_MAP_DH_MOI = {
        'ATT': 'ATT', 'MacStore': 'MacStore', 'iShop': 'iShop',
        'Liverpool': 'Liverpool', 'Coppel': 'Coppel',
    }

    oh_cat   = defaultdict(lambda: {'oh_retail': 0.0, 'oh_aca': 0.0,
                                    'oh_retail_val': 0.0, 'oh_aca_val': 0.0})
    oh_brand = defaultdict(lambda: {'oh_retail': 0.0, 'oh_aca': 0.0,
                                    'oh_retail_val': 0.0, 'oh_aca_val': 0.0})
    fc_cat   = defaultdict(lambda: {'fcst_pcs': [0.0]*12, 'fcst_val': [0.0]*12})
    fc_brand = defaultdict(lambda: {'fcst_pcs': [0.0]*12, 'fcst_val': [0.0]*12})

    # Fuente única: Datos HTML (OH + FCST en una sola pasada)
    for row in ws_datos_html.iter_rows(min_row=3, values_only=True):
        retail_raw = str(row[8]).strip() if row[8] else ''
        retail = RETAIL_MAP_DH_MOI.get(retail_raw)
        if not retail:
            continue
        cat   = str(row[5]).strip() if row[5] else 'Other'
        brand = str(row[0]).strip() if row[0] else 'Other'
        if brand.upper() == 'TECH21':
            brand = 'Tech21'
        oh_r = fl(row[14]); oh_a = fl(row[15])
        cost_unit = fl(row[11])
        rv = cost_unit * oh_r   # OH Retail valor MXN
        av = cost_unit * oh_a   # OH ACA valor MXN
        for r in [retail, 'all']:
            oh_cat[(r, cat)]['oh_retail']     += oh_r
            oh_cat[(r, cat)]['oh_aca']        += oh_a
            oh_cat[(r, cat)]['oh_retail_val'] += rv
            oh_cat[(r, cat)]['oh_aca_val']    += av
            oh_brand[(r, brand)]['oh_retail']     += oh_r
            oh_brand[(r, brand)]['oh_aca']        += oh_a
            oh_brand[(r, brand)]['oh_retail_val'] += rv
            oh_brand[(r, brand)]['oh_aca_val']    += av
        for i in range(12):
            p = fl(row[16 + i]); v = fl(row[28 + i])
            for r in [retail, 'all']:
                fc_cat[(r, cat)]['fcst_pcs'][i]    += p
                fc_cat[(r, cat)]['fcst_val'][i]    += v
                fc_brand[(r, brand)]['fcst_pcs'][i] += p
                fc_brand[(r, brand)]['fcst_val'][i] += v

    def _build_rows(oh_dict, fc_dict, client):
        keys = {k[1] for k in oh_dict if k[0] == client and oh_dict[k]['oh_retail'] > 0}
        rows = []
        for key in sorted(keys):
            oh  = oh_dict[(client, key)]
            fc  = fc_dict.get((client, key), {'fcst_pcs': [0.0]*12, 'fcst_val': [0.0]*12})
            # Ventas promedio Jul-Aug (índices 6-7)
            pcs_h2 = fc['fcst_pcs'][6:8]; val_h2 = fc['fcst_val'][6:8]
            mpcs = sum(pcs_h2) / 2 if any(pcs_h2) else 0.0
            mval = sum(val_h2) / 2 if any(val_h2) else 0.0
            orv = oh['oh_retail_val']; oav = oh['oh_aca_val']
            orp = oh['oh_retail'];     oap = oh['oh_aca']
            rows.append({
                'key':            key,
                'oh_retail':      round(orp, 1),
                'oh_aca':         round(oap, 1),
                'oh_retail_val':  round(orv, 2),
                'oh_aca_val':     round(oav, 2),
                'monthly_pcs':    round(mpcs, 1),
                'monthly_val':    round(mval, 1),
                'moi_retail_val': round(orv / mval, 1) if mval > 0 else None,
                'moi_aca_val':    round(oav / mval, 1) if mval > 0 else None,
                'moi_retail_pcs': round(orp / mpcs, 1) if mpcs > 0 else None,
                'moi_aca_pcs':    round(oap / mpcs, 1) if mpcs > 0 else None,
            })
        rows.sort(key=lambda x: -(x['oh_retail_val'] or 0))
        return rows

    # Construir resultado por cliente con vistas 'cat' y 'brand'
    result = {}
    for client in ['all'] + RETAILS_CONSENSO:
        result[client] = {
            'cat':   _build_rows(oh_cat,   fc_cat,   client),
            'brand': _build_rows(oh_brand, fc_brand, client),
        }
    return result


def build_moi_by_client_js(data):
    return 'const moiDataByClient = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';'


# ── Extracción: Inventario Lyde ───────────────────────────────────────────────
def extract_inventario_lyde(ws):
    """
    Hoja 'Inventario Lyde' — columnas (0-indexed, datos desde min_row=3):
      2  = Marca
      5  = Categoria
      6  = Sub Categoría
      7  = Clase
      10 = Disponible (col K)
      11 = Cuarentena (col L)
      13 = Costo USD
      14 = Costo Landed
      16 = Estatus  ('Activo', 'Activo2', 'New', 'Descontinuado')

    Estructura resultado: {status: {brand|brand_cat|brand_cat_sub|brand_cat_sub_cl: {...métricas}}}
    """
    STATUSES = {'Activo', 'Activo2', 'New', 'Descontinuado'}
    KEYS_VIEWS = ['brand', 'brand_cat', 'brand_cat_sub', 'brand_cat_sub_cl']

    def _empty():
        return {'disp': 0.0, 'cuar': 0.0,
                'usd_disp': 0.0, 'usd_cuar': 0.0,
                'landed_disp': 0.0, 'landed_cuar': 0.0}

    def _acc(d, keys, row_data):
        """Acumula row_data en d[key1][key2]..."""
        for lvl in range(1, len(keys) + 1):
            path_keys = keys[:lvl]
            node = d
            for k in path_keys[:-1]:
                if k not in node:
                    node[k] = {}
                node = node[k]
            leaf_k = path_keys[-1]
            if leaf_k not in node:
                node[leaf_k] = _empty()
            for f, v in row_data.items():
                node[leaf_k][f] += v

    # result[status_or_all][view] = nested dict
    result = {s: {'brand': {}, 'brand_cat': {}, 'brand_cat_sub': {}, 'brand_cat_sub_cl': {}}
              for s in ['all'] + sorted(STATUSES)}

    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or row[2] is None:
            continue
        status = str(row[16]).strip() if row[16] else ''
        # Incluir en 'all' aunque no tenga estatus definido; saltar fila de encabezado interno
        if status and status not in STATUSES:
            continue

        marca  = str(row[2]).strip()
        cat    = str(row[5]).strip() if row[5] else 'Sin Categoría'
        subcat = str(row[6]).strip() if row[6] else 'Sin SubCat'
        clase  = str(row[7]).strip() if row[7] else 'Sin Clase'
        disp   = fl(row[10])
        cuar   = fl(row[11])
        cusd   = fl(row[13])
        clnd   = fl(row[14])

        data = {
            'disp':        disp,
            'cuar':        cuar,
            'usd_disp':    round(cusd * disp, 2),
            'usd_cuar':    round(cusd * cuar, 2),
            'landed_disp': round(clnd * disp, 2),
            'landed_cuar': round(clnd * cuar, 2),
        }

        for grp in (['all'] + ([status] if status in STATUSES else [])):
            r = result[grp]
            # brand level
            if marca not in r['brand']:
                r['brand'][marca] = _empty()
            for f, v in data.items():
                r['brand'][marca][f] += v
            # brand_cat
            if marca not in r['brand_cat']:
                r['brand_cat'][marca] = {}
            if cat not in r['brand_cat'][marca]:
                r['brand_cat'][marca][cat] = _empty()
            for f, v in data.items():
                r['brand_cat'][marca][cat][f] += v
            # brand_cat_sub
            if marca not in r['brand_cat_sub']:
                r['brand_cat_sub'][marca] = {}
            if cat not in r['brand_cat_sub'][marca]:
                r['brand_cat_sub'][marca][cat] = {}
            if subcat not in r['brand_cat_sub'][marca][cat]:
                r['brand_cat_sub'][marca][cat][subcat] = _empty()
            for f, v in data.items():
                r['brand_cat_sub'][marca][cat][subcat][f] += v
            # brand_cat_sub_cl
            if marca not in r['brand_cat_sub_cl']:
                r['brand_cat_sub_cl'][marca] = {}
            if cat not in r['brand_cat_sub_cl'][marca]:
                r['brand_cat_sub_cl'][marca][cat] = {}
            if subcat not in r['brand_cat_sub_cl'][marca][cat]:
                r['brand_cat_sub_cl'][marca][cat][subcat] = {}
            if clase not in r['brand_cat_sub_cl'][marca][cat][subcat]:
                r['brand_cat_sub_cl'][marca][cat][subcat][clase] = _empty()
            for f, v in data.items():
                r['brand_cat_sub_cl'][marca][cat][subcat][clase][f] += v

    # Round all leaf values
    def _round_leaves(obj):
        if isinstance(obj, dict):
            if 'disp' in obj:
                return {k: round(v, 2) for k, v in obj.items()}
            return {k: _round_leaves(v) for k, v in obj.items()}
        return obj

    return _round_leaves(result)


def build_lyde_data_js(data):
    return 'const lydeData = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';'


# ── Genera JS de marcaRaw ─────────────────────────────────────────────────────
def build_marca_raw_js(marca_raw):
    parts_b = []
    for b in sorted(marca_raw.keys()):
        parts_r = []
        for r in ALL_RETAILS:
            if r not in marca_raw[b]:
                continue
            parts_c = []
            for c in sorted(marca_raw[b][r].keys()):
                cd = marca_raw[b][r][c]
                if sum(cd['val']) + sum(cd['pcs']) < 0.01:
                    continue
                v = [round(x) for x in cd['val']]
                p = [round(x) for x in cd['pcs']]
                parts_c.append(f'"{c}":{{"val":{json.dumps(v)},"pcs":{json.dumps(p)}}}')
            if parts_c:
                parts_r.append(f'"{r}":{{{",".join(parts_c)}}}')
        if parts_r:
            parts_b.append(f'"{b}":{{{",".join(parts_r)}}}')
    return 'const marcaRaw = {' + ','.join(parts_b) + '};'


# ── Genera JS de catData ──────────────────────────────────────────────────────
def build_cat_data_js(cat_data):
    result = {}
    for r in ALL_RETAILS:
        result[r] = {}
        for cat, cd in cat_data.get(r, {}).items():
            result[r][cat] = {
                'val': [round(v, 2) for v in cd['val']],
                'pcs': [round(p) for p in cd['pcs']],
            }
    return 'const catData = ' + json.dumps(result) + ';'


# ── Reemplaza una variable JS en el HTML ──────────────────────────────────────
def replace_js_var(html, var_name, new_line):
    lines = html.split('\n')
    new_lines = []
    replaced = 0
    for line in lines:
        if line.strip().startswith(f'const {var_name}') and '=' in line:
            indent = len(line) - len(line.lstrip())
            new_lines.append(' ' * indent + new_line)
            replaced += 1
        else:
            new_lines.append(line)
    return '\n'.join(new_lines), replaced


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    excel_path   = Path(sys.argv[1])
    html_path    = Path(sys.argv[2])
    output_path  = Path(sys.argv[3]) if len(sys.argv) >= 4 else html_path

    if not excel_path.exists():
        sys.exit(f"ERROR: No se encuentra el Excel: {excel_path}")
    if not html_path.exists():
        sys.exit(f"ERROR: No se encuentra el HTML: {html_path}")

    print(f"📂 Excel:    {excel_path.name}")
    print(f"📄 Template: {html_path.name}")
    print(f"💾 Output:   {output_path.name}")
    print()

    # Cargar Excel
    print("⏳ Cargando Excel…")
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)

    if 'Datos HTML' not in wb.sheetnames:
        sys.exit("ERROR: No se encontró la hoja 'Datos HTML' en el Excel.")
    if 'Ventas 2025' not in wb.sheetnames:
        sys.exit("ERROR: No se encontró la hoja 'Ventas 2025' en el Excel.")

    ws_dh = wb['Datos HTML']
    ws_25 = wb['Ventas 2025']

    # ── Datos HTML → sinMovData (fuente de inventario) ───────────────────────────
    print("⏳ Extrayendo SKUs sin movimiento (Datos HTML)…")
    sin_mov_records = extract_sin_movimiento_consenso(ws_dh)
    from collections import Counter
    clients_cnt = Counter(r['cliente'] for r in sin_mov_records)
    print(f"   ✅ sinMovData: {len(sin_mov_records)} registros → {dict(clients_cnt)}")

    # Extraer datos
    print("⏳ Extrayendo datos 2026 (Datos HTML)…")
    fcst_val, fcst_pcs, marca_raw, cat_data, all_brands = extract_datos_html(ws_dh)

    print("⏳ Extrayendo datos 2025 (Ventas 2025)…")
    actual_val, actual_pcs = extract_ventas_2025(ws_25)

    # Calcular trimestrales
    qv   = quarterly(fcst_val)
    qp   = quarterly(fcst_pcs)
    q25v = quarterly(actual_val)
    q25p = quarterly(actual_pcs)

    # Validar totales
    grand_total_2026 = round(sum(sum(v) for v in fcst_val.values()))
    grand_pcs_2026   = round(sum(sum(p) for p in fcst_pcs.values()))
    grand_total_2025 = round(sum(sum(v) for v in actual_val.values()))
    print(f"   ✅ 2026 MXN total:  ${grand_total_2026:>15,.0f}")
    print(f"   ✅ 2026 pzas total: {grand_pcs_2026:>15,.0f}")
    print(f"   ✅ 2025 MXN total:  ${grand_total_2025:>15,.0f}")
    print()

    # Construir líneas JS
    vars_to_replace = {
        'avb2026Val':  f'const avb2026Val = {js_retail_obj(fcst_val)};',
        'avb2026Pcs':  f'const avb2026Pcs = {js_retail_obj(fcst_pcs)};',
        'valorFull':   f'const valorFull = {js_retail_obj(fcst_val)};',
        'piezasAnual': f'const piezasAnual = {js_retail_obj(fcst_pcs)};',
        'avb2025Val':  f'const avb2025Val = {js_retail_obj(actual_val)};',
        'avb2025Pcs':  f'const avb2025Pcs = {js_retail_obj(actual_pcs)};',
        'qValorAll':   f'const qValorAll = {js_retail_obj(qv)};',
        'qPiezas':     f'const qPiezas   = {js_retail_obj(qp)};',
        'q2025Val':    f'const q2025Val  = {js_retail_obj(q25v)};',
        'q2025Pcs':    f'const q2025Pcs  = {js_retail_obj(q25p)};',
        'ALL_BRANDS':  f'const ALL_BRANDS = {json.dumps(all_brands)};',
    }

    # Leer HTML
    print("⏳ Actualizando HTML…")
    html = html_path.read_text(encoding='utf-8')

    # Reemplazar variables simples
    for var_name, new_line in vars_to_replace.items():
        html, n = replace_js_var(html, var_name, new_line)
        status = '✅' if n > 0 else '⚠️ NOT FOUND'
        print(f"   {status} {var_name}")

    # Reemplazar marcaRaw (puede ser muy larga)
    html, n = replace_js_var(html, 'marcaRaw', build_marca_raw_js(marca_raw))
    print(f"   {'✅' if n > 0 else '⚠️ NOT FOUND'} marcaRaw ({len(all_brands)} marcas)")

    # Reemplazar catData
    html, n = replace_js_var(html, 'catData', build_cat_data_js(cat_data))
    print(f"   {'✅' if n > 0 else '⚠️ NOT FOUND'} catData")

    # Reemplazar allSkuData (todos los SKUs con OH) desde Datos HTML
    all_sku_records = extract_all_skus_with_oh(ws_dh)
    ALL_SKU_RE = re.compile(r'const allSkuData\s*=\s*\[.*?\];', re.DOTALL)
    new_all_sku = build_all_sku_js(all_sku_records)
    html, n_ask = ALL_SKU_RE.subn(new_all_sku, html, count=1)
    print(f"   {'✅' if n_ask else '⚠️ NOT FOUND'} allSkuData ({len(all_sku_records)} SKUs totales)")

    # Reemplazar sinMovData desde Datos HTML
    if sin_mov_records:
        SINMOV_RE = re.compile(r'const sinMovData\s*=\s*\[.*?\];', re.DOTALL)
        new_sinmov = build_sin_mov_js(sin_mov_records)
        html, n_sm = SINMOV_RE.subn(new_sinmov, html, count=1)
        print(f"   {'✅' if n_sm else '⚠️ NOT FOUND'} sinMovData ({len(sin_mov_records)} SKUs desde Datos HTML)")

    # Reemplazar moiDataByClient desde Datos HTML
    print("⏳ Calculando moiDataByClient (OH por retail+categoría)…")
    moi_by_client = extract_moi_data_by_client(ws_dh)
    if True:
        MOI_BC_RE = re.compile(r'const moiDataByClient\s*=\s*\{.*?\};', re.DOTALL)
        new_mbc = build_moi_by_client_js(moi_by_client)
        html, n_mbc = MOI_BC_RE.subn(new_mbc, html, count=1)
        cats_all = len(moi_by_client.get('all', {}).get('cat', []))
        print(f"   {'✅' if n_mbc else '⚠️ NOT FOUND'} moiDataByClient ({cats_all} categorías)")

    # Reemplazar budgetVal / budgetPcs / budgetMarcaRaw (hoja Budget 120M)
    budget_sheet = 'Budget 120M' if 'Budget 120M' in wb.sheetnames else 'Budget 2026'
    if budget_sheet in wb.sheetnames:
        bval, bpcs, bmarca, bcat, bmc = extract_budget_2026(wb[budget_sheet])
        BUDGET_RE = re.compile(
            r'const budgetVal\s*=\s*\{.*?\};\nconst budgetPcs\s*=\s*\{.*?\};\n'
            r'const budgetMarcaRaw\s*=\s*\{.*?\};\nconst budgetCatRaw\s*=\s*\{.*?\};\n'
            r'const budgetMarcaCatRaw\s*=\s*\{.*?\};',
            re.DOTALL
        )
        new_budget = build_budget_js(bval, bpcs, bmarca, bcat, bmc)
        html, n_b = BUDGET_RE.subn(new_budget, html, count=1)
        total_b = sum(sum(v) for v in bval.values())
        print(f"   {'✅' if n_b else '⚠️ NOT FOUND'} budget vars (${total_b/1e6:.1f}M, {len(bmarca)} marcas, {len(bcat)} cats)")
    else:
        print("   ⚠️  Hoja 'Budget 120M' no encontrada — budget no actualizado")

    # Reemplazar lydeData desde Inventario Lyde (col K = Disponible)
    if 'Inventario Lyde' in wb.sheetnames:
        print("⏳ Extrayendo Inventario Lyde…")
        lyde_data = extract_inventario_lyde(wb['Inventario Lyde'])
        LYDE_RE = re.compile(r'const lydeData\s*=\s*\{.*?\};', re.DOTALL)
        new_lyde = build_lyde_data_js(lyde_data)
        html, n_lyde = LYDE_RE.subn(new_lyde, html, count=1)
        total_lyde_disp = lyde_data.get('all', {}).get('brand', {})
        total_disp = sum(v.get('disp', 0) for v in total_lyde_disp.values())
        print(f"   {'✅' if n_lyde else '⚠️ NOT FOUND'} lydeData ({total_disp:,.0f} unidades disponibles)")
    else:
        print("   ⚠️  Hoja 'Inventario Lyde' no encontrada — lydeData no actualizado")

    # Guardar
    output_path.write_text(html, encoding='utf-8')
    print()
    print(f"✅ Dashboard actualizado → {output_path}")
    print(f"   Tamaño: {output_path.stat().st_size / 1024:.0f} KB")


if __name__ == '__main__':
    main()

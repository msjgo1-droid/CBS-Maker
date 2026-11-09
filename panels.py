"""Fabric panel (marker) sheets: read MNMK, crop marker image, build/update buyer panel sheets."""
import io
import re
from copy import copy

import numpy as np
from PIL import Image

from xlutil import copy_sheet, add_image

LABELS = {'date': 'DATE', 'style': 'STYLE', 'fab': 'FAB DETAIL', 'size': 'BASIC SIZE', 'gqty': "GARMENT Q'TY",
          'width': 'CUTTABLE WIDTH', 'marker': 'TOTAL MARKER LENGTH', 'eff': 'EFFICIENCY',
          'shrink': 'SHRINKAGE', 'pqty': "PANEL Q'TY"}
SKIP_SHEETS = ('form', 'inter')          # MNMK sheets never turned into buyer panel sheets


def _s(v):
    return str(v).strip() if v is not None else ''


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def find_fields(ws, r0=1, r1=None):
    """{key: (row, value_col)} for the first block found between r0..r1."""
    r1 = r1 or min(ws.max_row, r0 + 12)
    out = {}
    for row in ws.iter_rows(min_row=r0, max_row=r1, max_col=14):
        for c in row:
            v = _s(c.value).upper()
            for k, lab in LABELS.items():
                if k not in out and v == lab:
                    out[k] = (c.row, c.column + 2)
    return out


def is_panel_sheet(ws):
    f = find_fields(ws, 1, 12)
    return 'marker' in f and 'gqty' in f


# ---------------- MNMK ----------------
def mnmk_blocks(ws):
    starts = [c.row for c in ws['E'] if _s(c.value).upper() == 'TOTAL MARKER LENGTH']
    blocks = []
    for i, s in enumerate(starts):
        r0 = max(1, s - 2)
        r1 = (starts[i + 1] - 3) if i + 1 < len(starts) else ws.max_row
        f = find_fields(ws, r0, min(r1, r0 + 10))
        val = lambda k, off=0: ws.cell(f[k][0], f[k][1] + off).value if k in f else None
        imgs = [im for im in ws._images if r0 - 1 <= im.anchor._from.row <= r1]
        blocks.append(dict(r0=r0, r1=r1, date=val('date'), width=val('width'), style=val('style'),
                           marker_y=val('marker'), marker_in=val('marker', 2), fab=val('fab'),
                           eff=val('eff'), size=val('size'), shrink=val('shrink'), gqty=val('gqty'),
                           pqty=val('pqty'), image=imgs[0] if imgs else None))
    return blocks


def pick_block(blocks, gender):
    want = {'M': 'SIZE M', 'W': 'SIZE S'}.get(gender)
    if want:
        for b in blocks:
            if want in _s(b['size']).upper():
                return b, None
    b = blocks[-1] if blocks else None
    return b, (dict(code='size_block', want=want) if want else None)


def read_mnmk(wb_values):
    """wb loaded data_only=True (keeps images). Returns sheets list + warnings."""
    out, warns = [], []
    hidden = [ws.title for ws in wb_values.worksheets if ws.sheet_state != 'visible']
    if hidden:
        warns.append(dict(code='hidden', src='MNMK', names=', '.join(hidden)))
    for ws in wb_values.worksheets:
        if ws.sheet_state != 'visible' or any(ws.title.strip().lower().startswith(s) for s in SKIP_SHEETS):
            continue
        blocks = mnmk_blocks(ws)
        if not blocks:
            continue
        out.append(dict(title=ws.title.strip(), blocks=blocks))
    return out, warns


def _longest_run(valid, gap=30):
    idx = np.where(valid)[0]
    if not len(idx):
        return None
    runs, s0, prev = [], idx[0], idx[0]
    for i in idx[1:]:
        if i - prev > gap:
            runs.append((s0, prev))
            s0 = i
        prev = i
    runs.append((s0, prev))
    return max(runs, key=lambda r: r[1] - r[0])


def crop_marker(data, pad=2):
    """Crop a CAD (Gerber/Optitex) screenshot down to the marker canvas (light-grey #F9F9F9 background)."""
    im = Image.open(io.BytesIO(data)).convert('RGB')
    a = np.asarray(im).astype(int)
    vals, counts = np.unique(a.reshape(-1, 3)[::7], axis=0, return_counts=True)
    order = np.argsort(-counts)
    bg = None
    for k in order[:6]:                              # most common light, non-white, neutral colour
        r, g, b = vals[k]
        if 235 <= r <= 252 and abs(r - g) < 4 and abs(r - b) < 4:
            bg = vals[k]
            break
    if bg is None:
        return data, im.size, im.size
    sat = (a.max(axis=2) - a.min(axis=2)) > 45
    m = (np.abs(a - bg).max(axis=2) <= 3) | sat          # canvas = background or coloured pieces
    rr = _longest_run(m.sum(axis=1) >= 0.25 * a.shape[1], gap=6)
    if not rr:
        return data, im.size, im.size
    band = m[rr[0]:rr[1] + 1]
    cr = _longest_run(band.sum(axis=0) >= 0.5 * band.shape[0], gap=6)
    if not cr or (cr[1] - cr[0]) < 50 or (rr[1] - rr[0]) < 20:
        return data, im.size, im.size
    x0, x1 = max(cr[0] - pad, 0), min(cr[1] + pad, a.shape[1] - 1)
    y0, y1 = max(rr[0] - pad, 0), min(rr[1] + pad, a.shape[0] - 1)
    c = im.crop((x0, y0, x1 + 1, y1 + 1))
    buf = io.BytesIO()
    c.save(buf, 'PNG')
    return buf.getvalue(), im.size, c.size


# ---------------- buyer panel sheets ----------------
def set_field(ws, f, key, value, off=0):
    if key in f:
        r, col = f[key]
        ws.cell(r, col + off).value = value


def build_from_mnmk(dst_wb, layout_ws, title, block, index=None, crop=True):
    """Create a buyer panel sheet from a layout sheet + MNMK block."""
    ws = copy_sheet(layout_ws, dst_wb, title[:31], index)
    for im in list(ws._images):
        ws._images.remove(im)
    f = find_fields(ws, 1, 12)
    for k in ('date', 'width', 'style', 'fab', 'eff', 'size', 'shrink', 'gqty', 'pqty'):
        set_field(ws, f, k, block[k])
    set_field(ws, f, 'marker', block['marker_y'])
    set_field(ws, f, 'marker', block['marker_in'], 2)
    # clear YY history values under the 'DATE / SAMPLE STAGE / YY' table
    if block['image'] is not None:
        data = block['image']._data()
        if crop:
            data, _, size = crop_marker(data)
        else:
            size = Image.open(io.BytesIO(data)).size
        anchor_row = (f['gqty'][0] + 1) if 'gqty' in f else 8
        w = min(size[0], 900)
        add_image(ws, data, f'A{anchor_row + 1}', w, int(size[1] * w / size[0]))
    return ws


def recompute_marker(ws, yy):
    """Back-calculate marker length so ((IN/36)+Y)/qty == yy. Returns (Y, IN)."""
    f = find_fields(ws, 1, 12)
    q = _num(ws.cell(*f['gqty']).value)
    if not q or yy is None:
        return None
    total_in = round(yy * q * 36, 4)
    y = int(total_in // 36)
    inch = round(total_in - y * 36, 4)
    r, col = f['marker']
    ws.cell(r, col).value = y
    ws.cell(r, col + 2).value = inch
    return y, inch


def panel_yy(ws):
    f = find_fields(ws, 1, 12)
    r, col = f['marker']
    y, i = _num(ws.cell(r, col).value) or 0, _num(ws.cell(r, col + 2).value) or 0
    q = _num(ws.cell(*f['gqty']).value)
    return round((i / 36 + y) / q, 4) if q else None


def panel_codes(ws):
    f = find_fields(ws, 1, 12)
    txt = _s(ws.cell(*f['fab']).value) if 'fab' in f else ''
    return {t for t in re.split(r'[\s/,()#]+', txt.upper()) if len(t) >= 5 and re.search(r'\d', t)}


def one_pc_one_way(text):
    t = re.sub(r'\s+', '', _s(text).upper())
    return '1PC1WAY' in t or ('1PC' in t and '1WAY' in t)

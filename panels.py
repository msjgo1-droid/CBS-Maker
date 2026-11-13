"""Fabric panel (marker) sheets: read MNMK, crop marker image, build/update buyer panel sheets."""
import io
import re
from copy import copy

import numpy as np
from PIL import Image

from xlutil import copy_sheet, add_image, img_bytes

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


def shrink_pct(v):
    """'L : 5%  W : 5%' / '5%' / 0.05 -> 5.0 (largest number found)."""
    if isinstance(v, (int, float)):
        return v * 100 if v < 1 else v
    nums = [float(x) for x in re.findall(r'(\d+(?:\.\d+)?)\s*%', _s(v))]
    return max(nums) if nums else None


def pick_block(blocks, gender):
    """MNMK usually has a 3% and a 5% shrinkage marker -> always the 5% one; then costing size (M men / S women)."""
    if not blocks:
        return None, None
    five = [b for b in blocks if shrink_pct(b['shrink']) == 5]
    pool = five or blocks
    want = {'M': 'SIZE M', 'W': 'SIZE S'}.get(gender)
    if want:
        for b in pool:
            if want in _s(b['size']).upper():
                return b, None
    warn = dict(code='no_5pct') if not five else (dict(code='size_block', want=want) if want else None)
    return pool[-1], warn


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


def _crop_by_background(a, im, pad):
    vals, counts = np.unique(a.reshape(-1, 3)[::7], axis=0, return_counts=True)
    bg = None
    for k in np.argsort(-counts)[:6]:                # most common light, non-white, neutral colour
        r, g, b = vals[k]
        if 235 <= r <= 252 and abs(r - g) < 4 and abs(r - b) < 4:
            bg = vals[k]
            break
    if bg is None:
        return None
    sat = (a.max(axis=2) - a.min(axis=2)) > 45
    m = (np.abs(a - bg).max(axis=2) <= 3) | sat
    rr = _longest_run(m.sum(axis=1) >= 0.25 * a.shape[1], gap=6)
    if not rr:
        return None
    band = m[rr[0]:rr[1] + 1]
    cr = _longest_run(band.sum(axis=0) >= 0.5 * band.shape[0], gap=6)
    if not cr or (cr[1] - cr[0]) < 50 or (rr[1] - rr[0]) < 20:
        return None
    return cr[0], rr[0], cr[1], rr[1]


def _crop_by_pieces(a, pad):
    """White-background CAD (Gerber AccuMark): box around the coloured pieces, ignoring thin grid lines,
    the status bar and the taskbar."""
    from scipy import ndimage
    sat = (a.max(axis=2) - a.min(axis=2)) > 40
    sat = ndimage.binary_opening(sat, iterations=2)          # drop 1-3px pink/blue grid lines
    rowc = sat.sum(axis=1)
    rr = _longest_run(rowc >= max(3, 0.01 * a.shape[1]), gap=8)
    if not rr or rr[1] - rr[0] < 20:
        return None
    rr = _trim_strips(rowc, rowc >= max(3, 0.01 * a.shape[1]), rr, 0.3) or rr
    band = sat[rr[0]:rr[1] + 1]
    col = band.sum(axis=0)
    cr = _longest_run(col >= max(2, 0.03 * band.shape[0]), gap=15)
    if not cr or cr[1] - cr[0] < 30:
        return None
    # drop a thin UI strip (blue side-panel header etc.) stuck to either end across a small gap
    cr = _trim_strips(col, col >= max(2, 0.03 * band.shape[0]), cr, 0.3) or cr
    return cr[0], rr[0], cr[1], rr[1]


def _trim_strips(cnt, on, run, strong):
    """Within run, split into on-segments; drop weak end segments (peak < frac x run peak) separated by a gap of >=5 px."""
    x0, x1 = int(run[0]), int(run[1])
    strong = strong * cnt[x0:x1 + 1].max()
    segs, s0 = [], None
    for x in range(x0, x1 + 2):
        v = x <= x1 and on[x]
        if v and s0 is None:
            s0 = x
        elif not v and s0 is not None:
            segs.append([s0, x - 1]); s0 = None
    if not segs:
        return None
    while len(segs) > 1 and cnt[segs[0][0]:segs[0][1] + 1].max() < strong and segs[1][0] - segs[0][1] >= 5:
        segs.pop(0)
    while len(segs) > 1 and cnt[segs[-1][0]:segs[-1][1] + 1].max() < strong and segs[-1][0] - segs[-2][1] >= 5:
        segs.pop()
    return segs[0][0], segs[-1][1]


def _colours(a, box):
    """Number of distinct saturated colours (each >= 1% of the box) - a marker has many piece colours."""
    x0, y0, x1, y1 = [int(v) for v in box]
    b = a[y0:y1 + 1, x0:x1 + 1]
    sat = (b.max(axis=2) - b.min(axis=2)) > 40
    if not sat.any():
        return 0
    q = ((b[..., 0] // 48) * 36 + (b[..., 1] // 48) * 6 + (b[..., 2] // 48))[sat]
    _, c = np.unique(q, return_counts=True)
    return int((c >= 0.01 * b.shape[0] * b.shape[1]).sum())


def crop_marker(data, pad=3):
    """Crop a CAD screenshot down to the marker only."""
    im = Image.open(io.BytesIO(data)).convert('RGB')
    a = np.asarray(im).astype(int)
    box = _crop_by_background(a, im, pad)
    if box is not None and _colours(a, box) < 3:
        box = None                    # a highlighted Excel cell / text box, not a marker
    if box is None or (box[2] - box[0]) * (box[3] - box[1]) > 0.85 * a.shape[0] * a.shape[1]:
        box = _crop_by_pieces(a, pad) or box
    else:
        # dark pieces (navy, dark green) can break the background run -> widen to the pieces on the same rows
        pb = _crop_by_pieces(a, pad)
        area_ = lambda b: max(0, b[2] - b[0]) * max(0, b[3] - b[1])
        if pb and pb[0] >= box[0] - 5 and pb[1] >= box[1] - 5 and pb[2] <= box[2] + 5 and pb[3] <= box[3] + 5 \
                and area_(pb) < 0.6 * area_(box) and pb[1] - box[1] > 40 and _colours(a, pb) >= 6 > _colours(a, box):
            box, pb = pb, None          # the 'background' box took the white table above the marker too
        inside = False
        if pb:
            area = lambda b: max(0, b[2] - b[0]) * max(0, b[3] - b[1])
            ov = area((max(pb[0], box[0]), max(pb[1], box[1]), min(pb[2], box[2]), min(pb[3], box[3])))
            inside = ov >= 0.7 * area(box) and area(pb) < 0.85 * a.shape[0] * a.shape[1]
        if pb and (inside or (abs(pb[1] - box[1]) <= 15 and abs(pb[3] - box[3]) <= 15)) and \
                (pb[0] < box[0] - 10 or pb[2] > box[2] + 10 or pb[1] < box[1] - 10 or pb[3] > box[3] + 10):
            box = (min(box[0], pb[0]), min(box[1], pb[1]), max(box[2], pb[2]), max(box[3], pb[3]))
    if box is None:
        return data, im.size, im.size
    x0, y0, x1, y1 = box
    x0, y0 = max(x0 - pad, 0), max(y0 - pad, 0)
    x1, y1 = min(x1 + pad, a.shape[1] - 1), min(y1 + pad, a.shape[0] - 1)
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
        if block.get(k) is not None:
            set_field(ws, f, k, block[k])
    if block.get('marker_y') is not None:
        set_field(ws, f, 'marker', block['marker_y'])
    if block.get('marker_in') is not None:
        set_field(ws, f, 'marker', block['marker_in'], 2)
    # clear YY history values under the 'DATE / SAMPLE STAGE / YY' table
    if block['image'] is not None:
        data = block['image'] if isinstance(block['image'], (bytes, bytearray)) else img_bytes(block['image'])
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

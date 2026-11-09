"""Fill a HOKA buyer CBS main sheet (template / Solid / Heather) from an internal CBS block."""
import re
from copy import copy

from cbs_format import map_columns, find_label
from xlutil import insert_row


def _s(v):
    return str(v).strip() if v is not None else ''


def _norm(v):
    return re.sub(r'[^a-z0-9]', '', _s(v).lower())


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


EXCLUDE_ADD = ('interlining',)           # never added as new rows (immaterial cost)

# buyer section -> internal category key
SECTION_KEY = {'body': 'fabric', 'lining - apparel': 'fabric', 'lining': 'fabric', 'trims': 'trim',
               'trim': 'trim', 'embellishments': 'emb', 'embellishment': 'emb', 'thread / stitching': 'thread',
               'thread': 'thread', 'finishing': 'any', 'labels': 'label', 'packaging': 'pack'}
FABRIC_SECTIONS = {k for k, v in SECTION_KEY.items() if v == 'fabric'}


def cat_key(cat):
    c = _s(cat).lower()
    for k, v in (('trim', 'trim'), ('embell', 'emb'), ('thread', 'thread'), ('stitch', 'thread'),
                 ('label', 'label'), ('pack', 'pack')):
        if k in c:
            return v
    return 'fabric'


# ---------------- internal side ----------------
def read_internal_block(ws, r0, r1):
    """ws loaded with data_only=True."""
    hdr = find_label(ws, r0, min(r0 + 5, r1), {'S/#'}).row
    cols = map_columns(ws, hdr)
    net = find_label(ws, hdr + 1, r1, {'NET'})
    others = find_label(ws, hdr + 1, r1, {'OTHERS'})
    stop = min(x.row for x in (net, others) if x)
    lines, cat, cm, after_cm = [], 'fabric', None, False
    for r in range(hdr + 1, stop):
        d = _s(ws.cell(r, cols['i_item']).value)
        e = _s(ws.cell(r, cols['i_detail']).value)
        if _norm(d) == 'cm':
            cm = _num(ws.cell(r, cols['b_ttl']).value)
            after_cm = True
            continue
        if d and after_cm:
            cat = d
        price = _num(ws.cell(r, cols['b_price']).value)
        if price is None or not (e or d):
            continue
        lines.append(dict(row=r, cat=cat if after_cm else 'fabric', key=cat_key(cat) if after_cm else 'fabric',
                          item=d, detail=e, remark=_s(ws.cell(r, cols['i_remark']).value),
                          price=price,
                          yy=_num(ws.cell(r, cols['b_yy']).value) or 0,
                          loss=_num(ws.cell(r, cols['b_loss']).value) or 1,
                          cif=_num(ws.cell(r, cols['b_cif']).value) if 'b_cif' in cols else None,
                          used=False))
    from cbs_format import profit_labels
    rates = {}
    for key, c in profit_labels(ws, hdr + 1, r1).items():
        if c:
            for col in range(c.column + 1, cols['b_ttl']):
                v = _num(ws.cell(c.row, col).value)
                if v is not None and v < 1:
                    rates[key] = v
                    break
    return dict(lines=lines, cm=cm, rates=rates)


# ---------------- buyer side ----------------
HDR = {'component': 'component', 'material': 'material', 'legacy_name': 'legacymaterialname',
       'legacy_id': 'legacymaterialid', 'category': 'costingcategory', 'supplier': 'supplier',
       'width': 'cuttablewidth', 'uom': 'uom', 'net': 'netusage', 'loss': 'loss', 'gross': 'grossusage',
       'price': 'price', 'override': 'priceoverride', 'freight': 'freight', 'freight_ovr': 'freightoverride',
       'unit': 'unitprice', 'comments': 'costingcomments', 'placement': 'placement'}
REQUIRED = ('component', 'material', 'net', 'loss', 'price', 'override', 'freight_ovr', 'unit')


def _hkey(v):
    return re.sub(r'[^a-z]', '', _s(v).lower())          # '% Loss' / '%Loss' -> 'loss'


def buyer_columns(ws):
    for row in ws.iter_rows(max_row=25, max_col=1):
        if _s(row[0].value) == 'Component':
            h = row[0].row
            m = {}
            for c in ws[h]:
                k = _hkey(c.value)
                if k and k not in m:
                    m[k] = c.column
            cols = {k: m.get(v) for k, v in HDR.items()}
            miss = [k for k in REQUIRED if not cols.get(k)]
            if miss:
                raise ValueError(f'[{ws.title}] buyer template columns not found: {miss}')
            return h, cols
    raise ValueError(f'[{ws.title}] "Component" header not found')


def material_text(ws, r, c):
    return ' '.join(_s(ws.cell(r, c[k]).value) for k in ('material', 'legacy_name', 'legacy_id') if c.get(k))


def is_main_sheet(ws):
    try:
        buyer_columns(ws)
        return True
    except ValueError:
        return False


def buyer_rows(ws):
    """[(row, component, section)] for component rows; plus section header rows & LOP start."""
    h, c = buyer_columns(ws)
    rows, sections, sec, lop = [], [], None, None
    for r in range(h + 1, ws.max_row + 1):
        a = _s(ws.cell(r, 1).value)
        al = a.lower()
        if al == 'lop':
            lop = r
            break
        u = _s(ws.cell(r, c['unit']).value).upper()
        if al in SECTION_KEY or (a and not material_text(ws, r, c).strip() and u.startswith('=SUM')):
            sec = al if al in SECTION_KEY else al
            sections.append((r, al))
            continue
        if not a and not material_text(ws, r, c).strip():
            continue
        rows.append((r, a, sec))
    return rows, sections, lop


def codes(text):
    return {t for t in re.split(r'[\s/,()#]+', _s(text).upper()) if len(t) >= 5 and re.search(r'\d', t)}


ALIAS = {'china': 'chinese', 'certification': 'certificate', 'resales': 'resale',
         'chain': 'thread', 'stitch': 'thread', 'contrast': 'thread', 'wb': 'waistband'}
STOP = {'the', 'for', 'on', 'of', 'with', 'order', 'label', 'sms', 'in', 'at'}


def _words(t):
    w = {ALIAS.get(x, x) for x in re.findall(r'[a-z]+', _s(t).lower())}
    return {x for x in w if len(x) > 1 and x not in STOP}


def _same_sec(ln, sec):
    k = SECTION_KEY.get(sec, 'any')
    return k == 'any' or ln['key'] == k or (k == 'fabric' and ln['key'] == 'fabric')


def match_line(comp, material, buyer_price, buyer_loss, sec, lines, manual=None):
    if manual is not None:
        return manual, 'manual'
    free = [l for l in lines if not l['used']]
    mc = codes(material)
    for ln in free:                                    # 1) material code
        if (codes(ln['detail']) | codes(ln['remark'])) & mc:
            return ln, 'code'
    comp = re.sub(r'^[A-Za-z /&-]+:\s*', '', comp)       # 'Packaging: Pin' -> 'Pin'
    nc = _norm(re.sub(r'\(.*?\)', '', comp))
    for ln in free:                                    # 2) component name
        for cand in (ln['detail'], ln['item']):
            n = _norm(cand)
            if n and (n == nc or (len(nc) > 3 and nc in n) or (len(n) > 3 and n in nc)):
                return ln, 'name'
    cw = _words(comp)                                  # 3) word overlap (same section)
    best, score = None, 0
    for ln in free:
        if not _same_sec(ln, sec):
            continue
        sc = len(cw & (_words(ln['detail']) | _words(ln['item']))) / max(len(cw), 1)
        if sc > score:
            best, score = ln, sc
    if best and score >= 0.5:
        return best, f'fuzzy {score:.0%}'
    mw = _words(material)                              # 3b) internal item name found in buyer material text
    best, score = None, 0
    for ln in free:
        if not _same_sec(ln, sec):
            continue
        iw = _words(ln['detail'])
        sc = len(iw & mw) / max(len(iw), 1)
        if sc > score:
            best, score = ln, sc
    if best and score >= 0.6:
        return best, f'material {score:.0%}'
    if buyer_price:                                    # 4) numbers: buyer price / loss (same section)
        cands = [l for l in free if _same_sec(l, sec) and abs(l['price'] - buyer_price) < 1e-6]
        if buyer_loss is not None and len(cands) > 1:
            cands = [l for l in cands if abs((l['loss'] - 1) - buyer_loss) < 1e-6] or cands
        if len(cands) == 1:
            return cands[0], 'price'
    return None, None


def _set(ws, r, col, val, changes, sheet):
    if not col:
        return
    c = ws.cell(r, col)
    old = c.value
    same = (old == val) or (_num(old) is not None and _num(val) is not None and abs(old - val) < 1e-9) \
        or (old in (None, '') and val in (None, ''))
    if not same:
        changes.append(dict(sheet=sheet, cell=c.coordinate, item=_s(ws.cell(r, 1).value), old=old, new=val))
    c.value = val


def fill_sheet(ws, blk, variant=None, manual=None, add_missing=True):
    """variant: 'solid'/'heather'/None. manual: {buyer_row: internal_line_index or -1 (skip)}.
    Returns dict(log, changes, unplaced, added)."""
    manual = manual or {}
    h, c = buyer_columns(ws)
    lines = blk['lines']
    for ln in lines:
        ln['used'] = False
    log, changes = [], []
    rows, sections, lop = buyer_rows(ws)
    other = {'solid': 'heather', 'heather': 'solid'}.get(variant)
    # manual assignments first so auto matching doesn't steal them
    for r, li in manual.items():
        if li is not None and li >= 0:
            lines[li]['used'] = True
    for r, comp, sec in rows:
        if other and f'({other})' in comp.lower():
            _set(ws, r, c['net'], 0, changes, ws.title)
            _set(ws, r, c['loss'], 0, changes, ws.title)
            log.append(dict(row=r, component=comp, match='—', how=f'zero ({other})'))
            continue
        if manual.get(r) == -1:
            log.append(dict(row=r, component=comp, match='—', how='skip (manual)'))
            continue
        bp = _num(ws.cell(r, c['price']).value) or 0
        bl = _num(ws.cell(r, c['loss']).value)
        man = lines[manual[r]] if manual.get(r) is not None and manual[r] >= 0 else None
        ln, how = match_line(comp, material_text(ws, r, c), bp, bl, sec, lines, man)
        if not ln:
            mc = codes(material_text(ws, r, c))
            twin = next((l for l in lines if l['used'] and (codes(l['detail']) | codes(l['remark'])) & mc), None)
            if twin and mc:
                # same fabric already put on another row (e.g. Shell 1 / Contrast 1): usage stays on the first row
                _set(ws, r, c['net'], 0, changes, ws.title)
                _set(ws, r, c['loss'], 0, changes, ws.title)
                log.append(dict(row=r, component=comp, match=f"{twin['item'] or twin['cat']} / {twin['detail']}",
                                how='shared (0)', line=lines.index(twin)))
                continue
            blank = not comp and not (_num(ws.cell(r, c['net']).value) or 0)
            log.append(dict(row=r, component=comp, match='', how='empty row' if blank else 'NO MATCH'))
            continue
        ln['used'] = True
        write_line(ws, r, c, ln, bp, changes)
        log.append(dict(row=r, component=comp, match=f"{ln['item'] or ln['cat']} / {ln['detail']}", how=how,
                        line=lines.index(ln)))
    # missing rows
    added = []
    unplaced = [l for l in lines if not l['used'] and l['yy'] and
                not any(x in (l['detail'] + l['item']).lower() for x in EXCLUDE_ADD)]
    if add_missing:
        for ln in unplaced:
            r = add_row(ws, ln, c, sections)
            if r:
                ln['used'] = True
                write_line(ws, r, c, ln, None, changes, new=True)
                added.append(dict(row=r, component=ln['detail'], section=ln['cat']))
                rows, sections, lop = buyer_rows(ws)
        unplaced = [l for l in unplaced if not l['used']]
    # LOP
    rows, sections, lop = buyer_rows(ws)
    set_lop(ws, c, lop or h, blk, changes)
    return dict(log=log, changes=changes, added=added,
                unplaced=[f"{l['cat']} / {l['detail']}" for l in unplaced])


def set_lop(ws, c, start, blk, changes):
    """Labor = CM; Overhead/Profit/Transportation = SUM(material:labor) * rate (standard layout)."""
    from openpyxl.utils import get_column_letter as L
    rate_col, uc = c['unit'] - 1, c['unit']
    rows = {}
    for rr in range(start, ws.max_row + 1):
        a = _s(ws.cell(rr, 1).value).lower()
        if a in ('total material cost', 'labor', 'overhead', 'profit', 'transportation') and a not in rows:
            rows[a] = rr
    if 'labor' in rows and blk['cm'] is not None:
        _set(ws, rows['labor'], uc, blk['cm'], changes, ws.title)
    mat, lab = rows.get('total material cost'), rows.get('labor')
    for a, key in (('overhead', 'oh'), ('profit', 'profit'), ('transportation', 'finance')):
        if a in rows and key in blk['rates']:
            r = rows[a]
            _set(ws, r, rate_col, blk['rates'][key], changes, ws.title)
            rc = ws.cell(r, rate_col)
            rc.number_format = '0%'
            al = copy(rc.alignment)
            al.horizontal = 'right'
            rc.alignment = al
            if mat and lab:
                _set(ws, r, uc, f'=SUM({L(uc)}{mat}:{L(uc)}{lab})*{L(rate_col)}{r}', changes, ws.title)


def write_line(ws, r, c, ln, buyer_price, changes, new=False):
    t = ws.title
    if not ln['yy']:
        _set(ws, r, c['net'], 0, changes, t)
        _set(ws, r, c['loss'], 0, changes, t)
        _set(ws, r, c['override'], None, changes, t)
        _set(ws, r, c['freight_ovr'], None, changes, t)
        return
    _set(ws, r, c['net'], round(ln['yy'], 4), changes, t)
    _set(ws, r, c['loss'], round(ln['loss'] - 1, 4), changes, t)
    if new:
        _set(ws, r, c['price'], ln['price'], changes, t)
        _set(ws, r, c['override'], None, changes, t)
    else:
        _set(ws, r, c['override'], None if abs((buyer_price or 0) - ln['price']) < 1e-9 else ln['price'], changes, t)
    _set(ws, r, c['freight_ovr'], ln['cif'] if ln['cif'] else None, changes, t)


def add_row(ws, ln, c, sections):
    """Insert a new component row at the end of the matching section."""
    want = ln['key']
    target = None
    for i, (sr, name) in enumerate(sections):
        if SECTION_KEY.get(name) == want or (want == 'fabric' and name == 'body'):
            target = i
            if SECTION_KEY.get(name) == want:
                break
    if target is None:
        return None
    sr, name = sections[target]
    nxt = sections[target + 1][0] if target + 1 < len(sections) else None
    last = sr
    end = (nxt or ws.max_row)
    for r in range(sr + 1, end):
        if _s(ws.cell(r, 1).value) or _s(ws.cell(r, c['material']).value):
            last = r
    at = last + 1
    insert_row(ws, at, style_from=last, extend_sum_end=last)
    for k in ('component', 'material', 'supplier', 'uom', 'comments', 'placement', 'width'):
        if c.get(k):
            ws.cell(at, c[k]).value = None
    ws.cell(at, 1).value = ln['detail'] or ln['item']
    ws.cell(at, c['material']).value = ln['detail']
    if c.get('uom'):
        ws.cell(at, c['uom']).value = ws.cell(last, c['uom']).value
    if c.get('freight'):
        ws.cell(at, c['freight']).value = 0
    # color columns (between supplier and cuttable width) left blank
    return at


# ---------------- FOB check (same math as HOKA template) ----------------
def calc_fob(ws):
    h, c = buyer_columns(ws)
    rows, sections, lop = buyer_rows(ws)
    mat = 0.0
    for r, comp, sec in rows:
        g = lambda k: _num(ws.cell(r, c[k]).value) if c.get(k) else None
        net, loss = g('net') or 0, g('loss') or 0
        price = g('override') if g('override') is not None else (g('price') or 0)
        fr = g('freight_ovr') if g('freight_ovr') is not None else (g('freight') or 0)
        mat += round(net * (1 + loss) * (price + fr), 4)
    labor = oh = pf = tr = 0
    rate_col = c['unit'] - 1
    for rr in range(lop or h, ws.max_row + 1):
        a = _s(ws.cell(rr, 1).value).lower()
        if a == 'labor':
            labor = _num(ws.cell(rr, c['unit']).value) or 0
        elif a == 'overhead':
            oh = _num(ws.cell(rr, rate_col).value) or 0
        elif a == 'profit':
            pf = _num(ws.cell(rr, rate_col).value) or 0
        elif a == 'transportation':
            tr = _num(ws.cell(rr, rate_col).value) or 0
    base = mat + labor
    return round(base * (1 + oh + pf + tr), 6), round(mat, 4)


FOB_FMT = '"$"#,##0.00'


def fob_cells(ws):
    """Value cells next to every 'FOB' label (top header + bottom), plus B8 ('FOB:')."""
    out = []
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.strip().upper() in ('FOB', 'FOB:'):
                v = ws.cell(c.row, c.column + 1)
                if v.value not in (None, ''):
                    out.append(v)
    return out


def format_fob(ws):
    for c in fob_cells(ws):
        c.number_format = FOB_FMT

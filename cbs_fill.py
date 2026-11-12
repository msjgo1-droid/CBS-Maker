"""Fill a HOKA buyer CBS main sheet (template / Solid / Heather) from an internal CBS block."""
import re
from openpyxl.utils import get_column_letter
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
    for k, v in (('pack', 'pack'), ('trim', 'trim'), ('embell', 'emb'), ('thread', 'thread'), ('stitch', 'thread'),
                 ('label', 'label')):
        if k in c:
            return v
    return 'fabric'


# ---------------- internal side ----------------
SPLIT = re.compile(r'(\d*\.\d+|\d+)(?:\s*\+\s*(\d*\.\d+|\d+))+')


def split_parts(*texts):
    """'0.029+0.027' / '=1.2+0.995' -> [0.029, 0.027]; first text that holds a split wins."""
    for t in texts:
        m = SPLIT.search(_s(t))
        if m:
            nums = [float(x) for x in re.findall(r'\d*\.\d+|\d+', m.group(0))]
            if len(nums) >= 2 and sum(nums) > 0:
                return nums
    return None


def read_internal_block(ws, r0, r1, wsf=None):
    """ws loaded with data_only=True; wsf = same sheet with formulas (to read '=a+b' YY)."""
    hdr = find_label(ws, r0, min(r0 + 5, r1), {'S/#'}).row
    cols = map_columns(ws, hdr)
    net = find_label(ws, hdr + 1, r1, {'NET'})
    others = find_label(ws, hdr + 1, r1, {'OTHERS'})
    stop = min(x.row for x in (net, others) if x)
    from cbs_format import profit_labels as _pl
    _pf = [c.row for c in _pl(ws, hdr + 1, r1).values() if c]
    if _pf:
        stop = min(stop, min(_pf))                    # never read Profit/OH/Finance or 'Others' rows as items
    lines, cat, cm, after_cm, cm_row = [], 'fabric', None, False, None
    for r in range(hdr + 1, stop):
        d = _s(ws.cell(r, cols['i_item']).value)
        e = _s(ws.cell(r, cols['i_detail']).value)
        if _norm(d) == 'cm':
            cm = _num(ws.cell(r, cols['b_ttl']).value)
            cm_row = r
            after_cm = True
            continue
        if d and after_cm:
            cat = d
        price = _num(ws.cell(r, cols['b_price']).value)
        lump = False
        if price is None and (e or d):
            ttl = _num(ws.cell(r, cols['b_ttl']).value)
            if ttl and not _s(e).upper().startswith('TTL'):  # e.g. 'Sublimation print' with only TTL $3.00 typed
                price, lump = ttl, True
        if price is None or not (e or d):
            continue
        nxt = r + 1
        below = [] if _num(ws.cell(nxt, cols['b_price']).value) is not None or _s(ws.cell(nxt, cols['i_item']).value) \
            else [ws.cell(nxt, k).value for k in range(cols['i_item'], cols['i_remark'] + 1)]
        parts = split_parts(wsf.cell(r, cols['i_yy']).value if wsf is not None and 'i_yy' in cols else None,
                            e, ws.cell(r, cols['i_remark']).value, *below) if not after_cm else None
        lines.append(dict(row=r, cat=cat if after_cm else 'fabric', key=cat_key(cat) if after_cm else 'fabric',
                          parts=parts,
                          item=d, detail=e, remark=_s(ws.cell(r, cols['i_remark']).value),
                          price=price,
                          iyy=_num(ws.cell(r, cols['i_yy']).value) if 'i_yy' in cols else None,
                          yy=1 if lump else (_num(ws.cell(r, cols['b_yy']).value) or 0),
                          loss=1 if lump else (_num(ws.cell(r, cols['b_loss']).value) or 1),
                          cif=None if lump else (_num(ws.cell(r, cols['b_cif']).value) if 'b_cif' in cols else None),
                          lump=lump,
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
    # how the internal rounds each amount (ROUND / ROUNDUP / ROUNDDOWN to cents) -> buyer uses the same
    rounding = {}
    if wsf is not None:
        for key, c in profit_labels(ws, hdr + 1, r1).items():
            fml = _s(wsf.cell(c.row, cols['b_ttl']).value).upper()
            m = re.search(r'\b(ROUNDUP|ROUNDDOWN|ROUND)\s*\(', fml)
            if m:
                rounding[key] = m.group(1)
    return dict(lines=lines, cm=cm, rates=rates, cols=cols, cm_row=cm_row, rounding=rounding)


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
    return {t for t in re.split(r'[\s/,()#]+', _s(text).upper()) if len(t) >= 5 and re.search(r'\d', t)
            and not t.startswith('$')}


def _line_codes(l):
    return codes(l['detail']) | codes(l['remark'])


def code_overlap(a, b):
    """'GK37794FW' ~ 'GOLDTEX-GK37794FW', 'PK-42' ~ 'PK-42/03' style matches."""
    for x in a:
        for y in b:
            if x == y or (len(min(x, y, key=len)) >= 6 and (x in y or y in x)):
                return True
    return False


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
    for ln in free:                                    # 3a) internal item name fully inside component name
        if not _same_sec(ln, sec):
            continue
        iw = {w for w in _words(ln['detail']) if len(w) >= 5}
        if iw and len(iw & cw) / len(iw) >= 0.5:
            return ln, 'fuzzy (item)'
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


def repair_sum_columns(ws):
    """Some PLM templates (wide ones, Unit Price in AA) come with sums pointing at the wrong column, e.g.
    Body =SUM(A12:A14), Sub-total =SUM(A57:A60), FOB =SUM(A61:A62) -> FOB shows 0. Re-point every SUM range
    in the Unit Price column at that column. Returns the cells fixed."""
    h, c = buyer_columns(ws)
    uc = c.get('unit')
    if not h or not uc:
        return []
    L = get_column_letter(uc)
    fixed = []
    pat = re.compile(r'(?<![A-Z$])(\$?)([A-Z]{1,3})(\$?)(\d+):(\$?)([A-Z]{1,3})(\$?)(\d+)')
    for r in range(h + 1, ws.max_row + 1):
        v = ws.cell(r, uc).value
        if not (isinstance(v, str) and v.startswith('=') and 'SUM' in v.upper()):
            continue
        def rep(m):
            if m.group(2) == m.group(6) and m.group(2) != L:
                return f'{m.group(1)}{L}{m.group(3)}{m.group(4)}:{m.group(5)}{L}{m.group(7)}{m.group(8)}'
            return m.group(0)
        nv = pat.sub(rep, v)
        if nv != v:
            ws.cell(r, uc).value = nv
            fixed.append(f'{L}{r}: {v} -> {nv}')
    return fixed


def fill_sheet(ws, blk, variant=None, manual=None, add_missing=True, prev_splits=None, include_minor=False):
    """variant: 'solid'/'heather'/None. manual: {buyer_row: internal_line_index or -1 (skip)}.
    Returns dict(log, changes, unplaced, added)."""
    manual = manual or {}
    repaired = repair_sum_columns(ws)
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
    # pass 1: material-code matches across the whole sheet first (so a name hit on an earlier row can't steal them)
    pre = {}
    for r, comp, sec in rows:
        if inactive_row(ws, r, c, variant) or (other and f'({other})' in comp.lower()) or r in manual:
            continue
        mc = codes(material_text(ws, r, c))
        ln = next((l for l in lines if not l['used'] and mc and code_overlap(_line_codes(l), mc)), None)
        if ln:
            ln['used'] = True
            pre[r] = ln
    for r, comp, sec in rows:
        why = inactive_row(ws, r, c, variant)
        if (other and f'({other})' in comp.lower()) or why:
            _set(ws, r, c['net'], 0, changes, ws.title)
            _set(ws, r, c['loss'], 0, changes, ws.title)
            _set(ws, r, c['override'], None, changes, ws.title)
            log.append(dict(row=r, component=comp, match='—', how=f'zero ({why or other})'))
            continue
        if manual.get(r) == -1:
            log.append(dict(row=r, component=comp, match='—', how='skip (manual)'))
            continue
        bp = _num(ws.cell(r, c['price']).value) or 0
        bl = _num(ws.cell(r, c['loss']).value)
        man = lines[manual[r]] if manual.get(r) is not None and manual[r] >= 0 else None
        mc = codes(material_text(ws, r, c))
        twin = None
        if r in pre:
            ln, how = pre[r], 'code'
        else:
            twin = next((l for l in lines if l['used'] and mc and code_overlap(_line_codes(l), mc)), None) \
                if man is None else None
            ln, how = (None, None) if twin else match_line(comp, material_text(ws, r, c), bp, bl, sec, lines, man)
        if not ln:
            if twin and mc:
                # same fabric already put on another row (e.g. Shell 1 / Contrast 1): usage stays on the first row
                _set(ws, r, c['net'], 0, changes, ws.title)
                _set(ws, r, c['loss'], 0, changes, ws.title)
                log.append(dict(row=r, component=comp, match=f"{twin['item'] or twin['cat']} / {twin['detail']}",
                                how='shared (0)', line=lines.index(twin)))
                twin.setdefault('rows', []).append(r)
                continue
            blank = not comp and not (_num(ws.cell(r, c['net']).value) or 0)
            if not blank:                                  # internal is the truth: no internal item -> no cost
                _set(ws, r, c['net'], 0, changes, ws.title)
                _set(ws, r, c['loss'], 0, changes, ws.title)
                _set(ws, r, c['override'], None, changes, ws.title)
            log.append(dict(row=r, component=comp, match='', how='empty row' if blank else 'NO MATCH'))
            continue
        ln['used'] = True
        ln['rows'] = [r]
        write_line(ws, r, c, ln, bp, changes)
        log.append(dict(row=r, component=comp, match=f"{ln['item'] or ln['cat']} / {ln['detail']}", how=how,
                        line=lines.index(ln)))
    # one internal fabric line on several buyer rows (Shell 1 / Contrast 1 ...): split the usage
    for ln in lines:
        rows_ = ln.get('rows') or []
        if len(rows_) < 2 or not ln['yy']:
            continue
        ratios, src = None, ''
        if ln.get('parts') and len(ln['parts']) == len(rows_):
            ratios, src = ln['parts'], 'internal ' + '+'.join(f'{x:g}' for x in ln['parts'])
        elif prev_splits:
            key = tuple(sorted(codes(material_text(ws, rows_[0], c))))
            pv = prev_splits.get(key)
            if pv and len(pv) == len(rows_):
                ratios, src = pv, 'previous CBS ' + '+'.join(f'{x:g}' for x in pv)
        if not ratios:
            continue                                      # never split before -> all on the first row
        tot = sum(ratios)
        for rr, part in zip(rows_, ratios):
            write_line(ws, rr, c, dict(ln, yy=ln['yy'] * part / tot), _num(ws.cell(rr, c['price']).value) or 0, changes)
        for l in log:
            if l['row'] in rows_:
                l['how'] = 'split'
                l['split'] = src
    # missing rows
    added = []
    excluded = [l for l in lines if not l['used'] and l['yy'] and
                any(x in (l['detail'] + l['item']).lower() for x in EXCLUDE_ADD)]
    unplaced = [l for l in lines if not l['used'] and l['yy'] and (include_minor or l not in excluded)]
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
    return dict(log=log, changes=changes, added=added, repaired=repaired, excluded=[] if include_minor else excluded,
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
                fx = f'SUM({L(uc)}{mat}:{L(uc)}{lab})*{L(rate_col)}{r}'
                fn = (blk.get('rounding') or {}).get(key)
                _set(ws, r, uc, f'={fn}({fx},2)' if fn else f'={fx}', changes, ws.title)


def colour_cells(ws, r, c):
    lo = (c.get('supplier') or c['material']) + 1
    hi = (c.get('width') or c['uom'] or c['net'])
    return [_s(ws.cell(r, k).value) for k in range(lo, hi)]


def inactive_row(ws, r, c, variant):
    """Row not used on this sheet: colourways all 'N/A', or heather fabric on the Solid sheet (and vice versa)."""
    cols_ = colour_cells(ws, r, c)
    if cols_ and all('N/A' in x.upper() for x in cols_):
        return 'N/A'
    mat = material_text(ws, r, c).upper()
    if variant == 'solid' and 'HEATHER' in mat and not any(x and 'N/A' not in x.upper() for x in cols_):
        return 'heather row'
    if variant == 'heather' and 'HEATHER' not in mat and re.search(r'\bSOLID\b', mat) \
            and not any(x and 'N/A' not in x.upper() for x in cols_):
        return 'solid row'
    return None


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
    if nxt is None:                                   # last section ends at the LOP block
        nxt = next((r for r in range(sr + 1, ws.max_row + 1) if _s(ws.cell(r, 1).value).lower() == 'lop'), None)
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
    lrow = {}
    for rr in range(lop or h, ws.max_row + 1):
        a = _s(ws.cell(rr, 1).value).lower()
        lrow.setdefault(a, rr)
        if a == 'labor':
            labor = _num(ws.cell(rr, c['unit']).value) or 0
        elif a == 'overhead':
            oh = _num(ws.cell(rr, rate_col).value) or 0
        elif a == 'profit':
            pf = _num(ws.cell(rr, rate_col).value) or 0
        elif a == 'transportation':
            tr = _num(ws.cell(rr, rate_col).value) or 0
    base = mat + labor
    import math
    def amt(rate, key):
        f = _s(ws.cell(lrow.get(key, 0), c['unit']).value).upper() if lrow.get(key) else ''
        v = base * rate
        if f.startswith('=ROUNDUP('):
            return math.ceil(round(v * 100, 6)) / 100
        if f.startswith('=ROUNDDOWN('):
            return math.floor(round(v * 100, 6)) / 100
        if f.startswith('=ROUND('):
            return round(v + 1e-12, 2)
        return v
    return round(base + amt(oh, 'overhead') + amt(pf, 'profit') + amt(tr, 'transportation'), 6), round(mat, 4)


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


def prev_split_map(wb):
    """{material codes: [net per row]} for fabrics the previous CBS spread over 2+ rows."""
    out = {}
    for ws in wb.worksheets:
        if not is_main_sheet(ws):
            continue
        h, c = buyer_columns(ws)
        groups = {}
        for r, comp, sec in buyer_rows(ws)[0]:
            if sec not in FABRIC_SECTIONS:
                continue
            k = tuple(sorted(codes(material_text(ws, r, c))))
            if k:
                groups.setdefault(k, []).append(_num(ws.cell(r, c['net']).value) or 0)
        for k, v in groups.items():
            if len(v) >= 2 and sum(1 for x in v if x) >= 2 and k not in out:
                out[k] = v
    return out


# ---------------- open price drop vs previous stage ----------------
def _line_key(l):
    return codes(l['detail']) | codes(l['remark'])


def match_prev_lines(cur, prev):
    """{cur_index: prev_index} by material code, then item/detail name."""
    out, used = {}, set()
    for i, l in enumerate(cur):
        k = _line_key(l)
        j = next((j for j, p in enumerate(prev) if j not in used and k and k & _line_key(p)), None)
        if j is None:
            n = _norm(l['detail']) or _norm(l['item'])
            j = next((j for j, p in enumerate(prev) if j not in used and n and n in (_norm(p['detail']), _norm(p['item']))), None)
        if j is not None:
            out[i] = j
            used.add(j)
    return out


def open_cost(l):
    return (l['yy'] or 0) * (l['loss'] or 1) * ((l['price'] or 0) + (l['cif'] or 0))


def price_drops(cur_blk, prev_blk):
    """Lines whose open (buyer-side) cost went down vs the previous stage, with the reason."""
    m = match_prev_lines(cur_blk['lines'], prev_blk['lines'])
    rows = []
    for i, j in m.items():
        a, b = cur_blk['lines'][i], prev_blk['lines'][j]
        d = open_cost(a) - open_cost(b)
        if d < -0.0005:
            why = [k for k, lo in (('yy', a['yy'] < b['yy'] - 1e-9), ('price', a['price'] < b['price'] - 1e-9),
                                   ('loss', a['loss'] < b['loss'] - 1e-9), ('cif', (a['cif'] or 0) < (b['cif'] or 0) - 1e-9)) if lo]
            rows.append(dict(line=i, item=a['item'] or a['cat'], detail=a['detail'], why=why,
                             cur=dict(yy=a['yy'], price=a['price'], loss=a['loss'], cif=a['cif']),
                             prev=dict(yy=b['yy'], price=b['price'], loss=b['loss'], cif=b['cif']),
                             delta=round(d, 4)))
    if (cur_blk['cm'] or 0) < (prev_blk['cm'] or 0) - 1e-9:
        rows.append(dict(line='CM', item='CM', detail='', why=['cm'], cur=dict(cm=cur_blk['cm']),
                         prev=dict(cm=prev_blk['cm']), delta=round(cur_blk['cm'] - prev_blk['cm'], 4)))
    return rows


def keep_previous(cur_blk, prev_blk, keep):
    """keep: {line_index or 'CM': ['yy','price','loss','cif'|'cm']} -> modifies cur_blk; returns internal cell edits."""
    m = match_prev_lines(cur_blk['lines'], prev_blk['lines'])
    from openpyxl.utils import get_column_letter as L
    cols, edits = cur_blk['cols'], {}
    for i, fields in keep.items():
        if i == 'CM':
            if 'cm' in fields and prev_blk['cm'] is not None:
                cur_blk['cm'] = prev_blk['cm']
                if cur_blk.get('cm_row'):
                    edits[f"{L(cols['b_ttl'])}{cur_blk['cm_row']}"] = prev_blk['cm']
            continue
        if i not in m:
            continue
        a, b = cur_blk['lines'][i], prev_blk['lines'][m[i]]
        for f in fields:
            col = {'yy': 'b_yy', 'price': 'b_price', 'loss': 'b_loss', 'cif': 'b_cif'}[f]
            if b.get(f) is None or col not in cols:
                continue
            a[f] = b[f]
            edits[f"{L(cols[col])}{a['row']}"] = b[f]
    return edits


def yy_drops(cur_blk, prev_blk):
    """Lines whose INTERNAL yy went down vs the previous-date tab."""
    m = match_prev_lines(cur_blk['lines'], prev_blk['lines'])
    out = []
    for i, j in m.items():
        a, b = cur_blk['lines'][i], prev_blk['lines'][j]
        if a.get('iyy') is None or b.get('iyy') is None:
            continue
        if a['iyy'] < b['iyy'] - 1e-6:
            out.append(dict(line=i, item=a['item'] or a['cat'], detail=a['detail'],
                            prev_yy=round(b['iyy'], 4), cur_yy=round(a['iyy'], 4),
                            diff=round(a['iyy'] - b['iyy'], 4),
                            pct=round((a['iyy'] - b['iyy']) / b['iyy'] * 100, 1) if b['iyy'] else None,
                            prev_open=round(b['yy'], 4), cur_open=round(a['yy'], 4)))
    return out

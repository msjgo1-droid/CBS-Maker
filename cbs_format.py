"""HOKA internal CBS -> standard number-format normalizer.

Rules:
  - final prices only (internal total, open/buyer FOB, header FOB): 2 decimals
  - every other number: up to 4 decimals ($0.20, $0.316, $0.5231, 1.03, 0.9852)
  - percentages: margin / Others rate 0.00%, buyer LOSS & Profit/OH/Finance rate 0%
Only number formats are changed; values, formulas, colors, layout stay untouched.
"""
import re
import openpyxl

PRICE4 = '"$"#,##0.0000'          # always 4 decimals (everything except the final prices)
PRICE2 = '"$"#,##0.00'
YY4 = '0.0000'
LOSS_INT = '0.0000'
PCT0 = '0%'
PCT2 = '0.00%'
DATE = 'd-mmm'


def _s(v):
    return str(v).strip().upper() if v is not None else ''


def _is_num_or_formula(c):
    return isinstance(c.value, (int, float)) or (isinstance(c.value, str) and c.value.startswith('='))


def find_blocks(ws):
    starts = [c.row for c in ws['A'] if _s(c.value) == 'COST DATE']
    ends = starts[1:] + [ws.max_row + 1]
    return list(zip(starts, [e - 1 for e in ends]))


def map_columns(ws, hdr_row):
    """Return dict of column indexes for internal (i_*) and buyer (b_*) sections."""
    cols, seen_remark = {}, False
    for c in ws[hdr_row]:
        h = _s(c.value)
        if not h:
            continue
        side = 'b_' if seen_remark else 'i_'
        if h == 'REMARK':
            if seen_remark:
                cols['b_remark'] = c.column
            else:
                cols['i_remark'] = c.column
                seen_remark = True
            continue
        key = {'PRICE': 'price', 'YY': 'yy', 'LOSS': 'loss', 'TTL': 'ttl', 'CIF': 'cif',
               'ITEM': 'item', 'DETAIL': 'detail'}.get(h)
        if 'LOSS' in h and h != 'LOSS':
            key, side = 'actual_loss', 'b_'
        if key and side + key not in cols:
            cols[side + key] = c.column
    return cols


def find_label(ws, r0, r1, labels, max_col=25):
    for row in ws.iter_rows(min_row=r0, max_row=r1, max_col=max_col):
        for c in row:
            if _s(c.value) in labels:
                return c
    return None


def find_contains(ws, r0, r1, word, max_col=25):
    """First cell whose text contains `word` (e.g. 'Transportation +Finance' for FINANCE)."""
    for row in ws.iter_rows(min_row=r0, max_row=r1, max_col=max_col):
        for c in row:
            if isinstance(c.value, str) and word in c.value.strip().upper() and len(c.value) < 40:
                return c
    return None


def profit_labels(ws, r0, r1):
    """{'profit': cell, 'oh': cell, 'finance': cell} — tolerant to label variants."""
    out = {}
    p = find_label(ws, r0, r1, {'PROFITS', 'PROFIT'})
    if p:
        out['profit'] = p
    o = find_label(ws, r0, r1, {'OH', 'O/H', 'OVERHEAD'})
    if o:
        out['oh'] = o
    f = find_contains(ws, r0, r1, 'FINANCE') or find_contains(ws, r0, r1, 'TRANSPORT')
    if f:
        out['finance'] = f
    return out


def normalize_block(ws, r0, r1, log):
    hdr = find_label(ws, r0, min(r0 + 5, r1), {'S/#'})
    if not hdr:
        log.append(f'  ! block r{r0}: header row (S/#) not found, skipped')
        return
    h = hdr.row
    cols = map_columns(ws, h)
    need = ['i_price', 'i_ttl', 'b_price', 'b_ttl']
    if any(k not in cols for k in need):
        log.append(f'  ! block r{r0}: missing columns {[k for k in need if k not in cols]}, skipped')
        return
    net = find_label(ws, h + 1, r1, {'NET'})
    others = find_label(ws, h + 1, r1, {'OTHERS'})
    pf = profit_labels(ws, h + 1, r1)
    pf_rows = {c.row for c in pf.values() if c}
    last_item = (others.row - 1) if others else (net.row - 1 if net else r1)
    item_col = cols.get('i_item', 4)

    def fmt(r, col_key, f):
        if col_key in cols:
            c = ws.cell(r, cols[col_key])
            if _is_num_or_formula(c):
                c.number_format = f

    # header area: date + FOB in row h-1
    ws.cell(h - 1, 1).number_format = DATE
    fob_hdr = ws.cell(h - 1, 2)
    if _is_num_or_formula(fob_hdr):
        fob_hdr.number_format = PRICE2

    # item lines
    for r in range(h + 1, last_item + 1):
        if r in pf_rows:
            continue
        for k in ('i_price', 'i_ttl', 'b_cif', 'b_price', 'b_ttl'):
            fmt(r, k, PRICE4)
        fmt(r, 'i_yy', YY4)
        fmt(r, 'b_yy', YY4)
        fmt(r, 'i_loss', LOSS_INT)
        fmt(r, 'b_loss', PCT0)
        fmt(r, 'actual_loss', PCT0)
        # category subtotals (=SUM in buyer remark column)
        if 'b_remark' in cols:
            c = ws.cell(r, cols['b_remark'])
            if isinstance(c.value, str) and c.value.upper().startswith('=SUM'):
                c.number_format = PRICE4

    # profit / OH / finance: rate % , amount 2 decimals
    for lbl, c in pf.items():
        if not c:
            continue
        r = c.row
        for col in range(c.column + 1, cols['b_ttl']):
            cc = ws.cell(r, col)
            if isinstance(cc.value, (int, float)) and cc.value < 1:
                cc.number_format = PCT0
        ws.cell(r, cols['b_ttl']).number_format = PRICE4

    # others: rate 0.00%, amount 2 decimals
    if others:
        r = others.row
        for col in range(others.column + 1, cols['i_ttl']):
            cc = ws.cell(r, col)
            if isinstance(cc.value, (int, float)) and cc.value < 1:
                cc.number_format = PCT2
        ws.cell(r, cols['i_ttl']).number_format = PRICE4

    # net + final row (margin, internal total, FOB)
    if net:
        ws.cell(net.row, cols['i_ttl']).number_format = PRICE4
        for r in range(net.row, min(net.row + 3, r1) + 1):
            fob = ws.cell(r, cols['b_ttl'])
            if isinstance(fob.value, str) and fob.value.upper().startswith('=SUM'):
                fob.number_format = PRICE2
                ws.cell(r, cols['i_ttl']).number_format = PRICE2
                ws.cell(r, cols['i_loss']).number_format = PCT2
                if 'b_remark' in cols:
                    q = ws.cell(r - 1, cols['b_remark'])
                    if isinstance(q.value, str) and q.value.upper().startswith('=SUM'):
                        q.number_format = PRICE4   # profit+OH+finance sum
                break
    log.append(f'  ok block r{r0}-{r1}: cols={ {k: openpyxl.utils.get_column_letter(v) for k, v in cols.items()} }')


def normalize_workbook(src, dst, all_sheets=False):
    wb = openpyxl.load_workbook(src)
    log = []
    sheets = wb.worksheets if all_sheets else wb.worksheets[:1]
    for ws in sheets:
        log.append(f'[{ws.title}]')
        for r0, r1 in find_blocks(ws):
            normalize_block(ws, r0, r1, log)
    wb.save(dst)
    return log


if __name__ == '__main__':
    import sys
    for line in normalize_workbook(sys.argv[1], sys.argv[2]):
        print(line)


# ---------------------------------------------------------------------------
# Lossless writer: only number formats are patched directly in the xlsx XML,
# so images (incl. WMF/EMF), shapes, comments etc. are kept byte-for-byte.
# ---------------------------------------------------------------------------
import io as _io
import re as _re
import zipfile as _zip
from xml.sax.saxutils import escape as _esc

_BUILTIN = {'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10}


def _sheet_paths(z):
    wb = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid2t = {m.group(1): m.group(2) for m in _re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    rid2t.update({m.group(2): m.group(1) for m in _re.finditer(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = []
    for m in _re.finditer(r'<sheet\b[^>]*?r:id="([^"]+)"', wb):
        tgt = rid2t[m.group(1)].lstrip('/')
        out.append(tgt if tgt.startswith('xl/') else 'xl/' + tgt)
    return out


def normalize_bytes(data, all_sheets=False):
    """Return (new_xlsx_bytes, log). Values/images/everything else untouched."""
    wb = openpyxl.load_workbook(_io.BytesIO(data))
    before = {}
    sheets = wb.worksheets if all_sheets else wb.worksheets[:1]
    for si, ws in enumerate(sheets):
        for row in ws.iter_rows():
            for c in row:
                before[(si, c.coordinate)] = c.number_format
    log = []
    for ws in sheets:
        log.append(f'[{ws.title}]')
        for r0, r1 in find_blocks(ws):
            normalize_block(ws, r0, r1, log)
    want = {}
    for si, ws in enumerate(sheets):
        for row in ws.iter_rows():
            for c in row:
                if c.number_format != before.get((si, c.coordinate)):
                    want.setdefault(si, {})[c.coordinate] = c.number_format
    zin = _zip.ZipFile(_io.BytesIO(data))
    paths = _sheet_paths(zin)
    styles = zin.read('xl/styles.xml').decode('utf8')
    # --- numFmts
    fmts = {m.group(2).replace('&quot;', '"').replace('&amp;', '&'): int(m.group(1))
            for m in _re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\s*/>', styles)}
    next_id = max([163] + list(fmts.values())) + 1
    new_fmts = []

    def fmt_id(code):
        nonlocal next_id
        if code in _BUILTIN:
            return _BUILTIN[code]
        if code not in fmts:
            fmts[code] = next_id
            new_fmts.append((next_id, code))
            next_id += 1
        return fmts[code]
    # --- cellXfs
    m = _re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', styles, _re.S)
    xfs = _re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', m.group(1), _re.S)
    added_xf, cache = [], {}

    def new_style(s, code):
        key = (s, code)
        if key in cache:
            return cache[key]
        x = xfs[s]
        fid = fmt_id(code)
        x2 = _re.sub(r'numFmtId="\d+"', f'numFmtId="{fid}"', x, count=1) if 'numFmtId=' in x \
            else x.replace('<xf ', f'<xf numFmtId="{fid}" ', 1)
        if 'applyNumberFormat=' in x2:
            x2 = _re.sub(r'applyNumberFormat="\w+"', 'applyNumberFormat="1"', x2)
        else:
            x2 = x2.replace('<xf ', '<xf applyNumberFormat="1" ', 1)
        added_xf.append(x2)
        cache[key] = len(xfs) + len(added_xf) - 1
        return cache[key]

    patched = {}
    for si, cells in want.items():
        path = paths[si]
        xml = zin.read(path).decode('utf8')

        def rep(mm):
            tag = mm.group(0)
            ref = mm.group(1)
            if ref not in cells:
                return tag
            sm = _re.search(r'\ss="(\d+)"', tag)
            s = int(sm.group(1)) if sm else 0
            ns = new_style(s, cells[ref])
            return _re.sub(r'\ss="\d+"', f' s="{ns}"', tag) if sm else tag.replace(f'r="{ref}"', f'r="{ref}" s="{ns}"')
        xml = _re.sub(r'<c r="([A-Z]+\d+)"[^>]*?>', rep, xml)
        patched[path] = xml
    # write styles
    if added_xf:
        body = m.group(1) + ''.join(added_xf)
        total = len(xfs) + len(added_xf)
        styles = styles[:m.start()] + _re.sub(r'count="\d+"', f'count="{total}"',
                                               styles[m.start():m.start() + 40], count=1).split('>')[0] + '>' + \
            body + '</cellXfs>' + styles[m.end():]
    if new_fmts:
        add = ''.join(f'<numFmt numFmtId="{i}" formatCode="{_esc(c, {chr(34): "&quot;"})}"/>' for i, c in new_fmts)
        if '<numFmts' in styles:
            styles = _re.sub(r'(<numFmts[^>]*>)(.*?)(</numFmts>)',
                             lambda mm: mm.group(1) + mm.group(2) + add + mm.group(3), styles, count=1, flags=_re.S)
            n = len(_re.findall(r'<numFmt ', styles))
            styles = _re.sub(r'<numFmts count="\d+"', f'<numFmts count="{n}"', styles, count=1)
        else:
            styles = styles.replace('<fonts', f'<numFmts count="{len(new_fmts)}">{add}</numFmts><fonts', 1)
    out = _io.BytesIO()
    with _zip.ZipFile(out, 'w', _zip.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            if it.filename == 'xl/styles.xml':
                zo.writestr(it, styles)
            elif it.filename in patched:
                zo.writestr(it, patched[it.filename])
            else:
                zo.writestr(it, zin.read(it.filename))
    return out.getvalue(), log


def set_text_cells(data, values, sheet_index=0):
    return set_cells(data, {k: ('t', v, None) for k, v in values.items()}, sheet_index)


def set_cells(data, values, sheet_index=0):
    """Write cells by patching sheet XML (keeps everything else byte-for-byte).
    values: {'Q84': ('t', 'CBS 11/9 $23.76', None), 'Q8': ('f', 'SUM(P4:P8)', 'Q15')}
    third item = cell whose style to copy (None keeps the target's own style)."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    zin = _zip.ZipFile(_io.BytesIO(data))
    path = _sheet_paths(zin)[sheet_index]
    xml = zin.read(path).decode('utf8')
    def style_of(ref):
        m = _re.search(rf'<c r="{ref}"(\s[^>]*?)?/?>', xml)
        sm = _re.search(r'\ss="(\d+)"', m.group(0)) if m else None
        return sm.group(1) if sm else None

    for ref, (kind, text, style_ref) in values.items():
        col, row = coordinate_from_string(ref)
        if kind == 'f':
            cell = f'<c r="{ref}"{{s}}><f>{_esc(text)}</f></c>'
        elif kind == 'n':
            cell = f'<c r="{ref}"{{s}}><v>{float(text)!r}</v></c>'
        elif kind == 'r':
            # rich text: keep the existing text as it is, add the new part in red
            old_t, new_t, fname, sz = text
            rpr = '<rPr>' + (f'<sz val="{sz}"/>' if sz else '') + '<color rgb="FFFF0000"/>' + \
                (f'<rFont val="{_esc(fname)}"/>' if fname else '') + '</rPr>'
            rpr0 = '<rPr>' + (f'<sz val="{sz}"/>' if sz else '') + (f'<rFont val="{_esc(fname)}"/>' if fname else '') + '</rPr>'
            runs = (f'<r>{rpr0}<t xml:space="preserve">{_esc(old_t)}</t></r>' if old_t else '') + \
                f'<r>{rpr}<t xml:space="preserve">{_esc(new_t)}</t></r>'
            cell = f'<c r="{ref}"{{s}} t="inlineStr"><is>{runs}</is></c>'
        else:
            cell = f'<c r="{ref}"{{s}} t="inlineStr"><is><t xml:space="preserve">{_esc(text)}</t></is></c>'
        st_ = style_of(style_ref) if style_ref else None
        m = _re.search(rf'<c r="{ref}"(?:\s[^>]*?)?(?:/>|>.*?</c>)', xml, _re.S)
        if m:
            own = _re.search(r'\ss="(\d+)"', m.group(0)[:m.group(0).find('>') + 1])
            st_ = st_ or (own.group(1) if own else None)
            xml = xml[:m.start()] + cell.format(s=f' s="{st_}"' if st_ else '') + xml[m.end():]
            continue
        rm = _re.search(rf'<row r="{row}"[^>]*?(?:/>|>(.*?)</row>)', xml, _re.S)
        if not rm:
            continue                                    # row missing: skip quietly
        new = cell.format(s=f' s="{st_}"' if st_ else '')
        if rm.group(0).endswith('/>'):
            xml = xml[:rm.start()] + rm.group(0)[:-2] + '>' + new + '</row>' + xml[rm.end():]
            continue
        body, ci = rm.group(1), column_index_from_string(col)
        pos = len(body)
        for cm in _re.finditer(r'<c r="([A-Z]+)\d+"', body):
            if column_index_from_string(cm.group(1)) > ci:
                pos = cm.start()
                break
        body = body[:pos] + new + body[pos:]
        start = rm.start(1)
        xml = xml[:start] + body + xml[rm.end(1):]
    if any(k == 'n' for k, _, _ in values.values()):
        # values changed underneath existing formulas -> drop their cached results so every app recalculates
        xml = _re.sub(r'(<f\b[^>]*?(?:/>|>.*?</f>))<v>[^<]*</v>', r'\1', xml, flags=_re.S)
        xml = _re.sub(r'(<c\b[^>]*?)\st="(?:str|e|b)"([^>]*>\s*<f)', r'\1\2', xml)
    wbx = zin.read('xl/workbook.xml').decode('utf8')
    if any(k in ('f', 'n') for k, _, _ in values.values()):          # make Excel recalculate new formulas
        if '<calcPr' in wbx:
            if 'fullCalcOnLoad' not in wbx:
                wbx = wbx.replace('<calcPr', '<calcPr fullCalcOnLoad="1"', 1)
        else:
            wbx = wbx.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
    out = _io.BytesIO()
    with _zip.ZipFile(out, 'w', _zip.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            if it.filename == path:
                zo.writestr(it, xml)
            elif it.filename == 'xl/workbook.xml':
                zo.writestr(it, wbx)
            else:
                zo.writestr(it, zin.read(it.filename))
    return out.getvalue()


# ---------------------------------------------------------------------------
# Section subtotals in the buyer REMARK column (one =SUM per cost section, like 1193712)
# ---------------------------------------------------------------------------
_CATS = ('TRIM', 'EMBELL', 'THREAD', 'STITCH', 'LABEL', 'LOGO', 'PACK', 'FINISH', 'WASH', 'PRINT')


def _sections(ws, hdr, stop, cols):
    """[(name, first_row, end_row, last_value_row)]: fabric block (rows before CM) + each cost category."""
    ic, tc = cols['i_item'], cols['b_ttl']
    has = lambda r: ws.cell(r, tc).value not in (None, '')
    cm = next((r for r in range(hdr + 1, stop) if str(ws.cell(r, ic).value or '').strip().upper() == 'CM'), None)
    secs = []
    if cm and cm > hdr + 1:
        vals = [r for r in range(hdr + 1, cm) if has(r)]
        if vals:
            secs.append(('Fabric', hdr + 1, cm - 1, vals[-1]))
    labels = [r for r in range((cm or hdr) + 1, stop) if str(ws.cell(r, ic).value or '').strip()]
    for i, r in enumerate(labels):
        name = str(ws.cell(r, ic).value).strip()
        end = (labels[i + 1] - 1) if i + 1 < len(labels) else stop - 1
        if not any(k in name.upper() for k in _CATS):
            continue
        vals = [x for x in range(r, end + 1) if has(x)]
        if vals:
            secs.append((name, r, end, vals[-1]))
    return secs


def subtotal_plan(data, sheet_index=0):
    """({ref: ('f', formula, style_ref)}, log) — adds only the section subtotals that are missing."""
    from openpyxl.utils import get_column_letter as L
    wb = openpyxl.load_workbook(_io.BytesIO(data))
    ws = wb.worksheets[sheet_index]
    plan, log = {}, []
    for r0, r1 in find_blocks(ws):
        h = find_label(ws, r0, min(r0 + 5, r1), {'S/#'})
        if not h:
            continue
        cols = map_columns(ws, h.row)
        if 'b_ttl' not in cols or 'i_item' not in cols:
            continue
        rc = cols.get('b_remark') or cols['b_ttl'] + 1
        others = find_label(ws, h.row + 1, r1, {'OTHERS'})
        net = find_label(ws, h.row + 1, r1, {'NET'})
        stop = min(x.row for x in (others, net) if x) if (others or net) else r1
        pf = sorted({c.row for c in profit_labels(ws, h.row + 1, r1).values()})
        if pf:
            stop = min(stop, pf[0])
        tcol = L(cols['b_ttl'])
        existing = {}
        for r in range(h.row + 1, r1 + 1):
            v = str(ws.cell(r, rc).value or '').replace('$', '').upper()
            m = _re.match(r'^=SUM\(([A-Z]+)(\d+)(?::[A-Z]+(\d+))?\)$', v)
            if m:
                existing[r] = (int(m.group(2)), int(m.group(3) or m.group(2)))
        style_ref = next((f'{L(rc)}{r}' for r in existing), None)
        for name, a, b, last in _sections(ws, h.row, stop, cols):
            first = next(x for x in range(a, b + 1) if ws.cell(x, cols['b_ttl']).value not in (None, ''))
            if any(x <= first and y >= last for x, y in existing.values()):
                continue          # already summed (even if one SUM covers several sections) -> never touch
            spot = next((x for x in range(last, b + 1) if ws.cell(x, rc).value in (None, '')), None)
            if spot is None:
                log.append(f'  ! {name}: {L(rc)}{last} not empty, subtotal not added')
                continue
            ref = f'{L(rc)}{spot}'
            plan[ref] = ('f', f'SUM({tcol}{a}:{tcol}{last})', style_ref or f'{tcol}{last}')
            log.append(f'  + subtotal {name} r{a}-{last} -> {ref}')
        if pf and not any(pf[0] <= r <= pf[-1] + 2 for r in existing) \
                and ws.cell(pf[-1], rc).value in (None, ''):
            ref = f'{L(rc)}{pf[-1]}'
            plan[ref] = ('f', f'SUM({tcol}{pf[0]}:{tcol}{pf[-1]})', style_ref or f'{tcol}{pf[-1]}')
            log.append(f'  + subtotal Profit/OH/Finance -> {ref}')
    return plan, log


def highlight_cells(data, refs, rgb='FFFF7C80', sheet_index=0):
    """Give cells a solid fill (value, number format, borders untouched) by patching styles.xml + sheet XML."""
    if not refs:
        return data
    zin = _zip.ZipFile(_io.BytesIO(data))
    path = _sheet_paths(zin)[sheet_index]
    styles = zin.read('xl/styles.xml').decode('utf8')
    fm = _re.search(r'<fills[^>]*>(.*?)</fills>', styles, _re.S)
    nf = len(_re.findall(r'<fill>|<fill\s', fm.group(1)))
    fill = f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill>'
    styles = styles[:fm.start()] + _re.sub(r'count="\d+"', f'count="{nf + 1}"', styles[fm.start():fm.start(1)], count=1) + \
        fm.group(1) + fill + '</fills>' + styles[fm.end():]
    m = _re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', styles, _re.S)
    xfs = _re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', m.group(1), _re.S)
    added, cache = [], {}

    def new_style(s):
        if s in cache:
            return cache[s]
        x = xfs[s]
        x2 = _re.sub(r'fillId="\d+"', f'fillId="{nf}"', x, count=1) if 'fillId=' in x else x.replace('<xf ', f'<xf fillId="{nf}" ', 1)
        x2 = _re.sub(r'applyFill="\w+"', 'applyFill="1"', x2) if 'applyFill=' in x2 else x2.replace('<xf ', '<xf applyFill="1" ', 1)
        added.append(x2)
        cache[s] = len(xfs) + len(added) - 1
        return cache[s]
    xml = zin.read(path).decode('utf8')
    want = set(refs)

    def rep(mm):
        tag, ref = mm.group(0), mm.group(1)
        if ref not in want:
            return tag
        sm = _re.search(r'\ss="(\d+)"', tag)
        ns = new_style(int(sm.group(1)) if sm else 0)
        return _re.sub(r'\ss="\d+"', f' s="{ns}"', tag) if sm else tag.replace(f'r="{ref}"', f'r="{ref}" s="{ns}"')
    xml = _re.sub(r'<c r="([A-Z]+\d+)"[^>]*?>', rep, xml)
    if added:
        m = _re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', styles, _re.S)
        total = len(xfs) + len(added)
        head = _re.sub(r'count="\d+"', f'count="{total}"', styles[m.start():m.start(1)], count=1)
        styles = styles[:m.start()] + head + m.group(1) + ''.join(added) + '</cellXfs>' + styles[m.end():]
    out = _io.BytesIO()
    with _zip.ZipFile(out, 'w', _zip.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            if it.filename == 'xl/styles.xml':
                zo.writestr(it, styles)
            elif it.filename == path:
                zo.writestr(it, xml)
            else:
                zo.writestr(it, zin.read(it.filename))
    return out.getvalue()

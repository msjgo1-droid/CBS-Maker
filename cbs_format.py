"""HOKA internal CBS -> standard number-format normalizer.

Rules (from 1193712 good example):
  - unit prices / line totals / subtotals : up to 4 decimals  ($0.20, $0.316, $0.5231)
  - margin, FOB, Net, Others, Profit/OH/Finance amounts : 2 decimals
Only number formats are changed; values, formulas, colors, layout stay untouched.
"""
import re
import openpyxl

PRICE4 = '"$"#,##0.00##'          # min 2, max 4 decimals
PRICE2 = '"$"#,##0.00'
YY4 = '0.00##'
LOSS_INT = '0.00'
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
    pf = {lbl: find_label(ws, h + 1, r1, {lbl}) for lbl in ('PROFITS', 'PROFIT', 'OH', 'FINANCE')}
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
        ws.cell(r, cols['b_ttl']).number_format = PRICE2

    # others: rate 0.00%, amount 2 decimals
    if others:
        r = others.row
        for col in range(others.column + 1, cols['i_ttl']):
            cc = ws.cell(r, col)
            if isinstance(cc.value, (int, float)) and cc.value < 1:
                cc.number_format = PCT2
        ws.cell(r, cols['i_ttl']).number_format = PRICE2

    # net + final row (margin, internal total, FOB)
    if net:
        ws.cell(net.row, cols['i_ttl']).number_format = PRICE2
        for r in range(net.row, min(net.row + 3, r1) + 1):
            fob = ws.cell(r, cols['b_ttl'])
            if isinstance(fob.value, str) and fob.value.upper().startswith('=SUM'):
                fob.number_format = PRICE2
                ws.cell(r, cols['i_ttl']).number_format = PRICE2
                ws.cell(r, cols['i_loss']).number_format = PCT2
                if 'b_remark' in cols:
                    q = ws.cell(r - 1, cols['b_remark'])
                    if isinstance(q.value, str) and q.value.upper().startswith('=SUM'):
                        q.number_format = PRICE2   # profit+OH+finance sum
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
        xml = _re.sub(r'<c r="([A-Z]+\d+)"[^>]*?/>', rep, xml)
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

"""Internal CBS history update: compare the latest tab (leftmost) with the previous stage tab and write what
changed (price / yy / loss / CIF / CM, internal and open side) into the REMARK column in red, keeping what is
already written there. Prices are never changed (an open price that stays while the internal cost went down is
the margin going up - on purpose). The file is also tidied like in 'CBS 만들기' (number formats, missing
section subtotals) and odd CIF are flagged.
"""
import datetime as dt
import io
import re

import openpyxl
from openpyxl.utils import get_column_letter as L

from cbs_fill import read_internal_block, match_prev_lines
from cbs_format import normalize_bytes, set_cells, subtotal_plan, highlight_cells
from internal import read_internal, previous_stage_block, tab_date

FIELDS = [  # (column key, label, money?) - i_* go to the internal REMARK, b_* to the open REMARK
    ('i_price', 'Price', True), ('i_yy', 'YY', False), ('i_loss', 'LOSS', False),
    ('b_cif', 'CIF', True), ('b_price', 'Price', True), ('b_yy', 'YY', False), ('b_loss', 'LOSS', False),
]


# notes this tool wrote before (Korean / English versions) - replaced on every run, never the user's text
APP_NOTE = re.compile(r'^(\[vs [^\]]+\]|\[[^\]]+ 대비\]|⚠ (Internal YY up|YY up|Open YY 0)\b)')


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).replace('$', '').replace(',', '').replace('%', '').strip())
    except (TypeError, ValueError):
        return None


def _fmt(v, money):
    if v is None:
        return '-'
    return f'${v:.4f}'.rstrip('0').rstrip('.') if money else f'{v:.4f}'.rstrip('0').rstrip('.')


def _short_tab(title):
    d = tab_date(title)
    stage = re.split(r'[\s\-_]+', title.strip())[0]
    return f'{stage} {d[0]}/{d[1]}' if d else title


def _vals(ws, blk, line):
    c = blk['cols']
    return {k: _num(ws.cell(line['row'], c[k]).value) if k in c else None for k, _, _ in FIELDS}


def compare(wv, ws_cur, cur_blk, heather):
    """[(row, note, kind)] for the current block vs the previous stage block."""
    pv = previous_stage_block(wv, heather)
    if not pv:
        return None, []
    si, title, r0, r1, _fob = pv
    ws_prev = wv.worksheets[si]
    prev_blk = read_internal_block(ws_prev, r0, r1)
    tag = f'[vs {_short_tab(title)}]'
    m = match_prev_lines(cur_blk['lines'], prev_blk['lines'])
    notes = []          # (row, text, kind, line, side)  side: 'i' internal REMARK / 'b' open REMARK
    for i, a in enumerate(cur_blk['lines']):
        if i not in m:
            if a['yy'] or a['price']:
                notes.append((a['row'], f'{tag} New item', 'new', a, 'i'))
            continue
        b = prev_blk['lines'][m[i]]
        va, vb = _vals(ws_cur, cur_blk, a), _vals(ws_prev, prev_blk, b)
        ch = {'i': [], 'b': []}
        for k, lab, money in FIELDS:
            x, y = va[k], vb[k]
            if x is None and y is None:
                continue
            if x is None or y is None or round(x, 4) != round(y, 4):
                arrow = '↑' if (x or 0) > (y or 0) else '↓'
                ch[k[0]].append(f'{lab} {_fmt(y, money)}→{_fmt(x, money)}{arrow}')
        if ch['i']:
            notes.append((a['row'], f"{tag} Internal: {', '.join(ch['i'])}", 'changed', a, 'i'))
        if ch['b']:
            notes.append((a['row'], f"{tag} Open: {', '.join(ch['b'])}", 'changed', a, 'b'))
        # internal YY up but open YY kept: fine for a closed MNMK, wrong if fit comment / spec made it bigger
        if va['i_yy'] and vb['i_yy'] and va['b_yy'] is not None and vb['b_yy'] is not None \
                and round(va['i_yy'], 4) > round(vb['i_yy'], 4) and round(va['b_yy'], 4) == round(vb['b_yy'], 4):
            ratio = vb['b_yy'] / vb['i_yy'] if vb['i_yy'] else None
            sug = f", last ratio x{ratio:.4f} -> Open YY {va['i_yy'] * ratio:.4f}" if ratio else ''
            notes.append((a['row'], f"⚠ Internal YY up vs {_short_tab(title)} but Open YY unchanged - double check "
                                    f"(if fit comment/spec, raise Open YY{sug})", 'yy_up', a, 'b'))
    # open side empty while the internal side is filled -> the open FOB is missing that cost (not something we did)
    for a in cur_blk['lines']:
        va = _vals(ws_cur, cur_blk, a)
        if (va['i_yy'] or 0) > 0 and (va['i_price'] or 0) > 0 and not (va['b_yy'] or 0):
            al = ws_cur.cell(a['row'], cur_blk['cols']['b_actual_loss']).value if 'b_actual_loss' in cur_blk['cols'] else 'x'
            hint = ' (ACTUAL LOSS empty)' if al in (None, '') else ''
            notes.append((a['row'], f'⚠ Open YY 0{hint} - open cost calculated as 0', 'open0', a, 'b'))
    gone = [b for j, b in enumerate(prev_blk['lines']) if j not in set(m.values()) and (b['yy'] or b['price'])]
    # CM - internal CM (TTL) and open CM (open TTL) separately
    cr, pr = cur_blk.get('cm_row'), prev_blk.get('cm_row')
    if cr and pr:
        for side, col, lab in (('i', 'i_ttl', 'Internal'), ('b', 'b_ttl', 'Open')):
            cc, pc = cur_blk['cols'].get(col), prev_blk['cols'].get(col)
            if not cc or not pc:
                continue
            x, y = _num(ws_cur.cell(cr, cc).value), _num(ws_prev.cell(pr, pc).value)
            if (x is not None or y is not None) and (x is None or y is None or round(x, 4) != round(y, 4)):
                arrow = '↑' if (x or 0) > (y or 0) else '↓'
                notes.append((cr, f"{tag} {lab}: CM {_fmt(y, True)}→{_fmt(x, True)}{arrow}", 'changed', None, side))
    return dict(title=title, gone=gone, tag=tag), notes


def update(data, name, today=None):
    """Return (new_bytes, report)."""
    from engine import missing_values, cif_marks, clean_name
    rep = dict(blocks=[], notes=[], skipped=[], warnings=[], open0=[])
    calc = data
    if missing_values(data) > 0.3:
        from xlimages import recalc_xlsx
        calc = recalc_xlsx(data) or data
    info, wf, wv = read_internal(io.BytesIO(calc), name)
    rep['tab'] = info['tab']
    wsv, wsf = wv.worksheets[0], wf.worksheets[0]
    cells, all_lines = {}, []
    for b in info['blocks']:
        blk = read_internal_block(wsv, b['r0'], b['r1'], wsf)
        all_lines += [dict(block='heather' if b['heather'] else 'solid', tab=info['tab'], row=l['row'], key=l['key'],
                           cat=l['cat'], item=l['item'], detail=l['detail'], price=l['price'], cif=l['cif'], yy=l['yy'],
                           cif_col=blk['cols'].get('b_cif'), note_col=(blk['cols'].get('b_ttl') or 0) + 1)
                      for l in blk['lines'] if not l['lump'] and l['yy']]
        prev, notes = compare(wv, wsv, blk, b['heather'])
        rep['blocks'].append(dict(variant='Heather' if b['heather'] else 'Solid',
                                  prev=prev['title'] if prev else None,
                                  gone=[f"{g['item'] or ''} {g['detail'] or ''}".strip() for g in (prev or {}).get('gone', [])]))
        rcs = {'i': blk['cols'].get('i_remark'), 'b': blk['cols'].get('b_remark')}
        if not rcs['i'] and not rcs['b']:
            continue
        # several notes for one cell are written together (internal and open REMARK are different cells)
        per = {}
        # cells that still hold notes from an earlier run of this tool -> rewritten (the user's own text stays)
        for side in ('i', 'b'):
            if rcs[side]:
                for r in range(b['r0'], b['r1'] + 1):
                    v = wsv.cell(r, rcs[side]).value
                    if isinstance(v, str) and any(APP_NOTE.match(x.strip()) for x in v.split('\n')):
                        per.setdefault(f'{L(rcs[side])}{r}', [])
        for row, note, kind, line, side in notes:
            rc = rcs[side] or rcs['i' if side == 'b' else 'b']
            if rc != rcs[side] and side == 'b' and not note.startswith(('⚠', '[')):
                note = 'Open: ' + note
            per.setdefault(f'{L(rc)}{row}', []).append((note, kind, line))
        for ref, items in per.items():
            fc = wsf[ref]
            old = wsv[ref].value
            if isinstance(fc.value, str) and fc.value.startswith('='):
                rep['skipped'] += [(ref, n) for n, _, _ in items]   # remark is a formula - leave it
                continue
            old_t = '' if old is None else (f'{old:g}' if isinstance(old, float) else str(old))
            kept = [x for x in old_t.split('\n') if not APP_NOTE.match(x.strip())]
            cleaned = len(kept) != len(old_t.split('\n'))
            old_t = '\n'.join(kept).strip('\n') if cleaned else old_t
            new_notes = [n for n, _, _ in items if n not in old_t]
            if not new_notes and not cleaned:
                continue
            if not new_notes:
                cells[ref] = ('t', old_t, None) if old_t else ('t', '', None)
                continue
            new_t = ('\n' if old_t else '') + '\n'.join(new_notes)
            font = fc.font
            cells[ref] = ('r', (old_t, new_t, font.name if font else None, font.sz if font else None), None)
            for note, kind, line in items:
                if note not in new_notes:
                    continue
                if kind == 'open0':
                    rep['open0'].append(f"{(line or {}).get('detail') or ''} ({ref})")
                rep['notes'].append(dict(variant='Heather' if b['heather'] else 'Solid', cell=ref,
                                         item=(line or {}).get('detail') or (line or {}).get('item') or 'CM',
                                         before=old_t, note=note))
    # tidy like 'CBS 만들기'
    out, fmt_log = normalize_bytes(data)
    plan, sub_log = subtotal_plan(calc)
    cells = {**dict(plan), **cells}
    rep_c = dict(cif_lines=all_lines, warnings=[])
    marks = []
    try:
        marks = cif_marks(rep_c, name, calc)
    except Exception as e:                                      # noqa
        rep['warnings'].append(f'CIF check: {e}')
    for mk in marks:
        if mk['note_ref'] and mk['note_ref'] not in cells:
            cells[mk['note_ref']] = ('t', mk['note'], None)
        if mk['empty']:
            cells.setdefault(mk['ref'], ('t', '', None))
    if cells:
        out = set_cells(out, cells)
    if marks:
        out = highlight_cells(out, [mk['ref'] for mk in marks])
    rep['cif_issues'] = rep_c.get('cif_issues') or []
    rep['subtotals_added'] = sub_log
    rep['format_log'] = fmt_log
    rep['filename'] = clean_name(name)
    return out, rep

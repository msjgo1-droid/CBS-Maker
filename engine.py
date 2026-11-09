"""End-to-end: internal CBS (+ blank template / previous submission / MNMK) -> buyer CBS workbook."""
import datetime as dt
import io
import os
import re

import openpyxl

from internal import read_internal
from cbs_fill import read_internal_block, fill_sheet, is_main_sheet, buyer_columns, buyer_rows, calc_fob, _norm, _s
from panels import (is_panel_sheet, read_mnmk, pick_block, build_from_mnmk, recompute_marker,
                    panel_yy, panel_codes)
from xlutil import copy_sheet

HERE = os.path.dirname(os.path.abspath(__file__))
FACTORY = 'PIVNVN'


def variant_of(title):
    t = title.lower()
    if 'heather' in t and 'solid' not in t:
        return 'heather'
    if 'solid' in t and 'heather' not in t:
        return 'solid'
    return None


def load(src):
    """src: path or bytes."""
    if isinstance(src, (bytes, bytearray)):
        return openpyxl.load_workbook(io.BytesIO(src)), openpyxl.load_workbook(io.BytesIO(src), data_only=True)
    return openpyxl.load_workbook(src), openpyxl.load_workbook(src, data_only=True)


def default_block_map(info, main_titles):
    act = [b for b in info['blocks'] if b['idx'] in info['active']] or info['blocks'][-1:]
    out = {}
    for t in main_titles:
        v = variant_of(t)
        if v == 'heather':
            cand = [b for b in act if b['heather']] or [b for b in info['blocks'] if b['heather']]
        elif v == 'solid':
            cand = [b for b in act if not b['heather']]
        else:
            cand = act
        out[t] = (cand or act)[0]['idx']
    return out


def make_filename(ref_name, info, today=None):
    today = today or dt.date.today()
    d8 = today.strftime('%m%d%Y')
    base = os.path.splitext(os.path.basename(ref_name or ''))[0]
    base = re.sub(r'^[0-9a-f]{8}-', '', base)            # upload prefix
    parts = base.split('_')
    if len(parts) >= 4 and re.fullmatch(r'\d{8}', parts[-1]):
        parts[-1] = d8
        parts = [FACTORY if re.fullmatch(r'P[IT][A-Z]*HJ|PTHJ', p) else p for p in parts]
        return '_'.join(parts) + '.xlsx'
    return f"{info.get('style') or 'STYLE'}_{info.get('gender') or ''}_NAME_SEASON_{FACTORY}_STAGE_{d8}.xlsx"


def generate(internal, internal_name, base=None, base_name='', prev=None, prev_name='', mnmk=None,
             block_map=None, manual=None, crop=True, today=None):
    rep = dict(warnings=[], sheets={}, panels=[], filename='')
    info, wf, wv = read_internal(io.BytesIO(internal) if isinstance(internal, bytes) else internal, internal_name)
    rep['internal'] = info
    if not info['gender']:
        from internal import gender_of
        info['gender'] = gender_of(re.sub(r'^[0-9a-f]{8}-', '', prev_name or ''), re.sub(r'^[0-9a-f]{8}-', '', base_name or ''))
    if info['hidden']:
        rep['warnings'].append('Internal hidden sheets: ' + ', '.join(info['hidden']))
    wsv = wv.worksheets[0]

    prev_wb = load(prev)[0] if prev else None
    if base:
        out, _ = load(base)
    elif prev:
        out, _ = load(prev)
    else:
        raise ValueError('need blank template or previous submission')
    import zipfile
    for raw, nm in ((base, 'Template'), (prev, 'Previous')):
        if raw:
            z = zipfile.ZipFile(io.BytesIO(raw) if isinstance(raw, bytes) else raw)
            if any(n.lower().endswith(('.wmf', '.emf')) for n in z.namelist()):
                rep['warnings'].append(f'{nm}: WMF/EMF image found — it will be dropped in the output, re-insert manually')
    for wb_, nm in ((out, 'Template'), (prev_wb, 'Previous')):
        if wb_:
            hid = [s.title for s in wb_.worksheets if s.sheet_state != 'visible']
            if hid:
                rep['warnings'].append(f'{nm} hidden sheets: ' + ', '.join(hid))

    # ---- main sheets (Solid/Heather duplication for CO) ----
    mains = [s for s in out.worksheets if is_main_sheet(s)]
    if prev_wb and base:
        prev_mains = [s.title for s in prev_wb.worksheets if is_main_sheet(s)]
        if len(prev_mains) > 1 and len(mains) == 1:
            src = mains[0]
            src.title = prev_mains[0]
            for k, t in enumerate(prev_mains[1:], 1):
                copy_sheet(src, out, t, out.worksheets.index(src) + k)
            mains = [s for s in out.worksheets if is_main_sheet(s)]
    titles = [s.title for s in mains]
    bmap = dict(default_block_map(info, titles), **(block_map or {}))

    # ---- fill ----
    for ws in mains:
        b = info['blocks'][bmap[ws.title]]
        blk = read_internal_block(wsv, b['r0'], b['r1'])
        res = fill_sheet(ws, blk, variant_of(ws.title), (manual or {}).get(ws.title))
        res['block'] = b['idx']
        res['internal_fob'] = b['fob']
        res['fob'], res['material'] = calc_fob(ws)
        res['lines'] = [f"{l['cat']} / {l['detail']}" for l in blk['lines']]
        rep['sheets'][ws.title] = res

    # ---- panel sheets ----
    if mnmk:
        mwb = openpyxl.load_workbook(io.BytesIO(mnmk) if isinstance(mnmk, bytes) else mnmk, data_only=True)
        msheets, w = read_mnmk(mwb)
        rep['warnings'] += w
        layout = next((s for s in (prev_wb.worksheets if prev_wb else []) if is_panel_sheet(s)), None) \
            or next((s for s in out.worksheets if is_panel_sheet(s)), None) \
            or openpyxl.load_workbook(os.path.join(HERE, 'assets', 'panel_layout.xlsx')).worksheets[0]
        for s in [s for s in out.worksheets if is_panel_sheet(s)]:
            out.remove(s)
        for ms in msheets:
            blk, warn = pick_block(ms['blocks'], info['gender'])
            if warn:
                rep['warnings'].append(f"[{ms['title']}] {warn}")
            build_from_mnmk(out, layout, ms['title'], blk, crop=crop)
    elif prev_wb and base:
        for s in prev_wb.worksheets:
            if is_panel_sheet(s) and s.title not in out.sheetnames:
                copy_sheet(s, out, s.title)
    rep['panels'] = sync_panels(out, mains, info, rep['warnings'])
    rep['filename'] = make_filename(prev_name or base_name, info, today)
    buf = io.BytesIO()
    out.save(buf)
    return buf.getvalue(), rep


def sync_panels(out, mains, info, warnings):
    """Match each panel sheet to a fabric row; back-calculate marker length from Net Usage."""
    rows = []                                  # (main, row, component, material_codes, net)
    for ws in mains:
        h, c = buyer_columns(ws)
        for r, comp, sec in buyer_rows(ws)[0]:
            if sec not in ('body', 'lining - apparel'):
                continue
            mat = _s(ws.cell(r, c['material']).value).upper()
            codes = {t for t in re.split(r'[\s/,()#]+', mat) if len(t) >= 5 and re.search(r'\d', t)}
            net = ws.cell(r, c['net']).value
            rows.append(dict(sheet=ws.title, row=r, comp=comp, codes=codes,
                             net=net if isinstance(net, (int, float)) else 0))
    panels = [s for s in out.worksheets if is_panel_sheet(s)]
    groups = {}
    for p in panels:
        pc, pt = panel_codes(p), _norm(re.sub(r'\(.*?\)', '', p.title))
        nm = lambda r: _norm(re.sub(r'\(.*?\)', '', r['comp']))
        by_name = [r for r in rows if r['net'] and nm(r) and (nm(r) in pt or pt in nm(r))]
        by_code = [r for r in rows if r['codes'] & pc and r['net']]
        both = [r for r in by_code if r in by_name]
        cand = both or (by_code if len({(r['comp'].lower()) for r in by_code}) == 1 else []) or by_name or by_code
        if not cand:
            warnings.append(f'[{p.title}] panel sheet: no fabric row matched → not recalculated')
            continue
        key = (cand[0]['comp'].lower(), round(cand[0]['net'], 4))
        groups.setdefault(key, []).append((p, cand[0]))
    out_rep = []
    for key, items in groups.items():
        net = items[0][1]['net']
        if len(items) == 1:
            main, aux = items[0][0], []
        else:
            comp = _norm(re.sub(r'\(.*?\)', '', items[0][1]['comp']))
            named = [p for p, _ in items if comp and comp in _norm(p.title)]
            main = named[0] if named else max(items, key=lambda x: panel_yy(x[0]) or 0)[0]
            aux = [p for p, _ in items if p is not main]
        aux_yy = sum(panel_yy(p) or 0 for p in aux)
        y = recompute_marker(main, round(net - aux_yy, 4))
        out_rep.append(dict(panel=main.title, component=items[0][1]['comp'], net=net,
                            aux=[p.title for p in aux], aux_yy=round(aux_yy, 4), marker=y,
                            check=panel_yy(main)))
    return out_rep

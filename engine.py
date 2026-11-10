"""End-to-end: internal CBS (+ blank template / previous submission / MNMK) -> buyer CBS + formatted internal."""
import datetime as dt
import io
import os
import re
import zipfile

import openpyxl
from PIL import Image

from internal import read_internal, gender_of, previous_stage_block, find_style
from cbs_fill import (price_drops, keep_previous, prev_split_map, FABRIC_SECTIONS, material_text, read_internal_block, fill_sheet, is_main_sheet, buyer_columns, buyer_rows, calc_fob,
                      format_fob, _norm, _s)
from cbs_format import normalize_bytes, set_cells, subtotal_plan
from panels import (is_panel_sheet, read_mnmk, pick_block, build_from_mnmk, recompute_marker,
                    panel_yy, panel_codes, find_fields, one_pc_one_way)
from xlimages import pictures, to_png
from orientation import judge
from xlutil import copy_sheet, add_image, img_bytes

HERE = os.path.dirname(os.path.abspath(__file__))
FACTORY = 'PIVNVN'


def W(code, **kw):
    return dict(code=code, **kw)


def clean_name(n):
    return re.sub(r'^[0-9a-f]{8}-', '', os.path.basename(n or ''))


def variant_of(title):
    t = title.lower()
    if 'heather' in t and 'solid' not in t:
        return 'heather'
    if 'solid' in t and 'heather' not in t:
        return 'solid'
    return None


def load(src):
    return openpyxl.load_workbook(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)


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
    base = os.path.splitext(clean_name(ref_name))[0]
    parts = base.split('_')
    if len(parts) >= 4 and re.fullmatch(r'\d{8}', parts[-1]):
        parts[-1] = d8
        parts = [re.sub(r'P[IT][A-Z]*HJ$', FACTORY, p) if re.search(r'P[IT][A-Z]*HJ$', p) else p for p in parts]
        return '_'.join(parts) + '.xlsx'
    return f"{info.get('style') or 'STYLE'}_{info.get('gender') or ''}_NAME_SEASON_{FACTORY}_STAGE_{d8}.xlsx"


# ---------------- sketch ----------------
def sketch_candidates(internal_bytes, info):
    """Pictures in columns A–C of the internal's first tab, topmost-in-active-block first."""
    pics = [p for p in pictures(internal_bytes, 0) if p['col'] <= 2]
    act = [b for b in info['blocks'] if b['idx'] in info['active']]
    r0 = min((b['r0'] for b in act), default=1) - 2
    r1 = max((b['r1'] for b in act), default=10 ** 6)
    pics.sort(key=lambda p: (not (r0 <= p['row'] <= r1), p['row']))
    return pics


def place_sketch(ws, png):
    """Put the garment sketch in the header, just left of the top 'FOB' label (rows 1-7)."""
    label = None
    for row in ws.iter_rows(max_row=9):
        for c in row:
            if _s(c.value).upper() == 'FOB' and c.column > 2:
                label = c
    if not label:
        return False
    for im in list(ws._images):                          # drop old header picture
        a = im.anchor
        if hasattr(a, '_from') and a._from.row < 8 and a._from.col >= 3:
            ws._images.remove(im)
    c0 = max(label.column - 4, 4)
    from openpyxl.utils import get_column_letter as L
    wpx = sum(int((ws.column_dimensions[L(c)].width or 8.43) * 7 + 5) for c in range(c0, label.column))
    hpx = sum(int((ws.row_dimensions[r].height or 15) * 96 / 72) for r in range(1, 8)) - 6
    im = Image.open(io.BytesIO(png))
    k = min(wpx / im.width, hpx / im.height)
    add_image(ws, png, f'{L(c0)}1', int(im.width * k), int(im.height * k))
    return True


# ---------------- main ----------------
def generate(internal, internal_name, base=None, base_name='', prev=None, prev_name='', mnmk=None,
             block_map=None, manual=None, crop=True, sketch=0, today=None, keep=None):
    rep = dict(warnings=[], sheets={}, panels=[], filename='', self_check=None, drops={}, internal_edits={})
    info, wf, wv = read_internal(io.BytesIO(internal), clean_name(internal_name))
    rep['internal'] = info
    if not info['gender']:
        info['gender'] = gender_of(clean_name(prev_name), clean_name(base_name))
    if info['hidden']:
        rep['warnings'].append(W('hidden', src='Internal', names=', '.join(info['hidden'])))
    for b in info['blocks']:
        if b['idx'] in info['active'] and b['warn']:
            rep['warnings'].append(W('block', block=b['idx'], text='; '.join(b['warn'])))
    wsv = wv.worksheets[0]

    prev_wb = load(prev) if prev else None
    if not (base or prev):
        raise ValueError('need blank template or previous submission')
    out = load(base or prev)
    for raw, nm in ((base, 'Template'), (prev, 'Previous')):
        if raw:
            z = zipfile.ZipFile(io.BytesIO(raw))
            if any(n.lower().endswith(('.wmf', '.emf')) for n in z.namelist()):
                rep['warnings'].append(W('wmf', src=nm))
            wb_ = load(raw)
            hid = [s.title for s in wb_.worksheets if s.sheet_state != 'visible']
            if hid:
                rep['warnings'].append(W('hidden', src=nm, names=', '.join(hid)))

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
    bmap = dict(default_block_map(info, [s.title for s in mains]), **(block_map or {}))
    psplit = prev_split_map(prev_wb) if prev_wb else {}

    # ---- sketch ----
    cands = sketch_candidates(internal, info)
    rep['sketches'] = [dict(row=p['row'] + 1, col=p['col'] + 1, ext=p['ext'], name=p['name']) for p in cands]
    png = None
    if cands:
        sel = cands[min(sketch or 0, len(cands) - 1)]
        png = to_png(sel['data'], sel['ext'])
        rep['sketch_png'] = png
        if not png:
            rep['warnings'].append(W('sketch_conv', ext=sel['ext']))
    else:
        rep['warnings'].append(W('no_sketch'))

    # ---- fill ----
    for ws in mains:
        b = info['blocks'][bmap[ws.title]]
        blk = read_internal_block(wsv, b['r0'], b['r1'], wf.worksheets[0])
        pv = previous_stage_block(wv, b['heather'])
        if pv:
            pblk = read_internal_block(wv.worksheets[pv[0]], pv[2], pv[3])
            rows_ = price_drops(blk, pblk)
            if b['fob'] is not None and pv[4] is not None and b['fob'] < pv[4] - 0.005:
                rep['drops'][ws.title] = dict(prev_tab=pv[1], prev_fob=pv[4], cur_fob=b['fob'], rows=rows_,
                                              block=b['idx'])
                rep['warnings'].append(W('price_drop', sheet=ws.title, tab=pv[1], prev=f'{pv[4]:.2f}',
                                         cur=f"{b['fob']:.2f}"))
            k = (keep or {}).get(ws.title)
            if k:
                rep['internal_edits'].update(keep_previous(blk, pblk, k))
        res = fill_sheet(ws, blk, variant_of(ws.title), (manual or {}).get(ws.title), prev_splits=psplit)
        format_fob(ws)
        if png:
            place_sketch(ws, png)
        res['block'] = b['idx']
        res['internal_fob'] = b['fob']
        res['fob'], res['material'] = calc_fob(ws)
        res['lines'] = [f"{l['cat']} / {l['detail']}" for l in blk['lines']]
        rep['sheets'][ws.title] = res
        for l in res['log']:
            if l['how'] == 'split':
                rep['warnings'].append(W('split', sheet=ws.title, row=l['row'], comp=l['component'][:60], src=l['split']))
            if l['how'] == 'shared (0)':
                rep['warnings'].append(W('shared', sheet=ws.title, row=l['row'], comp=l['component'][:60],
                                         match=l['match'][:60]))
            if l['how'] == 'NO MATCH':
                rep['warnings'].append(W('no_match', sheet=ws.title, row=l['row'], comp=l['component']))
        for a in res['added']:
            rep['warnings'].append(W('row_added', sheet=ws.title, row=a['row'], comp=a['component'], sec=a['section']))
        if res['unplaced']:
            rep['warnings'].append(W('not_placed', sheet=ws.title, items=', '.join(res['unplaced'])))

    # ---- panel sheets ----
    if mnmk:
        mwb = openpyxl.load_workbook(io.BytesIO(mnmk), data_only=True)
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
                rep['warnings'].append(dict(warn, sheet=ms['title']))
            build_from_mnmk(out, layout, ms['title'], blk, crop=crop)
    elif prev_wb and base:
        for s in prev_wb.worksheets:
            if is_panel_sheet(s) and s.title not in out.sheetnames:
                copy_sheet(s, out, s.title)
    rep['panels'] = sync_panels(out, mains, info, rep['warnings'])
    # fabric (marker) sheets: DATE = the day this CBS is made
    run_day = today or dt.date.today()
    for p in [s for s in out.worksheets if is_panel_sheet(s)]:
        f = find_fields(p, 1, 12)
        if 'date' in f:
            cell = p.cell(*f['date'])
            cell.value = dt.datetime(run_day.year, run_day.month, run_day.day)
            cell.number_format = 'yyyy-mm-dd'

    # ---- Self (body) must be 1pc / 1way, laid LENGTHWISE — judged from the marker picture ----
    sp = self_panel(out, rep['panels'])
    chk = dict(sheet=sp.title if sp else None, verdict='unknown', share=None, annotated=None)
    if sp is not None and sp._images:
        img = max(sp._images, key=lambda i: (i.width or 0) * (i.height or 0))
        j = judge(img_bytes(img))
        chk.update(verdict=j['verdict'], share=j['share'], annotated=j['annotated'])
    rep['self_check'] = chk
    if chk['verdict'] != 'lengthwise':
        rep['warnings'].insert(0, W('not_1pc', sheet=chk['sheet'] or '-', verdict=chk['verdict'],
                                    share=int((chk['share'] or 0) * 100)))

    rep['filename'] = make_filename(base_name or prev_name, info, today)
    buf = io.BytesIO()
    out.save(buf)
    return buf.getvalue(), rep


def self_panel(out, panels_rep):
    sheets = [s for s in out.worksheets if is_panel_sheet(s)]
    for s in sheets:
        if re.search(r'SELF|BODY', s.title.upper()):
            return s
    for p in panels_rep:
        if re.search(r'SELF|BODY', p['component'].upper()):
            return out[p['panel']]
    return sheets[0] if sheets else None


def sync_panels(out, mains, info, warnings):
    """Match each panel sheet to a fabric row; back-calculate marker length from Net Usage."""
    rows = []
    for ws in mains:
        h, c = buyer_columns(ws)
        for r, comp, sec in buyer_rows(ws)[0]:
            if sec not in FABRIC_SECTIONS:
                continue
            mat = material_text(ws, r, c).upper()
            codes = {t for t in re.split(r'[\s/,()#]+', mat) if len(t) >= 5 and re.search(r'\d', t)}
            net = ws.cell(r, c['net']).value
            rows.append(dict(sheet=ws.title, row=r, comp=comp, codes=codes, mat=_norm(mat),
                             net=net if isinstance(net, (int, float)) else 0))
    groups = {}
    for p in [s for s in out.worksheets if is_panel_sheet(s)]:
        pc, pt = panel_codes(p), _norm(re.sub(r'\(.*?\)', '', p.title))
        nm = lambda r: _norm(re.sub(r'\(.*?\)', '', r['comp']))
        by_name = [r for r in rows if r['net'] and nm(r) and (nm(r) in pt or pt in nm(r))] or \
                  [r for r in rows if r['net'] and len(pt) > 4 and pt in r['mat']]
        by_code = [r for r in rows if r['codes'] & pc and r['net']]
        both = [r for r in by_code if r in by_name]
        cand = both or (by_code if len({r['comp'].lower() for r in by_code}) == 1 else []) or by_name or by_code
        if not cand:
            warnings.append(W('panel_nomatch', sheet=p.title))
            continue
        c0 = cand[0]
        if c0['codes']:                    # fabric split over several rows (Shell 1 / Contrast 1): use the total
            same = [r for r in rows if r['sheet'] == c0['sheet'] and r['codes'] == c0['codes']]
            c0 = dict(c0, net=round(sum(r['net'] for r in same), 4))
        key = (c0['comp'].lower(), round(c0['net'], 4))
        groups.setdefault(key, []).append((p, c0))
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


# ---------------- one style = formatted internal + buyer CBS ----------------
def process_style(files, today=None, **opts):
    """files: dict(internal=(name,bytes), base=..., prev=..., mnmk=...)."""
    iname, idata = files['internal']
    kw = dict(internal=idata, internal_name=iname, today=today, **opts)
    for k in ('base', 'prev'):
        if files.get(k):
            kw[k], kw[k + '_name'] = files[k][1], files[k][0]
    if files.get('mnmk'):
        kw['mnmk'] = files['mnmk'][1]
    data, rep = generate(**kw)
    fmt_bytes, fmt_log = normalize_bytes(idata)
    # record what went to the buyer: "CBS m/d $FOB" right of the open (buyer) FOB of each block used
    d = today or dt.date.today()
    notes = {}
    for sh, r in rep['sheets'].items():
        b = rep['internal']['blocks'][r['block']]
        if b.get('fob_cell'):
            from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
            col, row = coordinate_from_string(b['fob_cell'])
            ref = f'{get_column_letter(column_index_from_string(col) + 1)}{row}'
            notes[ref] = f"CBS {d.month}/{d.day} ${r['fob']:.2f}"
    plan, sub_log = subtotal_plan(idata)                   # section subtotals that are missing
    cells = dict(plan)
    cells.update({k: ('t', v, None) for k, v in notes.items()})
    cells.update({k: ('n', v, None) for k, v in rep.get('internal_edits', {}).items()})
    if cells:
        fmt_bytes = set_cells(fmt_bytes, cells)
    rep['cbs_notes'] = notes
    rep['subtotals_added'] = sub_log
    rep['internal_formatted'] = (clean_name(iname), fmt_bytes)
    rep['format_log'] = fmt_log
    return data, rep


# ---------------- batch grouping ----------------
def style_of(name, data=None):
    st = find_style(clean_name(name))
    if st:
        return st
    if data:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            ws = wb.worksheets[0]
            for row in ws.iter_rows(max_row=12, max_col=6, values_only=True):
                for v in row:
                    st = find_style(v)
                    if st:
                        return st
        except Exception:                                 # noqa
            pass
    return None


def group_files(internals, bases, prevs, mnmks):
    """Each arg: list of (name, bytes). Returns ({style: files}, unmatched [(role, name)]).
    Styles are matched exactly (1187346A = Asia fit ≠ 1187346). A template/previous/MNMK whose number has no
    exact internal is attached to the one internal with the same 7 digits (e.g. MNMK named without the A)."""
    groups, loose, pending = {}, [], []
    for name, data in internals:
        st = style_of(name, data)
        if not st:
            loose.append(('internal', clean_name(name)))
        elif 'internal' in groups.get(st, {}):
            loose.append(('internal (duplicate)', clean_name(name)))
        else:
            groups.setdefault(st, {})['internal'] = (name, data)
    for role, lst in (('base', bases), ('prev', prevs), ('mnmk', mnmks)):
        for name, data in lst:
            st = style_of(name, data)
            if st and st not in groups:
                same = [k for k in groups if k[:7] == st[:7]]
                if len(same) == 1:
                    st = same[0]
            if not st:
                loose.append((role, clean_name(name)))
                continue
            g = groups.setdefault(st, {})
            if role in g:
                loose.append((role + ' (duplicate)', clean_name(name)))
                continue
            g[role] = (name, data)
    return groups, loose

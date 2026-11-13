"""End-to-end: internal CBS (+ blank template / previous submission / MNMK) -> buyer CBS + formatted internal."""
import datetime as dt
import io
import os
import re
import zipfile

import openpyxl
from PIL import Image

from internal import read_internal, gender_of, previous_stage_block, find_style
from cbs_fill import (code_overlap, yy_drops, price_drops, keep_previous, prev_split_map, FABRIC_SECTIONS, material_text, read_internal_block, fill_sheet, is_main_sheet, buyer_columns, buyer_rows, calc_fob,
                      format_fob, _norm, _s)
from cbs_format import normalize_bytes, set_cells, subtotal_plan, highlight_cells
from panels import (is_panel_sheet, read_mnmk, pick_block, build_from_mnmk, recompute_marker,
                    panel_yy, panel_codes, find_fields, one_pc_one_way, crop_marker)
from xlimages import pictures, to_png
from internal_mnmk import read_internal_mnmk
from orientation import judge
from xlutil import copy_sheet, add_image, img_bytes

HERE = os.path.dirname(os.path.abspath(__file__))
FACTORY = 'PIVNVN'


def W(code, **kw):
    return dict(code=code, **kw)


def clean_name(n):
    n = re.sub(r'^[0-9a-f]{8}-', '', os.path.basename(n or ''))
    return re.sub(r'^\d{10,}_', '', n)          # download/upload timestamp prefix (e.g. 1794456312671_)


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
    """Solid sheet -> solid block, Heather sheet -> heather block (submitted block of that kind first,
    otherwise the last one of that kind); single sheet -> active block."""
    blocks = info['blocks']
    act = [b for b in blocks if b['idx'] in info['active']] or blocks[-1:]

    def pick(kind):
        pool = [b for b in blocks if b['heather'] == kind]
        sub = [b for b in pool if b['submitted']]
        return (sub or pool[-1:] or act)[0]
    out = {}
    for t in main_titles:
        v = variant_of(t)
        if v == 'heather':
            out[t] = pick(True)['idx']
        elif v == 'solid':
            out[t] = pick(False)['idx']
        else:
            out[t] = act[0]['idx']
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
             block_map=None, manual=None, crop=True, sketch=0, today=None, keep=None, pics=None, qty=None):
    rep = dict(warnings=[], sheets={}, panels=[], filename='', self_check=None, drops={}, basis={}, pic_choice=[],
               yy_drops={})
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
        rep.setdefault('cif_lines', []).extend(
            dict(block=('heather' if b.get('heather') else 'solid'), tab=info.get('tab'), row=l['row'], key=l['key'],
                 cat=l['cat'], item=l['item'], detail=l['detail'], price=l['price'], cif=l['cif'], yy=l['yy'],
                 cif_col=blk['cols'].get('b_cif'), note_col=(blk['cols'].get('b_ttl') or 0) + 1)
            for l in blk['lines'] if not l['lump'] and l['yy'])      # internal CIF as typed (record + sanity check)
        pv = previous_stage_block(wv, b['heather'])
        if pv:
            pblk = read_internal_block(wv.worksheets[pv[0]], pv[2], pv[3])
            rows_ = price_drops(blk, pblk)
            yd = yy_drops(blk, pblk)
            if yd:
                rep['yy_drops'][ws.title] = dict(cur_tab=info['tab'], prev_tab=pv[1], rows=yd)
            fell = b['fob'] is not None and pv[4] is not None and b['fob'] < pv[4] - 0.005
            if rows_:
                rep['drops'][ws.title] = dict(prev_tab=pv[1], prev_fob=pv[4], cur_fob=b['fob'], rows=rows_,
                                              block=b['idx'], fell=fell)
            if fell:
                rep['warnings'].append(W('price_drop', sheet=ws.title, tab=pv[1], prev=f'{pv[4]:.2f}',
                                         cur=f"{b['fob']:.2f}"))
            # buyer open basis: 'current' (internal open as is) / 'prev_yy' / 'prev_all' / {line: [fields]}
            k = (keep or {}).get(ws.title)
            if isinstance(k, str):
                k = {r['line']: ([f for f in r['why'] if f == 'yy'] if k == 'prev_yy' else r['why'])
                     for r in rows_ if k != 'current'}
                k = {i: f for i, f in k.items() if f}
            if k:
                keep_previous(blk, pblk, k)          # changes the BUYER side only — the internal file is never touched
                rep['basis'][ws.title] = dict(prev_tab=pv[1], kept={str(i): f for i, f in k.items()})
        res = fill_sheet(ws, blk, variant_of(ws.title), (manual or {}).get(ws.title), prev_splits=psplit)
        if b['fob'] is not None and res.get('excluded') and abs(calc_fob(ws)[0] - b['fob']) > 0.01:
            # leaving the small interlining out would break the ±$0.01 rule -> put it in
            res = fill_sheet(ws, blk, variant_of(ws.title), (manual or {}).get(ws.title), prev_splits=psplit,
                             include_minor=True)
        format_fob(ws)
        if png:
            place_sketch(ws, png)
        res['block'] = b['idx']
        res['internal_fob'] = b['fob']
        res['fob'], res['material'] = calc_fob(ws)
        res['basis'] = 'previous' if ws.title in rep['basis'] else 'internal'
        if ws.title in rep['basis']:
            rep['warnings'].insert(0, W('basis_prev', sheet=ws.title, tab=rep['basis'][ws.title]['prev_tab'],
                                        buyer=f"{res['fob']:.2f}", internal=f"{(b['fob'] or 0):.2f}"))
        elif b['fob'] is not None and abs(res['fob'] - b['fob']) > 0.01:
            rep['warnings'].insert(0, W('fob_gap', sheet=ws.title, buyer=f"{res['fob']:.4f}", internal=f"{b['fob']:.4f}",
                                        diff=f"{res['fob'] - b['fob']:+.4f}"))
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
        if res.get('repaired'):
            rep['warnings'].append(W('sum_fixed', sheet=ws.title, n=len(res['repaired'])))
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
    elif (prev_wb and any(is_panel_sheet(s) for s in prev_wb.worksheets)) or \
            any(is_panel_sheet(s) for s in out.worksheets):
        # previous CBS fabric sheets (or ones already in the uploaded buyer file) -> check picture & fill blanks
        if base and prev_wb:
            for s in prev_wb.worksheets:
                if is_panel_sheet(s) and s.title not in out.sheetnames:
                    copy_sheet(s, out, s.title)
        # marker PICTURE: internal right-side MNMK vs the one in the previous CBS -> keep the correct one
        b0 = info['blocks'][bmap[mains[0].title]] if mains else None
        lines0 = read_internal_block(wsv, b0['r0'], b0['r1'])['lines'] if b0 else []
        # fabric sheets already exist (previous submission): only the CURRENT tab's MNMK may replace them
        im = read_internal_mnmk(internal, wsv, lines0, info['gender'], other_tabs=False)
        if im:
            rep['pic_choice'] = choose_pictures(out, im, pics or {}, crop)
            apply_parts(out, im, crop, rep['warnings'], info.get('style'))
            rep['internal_mnmk'] = [dict(title=x['title'], line=x['line'], ok=x['ok'],
                                         **{k: x['fields'].get(k) for k in ('width', 'eff', 'gqty', 'pqty', 'yy', 'direction')})
                                    for x in im]
        for s in [s for s in out.worksheets if is_panel_sheet(s)]:
            f = find_fields(s, 1, 12)
            miss = [n for k, n in (('gqty', "Garment Q'ty"), ('pqty', "Panel Q'ty"))
                    if k in f and s.cell(*f[k]).value in (None, '')]
            if miss:
                rep['warnings'].append(W('ocr_check', sheet=s.title, miss=', '.join(miss)))
    elif not any(is_panel_sheet(s) for s in out.worksheets):
        # no previous submission, no MNMK file -> MNMK pictures on the right side of the internal
        b0 = info['blocks'][bmap[mains[0].title]] if mains else None
        lines0 = read_internal_block(wsv, b0['r0'], b0['r1'])['lines'] if b0 else []
        im = read_internal_mnmk(internal, wsv, lines0, info['gender'])
        if im and im[0].get('tab'):
            rep['warnings'].append(W('mnmk_tab', tab=im[0]['tab']))
        rep['internal_mnmk'] = [dict(title=x['title'], line=x['line'], ok=x['ok'],
                                     **{k: x['fields'].get(k) for k in ('width', 'eff', 'gqty', 'pqty', 'yy', 'direction')})
                                for x in im]
        if im:
            layout = openpyxl.load_workbook(os.path.join(HERE, 'assets', 'panel_layout.xlsx')).worksheets[0]
            rep['warnings'].append(W('from_internal', n=len(im)))
            for x in im:
                x['block']['style'] = info.get('style')
                build_from_mnmk(out, layout, x['title'], x['block'], crop=crop is not False)
            apply_parts(out, im, crop, rep['warnings'], info.get('style'))
            for x in im:
                miss = [k for k in ('width', 'eff', 'gqty', 'pqty') if not x['fields'].get(k)]
                if miss or x['ok'] is False:
                    rep['warnings'].append(W('ocr_check', sheet=x['title'], miss=', '.join(miss) or "Garment Q'ty"))
        else:
            rep['warnings'].append(W('no_panels'))
    for p in [s for s in out.worksheets if is_panel_sheet(s)]:
        _clear_layout_junk(p)
        # values typed in the app for what the MNMK picture could not be read for
        for k, v in ((qty or {}).get(p.title) or {}).items():
            f = find_fields(p, 1, 12)
            if k in f and v not in (None, ''):
                p.cell(*f[k]).value = (f'{v}"' if k == 'width' and str(v).replace('.', '').isdigit() else v)
    rep['panels'] = sync_panels(out, mains, info, rep['warnings'])
    # what is still empty on the fabric sheets (asked for in the app, marker length needs Garment Q'ty)
    rep['warnings'] = [w for w in rep['warnings'] if w.get('code') != 'ocr_check']
    rep['panel_fields'] = []
    for p in [s for s in out.worksheets if is_panel_sheet(s)]:
        f = find_fields(p, 1, 12)
        vals = {k: p.cell(*f[k]).value if k in f else None for k in ('width', 'eff', 'gqty', 'pqty')}
        rep['panel_fields'].append(dict(sheet=p.title, **vals))
        miss = [n for k, n in (('width', 'Width'), ('gqty', "Garment Q'ty"), ('pqty', "Panel Q'ty"))
                if k in f and vals[k] in (None, '')]
        if miss:
            rep['warnings'].append(W('ocr_check', sheet=p.title, miss=', '.join(miss)))
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


def _pic_match(ws, cands):
    """Internal right-side picture that belongs to this fabric sheet (codes, then name)."""
    for x in cands:
        if x['title'] == ws.title:
            return x
    pc = panel_codes(ws)
    for x in cands:
        if x.get('part_main') is False:
            continue                       # extra part of a split fabric -> only by its own sheet title
        if code_overlap(pc, {t for t in re.split(r'[\s/,()#:]+', x['header'].upper()) if len(t) >= 5
                             and re.search(r'\d', t)}):
            return x
    t = _norm(re.sub(r'\(.*?\)', '', ws.title))
    for x in cands:
        n = _norm(re.sub(r'\(.*?\)', '', x['title']))
        if t and n and (t in n or n in t) and x.get('part_main') is not False:
            return x
    return None


def apply_parts(out, im, crop, warnings, style=None):
    """Same fabric with several MNMK (internal '0.459+1.458'): extra part sheets get their own yy
    (part x open/internal loss factor); the main sheet later gets the rest (sync_panels: main = net - aux)."""
    layout = None
    for x in im:
        if x.get('part_main') is not False or not x.get('part_yy'):
            continue
        n_parts = sum(1 for y in im if y['_hcol'] == x['_hcol'])
        hc = {t for t in re.split(r'[\s/,()#:]+', x['header'].upper()) if len(t) >= 5 and re.search(r'\d', t)}
        have = [s for s in out.worksheets if is_panel_sheet(s) and s.title != x['title'] and code_overlap(panel_codes(s), hc)]
        if x['title'] not in out.sheetnames and len(have) >= n_parts:
            continue          # the buyer file already splits this fabric (e.g. SELF 1 + SELF 1 (BINDING))
        if x['title'] not in out.sheetnames:
            if layout is None:
                layout = openpyxl.load_workbook(os.path.join(HERE, 'assets', 'panel_layout.xlsx')).worksheets[0]
            x['block']['style'] = style
            main = next((y for y in im if y.get('part_main') and y['_hcol'] == x['_hcol']), None)
            idx = out.sheetnames.index(main['title']) + 1 if main and main['title'] in out.sheetnames else None
            build_from_mnmk(out, layout, x['title'], x['block'], index=idx, crop=crop is not False)
        ws = out[x['title']]
        yy = round(x['part_yy'] * (x.get('scale') or 1), 4)
        recompute_marker(ws, yy)
        warnings.append(W('split_part', sheet=x['title'], part=x['part_yy'], yy=yy))


def _is_self(title):
    return bool(re.search(r'SELF|BODY|SHELL', title.upper()))


def choose_pictures(out, im, pics, crop):
    """Per fabric sheet: put the internal MNMK picture or keep the previous CBS one.
    auto: SELF -> whichever is 1pc/1way lengthwise (internal if both or neither); other fabrics -> internal."""
    from openpyxl.utils import get_column_letter as L
    rep, used = [], set()
    for ws in [s for s in out.worksheets if is_panel_sheet(s)]:
        cand = [x for x in im if id(x) not in used]
        x = _pic_match(ws, cand)
        if not x:
            continue
        used.add(id(x))
        old = [i for i in ws._images if hasattr(i.anchor, '_from') and i.anchor._from.row >= 6]
        prev_png = img_bytes(max(old, key=lambda i: (i.width or 0) * (i.height or 0))) if old else None
        int_png = x['block']['image']
        if int_png is not None and not isinstance(int_png, (bytes, bytearray)):
            int_png = img_bytes(int_png)
        if int_png and crop is not False:
            int_png = crop_marker(int_png)[0]          # only the marker, not the whole CAD screen
        v_int = judge(int_png)['verdict'] if int_png else 'unknown'
        v_prev = judge(prev_png)['verdict'] if prev_png else 'unknown'
        want = pics.get(ws.title, 'auto')
        if want == 'auto':
            if not int_png:
                want = 'prev'
            elif _is_self(ws.title) and v_prev == 'lengthwise' and v_int != 'lengthwise':
                want = 'prev'
            else:
                want = 'internal'
        f = find_fields(ws, 1, 12)
        fx = x['fields']
        if want != 'internal':
            # previous picture kept, but fill fields the previous sheet left empty
            for k in ('width', 'eff', 'pqty', 'gqty'):
                if fx.get(k) and k in f and ws.cell(*f[k]).value in (None, ''):
                    ws.cell(*f[k]).value = fx[k]
        if want == 'internal' and int_png:
            for i in old:
                ws._images.remove(i)
            for k in ('width', 'eff', 'pqty', 'gqty'):
                if fx.get(k) and k in f:
                    ws.cell(*f[k]).value = fx[k]
            if fx.get('direction') and 'size' in f:
                ws.cell(*f['size']).value = x['block']['size']
            w = min(Image.open(io.BytesIO(int_png)).size[0], 900)
            sz = Image.open(io.BytesIO(int_png)).size
            row = (f['gqty'][0] + 2) if 'gqty' in f else 8
            add_image(ws, int_png, f'A{row}', w, int(sz[1] * w / sz[0]))
        rep.append(dict(sheet=ws.title, used=want, internal=v_int, previous=v_prev,
                        internal_png=int_png, previous_png=prev_png))
    return rep


def _clear_layout_junk(ws):
    """Stray cells an older app version copied from its layout asset (K5 '#REF!', L5 36, C7 10, E7 0.712 ...)."""
    if ws['K5'].value == '#REF!':
        for ref in ('K5', 'L5', 'M5', 'N5', 'O5', 'P5', 'Q5', 'C7', 'E7', 'J7'):
            ws[ref].value = None
    for row in ws.iter_rows(min_row=1, max_row=12):
        for c in row:
            if c.value == '#REF!':
                c.value = None


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
            cw = ws.cell(r, c['width']).value if c.get('width') else None
            rows.append(dict(sheet=ws.title, row=r, comp=comp, codes=codes, mat=_norm(mat), sec=sec,
                             net=net if isinstance(net, (int, float)) else 0, cw=cw))
    groups = {}
    for p in [s for s in out.worksheets if is_panel_sheet(s)]:
        pc, pt = panel_codes(p), _norm(re.sub(r'\(.*?\)', '', p.title))
        nm = lambda r: _norm(re.sub(r'\(.*?\)', '', r['comp']))
        by_name = [r for r in rows if r['net'] and nm(r) and (nm(r) in pt or pt in nm(r))] or \
                  [r for r in rows if r['net'] and len(pt) > 4 and pt in r['mat']]
        by_code = [r for r in rows if r['net'] and code_overlap(r['codes'], pc)]
        mnum = re.match(r'\s*self\s*(\d+)\b', re.sub(r'\(.*?\)', '', p.title), re.I)
        if not by_name and not by_code and mnum:
            # 'SELF 2' sheet <-> 'Shell 2' fabric row
            by_name = [r for r in rows if r['net'] and re.search(r'\bshell\s*%s\b' % mnum.group(1), r['comp'], re.I)]
        if not by_name and not by_code and pt.startswith('self'):
            # 'SELF' sheet <-> the Shell 1 / first Body fabric row of S28 templates
            body = [r for r in rows if r['net'] and r['sec'] == 'body']
            by_name = [r for r in body if 'shell' in r['comp'].lower()] or body[:1]
        both = [r for r in by_code if r in by_name]
        cand = both or (by_code if len({r['comp'].lower() for r in by_code}) == 1 else []) or by_name or by_code
        vt = variant_of(p.title)
        if vt and cand:
            cand = [r for r in cand if variant_of(r['sheet']) == vt] or cand
        if not cand:
            warnings.append(W('panel_nomatch', sheet=p.title))
            continue
        c0 = cand[0]
        if c0['codes']:                    # fabric split over several rows (Shell 1 / Contrast 1): use the total
            same = [r for r in rows if r['sheet'] == c0['sheet'] and r['codes'] == c0['codes']]
            c0 = dict(c0, net=round(sum(r['net'] for r in same), 4))
        key = (c0['comp'].lower(), round(c0['net'], 4))
        groups.setdefault(key, []).append((p, c0))
        # cuttable width unreadable in the MNMK picture -> the buyer template row has it
        fw = find_fields(p, 1, 12).get('width')
        try:
            cwv = float(str(c0.get('cw') or '').replace('"', '').strip())
        except ValueError:
            cwv = 0
        if cwv > 80:
            cwv = int(cwv / 2.54)          # template width in cm (138 -> 54")
        if fw and p.cell(*fw).value in (None, '') and cwv > 0:
            p.cell(*fw).value = f'{cwv:g}"'
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
def missing_values(data):
    """Share of formula cells on the first tab that have no cached result (file saved without calculation)."""
    wf = openpyxl.load_workbook(io.BytesIO(data))
    wv = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    a, b = wf.worksheets[0], wv.worksheets[0]
    n = miss = 0
    for row in a.iter_rows(max_col=20):
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                n += 1
                if b[c.coordinate].value is None:
                    miss += 1
    return miss / n if n else 0


def process_style(files, today=None, **opts):
    """files: dict(internal=(name,bytes), base=..., prev=..., mnmk=...)."""
    iname, idata = files['internal']
    pre_warn = []
    if missing_values(idata) > 0.3:
        # internal saved without calculated values (e.g. an app output never opened in Excel) -> recalc to read it
        from xlimages import recalc_xlsx
        rc = recalc_xlsx(idata)
        if rc:
            calc_for_read = rc
            pre_warn.append(W('recalc'))
        else:
            raise ValueError('Internal CBS has no calculated values — open it in Excel, save, and upload again')
    else:
        calc_for_read = idata
    kw = dict(internal=calc_for_read, internal_name=iname, today=today, **opts)
    for k in ('base', 'prev'):
        if files.get(k):
            kw[k], kw[k + '_name'] = files[k][1], files[k][0]
    if files.get('mnmk'):
        kw['mnmk'] = files['mnmk'][1]
    cif_hist = kw.pop('cif_hist', None)
    data, rep = generate(**kw)
    rep['warnings'] = pre_warn + rep['warnings']
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
    plan, sub_log = subtotal_plan(calc_for_read)                   # section subtotals that are missing
    cells = dict(plan)
    cells.update({k: ('t', v, None) for k, v in notes.items()})
    # CIF that does not look like what is used for that kind of item (history of all internal CBS)
    marks = []
    try:
        marks = cif_marks(rep, iname, calc_for_read, cif_hist)
    except Exception as ex:                                      # noqa - never stop a CBS for this
        rep['warnings'].append(W('cif_err', err=f'{type(ex).__name__}: {ex}'))
    for m in marks:
        if m['note_ref'] and m['note_ref'] not in cells:
            cells[m['note_ref']] = ('t', m['note'], None)
        if m['empty']:
            cells.setdefault(m['ref'], ('t', '', None))
    if cells:
        fmt_bytes = set_cells(fmt_bytes, cells)
    if marks:
        fmt_bytes = highlight_cells(fmt_bytes, [m['ref'] for m in marks])
    rep['cbs_notes'] = notes
    rep['subtotals_added'] = sub_log
    rep['internal_formatted'] = (clean_name(iname), fmt_bytes)
    rep['format_log'] = fmt_log
    return data, rep


def cif_history(extra=None):
    """CIF of the internal CBS seen so far: starter set shipped with the app + records saved by the app."""
    import pandas as pd
    import cif as fcif
    p = os.path.join(HERE, 'assets', 'cif_seed.csv')
    h = pd.read_csv(p, dtype=str).pipe(fcif._types) if os.path.exists(p) else pd.DataFrame(columns=fcif.COLS)
    if extra is not None and len(extra):
        h = h[~h['style'].astype(str).isin(set(extra['style'].astype(str)))]
        h = pd.concat([h, extra], ignore_index=True)
    return h


def cif_marks(rep, iname, calc_bytes, extra_hist=None):
    """Internal CIF cells to flag: compared with the CIF other styles used for the same kind of item.
    Sets rep['cif_issues'] / rep['cif_recs']; returns [{ref, note_ref, note, empty}]."""
    import cif as fcif
    from openpyxl.utils import get_column_letter
    lines = rep.get('cif_lines') or []
    style = find_style(clean_name(iname)) or ''
    recs = fcif.records(style, fcif.fstore.season_code(clean_name(iname)) or '', clean_name(iname), lines)
    rep['cif_recs'] = recs
    issues = fcif.check(recs, cif_history(extra_hist))
    rep['cif_issues'] = issues
    if not issues:
        return []
    wv = openpyxl.load_workbook(io.BytesIO(calc_bytes), data_only=True).worksheets[0]
    out = []
    for x in issues:
        for l in lines:
            if (l['detail'] or '').strip() != (x['detail'] or '').strip() or (l['cif'] or 0) != (x['cif'] or 0) \
                    or not l.get('cif_col'):
                continue
            ref = f"{get_column_letter(l['cif_col'])}{l['row']}"
            nref = f"{get_column_letter(l['note_col'])}{l['row']}" if l.get('note_col', 1) > 1 else None
            if nref and wv[nref].value not in (None, ''):
                nref = None                                     # never overwrite a remark
            why = {'missing': 'CIF missing', 'differs': 'Check CIF', 'high': 'CIF high', 'low': 'CIF low'}[x['why']]
            out.append(dict(ref=ref, note_ref=nref, empty=not l['cif'],
                            note=f"⚠ {why}: {x['cif'] or 0:g} → usually {x['usual']:g} (other styles {x['seen']})"))
    rep['warnings'].append(W('cif_bad', n=len(issues)))
    return out


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

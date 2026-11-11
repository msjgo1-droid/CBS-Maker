"""Fabric (marker) sheets from the MNMK pictures pasted on the right side of the internal CBS.

Used when there is no previous submission and no MNMK file. Each picture is a screenshot of an MNMK block
(DATE / STYLE / FAB DETAIL ... table + CAD marker). Numbers are read with OCR (tesseract), the marker itself
is cropped out of the screenshot. Marker LENGTH is not taken from the picture: it is back-calculated from the
buyer (open) yy later, like every other fabric sheet.
"""
import io
import re

from PIL import Image

from xlimages import pictures
from panels import crop_marker


def _s(v):
    return str(v).strip() if v is not None else ''


def _codes(t):
    return {x for x in re.split(r'[\s/,()#:]+', _s(t).upper()) if len(x) >= 5 and re.search(r'\d', x)}


def ocr_text(png):
    """Whole picture text + (Gerber AccuMark) status-bar text read at 4x."""
    try:
        import pytesseract
        import numpy as np
    except ImportError:
        return None
    im = Image.open(io.BytesIO(png)).convert('L')
    try:
        txt = pytesseract.image_to_string(im.resize((im.width * 2, im.height * 2)))
        a = np.asarray(im)
        rows = np.where((a < 60).mean(axis=1) > 0.3)[0]
        rows = rows[rows > a.shape[0] * 0.5]
        if len(rows):
            from PIL import ImageFilter
            band0 = im.crop((0, int(rows.min()), im.width, int(rows.max()) + 1))
            for S in (3, 4, 5, 6):                       # several scales -> vote / try every reading
                band = band0.resize((band0.width * S, band0.height * S), Image.LANCZOS)
                for b in (band, band.filter(ImageFilter.SHARPEN)):
                    txt += '\n@@BAR@@\n' + pytesseract.image_to_string(b, config='--psm 6')
            txt += _ct_boxes(pytesseract, band0)
        return txt
    except Exception:                                        # tesseract binary missing
        return None


def _ct_boxes(pytesseract, band0):
    """Gerber 'CT: 0/84' box (total pieces in the marker): find the label, read only its box, digits only."""
    out = ''
    S = 4
    band = band0.resize((band0.width * S, band0.height * S), Image.LANCZOS)
    d = pytesseract.image_to_data(band, config='--psm 6', output_type=pytesseract.Output.DICT)
    for i, w in enumerate(d['text']):
        if re.fullmatch(r'C[TR][:;.]?', w.strip(), re.I):
            x0 = d['left'][i] + d['width'][i]
            y0, h = d['top'][i], d['height'][i]
            box = band.crop((x0, max(y0 - h // 2, 0), min(x0 + 70 * S, band.width), y0 + h + h // 2))
            from PIL import ImageOps
            inv = ImageOps.autocontrast(ImageOps.invert(box))
            for b in (box, inv, inv.point(lambda p: 255 if p > 128 else 0)):
                for psm in (7, 8):
                    v = pytesseract.image_to_string(b, config=f'--psm {psm} -c tessedit_char_whitelist=0123456789/')
                    out += f'\n@@CT@@ {v.strip()}'
    return out


def parse_fields(text):
    t = text or ''
    g = lambda pat, f=lambda x: x: (lambda m: f(m.group(1)) if m else None)(re.search(pat, t, re.I | re.S))
    out = dict(
        width=g(r'(?:WIDTH|IDTH)\W{0,4}(\d{2,3})\s*["”]?', lambda x: f'{x}"') or
        g(r'\bW\w{3,5}\s*[:=]\s*(\d{2,3})\s*["”]', lambda x: f'{x}"'),
        eff=g(r'Efficiency\s*[:\-]?\s*(\d{2}\.\d{1,2})', float),
        pqty=g(r"Panel\s*Q.?ty\s*(\d{1,4})", int),
        gqty=g(r"GARMENT\s*Q.?TY\s*(\d{1,3})(?!\.)", int),
        yy=g(r'(\d\.\d{3})\s*[\]}|]?\s*(?:Panel|\n)', float),
        inch=g(r'(\d{1,2}\.\d{2})\s*IN\b', float),
        shrink=g(r'L\s*[:;]\s*(\d{1,2})\s*%', int),
    )
    if not out['gqty']:                                      # qty printed just before the red yy
        out['gqty'] = g(r'(?<![\d.])(\d{1,3})\s+\d\.\d{3}', int)
    bars = t.split('@@BAR@@')[1:] if '@@BAR@@' in t else []
    if bars:                                                 # Gerber AccuMark status bar, read several times
        from collections import Counter

        def vote(pat, f):
            c = Counter()
            for bar in bars:
                for m in re.finditer(pat, bar, re.I):
                    try:
                        c[f(m.group(1))] += 1
                    except ValueError:
                        pass
            return c
        w = vote(r'\bW[Ii1l]\s*[:;.]?\s*[\[|({]?\s*(\d{2,3})\.\d{3}', int)
        cu = vote(r'(\d{2}\.\d{2})\s*%', float)
        ln = vote(r'(\d{1,2}\.\d{3})[I1l|\]]', float)
        bd = vote(r'\bB[DP]\W{0,3}(0?\.?\d{3})(?!\d)', lambda x: float(x if '.' in x else '0.' + x.lstrip('0')[-3:]))
        ct = vote(r'\bC[TR]\W{0,3}[0oO]\s*/\s*(\d{2,4})', int)
        if w and not out['width']:
            out['width'] = f'{w.most_common(1)[0][0]}"'
        if cu and not out['eff']:
            out['eff'] = cu.most_common(1)[0][0]
        for m in re.finditer(r'@@CT@@ \d{0,3}/(\d{2,4})\b', t):
            ct[int(m.group(1))] += 2
        if ct and not out['pqty']:
            out['pqty'] = ct.most_common(1)[0][0]
        if bd:
            out['bd'] = out['yy'] = bd.most_common(1)[0][0]
        base_inch = [out['inch']] if out['inch'] is not None else []
        out['inch_cands'] = base_inch + [x for x, _ in ln.most_common() if x >= 0.1 and x not in base_inch]
        if out['inch_cands']:
            out['inch'] = out['inch_cands'][0]
    up = t.upper()
    if 'RANDOM' in up:
        out['direction'] = '2 WAY RANDOM / 1PC 1 WAY'
    elif 'CLOSE' in up:
        out['direction'] = '2 WAY CLOSE / 1PC 1 WAY'
    elif '1 WAY' in up or '1WAY' in up:
        out['direction'] = '1 WAY'
    else:
        out['direction'] = ''
    return out


def marker_solutions(f, yy_hint=None):
    """[(err, qty, Y, inch)] with (Y + IN/36)/qty ≈ yy, for every inch reading."""
    yy = f.get('yy') or yy_hint
    inches = f.get('inch_cands') or ([f['inch']] if f.get('inch') is not None else [])
    if not (yy and inches):
        return []
    qtys = [f['gqty']] if f.get('gqty') else range(1, 121)
    sols = []
    for inch in inches:
        for q in qtys:
            y = round(yy * q - inch / 36)
            if y >= 0:
                sols.append((abs((y + inch / 36) / q - yy) / yy, q, y, inch))
    return [x for x in sols if x[0] < 0.03]


def solve_marker(f, yy_hint=None, ref_qty=None):
    """Garment Q'ty from the marker picture. Prefers the garment count of the other fabric markers (ref_qty),
    then the smallest qty that fits within 1%. Sets gqty / marker_y / inch; returns True / False / None."""
    sols = marker_solutions(f, yy_hint)
    if not sols:
        return None if not (f.get('inch_cands') or f.get('inch') is not None) else False
    pick = None
    if ref_qty:
        same = [x for x in sols if x[1] == ref_qty]
        pick = min(same) if same else None
    if pick is None:
        tight = [x for x in sols if x[0] < 0.01]
        pick = min(tight, key=lambda x: (x[1], x[0])) if tight else min(sols)
    f['gqty'], f['marker_y'], f['inch'] = pick[1], pick[2], pick[3]
    return True


def read_internal_mnmk(internal_bytes, ws_values, lines, gender):
    """Return [{title, block, fields, ok}] — block has the keys panels.build_from_mnmk expects."""
    pics = sorted([p for p in pictures(internal_bytes, 0) if p['col'] >= 16 and p['row'] <= 5],
                  key=lambda p: p['col'])
    if not pics:
        return []
    headers = {c.column: _s(c.value) for c in ws_values[1] if c.column >= 17 and _s(c.value)}
    fabrics = [l for l in lines if l['key'] == 'fabric']
    out, used_titles = [], {}
    for p in pics:
        hcol = max((c for c in headers if c <= p['col'] + 1), default=None)
        head = headers.get(hcol, '')
        ln = next((l for l in fabrics if _codes(head) & (_codes(l['detail']) | _codes(l['remark']))), None)
        if ln is None:
            hn = re.sub(r'[^a-z]', '', head.lower())
            ln = next((l for l in fabrics if re.sub(r'[^a-z]', '', (l['item'] or l['detail']).lower())
                       and re.sub(r'[^a-z]', '', (l['item'] or l['detail']).lower()) in hn), None)
        base = (ln['item'] or ln['detail']) if ln else (head.split('\n')[0][:25] or f'Fabric {len(out) + 1}')
        base = re.sub(r'[\[\]:*?/\\]', ' ', base).strip()[:28]
        n = used_titles.get(base, 0) + 1
        used_titles[base] = n
        title = base if n == 1 else f'{base} ({n})'
        png = p['data'] if p['ext'] in ('png', 'jpg', 'jpeg') else None
        if png is None:
            from xlimages import to_png
            png = to_png(p['data'], p['ext'])
        f = parse_fields(ocr_text(png)) if png else {}
        ok = None
        f['_yy_hint'] = ln.get('iyy') if ln else None
        marker_png = crop_marker(png)[0] if png else None
        size = f"SIZE {'M' if gender == 'M' else 'S'}" + (f" /YY DIRECTION: {f['direction']}" if f.get('direction') else '')
        shrink = f"[SHRINKAGE]\nL : {f.get('shrink') or 5}%\nW : {f.get('shrink') or 5}%"
        block = dict(date=None, width=f.get('width'), style=None, marker_y=f.get('marker_y'), marker_in=f.get('inch'),
                     fab=head.replace('\n', ' ')[:120], eff=f.get('eff'), size=size, shrink=shrink,
                     gqty=f.get('gqty'), pqty=f.get('pqty'), image=marker_png)
        out.append(dict(title=title, block=block, fields=f, ok=ok, header=head,
                        line=(ln['item'] or ln['detail']) if ln else None))
    # garment Q'ty: the marker whose reading is unambiguous sets the reference for the others
    ref = None
    for x in out:
        sols = marker_solutions(x['fields'], x['fields'].get('_yy_hint'))
        if sols and len({s_[1] for s_ in sols if s_[0] < 0.01}) == 1:
            ref = [s_[1] for s_ in sols if s_[0] < 0.01][0]
            break
    for x in out:
        f = x['fields']
        x['ok'] = solve_marker(f, f.get('_yy_hint'), ref) if f else None
        x['block'].update(gqty=f.get('gqty'), marker_y=f.get('marker_y'), marker_in=f.get('inch'),
                          width=f.get('width'), eff=f.get('eff'), pqty=f.get('pqty'))
    return out

"""Fabric (marker) sheets from the MNMK pictures pasted on the right side of the internal CBS.

Used when there is no previous submission and no MNMK file. Each picture is a screenshot of an MNMK block
(DATE / STYLE / FAB DETAIL ... table + CAD marker). Numbers are read with OCR (tesseract), the marker itself
is cropped out of the screenshot. Marker LENGTH is not taken from the picture: it is back-calculated from the
buyer (open) yy later, like every other fabric sheet.
"""
import io
import os
os.environ.setdefault('OMP_THREAD_LIMIT', '1')   # tesseract threads thrash on small cloud CPUs
import re

from PIL import Image

from xlimages import pictures
from panels import crop_marker


def _s(v):
    return str(v).strip() if v is not None else ''


def _codes(t):
    return {x for x in re.split(r'[\s/,()#:]+', _s(t).upper()) if len(x) >= 5 and re.search(r'\d', x)}



def _grid_cells(im, max_rows=6):
    """Cells of the MNMK table (rows between dark horizontal rules, split at dark vertical rules)."""
    import numpy as np
    g = np.asarray(im.convert('L')).astype(int)
    H, W = g.shape
    # dark or light grey table rules: keep the reading that finds more table rows in the upper part
    score = lambda rs: sum(1 for r in rs if len(r) >= 2 and r[0][1] < H * 0.4)
    cands = [_grid_rows(g < 110, max_rows), _grid_rows(g < 200, max_rows)]
    return max(cands, key=score)


def _grid_rows(dark, max_rows):
    H, W = dark.shape
    hl = [y for y in range(H) if dark[y].mean() > 0.45]
    groups = []
    for y in hl:
        if groups and y - groups[-1][-1] <= 2:
            groups[-1].append(y)
        else:
            groups.append([y])
    ys = [(q[0], q[-1]) for q in groups]
    rows = []
    for (a0, a1), (b0, b1) in zip(ys, ys[1:]):
        top, bot = a1 + 1, b0 - 1
        if bot - top < 6:
            continue
        band = dark[top:bot + 1]
        vx = [x for x in range(W) if band[:, x].mean() > 0.8]
        vg = []
        for x in vx:
            if vg and x - vg[-1][-1] <= 2:
                vg[-1].append(x)
            else:
                vg.append([x])
        xs = [(q[0], q[-1]) for q in vg]
        # screenshot cut through the table: the image edge closes the last / first cell
        if xs and xs[-1][1] < W - 20:
            xs.append((W - 1, W - 1))
        if xs and xs[0][0] > 20:
            xs.insert(0, (0, 0))
        row = [(c1 + 1, top, d0 - 1, bot) for (c0, c1), (d0, d1) in zip(xs, xs[1:]) if d0 - c1 > 5]
        if row:
            rows.append(row)
        if len(rows) >= max_rows:
            break
    return rows


def _cell_votes(pytesseract, im, box, pat, wl='0123456789.'):
    """Read one small value cell many ways (channel, scale, blur, resampling, threshold, red text) and count
    the answers that fit the pattern; stops early once one answer clearly wins."""
    import numpy as np
    from collections import Counter
    from PIL import ImageFilter
    c = im.crop(box)
    a = np.asarray(c.convert('RGB')).astype(int)
    red = (a[..., 0] > 140) & (a[..., 1] < 120) & (a[..., 2] < 120)
    bases = [Image.fromarray(a.min(axis=2).astype('uint8')), Image.fromarray(a.mean(axis=2).astype('uint8'))]
    if red.sum() > 3:
        bases.insert(0, Image.fromarray(np.where(red, 0, 255).astype('uint8')))
    bases.append(c.convert('L'))
    cnt = Counter()
    for S in (8, 10, 6):
        for base in bases:
            for rs in (Image.LANCZOS, Image.BICUBIC):
                g = base.resize((base.width * S, base.height * S), rs)
                for v in (g.filter(ImageFilter.GaussianBlur(S / 2)), g.filter(ImageFilter.GaussianBlur(S / 4)), g):
                    for th in (None, 170):
                        w = v.point(lambda p, th=th: 255 if p > th else 0) if th else v
                        pad = Image.new('L', (w.width + 60, w.height + 60), 255)
                        pad.paste(w, (30, 30))
                        t = pytesseract.image_to_string(
                            pad, config=f'--psm 7 --dpi 300 -c tessedit_char_whitelist={wl}')
                        m = re.fullmatch(pat, t.replace(' ', '').strip())
                        if m:
                            cnt[m.group(1) if m.lastindex == 1 else m.groups()] += 1
                top = cnt.most_common(2)
                if top and (top[0][1] >= 12 or (sum(cnt.values()) >= 10 and
                                                (len(top) < 2 or top[0][1] >= 2 * top[1][1]))):
                    return cnt
    return cnt


def _grid_read(pytesseract, png):
    """MNMK screenshot of our panel sheet: read every table cell on its own; the values right of the labels
    (CUTTABLE WIDTH, TOTAL MARKER, Efficiency, GARMENT Q'TY | yy, Panel Q'ty) are read many ways and voted.
    Returns text markers '@@G:key@@ value' (one per vote)."""
    im = Image.open(io.BytesIO(png)).convert('RGB')
    rows = _grid_cells(im)
    out = ''

    def label(box):
        c = im.crop(box).convert('L')
        c = c.resize((c.width * 4, c.height * 4), Image.LANCZOS)
        pad = Image.new('L', (c.width + 40, c.height + 40), 255)
        pad.paste(c, (20, 20))
        return pytesseract.image_to_string(pad, config='--psm 7').strip()

    def emit(key, cnt):
        nonlocal out
        for v, n in cnt.items():
            v = ' '.join(v) if isinstance(v, tuple) else v
            out += f'\n@@G:{key}@@ {v}' * n
    for ri, row in enumerate(rows):
        labs = [label(b) for b in row]
        low = [x.lower() for x in labs]
        for k, t in enumerate(low):
            nxt = row[k + 1] if k + 1 < len(row) else None
            if not nxt:
                # label and value in one cell ('Efficiency      80.29'): the right half holds the value
                x0, y0, x1, y1 = row[k]
                if x1 - x0 < 80:
                    continue
                nxt = (x0 + (x1 - x0) // 2, y0, x1, y1)
            if 'width' in t:
                emit('width', _cell_votes(pytesseract, im, nxt, r'(\d{2,3})'))
            elif re.search(r'ffic|ency|iency|olan|cienc', t):
                emit('eff', _cell_votes(pytesseract, im, nxt, r'(\d{2}\.\d{2})%?'))
            elif re.search(r'panel', t):
                emit('pq', _cell_votes(pytesseract, im, nxt, r'(\d{1,4})'))
            elif re.search(r'garm|q.?ty', t) and 'panel' not in t:
                emit('gq', _cell_votes(pytesseract, im, nxt, r'(\d{1,3})'))
                if k + 2 < len(row):
                    emit('yy', _cell_votes(pytesseract, im, row[k + 2], r'(\d?\.\d{2,4})'))
        # TOTAL MARKER LENGTH: '6 Y 7.45 IN' (one cell, or split in cells)
        for k, t in enumerate(labs):
            if re.search(r'\bIN\b', t) or re.search(r'\d\s*Y\b', t):
                emit('len', _cell_votes(pytesseract, im, row[k], r'(\d{1,2})Y(\d{1,2}\.\d{1,3})(?:IN|I|N)?',
                                        '0123456789.YIN'))
    return out


_OCR_CACHE = {}
PROGRESS = None          # app sets a callback(i, n) to show 'reading MNMK picture i/n'


def ocr_text(png):
    """Whole picture text + Garment/Panel Q'ty row + (Gerber AccuMark) status bar - the slow extra passes
    only run when the fields are still missing. Results cached per picture (re-runs are instant)."""
    import hashlib
    key = hashlib.sha1(png).hexdigest()
    if key in _OCR_CACHE:
        return _OCR_CACHE[key]
    try:
        import pytesseract
        import numpy as np
    except ImportError:
        return None
    im = Image.open(io.BytesIO(png)).convert('L')
    try:
        txt = pytesseract.image_to_string(im.resize((im.width * 2, im.height * 2)))
        txt += _qty_boxes(pytesseract, im)

        def need():
            f = parse_fields(txt)
            return [k for k in ('width', 'eff', 'pqty') if not f.get(k)] + \
                (['eff2'] if f.get('eff') and len(str(f['eff']).split('.')[-1]) < 2 else []) + \
                ([] if f.get('inch') is not None or f.get('yy') else ['len'])
        if need() or not parse_fields(txt).get('gqty'):
            try:
                txt += _grid_read(pytesseract, png)            # table cells one by one, voted
            except Exception:                                  # noqa
                pass
        if [k for k in need() if k in ('width', 'eff', 'pqty')] or not parse_fields(txt).get('gqty'):
            # small screenshot: read the table above the marker on its own, bigger
            try:
                from panels import _crop_by_pieces
                pb = _crop_by_pieces(np.asarray(Image.open(io.BytesIO(png)).convert('RGB')).astype(int), 0)
            except Exception:
                pb = None
            if pb and pb[1] > 40:
                tab = im.crop((0, 0, im.width, int(pb[1])))
                big = tab.resize((tab.width * 4, tab.height * 4), Image.LANCZOS)
                for b in (big, big.point(lambda p: 255 if p > 160 else 0)):
                    txt += '\n@@TAB@@\n' + pytesseract.image_to_string(b, config='--psm 6')
                txt += _qty_boxes(pytesseract, tab, S=5)
        a = np.asarray(im)
        rows = np.where((a < 60).mean(axis=1) > 0.3)[0]
        rows = rows[rows > a.shape[0] * 0.5]
        if len(rows) and need():
            from PIL import ImageFilter
            band0 = im.crop((0, int(rows.min()), im.width, int(rows.max()) + 1))
            for S in (4, 3, 5, 6):                       # several scales -> vote; stop once fields are found
                band = band0.resize((band0.width * S, band0.height * S), Image.LANCZOS)
                for b in (band, band.filter(ImageFilter.SHARPEN)):
                    txt += '\n@@BAR@@\n' + pytesseract.image_to_string(b, config='--psm 6')
                if not need():
                    break
            if 'pqty' in need():
                txt += _ct_boxes(pytesseract, band0)
        _OCR_CACHE[key] = txt
        return txt
    except Exception:                                        # tesseract binary missing
        return None


def _qty_boxes(pytesseract, im, S=3):
    """MNMK table row 'GARMENT Q'TY | 2 | 0.164 | Panel Q'ty | 14': read that row on its own (table lines removed)
    plus the sparse-text tokens on the same row; integer left of the yy = garments, right of it = panels."""
    import numpy as np
    big = im.resize((im.width * S, im.height * S), Image.LANCZOS)
    d = pytesseract.image_to_data(big, config='--psm 11', output_type=pytesseract.Output.DICT)
    words = [(w.strip(), d['left'][i], d['top'][i], d['width'][i], d['height'][i])
             for i, w in enumerate(d['text']) if w.strip()]
    out = ''
    # TOTAL MARKER LENGTH | 3 | Y | 35.05 | IN  -> the yards digit is often lost in the whole-picture read
    for w, x, y, ww, h in words:
        if re.search(r'TOTAL', w, re.I):
            band = big.crop((x, max(0, y - h), big.width, y + 2 * h))
            a = np.asarray(band).copy()
            a[:, (a < 100).mean(axis=0) > 0.6] = 255
            a[(a < 100).mean(axis=1) > 0.5, :] = 255
            for psm in (6, 7):
                t = pytesseract.image_to_string(Image.fromarray(a), config=f'--psm {psm}')
                m = re.search(r'(\d{1,2})\s*[|\]\[]?\s*Y\b\W*(\d{1,2}\.\d{1,3})', t)
                if m:
                    out += f'\n@@LN@@ {m.group(1)} {m.group(2)}'
            break
    for w, x, y, ww, h in words:
        if not re.search(r'GARM', w, re.I):
            continue
        band = big.crop((0, max(0, y - h), big.width, y + 2 * h))
        a = np.asarray(band).copy()
        a[:, (a < 100).mean(axis=0) > 0.6] = 255
        a[(a < 100).mean(axis=1) > 0.5, :] = 255
        for psm in (6, 7):
            t = pytesseract.image_to_string(Image.fromarray(a), config=f'--psm {psm}')
            m = re.search(r"GARMENT\s*Q.?TY\W*(\d{1,3})\W+(\d?\.\d{2,4})\W*Panel\s*Q.?ty\W*(\d{1,4})", t, re.I)
            if m:
                out += f'\n@@GQ@@ {m.group(1)}\n@@GQ@@ {m.group(1)}\n@@PQ@@ {m.group(3)}\n@@PQ@@ {m.group(3)}'
        # same-row tokens from the sparse read
        row = [(w2, x2) for w2, x2, y2, ww2, h2 in words if abs((y2 + h2 / 2) - (y + h / 2)) < h and x2 > x]
        dec = [x2 for w2, x2 in row if re.fullmatch(r'\(?\d?\.\d{2,4}\)?', w2)]
        if dec:
            for w2, x2 in row:
                if re.fullmatch(r'\d{1,4}', w2):
                    out += f'\n@@{"GQ" if x2 < dec[0] else "PQ"}@@ {w2}'
        break
    return out


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


def _ln_vote(t):
    from collections import Counter
    c = Counter(int(y) for y, _ in re.findall(r'@@LN@@ (\d{1,2}) (\d{1,2}\.\d{1,3})', t or ''))
    return c.most_common(1)[0][0] if c else None


def _vote(t, lab):
    from collections import Counter
    c = Counter(int(v) for v in re.findall(r'@@%s@@ (\d{1,4})' % lab, t or '') if int(v) > 0)
    return c.most_common(1)[0][0] if c else None


def _gvotes(t, key):
    from collections import Counter
    return Counter(re.findall(r'@@G:%s@@ ([^\n]+)' % key, t or ''))


def parse_fields(text):
    t = text or ''
    g = lambda pat, f=lambda x: x: (lambda m: f(m.group(1)) if m else None)(re.search(pat, t, re.I | re.S))
    out = dict(
        width=g(r'(?:WIDTH|IDTH)\W{0,4}(\d{2,3})(?![\d.])\s*["”]?', lambda x: f'{x}"') or
        g(r'\bW\w{3,5}\s*[:=]\s*(\d{2,3})\s*["”]', lambda x: f'{x}"'),
        eff=g(r'Effic\w*\s*[:\-]?\s*(\d{2}\.\d{1,2})', float) or
            g(r'(?m)FAB DETAIL[^\n]*?\s([5-9]\d\.\d{2})\s*$', float),
        pqty=_vote(t, 'PQ') or g(r"Panel\s*Q.?ty\s*(\d{1,4})", int),
        gqty=_vote(t, 'GQ') or g(r"GARMENT\s*Q.?TY[\s\[|(]{0,4}(\d{1,3})(?![\d.])", int),
        yy=g(r'(\d\.\d{3})\s*[\]}|]?\s*(?:Panel|\n)', float),
        inch=g(r'(\d{1,2}\.\d{2})\s*IN\b', float),
        marker_y_read=_ln_vote(t),
        shrink=g(r'L\s*[:;]\s*(\d{1,2})\s*%', int),
    )
    if not out['gqty']:                                      # qty printed just before the red yy
        # only on the GARMENT Q'TY / Panel Q'ty row - elsewhere 'n 0.136' is anything (dates, %, sizes)
        qrow = '\n'.join(l for l in t.splitlines() if re.search(r'GARM|Panel', l, re.I))
        mq = re.search(r'(?<![\d.])(\d{1,3})\s+\d\.\d{3}', qrow)
        out['gqty'] = int(mq.group(1)) if mq else None
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
    # impossible readings count as not read (17.37% efficiency = '7' read as '1')
    if out['eff'] is not None and not 35 <= out['eff'] <= 99.5:
        out['eff'] = None
    if out['pqty'] and out['gqty'] and (out['pqty'] < out['gqty'] or out['pqty'] % out['gqty']):
        out['pqty'] = None
    # table cells read one by one (grid) - fills what the whole-picture read did not find
    from collections import Counter
    gw, ge, gg, gp, gy, gl = (_gvotes(t, k) for k in ('width', 'eff', 'gq', 'pq', 'yy', 'len'))
    # cuttable width is 30"..200" and must be read more than once ('54"' in red read as 04 / 34)
    gw = Counter({k: v for k, v in gw.items() if k.isdigit() and 30 <= int(k) <= 200 and v >= 2})
    if out['width'] and not 30 <= int(out['width'].rstrip('"')) <= 200:
        out['width'] = None
    if gw and not out['width']:
        out['width'] = f'{gw.most_common(1)[0][0]}"'
    # efficiency: every plausible 'Efficiency nn.nn' in the text reads (weight 5) + the cell votes
    from collections import Counter
    ce = Counter()
    for v in re.findall(r'Effic\w*\s*[:\-]?\s*(\d{2}\.\d{2})', t):
        if 35 <= float(v) <= 99.5:
            ce[v] += 5
    for k, v in ge.items():
        if 35 <= float(k) <= 99.5:
            ce[k] += v
    if ce and (max(ce.values()) >= 2):
        top = max(ce, key=ce.get)
        if not out['eff']:
            out['eff'] = float(top)
        elif len(str(out['eff']).split('.')[-1]) < 2:
            # 80.2 (last digit lost) -> the 80.2x reading with most votes
            same = {k: v for k, v in ce.items() if k.startswith(f"{out['eff']:.1f}")}
            if same:
                out['eff'] = float(max(same, key=same.get))
    if gy and not out['yy']:
        out['yy'] = float(gy.most_common(1)[0][0])
    if gp:
        out['pqty_cands'] = ([out['pqty']] if out['pqty'] else []) + \
            [int(v) for v, _ in gp.most_common() if int(v) != out['pqty']]
        out['pqty'] = out['pqty'] or out['pqty_cands'][0]
    if gg:
        out['gqty_cands'] = ([out['gqty']] if out['gqty'] else []) + \
            [int(v) for v, _ in gg.most_common() if int(v) > 0 and int(v) != out['gqty']]
        if not out['gqty'] and gg.most_common(1)[0][1] >= 3:
            out['gqty'] = int(gg.most_common(1)[0][0])
    lens = None
    if gl:
        lens = [tuple(v.split()) for v, _ in gl.most_common()]
        lens = [(y, i) for y, i in lens if int(y) <= 40 and 0 <= float(i) < 36] or None
    if lens:
        base = list(out.get('inch_cands') or ([out['inch']] if out['inch'] is not None else []))
        out['inch_cands'] = base + [float(i) for _, i in lens if float(i) not in base]
        if out['inch'] is None:
            out['inch'] = out['inch_cands'][0]
            out['yards_cands'] = sorted({int(y) for y, _ in lens})
            if out['marker_y_read'] is None and len(out['yards_cands']) == 1:
                out['marker_y_read'] = out['yards_cands'][0]
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
    sols = [x for x in sols if x[0] < 0.03]
    if f.get('yards_cands') and any(x[2] in f['yards_cands'] for x in sols):
        sols = [x for x in sols if x[2] in f['yards_cands']]       # yards digit read (maybe several readings)
    if f.get('gqty_cands') and not f.get('gqty') and any(x[1] in f['gqty_cands'] for x in sols):
        sols = [x for x in sols if x[1] in f['gqty_cands']]
    if f.get('marker_y_read') is not None and any(x[2] == f['marker_y_read'] for x in sols):
        sols = [x for x in sols if x[2] == f['marker_y_read']]      # yards read in TOTAL MARKER row
    return sols


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


def _area(x):
    """Marker picture size (pixels of the cropped marker) - fallback measure of how big a marker is."""
    try:
        im = Image.open(io.BytesIO(crop_marker(x['block']['image'])[0]))
        return im.width * im.height
    except Exception:
        return 0


def _is_marker_pic(p):
    """MNMK picture (marker layout of coloured pieces) - not a fabric spec sheet / text screenshot."""
    import numpy as np
    from panels import _colours
    data = p['data'] if p['ext'] in ('png', 'jpg', 'jpeg') else None
    if data is None:
        try:
            from xlimages import to_png
            data = to_png(p['data'], p['ext'])
        except Exception:
            return True
    if not data:
        return True
    p['_png'] = data
    cb, full, cropped = crop_marker(data)
    if cropped == full:
        return False
    a = np.asarray(Image.open(io.BytesIO(cb)).convert('RGB')).astype(int)
    return _colours(a, (0, 0, a.shape[1] - 1, a.shape[0] - 1)) >= 6


def _yy_fits(l, y):
    vals = [l.get('iyy')] + list(l.get('parts') or [])
    return any(v and abs(v - y) <= 0.01 * v + 0.0015 for v in vals)


def read_internal_mnmk(internal_bytes, ws_values, lines, gender, sheet_index=None, other_tabs=True):
    """Return [{title, block, fields, ok, tab}] — block has the keys panels.build_from_mnmk expects.
    MNMK pictures sit right of the costing, top rows. When the current tab has none, the newest
    other tab that has them is used (rep shows which tab)."""
    wb = ws_values.parent
    if sheet_index is None:
        cur = wb.worksheets.index(ws_values)
        order = [cur] + ([i for i in range(len(wb.worksheets)) if i != cur] if other_tabs else [])
        for i in order:
            if any(p['col'] >= 16 and p['row'] <= 60 and _is_marker_pic(p) for p in pictures(internal_bytes, i)):
                res = read_internal_mnmk(internal_bytes, wb.worksheets[i], lines, gender, sheet_index=i)
                for x in res:
                    x['tab'] = wb.worksheets[i].title if i != cur else None
                return res
        return []
    pics = sorted([p for p in pictures(internal_bytes, sheet_index) if p['col'] >= 16 and p['row'] <= 60
                   and _is_marker_pic(p)], key=lambda p: (p['col'], p['row']))
    if not pics:
        return []
    headers = {c.column: _s(c.value) for c in ws_values[1] if c.column >= 17 and _s(c.value)}
    fabrics = [l for l in lines if l['key'] == 'fabric']
    out, used_titles = [], {}
    for n_, p in enumerate(pics):
        if PROGRESS:
            PROGRESS(n_, len(pics))
        hcol = max((c for c in headers if c <= p['col'] + 1), default=None) or \
            min((c for c in headers if c <= p['col'] + 4), default=None)
        head = headers.get(hcol, '')
        ln = next((l for l in fabrics if _codes(head) & (_codes(l['detail']) | _codes(l['remark']))), None)
        if ln is None and len(pics) == 1 and len(fabrics) == 1:
            ln = fabrics[0]
        if ln is None:
            hn = re.sub(r'[^a-z]', '', head.lower())
            ln = next((l for l in fabrics if re.sub(r'[^a-z]', '', (l['item'] or l['detail']).lower())
                       and re.sub(r'[^a-z]', '', (l['item'] or l['detail']).lower()) in hn), None)
        png = p.get('_png') or (p['data'] if p['ext'] in ('png', 'jpg', 'jpeg') else None)
        if png is None:
            from xlimages import to_png
            png = to_png(p['data'], p['ext'])
        _t = ocr_text(png) if png else None
        f = parse_fields(_t) if png else {}
        if _t:
            f['_binding'] = bool(re.search(r'BINDING', _t.split('@@')[0], re.I))
        if ln is None and f.get('yy'):
            # picture YY (e.g. 0.776) = internal YY of one fabric line -> that line (Self), not the fabric code
            near = [l for l in fabrics if l.get('iyy') and abs(l['iyy'] - f['yy']) <= 0.01 * l['iyy'] + 0.0015]
            near += [l for l in fabrics if any(abs(pp - f['yy']) <= 0.01 * pp + 0.0015 for pp in (l.get('parts') or []))
                     and l not in near]
            ln = near[0] if near else None
        if ln is None and len(fabrics) == 1:
            ln = fabrics[0]
        if ln is not None and f.get('yy') and not _yy_fits(ln, f['yy']):
            # picture placed under another fabric's header (Self + CF facing pasted side by side in col R):
            # its own YY (0.034) names the fabric line it belongs to
            alt = [l for l in fabrics if l is not ln and _yy_fits(l, f['yy'])]
            if alt:
                ln = alt[0]
                hcol = ('line', fabrics.index(ln))
                lc = _codes(ln['detail']) | _codes(ln['remark'])
                head = headers.get(next((c for c in headers if any(
                    a == b or (min(len(a), len(b)) >= 6 and (a.startswith(b) or b.startswith(a)))
                    for a in _codes(headers[c]) for b in lc)), None), head)
        base = (ln['item'] or ln['detail']) if ln else (head.split('\n')[0][:25] or f'Fabric {len(out) + 1}')
        base = re.sub(r'\s+', ' ', re.sub(r'[\[\]:*?/\\]', ' ', base)).strip()[:28]
        n = used_titles.get(base, 0) + 1
        used_titles[base] = n
        title = base if n == 1 else f'{base} ({n})'
        if not f.get('width'):
            # tiny CAD status bar unreadable -> cuttable width written in the fabric header ('60" CW')
            mw = re.search(r'CUTTABLE\W*\d{2,3}(?:\.\d)?\s*CM\W{0,3}(\d{2,3})', head.upper()) or \
                re.search(r'(\d{2,3})\s*(?:"|”|\'\'|INCH|IN)?\s*C\.?W\b', head.upper()) or \
                re.search(r'\bC\.?W\.?\s*:?\s*(\d{2,3})\s*(?:"|”)', head.upper()) or \
                re.search(r'(\d{2,3})\s*C\.?W\b', ((ln or {}).get('remark') or '') + ' ' + ((ln or {}).get('detail') or ''), re.I)
            if mw:
                f['width'] = f'{mw.group(1)}"'
        ok = None
        f['_yy_hint'] = ln.get('iyy') if ln else None
        marker_png = crop_marker(png)[0] if png else None
        size = f"SIZE {'M' if gender == 'M' else 'S'}" + (f" /YY DIRECTION: {f['direction']}" if f.get('direction') else '')
        shrink = f"[SHRINKAGE]\nL : {f.get('shrink') or 5}%\nW : {f.get('shrink') or 5}%"
        block = dict(date=None, width=f.get('width'), style=None, marker_y=f.get('marker_y'), marker_in=f.get('inch'),
                     fab=head.replace('\n', ' ')[:120], eff=f.get('eff'), size=size, shrink=shrink,
                     gqty=f.get('gqty'), pqty=f.get('pqty'), image=marker_png)
        out.append(dict(title=title, block=block, fields=f, ok=ok, header=head, _hcol=hcol, _line=ln, _row=p['row'],
                        line=(ln['item'] or ln['detail']) if ln else None))
    # several MNMK under one fabric header = same fabric cut for different parts (e.g. 0.459 + 1.458 = 1.917)
    groups = {}
    for x in out:
        groups.setdefault(x['_hcol'], []).append(x)
    for xs in groups.values():
        if len(xs) < 2:
            continue
        ln = xs[0]['_line']
        xs.sort(key=lambda x: x['_row'])
        parts = list((ln or {}).get('parts') or [])
        ys = [x['fields'].get('yy') for x in xs]
        iyy = (ln or {}).get('iyy')
        sums = iyy and all(ys) and abs(sum(ys) - iyy) <= 0.03 * iyy + 0.002
        if not (len(parts) == len(xs) or sums):
            # not a split of one fabric (e.g. solid PK-42/03 0.166 and heather RF-5901 0.134 stacked):
            # keep the top picture only
            for x in xs[1:]:
                x['_drop'] = True
            continue
        # a yy read from the picture claims its part; pictures without one take what is left
        left = list(parts)
        for x in xs:
            py = x['fields'].get('yy')
            if py and left:
                near = min(left, key=lambda v: abs(v - py))
                if abs(near - py) <= 0.01 + 0.05 * near:
                    left.remove(near)
            x['part_yy'] = py
        rest = [x for x in xs if not x['part_yy']]
        if rest and len(left) == len(rest):
            # bigger marker (length per garment) = bigger part
            def size(x):
                f = x['fields']
                if f.get('inch') is not None and f.get('gqty'):
                    return ((f.get('marker_y') or 0) * 36 + f['inch']) / f['gqty'] / 36
                return _area(x) / 1e6

            for x, v in zip(sorted(rest, key=size, reverse=True), sorted(left, reverse=True)):
                x['part_yy'] = v
        for x in xs:
            x['fields']['_yy_hint'] = x['part_yy']
        scale = (ln['yy'] / ln['iyy']) if ln and ln.get('iyy') and ln.get('yy') else 1
        xs.sort(key=lambda x: -(x.get('part_yy') or 0))       # biggest part = main sheet
        base = xs[0]['title'].split(' (')[0]
        for k, x in enumerate(xs):
            x['title'] = base if k == 0 else (f'{base} (BINDING)' if x['fields'].get('_binding') and
                                              not xs[0]['fields'].get('_binding') else f'{base} ({k + 1})')
            x['scale'] = scale
            x['part_main'] = k == 0
    order = []
    for xs in groups.values():
        order += [x for x in xs if not x.get('_drop')]
    # sheets in the internal's fabric order (Self first), whatever column the picture was pasted in
    order.sort(key=lambda x: fabrics.index(x['_line']) if x.get('_line') in fabrics else len(fabrics))
    out[:] = order
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
        if f and f.get('gqty') and f.get('pqty_cands'):
            # panels = garments x pieces per garment -> the reading divisible by Garment Q'ty (176 not 178)
            ok = [p for p in f['pqty_cands'] if p % f['gqty'] == 0 and p >= f['gqty']]
            f['pqty'] = ok[0] if ok else None
        if f and f.get('gqty') and f.get('pqty') and (f['pqty'] % f['gqty'] or f['pqty'] < f['gqty']):
            f['pqty'] = None            # can't be right - left empty for the user (warning)
        x['block'].update(gqty=f.get('gqty'), marker_y=f.get('marker_y'), marker_in=f.get('inch'),
                          width=f.get('width'), eff=f.get('eff'), pqty=f.get('pqty'))
    return out

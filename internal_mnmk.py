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
    try:
        import pytesseract
    except ImportError:
        return None
    im = Image.open(io.BytesIO(png)).convert('L')
    im = im.resize((im.width * 2, im.height * 2))
    try:
        return pytesseract.image_to_string(im)
    except Exception:                                        # tesseract binary missing
        return None


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


def solve_marker(f):
    """Find whole Y (and qty if missing) so that (Y + IN/36)/qty ≈ yy. Sets f['marker_y'], returns ok flag."""
    if not (f.get('yy') and f.get('inch') is not None):
        return None
    qtys = [f['gqty']] if f.get('gqty') else range(1, 81)
    best = None
    for q in qtys:
        y = round(f['yy'] * q - f['inch'] / 36)
        if y < 0:
            continue
        err = abs((y + f['inch'] / 36) / q - f['yy']) / f['yy']
        if best is None or err < best[0]:
            best = (err, q, y)
    if best and best[0] < 0.03:
        f['gqty'], f['marker_y'] = best[1], best[2]
        return True
    return False


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
        ok = solve_marker(f) if f else None
        marker_png = crop_marker(png)[0] if png else None
        size = f"SIZE {'M' if gender == 'M' else 'S'}" + (f" /YY DIRECTION: {f['direction']}" if f.get('direction') else '')
        shrink = f"[SHRINKAGE]\nL : {f.get('shrink') or 5}%\nW : {f.get('shrink') or 5}%"
        block = dict(date=None, width=f.get('width'), style=None, marker_y=f.get('marker_y'), marker_in=f.get('inch'),
                     fab=head.replace('\n', ' ')[:120], eff=f.get('eff'), size=size, shrink=shrink,
                     gqty=f.get('gqty'), pqty=f.get('pqty'), image=marker_png)
        out.append(dict(title=title, block=block, fields=f, ok=ok, header=head,
                        line=(ln['item'] or ln['detail']) if ln else None))
    return out

"""CIF history: the CIF typed in every internal CBS, per item type, and a check of new CBS against it.

CIF depends on what the item is (fabric per yard, zipper, label, poly bag ...), so every internal line is
classified into an item type and compared with what was typed for that type in the other styles.

Storage: same as the FOB history (store.Store), file 'cif_history.csv' in the same data repo / local folder.
"""
import datetime as dt
import re

import pandas as pd

import store as fstore

COLS = ['season', 'style', 'tab', 'section', 'type', 'item', 'detail', 'price', 'cif', 'file', 'saved_at']
KEY = ['style', 'tab', 'type', 'detail']

# (item type, regex on 'item + detail'), first match wins - order matters (puller before zipper ...)
TYPES = [
    ('interlining', r'interlin|fusing|fusible'),
    ('adhesive / bonding', r'adhesive|bonding|bonded|sew\s*free'),
    ('zipper puller', r'pull(er)?\b|zip(per)?\s*pull'),
    ('zipper', r'zip'),
    ('seal tape', r'seal\s*tape'),
    ('elastic', r'elastic'),
    ('drawcord', r'draw\s*cord|drawstring|\bcord\b(?!.*lock)'),
    ('cord lock / toggle', r'cord\s*lock|toggle|stopper|\bend\b|aglet|tip'),
    ('snap / button', r'snap|button|eyelet|grommet|rivet'),
    ('piping / binding / tape', r'piping|binding|tape|webbing|loop|bias'),
    ('logo / print', r'logo|heat\s*transfer|\bht\b|print|reflect|emb|patch|silicon'),
    ('thread', r'thread'),
    ('size / care label', r'care|size|coo'),
    ('label', r'label'),
    ('hangtag', r'hang\s*tag|hangtag|\btag\b'),
    ('sticker', r'sticker|upc'),
    ('poly bag', r'poly\s*bag|polybag'),
    ('carton', r'carton|box'),
    ('tissue', r'tissue|paper'),
    ('anti-mildew', r'mildew|silica|desiccant'),
    ('pin / string', r'\bpin\b|string|cord|clip|plastic\s*fastener|tag\s*pin'),
]


PACK = {'seal tape', 'hangtag', 'sticker', 'poly bag', 'carton', 'tissue', 'anti-mildew', 'pin / string'}


def item_type(line):
    """Item type of an internal line (fabric rows = everything above CM; packing items only in Packing)."""
    txt = f"{line.get('item') or ''} {line.get('detail') or ''}".lower()
    key = line.get('key')
    if key == 'fabric':
        return 'interlining' if re.search(TYPES[0][1], txt) else 'fabric'
    for name, pat in TYPES[1:]:
        if (name in PACK) != (key == 'pack'):
            continue
        if re.search(pat, txt):
            return name
    return {'trim': 'other trim', 'emb': 'logo / print', 'thread': 'thread', 'label': 'label',
            'pack': 'other packing'}.get(key, 'other')


def records(style, season, file, cif_lines):
    """Rows to save from engine rep['cif_lines'] (solid/heather blocks are the same items -> once)."""
    now = dt.datetime.now().strftime('%Y-%m-%d %H:%M')
    out, seen = [], set()
    for l in cif_lines or []:
        kk = (l.get('tab'), (l.get('item') or '').strip().lower(), (l.get('detail') or '').strip().lower(), l.get('price'))
        if kk in seen:
            continue
        seen.add(kk)
        out.append(dict(season=season or '', style=str(style), tab=l.get('tab') or '', section=l.get('key'),
                        type=item_type(l),
                        item=(l.get('item') or '').strip(), detail=(l.get('detail') or '').strip(),
                        price=l.get('price'), cif=l.get('cif'), file=file or '', saved_at=now))
    return out


def name_key(detail):
    """'Elastic 1" (K&K)' -> 'elastic', 'Drawcord 9MM' -> 'drawcord', 'Adhesive Tape Bonding 3913' -> 'adhesive tape bonding'."""
    t = re.sub(r'\(.*?\)|@.*$', ' ', str(detail or '').lower())
    t = re.sub(r'[\d.,/"\'#+-]+\s*(mm|cm|in|inch|yds?)?\b', ' ', t)
    return re.sub(r'[^a-z ]+', ' ', t).split() and ' '.join(re.sub(r'[^a-z ]+', ' ', t).split()) or ''


def _common(vals):
    c = vals.round(4).value_counts()
    return ', '.join(f'{v:g}×{n}' for v, n in c.head(3).items())


def check(new, hist, min_n=3):
    """Compare each line's CIF with the other styles: first the same item (same name), else the same item type.
    Returns [dict(type, item, detail, cif, usual, seen, n, why)] - why: missing / differs / high / low."""
    if hist is None or hist.empty or not new:
        return []
    style = str(new[0]['style'])
    h = hist[(hist['style'].astype(str) != style)].copy()
    h['cif'] = pd.to_numeric(h['cif'], errors='coerce').fillna(0)
    h['_nk'] = h['detail'].map(name_key)
    out = []
    for r in new:
        cif = float(r['cif'] or 0)
        nk = name_key(r['detail'])
        same = h[(h['_nk'] == nk) & (h['type'] == r['type'])] if nk and r['type'] not in ('fabric',) else h.iloc[0:0]
        price = float(r.get('price') or 0)
        if same['style'].nunique() < 2 and price > 0:
            # same kind of item at a similar price (a $10/yd bonding roll vs a $0.05 bar tack)
            hp = pd.to_numeric(h['price'], errors='coerce').fillna(0)
            same = h[(h['type'] == r['type']) & (hp >= price / 2) & (hp <= price * 2)]
            if same['style'].nunique() < 3:
                same = h.iloc[0:0]
        why, base = None, None
        if same['style'].nunique() >= 2:
            base = same['cif']
            top = base.round(4).value_counts()
            usual, share = float(top.index[0]), top.iloc[0] / len(base)
            if cif <= 0 < usual and share >= 0.6:
                why = 'missing'
            elif cif > 0 and share >= 0.6 and abs(cif - usual) > 1e-9 and top.get(round(cif, 4), 0) / len(base) <= 0.3:
                why = 'differs'
        else:
            base = h[h['type'] == r['type']]['cif']
            if len(base) < min_n:
                continue
            filled = base[base > 0]
            usual = float(filled.round(4).mode().iloc[0]) if len(filled) else 0.0
            if cif <= 0 and len(filled) / len(base) >= 0.7:
                why = 'missing'
            elif cif > 0 and len(filled) >= min_n:
                if cif > filled.max() * 1.5 + 1e-9:
                    why = 'high'
                elif cif < filled.min() / 1.5 - 1e-9:
                    why = 'low'
        if why:
            out.append(dict(type=r['type'], item=r['item'], detail=r['detail'], price=r.get('price'), cif=r['cif'],
                            usual=usual, seen=_common(base[base > 0] if why != 'missing' else base), n=int(len(base)),
                            why=why))
    return out


class CifStore(fstore.Store):
    """Same backend as the FOB history, another CSV."""

    def __init__(self, secrets=None):
        super().__init__(secrets)
        self.path = 'cif_history.csv'
        self.local = fstore.LOCAL.replace('fob_history.csv', 'cif_history.csv')

    def load(self):
        import io as _io
        import os
        if self.kind == 'local':
            if os.path.exists(self.local):
                return pd.read_csv(self.local, dtype=str).pipe(_types)
            return pd.DataFrame(columns=COLS)
        import base64
        import requests
        r = requests.get(self._url(), headers=self._hdr(), params={'ref': self.branch}, timeout=30)
        if r.status_code == 404:
            self._sha = None
            return pd.DataFrame(columns=COLS)
        r.raise_for_status()
        j = r.json()
        self._sha = j['sha']
        raw = base64.b64decode(j['content']).decode('utf-8-sig')
        return pd.read_csv(_io.StringIO(raw), dtype=str).pipe(_types)

    def save(self, df, message='update CIF history'):
        import os
        csv = df[COLS].to_csv(index=False)
        if self.kind == 'local':
            os.makedirs(os.path.dirname(self.local), exist_ok=True)
            open(self.local, 'w', encoding='utf-8-sig').write(csv)
            return
        import base64
        import requests
        body = {'message': message, 'branch': self.branch,
                'content': base64.b64encode(csv.encode('utf-8-sig')).decode()}
        if self._sha:
            body['sha'] = self._sha
        r = requests.put(self._url(), headers=self._hdr(), json=body, timeout=30)
        r.raise_for_status()

    def upsert(self, recs):
        """Same style + tab + item type + detail -> newest wins."""
        cur = self.load()
        new = pd.DataFrame(recs, columns=COLS)
        allr = pd.concat([cur, new], ignore_index=True) if not cur.empty else new
        allr['_k'] = allr[KEY].astype(str).agg('|'.join, axis=1)
        allr = allr.drop_duplicates('_k', keep='last').drop(columns='_k')
        self.save(allr)
        return len(new)


def _types(df):
    for c in COLS:
        if c not in df.columns:
            df[c] = None
    for c in ('price', 'cif'):
        df[c] = pd.to_numeric(df[c], errors='coerce')
    return df[COLS]


def from_internal(data, filename=''):
    """(style, season, tab, cif_lines) of an internal CBS file (current tab, every block) - for importing
    old internal CBS into the history without building a buyer CBS."""
    import io
    from internal import read_internal
    from cbs_fill import read_internal_block
    info, wf, wv = read_internal(io.BytesIO(data), filename)
    wsv = wv.worksheets[0]
    lines = []
    for b in info['blocks']:
        blk = read_internal_block(wsv, b['r0'], b['r1'], wf.worksheets[0])
        lines += [dict(block='heather' if b['heather'] else 'solid', tab=info['tab'], row=l['row'], key=l['key'],
                       cat=l['cat'], item=l['item'], detail=l['detail'], price=l['price'], cif=l['cif'], yy=l['yy'])
                  for l in blk['lines'] if not l['lump'] and l['yy']]
    season = fstore.season_code(f"{info['tab']} {filename}")
    return info['style'], season, info['tab'], lines

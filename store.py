"""FOB history: what went to the buyer, per season / stage / style / colour.

Storage
- GitHub (recommended on Streamlit Cloud): a CSV in a *separate private data repo*, written through the
  GitHub API. Set in Streamlit secrets:
      [fob_store]
      github_token = "ghp_..."            # fine-grained token, Contents: read & write on the data repo
      github_repo  = "owner/hoka-cbs-data"
      path         = "fob_history.csv"    # optional
      branch       = "main"               # optional
  (Use a different repo from the app repo, otherwise every save redeploys the app.)
- Local CSV (data/fob_history.csv) when no secrets are set — fine on your own PC, but on Streamlit Cloud
  it is wiped whenever the app restarts.
"""
import base64
import datetime as dt
import io
import os
import re

import openpyxl
import pandas as pd

COLS = ['season', 'stage', 'style', 'color', 'fob', 'date', 'file', 'saved_at']
KEY = ['season', 'stage', 'style', 'color']
HERE = os.path.dirname(os.path.abspath(__file__))
LOCAL = os.path.join(HERE, 'data', 'fob_history.csv')


# ---------------- backend ----------------
class Store:
    def __init__(self, secrets=None):
        cfg = {}
        try:
            cfg = dict(secrets.get('fob_store', {})) if secrets is not None else {}
        except Exception:                                    # noqa
            cfg = {}
        self.token = cfg.get('github_token')
        self.repo = cfg.get('github_repo')
        self.path = cfg.get('path', 'fob_history.csv')
        self.branch = cfg.get('branch', 'main')
        self.kind = 'github' if self.token and self.repo else 'local'
        self._sha = None

    # --- github
    def _url(self):
        return f'https://api.github.com/repos/{self.repo}/contents/{self.path}'

    def _hdr(self):
        return {'Authorization': f'Bearer {self.token}', 'Accept': 'application/vnd.github+json'}

    def load(self):
        if self.kind == 'local':
            if os.path.exists(LOCAL):
                return pd.read_csv(LOCAL, dtype=str).pipe(_types)
            return pd.DataFrame(columns=COLS)
        import requests
        r = requests.get(self._url(), headers=self._hdr(), params={'ref': self.branch}, timeout=30)
        if r.status_code == 404:
            self._sha = None
            return pd.DataFrame(columns=COLS)
        r.raise_for_status()
        j = r.json()
        self._sha = j['sha']
        raw = base64.b64decode(j['content']).decode('utf-8-sig')
        return pd.read_csv(io.StringIO(raw), dtype=str).pipe(_types)

    def save(self, df, message='update FOB history'):
        csv = df[COLS].to_csv(index=False)
        if self.kind == 'local':
            os.makedirs(os.path.dirname(LOCAL), exist_ok=True)
            open(LOCAL, 'w', encoding='utf-8-sig').write(csv)
            return
        import requests
        body = {'message': message, 'branch': self.branch,
                'content': base64.b64encode(csv.encode('utf-8-sig')).decode()}
        if self._sha:
            body['sha'] = self._sha
        r = requests.put(self._url(), headers=self._hdr(), json=body, timeout=30)
        r.raise_for_status()
        self._sha = r.json()['content']['sha']

    def upsert(self, records):
        """Add/replace rows (same season+stage+style+colour -> newest wins). Returns (added, replaced)."""
        for attempt in range(2):
            df = self.load()
            new = pd.DataFrame(records)[COLS[:-1]]
            new = new.sort_values('date').drop_duplicates(KEY, keep='last')   # same key in one batch: newest
            new['saved_at'] = dt.datetime.now().strftime('%Y-%m-%d %H:%M')
            keys = set(map(tuple, new[KEY].astype(str).values))
            old_keys = set(map(tuple, df[KEY].astype(str).values)) if len(df) else set()
            replaced = len(keys & old_keys)
            if len(df):
                df = df[[tuple(x) not in keys for x in df[KEY].astype(str).values]]
            out = pd.concat([df, new], ignore_index=True).pipe(_types)
            try:
                self.save(out, f'FOB history: {len(new)} rows')
                return len(new) - replaced, replaced
            except Exception:                                # sha conflict -> reload once
                if attempt:
                    raise
        return 0, 0


def _types(df):
    for c in COLS:
        if c not in df:
            df[c] = None
    df['fob'] = pd.to_numeric(df['fob'], errors='coerce')
    df['style'] = df['style'].astype(str)
    return df[COLS]


# ---------------- reading buyer CBS files ----------------
def _s(v):
    return str(v).strip() if v is not None else ''


def season_code(text):
    t = _s(text).upper()
    m = re.search(r'(?<![A-Z])([SF])(\d{2})(?!\d)', t)
    if m:
        return m.group(1) + m.group(2)
    m = re.search(r'(SPRING|FALL)\s*20(\d{2})', t)
    if m:
        return ('S' if m.group(1) == 'SPRING' else 'F') + m.group(2)
    return ''


def parse_name(name):
    """1187308_M_HRCTrackJacket_S28_A086PIVNVN_GTM2_11092026.xlsx -> style, season, stage, date."""
    base = re.sub(r'^[0-9a-f]{8}-', '', os.path.splitext(os.path.basename(name))[0])
    parts = base.split('_')
    out = dict(style='', season='', stage='', date='')
    m = re.search(r'(?<!\d)(\d{7})(?!\d)', base)
    out['style'] = m.group(1) if m else ''
    for p in parts:
        if not out['season']:
            out['season'] = season_code(p) if re.match(r'^[SF]\d{2}', p.upper()) else ''
    if len(parts) >= 3 and re.fullmatch(r'\d{8}', parts[-1]):
        d = parts[-1]
        try:
            out['date'] = dt.date(int(d[4:]), int(d[:2]), int(d[2:4])).isoformat()
        except ValueError:
            pass
        out['stage'] = parts[-2].upper()
    return out


def color_of(title, ws=None):
    t = title.lower()
    if 'heather' in t and 'solid' not in t:
        return 'HEATHER'
    return 'SOLID'


def records_from_buyer_file(name, data, fobs=None):
    """One record per main sheet (template / Solid / Heather). fobs: {sheet: fob} when already known."""
    from cbs_fill import is_main_sheet, calc_fob
    info = parse_name(name)
    wf = openpyxl.load_workbook(io.BytesIO(data))
    wv = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    recs = []
    from cbs_fill import buyer_columns, buyer_rows, FABRIC_SECTIONS
    for ws in wf.worksheets:
        if not is_main_sheet(ws):
            continue
        vs = wv[ws.title]
        h, c = buyer_columns(ws)
        nets = [ws.cell(r, c['net']).value for r, comp, sec in buyer_rows(ws)[0] if sec in FABRIC_SECTIONS]
        if not any(isinstance(v, (int, float)) and v > 0 for v in nets):
            raise ValueError(f'[{ws.title}] fabric Net Usage all 0 — looks like a blank buyer template, skipped')
        fob = (fobs or {}).get(ws.title)
        if fob is None:
            v = vs['B8'].value
            fob = v if isinstance(v, (int, float)) else calc_fob(ws)[0]
        season = info['season'] or season_code(vs['B4'].value)
        stage = info['stage'] or _s(vs['B5'].value).upper()
        style = info['style'] or (re.search(r'\d{7}', _s(vs['B1'].value)) or [''])[0]
        recs.append(dict(season=season, stage=stage, style=style, color=color_of(ws.title),
                         fob=round(float(fob), 2), date=info['date'] or dt.date.today().isoformat(),
                         file=re.sub(r'^[0-9a-f]{8}-', '', os.path.basename(name))))
    return recs


# ---------------- report ----------------
def report_xlsx(df):
    """Excel in the same look as the FOB summary table."""
    from openpyxl.styles import Alignment, Border, Font, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'FOB'
    hdr = ['SEASON', 'STAGE', 'STYLE#', 'COLOR', 'FOB (USD)']
    ws.append(hdr)
    for _, r in df.iterrows():
        ws.append([r['season'], r['stage'], int(r['style']) if str(r['style']).isdigit() else r['style'],
                   r['color'], float(r['fob'])])
    thin = Side(style='thin')
    for row in ws.iter_rows():
        for c in row:
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            c.alignment = Alignment(horizontal='center', vertical='center')
            if c.row == 1:
                c.font = Font(bold=True)
            elif c.column == 5:
                c.number_format = '"$"#,##0.00'
    for col, w in zip('ABCDE', (10, 10, 12, 22, 14)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def sort_table(df):
    order = {'SOLID': 0, 'HEATHER': 1}
    return df.assign(_c=df['color'].map(order).fillna(2)).sort_values(['season', 'stage', 'style', '_c']).drop(columns='_c')

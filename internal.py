"""Read a HOKA internal CBS workbook: latest tab, cost blocks, submitted check, stage, gender."""
import re
import openpyxl

from cbs_format import find_blocks, find_label, map_columns

STAGE_ORDER = ['VP1', 'P1', 'P2', 'P3', 'P4', 'P5', 'JSS', 'PP']           # fit stages
BUY_ORDER = ['GTM1', 'GTM2', 'SMS']                                         # buying stages


def _s(v):
    return str(v).strip() if v is not None else ''


def _num(v):
    return v if isinstance(v, (int, float)) else None


def parse_tab(title):
    """'GTM2 S28-10.26' -> {'buy':'GTM2','season':'S28','date':'10.26'}"""
    t = title.upper().replace(' ', '')
    buy = next((b for b in BUY_ORDER if b in t), None)
    m = re.search(r'GTM(\d)', t)
    if m:
        buy = 'GTM' + m.group(1)
    season = (re.search(r'([SF]\d{2})', t) or [None, None])[1]
    date = (re.search(r'(\d{1,2}\.\d{1,2})', title) or [None, None])[1]
    fit = (re.search(r'\b(P\d|PP|JSS|VP\d)\b', title.upper()) or [None, None])[1]
    return dict(buy=buy, fit=fit, season=season, date=date)


def latest_fit(option_text):
    found = [s for s in STAGE_ORDER if re.search(rf'(^|[\s*/]){s}\b', _s(option_text).upper())]
    return found[-1] if found else None


def gender_of(*texts):
    for t in texts:
        m = re.search(r'(?:^|[\s_\-])([MW])(?:[\s_\-]|$)', _s(t))
        if m:
            return m.group(1)
    return None


STYLE_RE = re.compile(r'(?<!\d)(\d{7})(A?)(?![0-9A-Za-z])|(?<!\d)(\d{7})(?!\d)')


def find_style(text):
    """'1187346A_W...' -> '1187346A' (Asia fit keeps its A), '1187348_W' -> '1187348'."""
    m = STYLE_RE.search(_s(text))
    if not m:
        return None
    return (m.group(1) + m.group(2)) if m.group(1) else m.group(3)


def style_no(*texts):
    for t in texts:
        st = find_style(t)
        if st:
            return st
    return None


def read_internal(path, filename=''):
    wf = openpyxl.load_workbook(path)                    # formulas (for later edits/format)
    wv = openpyxl.load_workbook(path, data_only=True)    # cached values
    ws, wsv = wf.worksheets[0], wv.worksheets[0]
    tab = parse_tab(ws.title)
    blocks = []
    for i, (r0, r1) in enumerate(find_blocks(wsv)):
        hdr = find_label(wsv, r0, min(r0 + 5, r1), {'S/#'})
        if not hdr:
            continue
        cols = map_columns(wsv, hdr.row)
        net = find_label(wsv, hdr.row + 1, r1, {'NET'})
        fob, fob_cell = None, None
        if net and 'b_ttl' in cols:
            for r in range(net.row, min(net.row + 3, r1) + 1):
                c = ws.cell(r, cols['b_ttl'])
                if _s(c.value).upper().startswith('=SUM'):
                    fob, fob_cell = _num(wsv.cell(r, cols['b_ttl']).value), c.coordinate
                    break
            if fob is None:                      # FOB typed / other formula: last number > 1 under Net
                for r in range(net.row, min(net.row + 3, r1) + 1):
                    v = _num(wsv.cell(r, cols['b_ttl']).value)
                    if v is not None and v > 1:
                        fob, fob_cell = v, ws.cell(r, cols['b_ttl']).coordinate
        notes = []
        for row in wsv.iter_rows(min_row=r0, max_row=r1, max_col=22):
            for c in row:
                m = re.match(r'^\s*CBS\s*(\d{1,2}/\d{1,2})\s*\$\s*([\d.]+)', _s(c.value))
                if m:
                    notes.append((m.group(1), float(m.group(2))))
        submitted = any(fob is not None and abs(round(fob, 2) - p) < 0.011 for _, p in notes)
        hd_stage = _s(wsv.cell(hdr.row - 1, 4).value)
        hd_date = wsv.cell(hdr.row - 1, 1).value
        body = ' '.join(_s(wsv.cell(r, c).value) for r in range(hdr.row + 1, (net.row if net else r1))
                        for c in (3, 4, 5)).upper()
        warn = []
        if tab['buy'] and tab['buy'] not in hd_stage.upper().replace(' ', ''):
            warn.append(f'block stage "{hd_stage}" ≠ tab "{ws.title}"')
        if notes and not submitted:
            warn.append(f'CBS note {notes[-1][1]} ≠ FOB {fob and round(fob, 2)}')
        label = ''
        m = re.search(r'COLORWAY\s+([A-Z ]+)', hd_stage.upper())
        if m:
            label = m.group(1).strip()
        blocks.append(dict(idx=i, r0=r0, r1=r1, fob=fob, fob_cell=fob_cell, notes=notes,
                           submitted=submitted, stage_hdr=hd_stage, date_hdr=hd_date,
                           heather='HEATHER' in body, label=label, warn=warn,
                           option=_s(wsv.cell(hdr.row - 1, 10).value)))
    first_txt = _s(wsv['A4'].value) or _s(wsv['B4'].value)
    info = dict(
        tab=ws.title, tab_info=tab,
        style=style_no(filename, first_txt),
        gender=gender_of(filename, first_txt),
        fit=latest_fit(' '.join(b['option'] for b in blocks)),
        hidden=[s.title for s in wf.worksheets if s.sheet_state != 'visible'],
        blocks=blocks)
    # pick the blocks that matter: submitted ones; else the last block
    act = [b for b in blocks if b['submitted']] or blocks[-1:]
    info['active'] = [b['idx'] for b in act]
    return info, wf, wv


def tab_date(title):
    """'GTM2 S28-11.10(OPT2)' -> (11, 10); 'SMS-F27 9.7' -> (9, 7); None if no m.d in the name."""
    m = re.findall(r'(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)', _s(title))
    return (int(m[-1][0]), int(m[-1][1])) if m else None


def previous_stage_block(wv, heather=False):
    """(sheet_index, title, r0, r1, fob) of the previous stage: the next tab to the right whose sheet date is
    different from the current tab's (tabs of the same date, e.g. 11.10(OG) / 11.10(OPT2), are options of the
    same stage, not the previous one)."""
    cur = tab_date(wv.worksheets[0].title)
    for si, ws in enumerate(wv.worksheets[1:], 1):
        if ws.sheet_state != 'visible':
            continue
        if cur and tab_date(ws.title) == cur:
            continue
        bl = find_blocks(ws)
        if not bl:
            continue
        cands = []
        for r0, r1 in bl:
            hdr = find_label(ws, r0, min(r0 + 5, r1), {'S/#'})
            if not hdr:
                continue
            body = ' '.join(_s(ws.cell(r, c).value) for r in range(hdr.row + 1, r1) for c in (3, 5)).upper()
            cols = map_columns(ws, hdr.row)
            net = find_label(ws, hdr.row + 1, r1, {'NET'})
            fob = None
            if net and 'b_ttl' in cols:
                for r in range(net.row, min(net.row + 3, r1) + 1):
                    v = ws.cell(r, cols['b_ttl']).value
                    if isinstance(v, (int, float)) and v > 1:
                        fob = v
            cands.append((('HEATHER' in body) == heather, si, ws.title, r0, r1, fob))
        if cands:
            best = sorted(cands, key=lambda x: (not x[0], x[3]))[0]     # top block = that stage's latest
            return best[1:]
    return None

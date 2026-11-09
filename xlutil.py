"""Excel helpers: cross-workbook sheet copy, row insert with formula shifting, images."""
import io
import re
from copy import copy

from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

REF = re.compile(r"(?<![A-Za-z_\d'!])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")


def copy_sheet(src, dst_wb, title=None, index=None):
    ws = dst_wb.create_sheet(title or src.title, index)
    for row in src.iter_rows():
        for c in row:
            if c.__class__.__name__ == 'MergedCell':
                continue
            n = ws.cell(c.row, c.column, c.value)
            if c.has_style:
                n.font, n.fill, n.border = copy(c.font), copy(c.fill), copy(c.border)
                n.alignment, n.protection = copy(c.alignment), copy(c.protection)
                n.number_format = c.number_format
    for rng in src.merged_cells.ranges:
        ws.merge_cells(str(rng))
    for k, d in src.column_dimensions.items():
        ws.column_dimensions[k].width = d.width
        ws.column_dimensions[k].hidden = d.hidden
    for k, d in src.row_dimensions.items():
        if d.height:
            ws.row_dimensions[k].height = d.height
    ws.sheet_view.zoomScale = src.sheet_view.zoomScale
    ws.sheet_properties.tabColor = src.sheet_properties.tabColor
    for im in src._images:
        add_image(ws, img_bytes(im), im.anchor, im.width, im.height)
    return ws


def img_bytes(im):
    """Image bytes without closing the underlying stream (openpyxl's _data() closes it)."""
    ref = im.ref
    if hasattr(ref, 'getvalue'):
        return ref.getvalue()
    if hasattr(ref, 'fp') and hasattr(ref.fp, 'getvalue'):
        return ref.fp.getvalue()
    data = im._data()
    im.ref = io.BytesIO(data)
    return data


def add_image(ws, data, anchor, width=None, height=None):
    img = XLImage(io.BytesIO(data))
    if width:
        img.width, img.height = width, height
    img.anchor = copy(anchor) if not isinstance(anchor, str) else anchor
    ws.add_image(img)
    return img


def _shift_formula(f, at, n):
    def rep(m):
        r = int(m.group(4))
        if r >= at:
            r += n
        return f'{m.group(1)}{m.group(2)}{m.group(3)}{r}'
    return REF.sub(rep, f)


def insert_row(ws, at, style_from=None, extend_sum_end=None):
    """Insert one row at `at`, shift formulas/merges/images; optionally extend SUM ranges ending at extend_sum_end."""
    merges = [str(m) for m in ws.merged_cells.ranges]
    for m in merges:
        ws.unmerge_cells(m)
    heights = {k: d.height for k, d in ws.row_dimensions.items() if d.height}
    ws.insert_rows(at)
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                c.value = _shift_formula(c.value, at, 1)
                if extend_sum_end is not None:
                    c.value = re.sub(
                        rf'([A-Z]{{1,3}}\$?\d+:\$?[A-Z]{{1,3}}\$?){extend_sum_end}\b',
                        lambda m: f'{m.group(1)}{at}', c.value)
    for m in merges:
        a, b = m.split(':')
        ws.merge_cells(':'.join(_shift_formula(x, at, 1) for x in (a, b)))
    for k, h in heights.items():
        ws.row_dimensions[k + 1 if k >= at else k].height = h
    for im in ws._images:
        a = im.anchor
        if hasattr(a, '_from') and a._from.row >= at - 1:
            a._from.row += 1
            if getattr(a, 'to', None) is not None:
                a.to.row += 1
    if style_from:
        src = style_from if style_from < at else style_from + 1
        for col in range(1, ws.max_column + 1):
            s, d = ws.cell(src, col), ws.cell(at, col)
            if s.has_style:
                d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
                d.alignment, d.number_format = copy(s.alignment), s.number_format
            if isinstance(s.value, str) and s.value.startswith('='):
                d.value = re.sub(rf'(\$?[A-Z]{{1,3}}\$?){src}(?!\d)', rf'\g<1>{at}', s.value)
        if ws.row_dimensions[src].height:
            ws.row_dimensions[at].height = ws.row_dimensions[src].height


def col_letter(i):
    return get_column_letter(i)

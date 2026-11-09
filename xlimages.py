"""Read pictures straight from the xlsx zip (keeps WMF/EMF that openpyxl drops) + convert to PNG."""
import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile


def _rels(z, path):
    p = os.path.dirname(path) + '/_rels/' + os.path.basename(path) + '.rels'
    if p not in z.namelist():
        return {}
    x = z.read(p).decode('utf8')
    out = {}
    for m in re.finditer(r'<Relationship\b([^>]*)/?>', x):
        a = m.group(1)
        i = re.search(r'Id="([^"]+)"', a)
        t = re.search(r'Target="([^"]+)"', a)
        if i and t:
            out[i.group(1)] = t.group(1)
    return out


def _resolve(base_path, target):
    if target.startswith('/'):
        return target.lstrip('/')
    return os.path.normpath(os.path.join(os.path.dirname(base_path), target)).replace('\\', '/')


def sheet_path(z, index=0):
    wb = z.read('xl/workbook.xml').decode('utf8')
    rid = re.findall(r'<sheet\b[^>]*?r:id="([^"]+)"', wb)[index]
    return _resolve('xl/workbook.xml', _rels(z, 'xl/workbook.xml')[rid])


def pictures(data, sheet_index=0):
    """[{col,row,col2,row2,name,ext,data}] (0-based anchors)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    sp = sheet_path(z, sheet_index)
    srel = _rels(z, sp)
    out = []
    for rid, tgt in srel.items():
        p = _resolve(sp, tgt)
        if not re.search(r'drawings/drawing\d+\.xml$', p):
            continue
        d = z.read(p).decode('utf8')
        drel = _rels(z, p)
        for a in re.findall(r'<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?</xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>', d):
            e = re.search(r'r:embed="([^"]+)"', a)
            if not e or e.group(1) not in drel:
                continue
            fr = re.search(r'<xdr:from>\s*<xdr:col>(\d+)</xdr:col>[\s\S]*?<xdr:row>(\d+)</xdr:row>', a)
            to = re.search(r'<xdr:to>\s*<xdr:col>(\d+)</xdr:col>[\s\S]*?<xdr:row>(\d+)</xdr:row>', a)
            ext = re.search(r'<xdr:ext cx="(\d+)" cy="(\d+)"', a) or re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', a)
            nm = re.search(r'<xdr:cNvPr\b[^>]*name="([^"]*)"', a)
            mp = _resolve(p, drel[e.group(1)])
            out.append(dict(col=int(fr.group(1)) if fr else 0, row=int(fr.group(2)) if fr else 0,
                            col2=int(to.group(1)) if to else None, row2=int(to.group(2)) if to else None,
                            cx=int(ext.group(1)) if ext else None, cy=int(ext.group(2)) if ext else None,
                            name=nm.group(1) if nm else '', ext=os.path.splitext(mp)[1].lower().lstrip('.'),
                            data=z.read(mp)))
    return out


def to_png(data, ext):
    """PNG bytes. WMF/EMF are converted with LibreOffice (or ImageMagick) when available."""
    if ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp'):
        return data
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, f'img.{ext}')
        open(src, 'wb').write(data)
        if soffice:
            subprocess.run([soffice, '--headless', '--convert-to', 'png', '--outdir', td, src],
                           capture_output=True, timeout=90)
        elif shutil.which('convert'):
            subprocess.run(['convert', src, os.path.join(td, 'img.png')], capture_output=True, timeout=60)
        p = os.path.join(td, 'img.png')
        return open(p, 'rb').read() if os.path.exists(p) else None

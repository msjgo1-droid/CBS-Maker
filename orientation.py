"""Judge from the marker picture whether garments are laid lengthwise (1pc/1way rule for Self).

Each garment is drawn in one colour in the CAD marker. Lengthwise = the same-colour panels sit side by side
along the marker length (a horizontal band). Widthwise = same-colour panels stacked across the fabric width
(a vertical column). Random = one colour scattered all over the marker.
"""
import io

import numpy as np
from scipy import ndimage
from PIL import Image, ImageDraw


def _colour_masks(a, max_colours=14, min_share=0.03):
    sat = (a.max(axis=2) - a.min(axis=2)) > 45
    pts = a[sat]
    if len(pts) < 500:
        return [], sat
    sample = pts[np.random.default_rng(0).choice(len(pts), min(len(pts), 40000), replace=False)]
    # simple k-means on RGB
    k = max_colours
    cent = sample[np.linspace(0, len(sample) - 1, k).astype(int)].astype(float)
    for _ in range(15):
        d = ((sample[:, None, :] - cent[None]) ** 2).sum(-1)
        lab = d.argmin(1)
        for j in range(k):
            if (lab == j).any():
                cent[j] = sample[lab == j].mean(0)
    # merge near-identical centres
    merged = []
    for c in cent:
        if all(np.linalg.norm(c - m) > 30 for m in merged):
            merged.append(c)
    cent = np.array(merged)
    d = ((a[..., None, :].astype(float) - cent[None, None]) ** 2).sum(-1)
    lab = d.argmin(-1)
    near = d.min(-1) < 40 ** 2
    masks = []
    total = sat.sum()
    for j, c in enumerate(cent):
        m = (lab == j) & near & sat
        m = ndimage.binary_opening(m, iterations=2)          # drop outlines, grain lines, text
        if m.sum() >= min_share * total:
            masks.append((tuple(int(x) for x in c), m))
    return masks, sat


def judge(png, ratio_len=1.0, max_height_share=0.6):
    """Returns dict(verdict='lengthwise'|'widthwise'|'mixed'|'unknown', share, colours, annotated)."""
    im = Image.open(io.BytesIO(png)).convert('RGB')
    if im.width > 800:
        im = im.resize((800, max(1, int(im.height * 800 / im.width))), Image.NEAREST)
    a = np.asarray(im).astype(int)
    H, Wd = a.shape[:2]
    masks, sat = _colour_masks(a)
    if not masks:
        return dict(verdict='unknown', share=None, colours=[], annotated=png)
    out, area_len, area_all = [], 0, 0
    draw_im = im.copy()
    dr = ImageDraw.Draw(draw_im)
    r = max(3, int(0.06 * H))
    for col, m in masks:
        lab, n = ndimage.label(ndimage.binary_dilation(m, iterations=r))
        sizes = ndimage.sum(m, lab, range(1, n + 1))
        tot = m.sum()
        for g in range(1, n + 1):
            area = sizes[g - 1]
            if area < 0.15 * tot:                          # stray single pieces: judged with the colour below
                continue
            ys, xs = np.where((lab == g) & m)
            x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
            xe, ye = x1 - x0, y1 - y0
            if ye <= max_height_share * H and xe >= ratio_len * ye:
                kind = 'L'
            elif ye > max_height_share * H and xe < ye:
                kind = 'W'
            else:
                kind = 'M'
            out.append(dict(rgb=col, kind=kind, x_ext=round(xe / H, 2), y_ext=round(ye / H, 2), area=int(area)))
            clr = {'L': (0, 170, 0), 'W': (220, 0, 0), 'M': (255, 140, 0)}[kind]
            dr.rectangle([x0, y0, x1, y1], outline=clr, width=3)
            dr.text((x0 + 4, y0 + 2), kind, fill=clr)
            area_all += area
            if kind == 'L':
                area_len += area
        big = sizes.max() if n else 0
        scattered = tot - sum(a for a in sizes if a >= 0.15 * tot)
        if scattered > 0.4 * tot:                          # colour broken into many far-apart bits = random
            out.append(dict(rgb=col, kind='R', x_ext=None, y_ext=None, area=int(scattered)))
            area_all += scattered
    if not area_all:
        return dict(verdict='unknown', share=None, colours=out, annotated=png)
    share = area_len / area_all if area_all else 0
    nW = sum(o['area'] for o in out if o['kind'] == 'W') / area_all
    verdict = 'lengthwise' if share >= 0.75 else ('widthwise' if nW >= 0.5 else 'mixed')
    buf = io.BytesIO()
    draw_im.save(buf, 'PNG')
    return dict(verdict=verdict, share=round(share, 2), colours=out, annotated=buf.getvalue())

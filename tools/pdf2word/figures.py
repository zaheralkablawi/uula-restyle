#!/usr/bin/env python3
"""Step 2 — the figures of each page as pictures (graphs, drawings, photos, equations you want as pictures).

  python3 figures.py find in.pdf work/                       # every candidate figure, cropped
  python3 figures.py crop in.pdf PAGE X0 TOP X1 BOTTOM out.png  # crop a region yourself (points, from the page's top-left)

find writes work/fig-p01-1.png ... and work/figures.json: [{"file", "page", "bbox": [x0, top, x1, bottom], "w", "h", "parts"}]
("parts": the pieces a crop was merged from — use them with crop when a figure was joined to its neighbour)
(look at every crop: drop page decorations, merge or re-crop where needed)
"""
import sys, os, json, subprocess
import pdfplumber

DPI = 250


def crop(src, page, x0, top, x1, bottom, out, pad=4):
    x0, top = max(0, x0 - pad), max(0, top - pad)
    x1, bottom = x1 + pad, bottom + pad
    s = DPI / 72.0
    base = out[:-4]
    subprocess.run(['pdftoppm', '-r', str(DPI), '-f', str(page), '-l', str(page), '-png', '-singlefile',
                    '-x', str(int(x0 * s)), '-y', str(int(top * s)), '-W', str(int((x1 - x0) * s)), '-H', str(int((bottom - top) * s)),
                    src, base], check=True)
    return out


def merge(boxes, gap=8):
    boxes = [list(b) for b in boxes]
    changed = True
    while changed:
        changed = False
        out = []
        while boxes:
            b = boxes.pop()
            for o in boxes[:]:
                if b[0] - gap <= o[2] and o[0] - gap <= b[2] and b[1] - gap <= o[3] and o[1] - gap <= b[3]:
                    b = [min(b[0], o[0]), min(b[1], o[1]), max(b[2], o[2]), max(b[3], o[3])]
                    boxes.remove(o)
                    changed = True
            out.append(b)
        boxes = out
    return boxes


def find(src, out):
    os.makedirs(out, exist_ok=True)
    res = []
    with pdfplumber.open(src) as pdf:
        for pn, pg in enumerate(pdf.pages, 1):
            W, H = pg.width, pg.height
            boxes, imgs = [], []
            for im in pg.images:
                imgs.append((im['x0'], im['top'], im['x1'], im['bottom']))
            def light(c):
                if c is None or isinstance(c, str):
                    return False
                c = c if isinstance(c, (list, tuple)) else [c]
                if len(c) == 4:                                                    # CMYK
                    return sum(c) < 0.6
                return all(v > 0.75 for v in c)
            for o in pg.curves + pg.lines + pg.rects:
                w, h = o['x1'] - o['x0'], o['bottom'] - o['top']
                # watermarks, light backgrounds of frames, light grid / border lines are not figures
                if (o.get('fill') and not o.get('stroke') and light(o.get('non_stroking_color'))) or \
                        (o.get('stroke') and not o.get('fill') and light(o.get('stroking_color'))):
                    continue
                if w > W * 0.8 or h > H * 0.5 or w * h > W * H * 0.25:      # page frames, backgrounds, rules across the page
                    continue
                boxes.append((o['x0'], o['top'], o['x1'], o['bottom']))
            k = 0
            # drawings close together are one figure; pictures join a drawing (or another picture) only where they overlap
            for b in sorted(merge(merge(boxes) + imgs, gap=1), key=lambda b: (b[1], -b[2])):
                w, h = b[2] - b[0], b[3] - b[1]
                if w < 20 or h < 14 or (h < 3 or w < 3):                     # underlines, bullets, tiny marks
                    continue
                k += 1
                f = os.path.join(out, 'fig-p%02d-%d.png' % (pn, k))
                crop(src, pn, *b, f)
                parts = [[round(v, 1) for v in o] for o in merge(boxes) + imgs if o[0] >= b[0] - 1 and o[2] <= b[2] + 1 and o[1] >= b[1] - 1 and o[3] <= b[3] + 1]
                res.append({'file': f, 'page': pn, 'bbox': [round(v, 1) for v in b], 'w': round(w, 1), 'h': round(h, 1),
                            'parts': parts if len(parts) > 1 else []})
    json.dump(res, open(os.path.join(out, 'figures.json'), 'w'), indent=1)
    for r in res:
        print('page %d  %-28s %5.0f x %-5.0f at x=%.0f top=%.0f' % (r['page'], os.path.basename(r['file']), r['w'], r['h'], r['bbox'][0], r['bbox'][1]))


if __name__ == '__main__':
    if sys.argv[1] == 'find':
        find(sys.argv[2], sys.argv[3])
    elif sys.argv[1] == 'crop':
        a = sys.argv
        print(crop(a[2], int(a[3]), float(a[4]), float(a[5]), float(a[6]), float(a[7]), a[8]))

#!/usr/bin/env python3
"""Step 1 — look at the PDF: one picture and one text draft per page, and a report.

  python3 pdf_pages.py in.pdf work/

work/page-01.png ...   the page as a picture (read these to see the layout, colours, figures, equations)
work/page-01.txt ...   the text the PDF holds, if any (exact characters; Arabic order may be off)
prints per page: characters found, whether the text looks usable, pictures and drawings on it
"""
import sys, os, subprocess, re
import pdfplumber

src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
subprocess.run(['pdftoppm', '-r', '110', '-png', src, os.path.join(out, 'page')], check=True)
# pdftoppm names pages page-1.png or page-01.png depending on the count: make them two digits
for f in os.listdir(out):
    m = re.match(r'page-(\d+)\.png$', f)
    if m and len(m.group(1)) == 1:
        os.rename(os.path.join(out, f), os.path.join(out, 'page-0%s.png' % m.group(1)))

with pdfplumber.open(src) as pdf:
    print('pages:', len(pdf.pages), ' size (pt): %.0f x %.0f' % (pdf.pages[0].width, pdf.pages[0].height))
    for i, pg in enumerate(pdf.pages, 1):
        txt = subprocess.run(['pdftotext', '-layout', '-f', str(i), '-l', str(i), src, '-'], capture_output=True, text=True).stdout
        txt = re.sub('[\u200e\u200f\u202a-\u202e\u2066-\u2069]', '', txt)          # direction marks
        open(os.path.join(out, 'page-%02d.txt' % i), 'w', encoding='utf-8').write(txt)
        chars = [c['text'] for c in pg.chars]
        n = len(chars)
        arabic = sum(1 for c in chars if '\u0600' <= c <= '\u06FF')
        forms = sum(1 for c in chars if '\uFB50' <= c <= '\uFEFF')          # presentation forms: text may need fixing
        bad = sum(1 for c in chars if '\uE000' <= c <= '\uF8FF' or c == '\ufffd' or c.startswith('(cid:'))
        if n < 20:
            state = 'NO TEXT (scanned page or pictures only) -> read the page picture'
        elif bad > n * 0.1:
            state = 'BROKEN TEXT (unreadable glyphs) -> read the page picture'
        elif forms > n * 0.3:
            state = 'text in presentation forms -> usable after normalising, check against the picture'
        else:
            state = 'text OK -> use it, checked against the picture'
        print('page %d: %d chars (%d Arabic), %d pictures, %d drawings: %s' % (i, n, arabic, len(pg.images), len(pg.curves) + len(pg.lines) + len(pg.rects), state))

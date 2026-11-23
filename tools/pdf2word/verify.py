#!/usr/bin/env python3
"""Step 5 — check the Word file against the PDF: same number of pages, each page starting where the PDF page starts.

  python3 verify.py in.pdf out.docx work/

prints OK, or the pages that ran over (shrink those pages' gaps / font size in spec.json and build again)
and writes work/check-01.png ... (the Word pages as pictures, to compare with page-01.png ...)
"""
import sys, os, subprocess, re, shutil

src, docx, work = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(work, exist_ok=True)
tmp = os.path.join(work, '_check')
shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, docx], capture_output=True, timeout=180)
out = os.path.join(tmp, os.path.splitext(os.path.basename(docx))[0] + '.pdf')
if not os.path.exists(out):
    sys.exit('could not open the Word file (LibreOffice)')


def pages(p):
    r = subprocess.run(['pdfinfo', p], capture_output=True, text=True).stdout
    return int(re.search(r'Pages:\s+(\d+)', r).group(1))


def words(p, n):
    t = subprocess.run(['pdftotext', '-f', str(n), '-l', str(n), p, '-'], capture_output=True, text=True).stdout
    return [w for w in re.findall(r'[\u0600-\u06FFA-Za-z0-9]{2,}', t)]


a, b = pages(src), pages(out)
subprocess.run(['pdftoppm', '-r', '60', '-png', out, os.path.join(work, 'check')])
for f in os.listdir(work):
    m = re.match(r'check-(\d+)\.png$', f)
    if m and len(m.group(1)) == 1:
        os.rename(os.path.join(work, f), os.path.join(work, 'check-0%s.png' % m.group(1)))
print('PDF pages: %d   Word pages: %d' % (a, b))
bad = []
for n in range(1, min(a, b) + 1):
    wa, wb = set(words(src, n)), set(words(out, n))
    if wa and len(wa & wb) < len(wa) * 0.6:
        bad.append(n)
if a == b and not bad:
    print('OK: page for page')
else:
    if b > a:
        print('Word has %d extra page(s): a page ran over' % (b - a))
    if bad:
        print('pages whose words differ from the PDF page (look at check-NN.png): %s' % bad)
print('(LibreOffice substitutes fonts it lacks, e.g. UULA Sans, so lines can wrap a little differently than in Word)')

---
name: uula-pdf-to-word
description: PDF إلى وورد مميز: يحوّل ملف PDF (محوَّل أو نسخة ضوئية) إلى ملف وورد مميَّز بإطارات UULA، صفحة مقابل صفحة وبنفس الترتيب والصور والنص حرفيًا، جاهز للإدراج في الباوربوينت بأداة Restyle.
---

# PDF ← وورد مميَّز (لأداة UULA Restyle)

المستخدم يرفق ملف PDF من مذكرات علا، محوَّلًا من وورد أو نسخة ضوئية. المطلوب: ملف وورد فيه **كل شيء حرفيًا وبنفس الترتيب**، وكل **صفحة PDF تقابلها صفحة وورد واحدة**، لا زيادة ولا نقص. وكل جزء مميَّز بإطار `UULA:<type>` كما تفعل لوحة الوورد في أداة UULA Restyle، فيدرجه المستخدم بعدها في الباوربوينت من «إدراج ← إدراج ملف وورد».

الردود للمستخدم دائمًا بالعربية.

## الشروط (لا تُخالف)

1. **حرفيًا:** كل كلمة ورقم وعلامة كما في الـPDF. لا تلخّص، ولا تصحّح أخطاء الكاتب، ولا تعِد الصياغة، ولا تضف شيئًا من عندك. الأرقام بنفس شكلها (هندية أو إنجليزية).
2. **نفس الترتيب:** ترتيب القراءة في الصفحة من الأعلى للأسفل، ومن اليمين لليسار في العربية. ما كان بجانب شيء آخر يبقى بجانبه (صورة بجانب سؤال تُكتب `"float": "left"` أو `"right"` حسب جهتها).
3. **صفحة مقابل صفحة:** لكل صفحة PDF عنصر واحد في `pages`، والسكربت يبدأ كل صفحة في صفحة وورد جديدة. لا تنقل محتوى من صفحة إلى أخرى أبدًا.
4. **الصور موجودة:** كل رسم أو شكل أو تمثيل بياني أو صورة يُقصّ من الـPDF ويوضع في مكانه بنفس عرضه تقريبًا.
5. **الألوان والغامق كما هي:** الحل الأزرق أو الأحمر يبقى بلونه، والعنوان الغامق غامق.

## الخطوات

اكتب السكربتات الأربعة في آخر هذا الملف إلى مجلد العمل كما هي (تحتاج pdfplumber وpython-docx وlxml، وأدوات poppler: pdftoppm وpdftotext وpdfinfo، وLibreOffice للتحقق). ثم انسخ الـPDF إلى مجلد العمل باسم `in.pdf` (لا تعدّل الأصل).

1. **النظر إلى الصفحات:** `python3 pdf_pages.py in.pdf work/`
   - **اقرأ كل صورة صفحة** `work/page-NN.png` بأداة Read. هي المرجع للترتيب والألوان والمعادلات والصور.
   - إذا قال التقرير `text OK`: خذ النص من `work/page-NN.txt` لأنه حروف الملف الحقيقية، لكن صحّح ترتيبه وأخطاءه بالمقارنة مع الصورة. أخطاء الاستخراج المعروفة في العربية:
     - لام ألف تأتي مقلوبة: «ألن» صوابها «لأن»، و«ال» صوابها «لا».
     - التنوين يتحرك: «أوًال» صوابها «أولاً».
     - أقواس وأرقام في غير مكانها، والمعادلات تضيع.
   - إذا قال `NO TEXT` (نسخة ضوئية) أو `BROKEN TEXT`: **اقرأ النص من الصورة نفسها.** للخط الصغير أو المعادلات كبّر الجزء: `pdftoppm -r 220 -f N -l N -x X -y Y -W W -H H -png in.pdf work/zoom`. البكسل = النقطة × 220/72. اقرأ الصورة المكبّرة.
2. **الأشكال:** `python3 figures.py find in.pdf work/`، ثم **انظر إلى كل قصّة** `work/fig-pNN-K.png`:
   - **أبقِ** الرسوم والأشكال والصور والتمثيلات البيانية الحقيقية.
   - **أسقط** زخارف القالب: الشعار، رمز QR، العلامة المائية، الخلفيات، الإطارات، ورسوم رأس الصفحة وتذييلها.
   - **حدود الجداول** قد تظهر كشكل: الجدول يُكتب جدولًا نصيًا (`rows`) وليس صورة.
   - إذا التصق شكل بجاره، أو قُصّ ناقصًا، أعِد القص بنفسك: `python3 figures.py crop in.pdf صفحة x0 أعلى x1 أسفل work/fig-...png` (بالنقاط من أعلى يسار الصفحة؛ `parts` في figures.json تعطي القطع المنفصلة).
   - النص المكتوب **داخل** الرسم (أسماء النقاط، الأرقام على الأسهم) يبقى داخل الصورة، ولا يُكتب مرة ثانية في الوورد.
3. **اكتب `spec.json`** صفحة صفحة. الصيغة مشروحة في رأس `build_docx.py`. ملاحظات:
   - `page`: مقاس صفحة الـPDF من التقرير (نقاط)، و`margin` قريب من هامش الصفحة.
   - كل كتلة لها `type` حسب جدول الأنواع أدناه. النص الذي لا نوع له، مثل رأس الصفحة وتذييلها ورقمها، يُكتب بـ `"type": null` إذا أردت الحفاظ عليه، ولن يُنقل إلى الباوربوينت.
   - المعادلات تُكتب بـ `math` (LaTeX مبسّط مشروح في السكربت): معادلة في سطرها `{"math": "..."}`، وداخل سطر `{"runs": [..., {"math": "..."}]}`. اكتبها كما تظهر بالضبط، وفي المعادلات العربية حافظ على ترتيب الحروف والأرقام كما يظهر في الصفحة. إذا كانت المعادلة معقدة جدًا لا يغطيها الـLaTeX المبسّط، قصّها صورة بدل كتابتها خطأ، واذكرها في التقرير.
   - الإجابة من كلمة أو كلمتين داخل سطر السؤال: run عليه `"answer": true` بلونها الأصلي. والخيار الصحيح داخل سطر الخيارات: `"correct": true`.
4. **ابنِ:** `python3 build_docx.py spec.json out.docx`
5. **تحقّق:** `python3 verify.py in.pdf out.docx work/`
   - يجب أن يطبع `OK: page for page`.
   - إذا زادت صفحات الوورد، صفحةٌ تجاوزت: قلّل في تلك الصفحة فقط `gap` أو `size` أو عرض الصورة أو `margin` العام، وأعد البناء. **لا تنقل محتوى لصفحة أخرى.**
   - انظر إلى `work/check-NN.png` بجانب `work/page-NN.png`، وتأكد أن كل شيء موجود وبنفس الترتيب. الخط قد يختلف قليلًا لأن LibreOffice لا يملك خط UULA Sans.
6. **سلّم الملف:** انسخه باسم الـPDF الأصلي مع ` - مميز.docx` وأرسله (SendUserFile). ومعه تقرير قصير بالعربية:
   - عدد الصفحات، وعدد كل نوع.
   - ما قرأته من صورة لأن النص غير موجود أو تالف (نسخة ضوئية).
   - كل كلمة أو معادلة لم تكن متأكدًا من قراءتها، برقم الصفحة، ليراجعها.
   - المعادلات التي وضعتها صورًا.
   - وذكّره بالخطوة التالية: يفتحه في الوورد للمراجعة (تبويب «مراجعة» في الأداة)، ثم «إدراج ملف وورد» في الباوربوينت.

## الأنواع (`type`)

| type | الاسم | كيف تعرفه |
|---|---|---|
| `chapter` | اسم الوحدة | «الوحدة …» أعلى الدرس، قصير |
| `lesson` | اسم الدرس | «الدرس …» أو عنوان الدرس الكبير، أو رقم مثل (1 – 2) ثم اسم |
| `topic` | موضوع | عنوان قسم داخل الدرس، قصير غامق غالبًا، يليه شرح |
| `body` | فقرة | شرح عادي (فقرات الشرح المتتالية لنفس الفكرة كتلة واحدة) |
| `bullets` | نقاط | قائمة نقاط في الشرح |
| `qhead` | عنوان أسئلة | «أسئلة الدرس»، «أجب عن …»، «تمارين»، «تحقق من فهمك»، «حاول أن تحل»، عنوان مجموعة أسئلة |
| `question` | سؤال | استفهام أو أمر (أوجد، احسب، أثبت، علل، اذكر، اختر …)، «مثال 1:»، «أولاً: … المطلوب:»، رقم ثم نص |
| `choice` | خيارات | أسطر أ) ب) ج) د) |
| `answer` | حل | الحل/الجواب/البرهان/الإجابة، أو نص الحل الملوّن تحت السؤال، كل أسطره حتى السؤال التالي |
| `note` | ملاحظة | ملاحظة، تنبيه، انتبه، تذكّر، هام |
| `def` | تعريف | تعريف، المفهوم، «X هو …» لمصطلح |
| `quote` | نص مقتطف | آية، حديث، بيت شعر، اقتباس |
| `formula` | قوانين | قانون، نظرية، قاعدة، خاصية، الصيغة العامة |
| `other` | صيغة أخرى | صيغة أخرى، طريقة أخرى، حل آخر |
| `activity` | نشاط | نشاط |
| `mars` | من المريخ | إثراء، تحدٍّ، فكّر، سؤال ذكاء، من المريخ |
| `table` | جدول | جدول مستقل في الشرح |
| `figure` | شكل | صورة أو رسم مستقل في الشرح |

**قواعد التجميع:**
- السؤال يأخذ معه أسطره الفرعية («- السهم رقم (1)…») ومعادلاته وشكله (حتى لو جاء الشكل بعده). فتبقى الصورة جزءًا من السؤال ولا تُخفى مع الحلول.
- الخيارات كتلة `choice` بعد السؤال.
- الحل كتلة `answer` منفصلة، تشمل كل أسطره وصوره (التمثيل البياني للحل داخل كتلة الحل).
- الإطار في الـPDF (ملاحظة، تعريف، نظرية…) كتلة واحدة بكل ما بداخله.
- الجدول داخل سؤال أو حل يدخل في كتلته. والجدول المستقل `table`.
- كتلة واحدة لكل سؤال. وتحت عنوان الأسئلة: العنوان `qhead`، ثم كل سؤال وحله كل منهما كتلة.
- عند الشك اختر الأقرب واذكره في التقرير.

## السكربتات

### `pdf_pages.py`

```python
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
```

### `figures.py`

```python
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
```

### `build_docx.py`

```python
#!/usr/bin/env python3
"""Step 4 — build the tagged Word file from spec.json (page for page).

  python3 build_docx.py spec.json out.docx

spec.json:
{
  "page": {"width": 595.3, "height": 841.9, "margin": 36},          # points, from the PDF page
  "font": "UULA Sans", "size": 11,
  "pages": [
    {"blocks": [
      {"type": "lesson",   "paras": ["الدرس الأول: المعادلات"]},
      {"type": "question", "paras": [
          "مثال 1: حل المعادلة الآتية:",
          {"math": "x^{2}-5x+6=0"},
          {"runs": [{"t": "السهم (1) يشير إلى: "}, {"t": "فجوة مركزية.", "answer": true, "color": "FF0000"}]},
          {"img": "work/fig-p01-1.png", "w": 150, "float": "left"}]},
      {"type": "answer",   "paras": [{"runs": [{"t": "الحل: ", "bold": true}, {"t": "س = ٢"}], "color": "1F71DE"}]},
      {"type": "table", "head": true, "rows": [["العبارة", "صح أم خطأ"], ["...", {"runs": [{"t": "صح", "answer": true}]}]]}
    ]}
  ]
}
block "type": chapter lesson topic body bullets qhead question choice answer note def quote formula other activity mars table figure
          (null = left untagged: it will not be brought into PowerPoint)
paragraph: a string, or {"runs": [...]} / {"math": "..."} / {"img": file, "w": pt, "float": "left"|"right"} with optional
           "align": "center"|"left" (default: the reading side), "bold", "color", "size", "bullet": true, "gap": pt before
run: {"t": text, "bold", "italic", "color": "RRGGBB", "size", "sup", "sub", "answer": true | "correct": true, "math": "..."}
math (a small LaTeX): \\frac{a}{b}  x^{2}  x_{1}  \\sqrt{x}  \\sqrt[3]{x}  \\overline{AB}  \\vec{v}  \\pi \\times \\div \\pm \\le \\ge \\ne
                       \\infty \\theta \\alpha \\beta \\angle \\parallel \\perp \\in \\cdot \\degree; anything else is written as it is
"""
import sys, json, re, copy
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
NS = 'xmlns:w="%s" xmlns:m="%s"' % (W, M)
AR = re.compile('[\u0600-\u06FF]')

TYPES = {
    'chapter': ('اسم الوحدة', 'A6A6A6'), 'lesson': ('اسم الدرس', '1F71DD'), 'topic': ('موضوع', '404040'),
    'body': ('فقرة', '8C8C8C'), 'bullets': ('نقاط', '8C8C8C'), 'qhead': ('عنوان أسئلة', 'C00000'),
    'question': ('سؤال', 'FF0000'), 'choice': ('خيارات', '7030A0'), 'answer': ('حل', '1F71DE'),
    'correct': ('الخيار الصحيح', '2070DF'), 'note': ('ملاحظة', 'E69138'), 'def': ('تعريف', '2AB672'),
    'quote': ('نص مقتطف', '2AB672'), 'formula': ('قوانين', '5B9BD5'), 'other': ('صيغة أخرى', 'C45A11'),
    'activity': ('نشاط', '7F7F7F'), 'mars': ('من المريخ', '595959'), 'table': ('جدول', '44546A'),
    'figure': ('شكل', '00B050'), 'given': ('شكل من السؤال', 'A6A6A6'),
}

# ---------- a small LaTeX -> OMML ----------
SYM = {'pi': 'π', 'times': '×', 'div': '÷', 'pm': '±', 'mp': '∓', 'le': '≤', 'leq': '≤', 'ge': '≥', 'geq': '≥', 'ne': '≠', 'neq': '≠',
       'infty': '∞', 'theta': 'θ', 'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'Delta': 'Δ', 'lambda': 'λ', 'mu': 'μ',
       'sigma': 'σ', 'Sigma': 'Σ', 'omega': 'ω', 'phi': 'φ', 'angle': '∠', 'parallel': '∥', 'perp': '⊥', 'in': '∈', 'notin': '∉',
       'subset': '⊂', 'cup': '∪', 'cap': '∩', 'cdot': '·', 'degree': '°', 'circ': '°', 'to': '→', 'rightarrow': '→', 'Rightarrow': '⇒',
       'Leftrightarrow': '⇔', 'approx': '≈', 'equiv': '≡', 'triangle': '△', 'therefore': '∴', 'because': '∵', 'emptyset': '∅',
       'sum': '∑', 'int': '∫', 'quad': ' ', ',': ' ', ';': ' ', ' ': ' ', '{': '{', '}': '}', '%': '%'}


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def mr(t):
    if not t:
        return ''
    return '<m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t xml:space="preserve">%s</m:t></m:r>' % esc(t)


class TeX:
    def __init__(self, s):
        self.s, self.i = s, 0

    def group(self):
        """{...} or one token"""
        self.ws()
        if self.i < len(self.s) and self.s[self.i] == '{':
            self.i += 1
            out = self.seq('}')
            self.i += 1
            return out
        return self.atom()

    def ws(self):
        while self.i < len(self.s) and self.s[self.i] == ' ':
            self.i += 1

    def atom(self):
        s = self.s
        if self.i >= len(s):
            return ''
        c = s[self.i]
        if c == '\\':
            m = re.match(r'\\([A-Za-z]+|.)', s[self.i:])
            name = m.group(1)
            self.i += len(m.group(0))
            if name == 'frac':
                a = self.group(); b = self.group()
                return '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (a, b)
            if name == 'sqrt':
                deg = ''
                if self.i < len(s) and s[self.i] == '[':
                    j = s.index(']', self.i)
                    deg = TeX(s[self.i + 1:j]).seq(None)
                    self.i = j + 1
                e = self.group()
                pr = '' if deg else '<m:radPr><m:degHide m:val="1"/></m:radPr>'
                return '<m:rad>%s<m:deg>%s</m:deg><m:e>%s</m:e></m:rad>' % (pr, deg, e)
            if name == 'overline':
                return '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>%s</m:e></m:bar>' % self.group()
            if name in ('vec', 'overrightarrow'):
                return '<m:acc><m:accPr><m:chr m:val="\u20d7"/></m:accPr><m:e>%s</m:e></m:acc>' % self.group()
            if name == 'hat':
                return '<m:acc><m:accPr><m:chr m:val="\u0302"/></m:accPr><m:e>%s</m:e></m:acc>' % self.group()
            if name in ('text', 'mathrm', 'textbf'):
                self.ws()
                if self.i < len(s) and s[self.i] == '{':
                    j = s.index('}', self.i)
                    t = s[self.i + 1:j]; self.i = j + 1
                    return mr(t)
            if name in ('left', 'right'):
                return self.atom()
            return mr(SYM.get(name, name))
        self.i += 1
        return mr(c)

    def seq(self, end):
        out = []
        buf = ''
        while self.i < len(self.s) and (end is None or self.s[self.i] != end):
            c = self.s[self.i]
            if c in '^_':
                self.i += 1
                # the base is the last item written
                base = out.pop() if out else mr('')
                if buf:
                    out.append(mr(buf[:-1])); base = mr(buf[-1]); buf = ''
                arg = self.group()
                if c == '^':
                    out.append('<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>' % (base, arg))
                else:
                    out.append('<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>' % (base, arg))
                continue
            if c in '\\{':
                if buf:
                    out.append(mr(buf)); buf = ''
                out.append(self.group() if c == '{' else self.atom())
                continue
            buf += c
            self.i += 1
        if buf:
            out.append(mr(buf))
        return ''.join(out)


def omml(tex, display=False):
    inner = TeX(tex).seq(None)
    x = '<m:oMath %s>%s</m:oMath>' % (NS, inner)
    if display:
        x = '<m:oMathPara %s><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>%s</m:oMathPara>' % (NS, x.replace(' ' + NS, '', 1))
    return parse_xml(x)


# ---------- tags ----------
def sdt(type_):
    title, col = TYPES[type_]
    return parse_xml('<w:sdt xmlns:w="%s" xmlns:w15="%s"><w:sdtPr><w:alias w:val="%s"/><w:tag w:val="UULA:%s"/>'
                     '<w15:color w:val="%s"/><w15:appearance w15:val="boundingBox"/></w:sdtPr><w:sdtContent/></w:sdt>' % (W, W15, title, type_, col))


def wrap(elements, type_):
    s = sdt(type_)
    elements[0].addprevious(s)
    c = s.find(qn('w:sdtContent'))
    for e in elements:
        c.append(e)
    return s


# ---------- building ----------
class Builder:
    def __init__(self, spec):
        self.spec = spec
        self.doc = Document()
        pg = spec.get('page', {})
        sec = self.doc.sections[0]
        sec.page_width = Pt(pg.get('width', 595.3)); sec.page_height = Pt(pg.get('height', 841.9))
        if pg.get('width', 595.3) > pg.get('height', 841.9):
            sec.orientation = WD_ORIENT.LANDSCAPE
        m = Pt(pg.get('margin', 36))
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = m
        sec._sectPr.append(parse_xml('<w:bidi %s/>' % NS))
        st = self.doc.styles['Normal']
        font = spec.get('font', 'UULA Sans')
        st.font.name = font; st.font.size = Pt(spec.get('size', 11))
        rf = st.element.get_or_add_rPr().get_or_add_rFonts()
        for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rf.set(qn(a), font)
        st.element.get_or_add_rPr().append(parse_xml('<w:szCs %s w:val="%d"/>' % (NS, spec.get('size', 11) * 2)))
        st.paragraph_format.space_after = Pt(2); st.paragraph_format.space_before = Pt(0)
        self.body = self.doc.element.body
        self.content_w = pg.get('width', 595.3) - 2 * pg.get('margin', 36)

    def run(self, p, r, para):
        if r.get('math'):
            p._p.append(omml(r['math']))
            return p._p[-1]
        t = r.get('t', '')
        run = p.add_run(t)
        rtl = bool(AR.search(t)) or bool(AR.search(para.get('_text', '')))
        if rtl:
            run._r.get_or_add_rPr().append(parse_xml('<w:rtl %s/>' % NS))
        b = r.get('bold', para.get('bold'))
        if b:
            run.bold = True
            run._r.get_or_add_rPr().append(parse_xml('<w:bCs %s/>' % NS))
        if r.get('italic'):
            run.italic = True
        col = r.get('color', para.get('color'))
        if col:
            run.font.color.rgb = RGBColor.from_string(col.upper())
        sz = r.get('size', para.get('size'))
        if sz:
            run.font.size = Pt(sz)
            run._r.get_or_add_rPr().append(parse_xml('<w:szCs %s w:val="%d"/>' % (NS, int(sz * 2))))
        if r.get('sup'):
            run.font.superscript = True
        if r.get('sub'):
            run.font.subscript = True
        return run._r

    def para(self, cont, spec):
        """one paragraph into container (body or a table cell); returns the w:p"""
        if isinstance(spec, str):
            spec = {'runs': [{'t': spec}]}
        p = cont.add_paragraph()
        ppr = p._p.get_or_add_pPr()
        text = ''.join(r.get('t', '') for r in spec.get('runs', []))
        spec['_text'] = text
        if AR.search(text) or (not text and 'math' not in spec and 'img' not in spec):
            ppr.append(parse_xml('<w:bidi %s/>' % NS))
        al = spec.get('align')
        if al == 'center':
            p.alignment = 1
        elif al == 'left':
            ppr.append(parse_xml('<w:jc %s w:val="%s"/>' % (NS, 'right' if AR.search(text) else 'left')))
        if spec.get('gap'):
            p.paragraph_format.space_before = Pt(spec['gap'])
        if spec.get('bullet'):
            p.paragraph_format.left_indent = None
            spec['runs'] = [{'t': '▪ '}] + spec.get('runs', [])
        if 'math' in spec:
            p._p.append(omml(spec['math'], display=True))
        elif 'img' in spec:
            w = min(spec.get('w', 150), self.content_w)
            p.add_run().add_picture(spec['img'], width=Pt(w))
            if spec.get('float') in ('left', 'right'):
                self.floating(p, spec['float'])
            else:
                p.alignment = 1
        for r in spec.get('runs', []):
            el = self.run(p, r, spec)
            if r.get('answer') or r.get('correct'):
                wrap([el], 'correct' if r.get('correct') else 'answer')
        return p._p

    def floating(self, p, side):
        """inline picture -> picture on the left / right with the text wrapping around it"""
        inl = p._p.find('.//' + qn('wp:inline'))
        if inl is None:
            return
        WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        anc = etree.SubElement(inl.getparent(), '{%s}anchor' % WP)
        for k, v in dict(distT='0', distB='0', distL='114300', distR='114300', simplePos='0', relativeHeight='2', behindDoc='0',
                         locked='0', layoutInCell='1', allowOverlap='1').items():
            anc.set(k, v)
        anc.append(parse_xml('<wp:simplePos xmlns:wp="%s" x="0" y="0"/>' % WP))
        anc.append(parse_xml('<wp:positionH xmlns:wp="%s" relativeFrom="margin"><wp:align>%s</wp:align></wp:positionH>' % (WP, side)))
        anc.append(parse_xml('<wp:positionV xmlns:wp="%s" relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' % WP))
        for tag in ('extent', 'effectExtent'):
            e = inl.find('{%s}%s' % (WP, tag))
            if e is not None:
                anc.append(e)
        anc.append(parse_xml('<wp:wrapSquare xmlns:wp="%s" wrapText="bothSides"/>' % WP))
        for tag in ('docPr', 'cNvGraphicFramePr'):
            e = inl.find('{%s}%s' % (WP, tag))
            if e is not None:
                anc.append(e)
        anc.append(inl.find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphic'))
        inl.getparent().remove(inl)

    def table(self, b):
        rows = b['rows']
        ncol = max(len(r) for r in rows)
        t = self.doc.add_table(rows=len(rows), cols=ncol)
        t.style = 'Table Grid'
        txt = json.dumps(rows, ensure_ascii=False)
        if AR.search(txt):
            t._tbl.tblPr.append(parse_xml('<w:bidiVisual %s/>' % NS))
        if b.get('widths'):
            for row in t.rows:
                for k, c in enumerate(row.cells):
                    if k < len(b['widths']):
                        c.width = Pt(b['widths'][k])
        for i, r in enumerate(rows):
            for j, val in enumerate(r):
                cell = t.cell(i, j)
                first = cell.paragraphs[0]._p
                vals = val if isinstance(val, list) else [val]
                for v in vals:
                    if isinstance(v, str):
                        v = {'runs': [{'t': v, 'bold': bool(b.get('head') and i == 0)}]}
                    self.para(cell, v)
                first.getparent().remove(first)
        return t._tbl

    def build(self):
        first_page = True
        for page in self.spec['pages']:
            page_start = None
            for b in page.get('blocks', []):
                start = len(self.body) - 1           # before sectPr
                if b.get('type') == 'table' or 'rows' in b:
                    els = [self.table(b)]
                else:
                    els = [self.para(self.doc, p) for p in b.get('paras', [])]
                if not els:
                    continue
                if page_start is None:
                    page_start = els[0]
                ty = b.get('type')
                if ty and ty in TYPES and not (ty == 'table'):
                    wrap(els, ty)
                elif ty == 'table':
                    wrap(els, 'table')
            # page for page: every PDF page starts a new Word page
            if not first_page and page_start is not None:
                if page_start.tag == qn('w:tbl'):
                    br = parse_xml('<w:p %s><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>' % NS)
                    holder = page_start.getparent() if page_start.getparent().tag == qn('w:sdtContent') else None
                    (page_start.getparent().getparent() if holder is not None else page_start).addprevious(br)
                else:
                    ppr = page_start.find(qn('w:pPr'))
                    if ppr is None:
                        ppr = etree.SubElement(page_start, qn('w:pPr')); page_start.insert(0, ppr)
                    ppr.insert(0, parse_xml('<w:pageBreakBefore %s/>' % NS))
            first_page = False
        return self.doc


if __name__ == '__main__':
    spec = json.load(open(sys.argv[1], encoding='utf-8'))
    Builder(spec).build().save(sys.argv[2])
    print('saved', sys.argv[2], 'pages:', len(spec['pages']))
```

### `verify.py`

```python
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
```

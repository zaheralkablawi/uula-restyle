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

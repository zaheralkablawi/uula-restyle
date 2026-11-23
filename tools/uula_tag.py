#!/usr/bin/env python3
"""UULA Restyle — tag a Word file's parts (content controls "UULA:<type>").

  python3 uula_tag.py dump  in.docx            > blocks.txt   # numbered blocks to classify
  python3 uula_tag.py apply in.docx plan.json out.docx        # wrap the blocks as planned

plan.json:
  {"groups": [{"from": 0, "to": 0, "type": "lesson"}, {"from": 3, "to": 5, "type": "question"}, ...],
   "inline": [{"block": 4, "text": "فجوة مركزية.", "type": "answer"}, ...]}
"""
import sys, json, re, zipfile, shutil, copy
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
q = lambda t: '{%s}%s' % (W, t)

TYPES = {  # id: (Arabic title, frame colour)
    'chapter': ('اسم الوحدة', 'A6A6A6'), 'lesson': ('اسم الدرس', '1F71DD'), 'topic': ('موضوع', '404040'),
    'body': ('فقرة', '8C8C8C'), 'bullets': ('نقاط', '8C8C8C'), 'qhead': ('عنوان أسئلة', 'C00000'),
    'question': ('سؤال', 'FF0000'), 'choice': ('خيارات', '7030A0'), 'answer': ('حل', '1F71DE'),
    'correct': ('الخيار الصحيح', '2070DF'), 'note': ('ملاحظة', 'E69138'), 'def': ('تعريف', '2AB672'),
    'quote': ('نص مقتطف', '2AB672'), 'formula': ('قوانين', '5B9BD5'), 'other': ('صيغة أخرى', 'C45A11'),
    'activity': ('نشاط', '7F7F7F'), 'mars': ('من المريخ', '595959'), 'table': ('جدول', '44546A'),
    'figure': ('شكل', '00B050'), 'given': ('شكل من السؤال', 'A6A6A6'),
}


def load(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('word/document.xml'))
    return z, root


def body_of(root):
    return root.find(q('body'))


def blocks(body):
    """body-level items in order (paragraphs, tables, existing content controls); sectPr excluded"""
    return [el for el in body if el.tag in (q('p'), q('tbl'), q('sdt'))]


def ptext(el):
    out = []
    for n in el.iter():
        if n.tag == q('t'):
            out.append(n.text or '')
        elif n.tag == q('tab'):
            out.append('\t')
        elif n.tag == '{%s}t' % M:
            out.append(n.text or '')
    return ''.join(out)


def run_colors(p):
    cols = []
    for r in p.iter(q('r')):
        t = ''.join(x.text or '' for x in r.iter(q('t'))).strip()
        if not t:
            continue
        c = r.find(q('rPr') + '/' + q('color'))
        col = c.get(q('val')) if c is not None else 'auto'
        if cols and cols[-1][0] == col:
            cols[-1][1] += ' ' + t
        else:
            cols.append([col, t])
    return cols


def tag_of(sdt):
    t = sdt.find(q('sdtPr') + '/' + q('tag'))
    v = t.get(q('val')) if t is not None else ''
    return v


def describe(el):
    if el.tag == q('tbl'):
        rows = el.findall(q('tr'))
        cells = [' | '.join(ptext(c).strip()[:25] for c in r.findall(q('tc'))) for r in rows[:3]]
        return 'TABLE %d rows: %s' % (len(rows), ' / '.join(cells))
    if el.tag == q('sdt'):
        return 'TAGGED %s: %s' % (tag_of(el), ptext(el).strip()[:120])
    pPr = el.find(q('pPr'))
    style = ''
    if pPr is not None and pPr.find(q('pStyle')) is not None:
        style = pPr.find(q('pStyle')).get(q('val'))
    num = pPr is not None and pPr.find(q('numPr')) is not None
    pic = any(True for _ in el.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip')) or \
        any(True for _ in el.iter('{urn:schemas-microsoft-com:vml}imagedata'))
    math = any(True for _ in el.iter('{%s}oMath' % M))
    def is_b(r):
        b = r.find(q('rPr') + '/' + q('b'))
        return b is not None and b.get(q('val'), 'true') not in ('0', 'false', 'off')
    bold = any(is_b(r) for r in el.iter(q('r')) if ''.join(x.text or '' for x in r.iter(q('t'))).strip())
    text = ptext(el).strip()
    flags = [f for f, on in (('style=' + style, bool(style)), ('list', num), ('PICTURE', pic), ('MATH', math), ('bold', bold)) if on]
    cols = run_colors(el)
    colour = ''
    if len(cols) > 1 or (cols and cols[0][0] not in ('auto', '000000')):
        colour = ' colours: ' + ' '.join('[%s]%s' % (c, t[:40]) for c, t in cols)
    if not text and not pic:
        return '(empty)'
    return '%s%s%s' % (('{' + ', '.join(flags) + '} ') if flags else '', text[:200], colour)


def dump(path):
    z, root = load(path)
    for i, el in enumerate(blocks(body_of(root))):
        print('%d\t%s' % (i, describe(el)))


def sdt_xml(type_):
    title, col = TYPES[type_]
    sdt = etree.Element(q('sdt'))
    pr = etree.SubElement(sdt, q('sdtPr'))
    etree.SubElement(pr, q('alias')).set(q('val'), title)
    etree.SubElement(pr, q('tag')).set(q('val'), 'UULA:' + type_)
    etree.SubElement(pr, '{%s}color' % W15).set(q('val'), col)
    etree.SubElement(pr, '{%s}appearance' % W15).set('{%s}val' % W15, 'boundingBox')
    etree.SubElement(sdt, q('sdtContent'))
    return sdt


def spacer_p():
    p = etree.Element(q('p'))
    ppr = etree.SubElement(p, q('pPr'))
    sp = etree.SubElement(ppr, q('spacing'))
    sp.set(q('before'), '0'); sp.set(q('after'), '0')
    rpr = etree.SubElement(ppr, q('rPr'))
    etree.SubElement(rpr, q('sz')).set(q('val'), '12')
    return p


def split_run(r, at):
    """split run r at character offset `at` of its text; returns (left, right)"""
    t = r.find(q('t'))
    s = t.text or ''
    right = copy.deepcopy(r)
    t.text = s[:at]
    right.find(q('t')).text = s[at:]
    for x in (t, right.find(q('t'))):
        x.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.addnext(right)
    return r, right


def wrap_inline(p, text, type_):
    """put the exact `text` of paragraph p into an inline content control"""
    runs = [r for r in p.iter(q('r')) if r.find(q('t')) is not None and r.getparent().tag != q('sdtContent')]
    # one w:t per run to keep offsets simple
    full = ''.join(r.find(q('t')).text or '' for r in runs)
    k = full.find(text)
    if k < 0:
        k = full.replace('ـ', '').find(text.replace('ـ', ''))
        if k < 0:
            return False
    end = k + len(text)
    pos = 0
    chosen = []
    for r in list(runs):
        s = r.find(q('t')).text or ''
        a, b = pos, pos + len(s)
        pos = b
        if b <= k or a >= end:
            continue
        if a < k:
            r = split_run(r, k - a)[1]; a = k
            s = r.find(q('t')).text or ''
        if b > end:
            split_run(r, end - a)
        chosen.append(r)
    if not chosen:
        return False
    parent = chosen[0].getparent()
    if any(c.getparent() is not parent for c in chosen):
        return False
    sdt = sdt_xml(type_)
    chosen[0].addprevious(sdt)
    content = sdt.find(q('sdtContent'))
    for c in chosen:
        content.append(c)
    return True


def apply(path, plan_path, out):
    plan = json.load(open(plan_path, encoding='utf-8'))
    z, root = load(path)
    body = body_of(root)
    bl = blocks(body)
    report = {'groups': 0, 'inline': 0, 'inline_missed': [], 'spacers': 0}
    # inline answers first (block indexes still valid)
    for it in plan.get('inline', []):
        el = bl[it['block']]
        ps = [el] if el.tag == q('p') else list(el.iter(q('p')))
        if any(wrap_inline(p, it['text'], it.get('type', 'answer')) for p in ps):
            report['inline'] += 1
        else:
            report['inline_missed'].append(it['text'])
    groups = sorted(plan.get('groups', []), key=lambda g: g['from'])
    used = set()
    prev_to = None
    for g in groups:
        items = [bl[i] for i in range(g['from'], g['to'] + 1) if i not in used]
        if not items:
            continue
        used.update(range(g['from'], g['to'] + 1))
        # an already tagged block: retag it instead of nesting
        if len(items) == 1 and items[0].tag == q('sdt') and tag_of(items[0]).startswith('UULA:'):
            pr = items[0].find(q('sdtPr'))
            pr.getparent().remove(pr)
            sdt = sdt_xml(g['type'])
            items[0].insert(0, sdt.find(q('sdtPr')))
            report['groups'] += 1
            prev_to = g['to']
            continue
        sdt = sdt_xml(g['type'])
        items[0].addprevious(sdt)
        content = sdt.find(q('sdtContent'))
        for el in items:
            content.append(el)
        # two tagged parts that touch get an empty line between them (their frames do not stick together)
        if prev_to is not None and g['from'] == prev_to + 1:
            sdt.addprevious(spacer_p()); report['spacers'] += 1
        report['groups'] += 1
        prev_to = g['to']
    # untagged leftovers (not empty)
    left = [i for i, el in enumerate(bl) if i not in used and el.tag != q('sdt') and describe(el) != '(empty)']
    report['untagged'] = left
    # make sure w15 is declared (Word 2013+ files always have it)
    if W15 not in root.nsmap.values():
        nroot = etree.Element(root.tag, nsmap=dict(root.nsmap, w15=W15))
        nroot[:] = root[:]
        for k, v in root.attrib.items():
            nroot.set(k, v)
        root = nroot
    xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            data = xml if item.filename == 'word/document.xml' else z.read(item.filename)
            zo.writestr(item, data)
    print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else ''
    if cmd == 'dump':
        dump(sys.argv[2])
    elif cmd == 'apply':
        apply(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        print(__doc__)

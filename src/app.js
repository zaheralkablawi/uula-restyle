(() => {
'use strict';
const $ = id => document.getElementById(id);
const VER = '1.1';
const PT = UR.PT;
const save = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
const load = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } };
let lang = load('ur-lang', 'ar');
const L = (ar, en) => lang === 'en' ? en : ar;
// the tool works on the current slide only (all slides at once was slow and risky)
let scope = 'cur';
let busy = false, supported = false;

// ---------- status / results ----------
const status = $('status'), res = $('res');
const say = (m, cls) => { status.textContent = m || ''; status.className = 'status' + (cls ? ' ' + cls : ''); };
const tsay = m => say(m);
const prog = f => { $('prog').hidden = f == null; if (f != null) $('prog').firstElementChild.style.width = Math.round(f * 100) + '%'; };
function line(cls, html) { const d = document.createElement('div'); d.className = 'ln ' + cls; d.innerHTML = html; res.appendChild(d); return d; }
const slideLinks = (ids, num) => ids.map(id => '<a data-id="' + id + '">' + num(id) + '</a>').join('، ');
res.addEventListener('click', e => {
  const a = e.target.closest('a[data-id]'); if (!a) return;
  PowerPoint.run(async ctx => { ctx.presentation.setSelectedSlides([a.dataset.id]); await ctx.sync(); }).catch(() => {});
});
const errText = e => { const loc = e && e.debugInfo && e.debugInfo.errorLocation; return (e && (e.code || e.message) || String(e)) + (loc ? ' (' + loc + ')' : ''); };
const cnt = n => lang === 'en' ? (n === 1 ? '1 slide' : n + ' slides') : n === 1 ? 'شريحة واحدة' : n === 2 ? 'شريحتان' : n <= 10 ? n + ' شرائح' : n + ' شريحة';
const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

// ---------- Office helpers ----------
const run = f => PowerPoint.run(f);
async function targetIds() {
  return run(async ctx => {
    const col = scope === 'cur' ? ctx.presentation.getSelectedSlides() : ctx.presentation.slides;
    col.load('items/id'); await ctx.sync();
    const ids = col.items.map(s => s.id);
    return scope === 'cur' ? ids.slice(0, 1) : ids;
  });
}
// every selected slide (the solutions work on all of them at once)
async function selIds() { return run(async ctx => { const s = ctx.presentation.getSelectedSlides(); s.load('items/id'); await ctx.sync(); return s.items.map(x => x.id); }); }
async function allIds() { return run(async ctx => { const s = ctx.presentation.slides; s.load('items/id'); await ctx.sync(); return s.items.map(x => x.id); }); }
async function currentId() { return run(async ctx => { const s = ctx.presentation.getSelectedSlides(); s.load('items/id'); await ctx.sync(); return s.items.length ? s.items[0].id : null; }); }
async function exportSlide(id) { return run(async ctx => { const r = ctx.presentation.slides.getItem(id).exportAsBase64(); await ctx.sync(); return r.value; }); }
async function deleteSlide(id) { return run(async ctx => { ctx.presentation.slides.getItem(id).delete(); await ctx.sync(); }); }
async function insertAfter(b64, afterId) {
  const before = await allIds();
  const tryIns = async target => run(async ctx => {
    ctx.presentation.insertSlidesFromBase64(b64, { formatting: 'UseDestinationTheme', targetSlideId: target });
    await ctx.sync();
  });
  try { await tryIns(afterId); }
  catch (e) { await tryIns(String(afterId).split('#')[0] + '#'); }
  const ids = await allIds();
  const fresh = ids.find(i => !before.includes(i));
  return fresh || ids[ids.indexOf(afterId) + 1];
}
// the shapes of a slide by their " #ur<n>" mark
async function shapesByMark(ctx, slideId) {
  const shapes = ctx.presentation.slides.getItem(slideId).shapes;
  shapes.load('items/id,items/name,items/type'); await ctx.sync();
  const map = {};
  for (const s of shapes.items) { const m = / #ur(\d+)$/.exec(s.name || ''); if (m) map[+m[1]] = s; }
  return { map, shapes };
}
// the text items of a plan (with the width they get from a first layout)
function textItems(plan) {
  const out = [];
  const add = it => { if (it.kind === 'text') out.push(it); };
  plan.header.forEach(add);
  plan.blocks.forEach(b => (b.type === 'box' ? b.rows : [b]).forEach(r => r.cols.forEach(c => c.forEach(add))));
  return out;
}
// real text heights from PowerPoint, then the final positions
async function refine(slideId, plan) {
  return run(async ctx => {
    const { map } = await shapesByMark(ctx, slideId);
    const first = UR.layout(plan, null);
    const texts = textItems(plan).filter(it => map[it.idx]);
    for (const it of texts) {
      const s = map[it.idx], p = first.pos[it.idx];
      if (p && p.w) { s.width = p.w / PT; }
      try { s.textFrame.wordWrap = true; s.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; } catch (e) {}
    }
    await ctx.sync();
    const tables = [];
    plan.blocks.forEach(b => (b.type === 'box' ? b.rows : [b]).forEach(r => r.cols.forEach(c => c.forEach(it => { if (it.kind === 'table' && map[it.idx]) tables.push(it); }))));
    [...texts, ...tables].forEach(it => map[it.idx].load('height'));
    await ctx.sync();
    const heights = {};
    for (const it of [...texts, ...tables]) heights[it.idx] = Math.round(map[it.idx].height * PT);
    const lay = UR.layout(plan, heights);
    for (const [k, p] of Object.entries(lay.pos)) {
      const s = map[+k]; if (!s) continue;
      if (p.hide) { s.delete(); continue; }
      if (p.x != null) s.left = p.x / PT;
      if (p.y != null) s.top = p.y / PT;
      if (!p.keep) {
        if (p.w != null && p.w > 0) s.width = p.w / PT;
        if (p.h != null && p.h > 0) s.height = p.h / PT;
      }
    }
    await ctx.sync();
    return lay;
  });
}
// remove the " #ur<n>" marks
async function unmark(slideId) {
  await run(async ctx => {
    const shapes = ctx.presentation.slides.getItem(slideId).shapes;
    shapes.load('items/name'); await ctx.sync();
    for (const s of shapes.items) if (/ #ur\d+$/.test(s.name || '')) s.name = s.name.replace(/ #ur\d+$/, '');
    await ctx.sync();
  }).catch(() => {});
  await frameTop(slideId);
}
// a frame (note, formulas, definition…) that is the first thing on the page, near the top: its top edge goes up to the
// top edge of the slide (its bottom and everything inside it stay where they are)
async function frameTop(slideId) {
  const { W, H } = await getSize().catch(() => ({ W: 0, H: 0 }));
  if (!W) return;
  return run(async ctx => {
    const sh = ctx.presentation.slides.getItem(slideId).shapes;
    sh.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height'); await ctx.sync();
    const fixed = x => /^UULA (Footer|Watermark)|^CustomPageNumber_/.test(x.name || '') || x.top >= H - 54 || x.left + x.width <= 0 || x.left >= W || (x.width > W * 0.85 && x.height > H * 0.9);
    const cand = sh.items.filter(x => x.type === 'GeometricShape' && x.width >= W * 0.85 && x.top > 0.3 && x.top <= 80 && !fixed(x) && !/Line$/.test(x.name || ''));
    if (!cand.length) return;
    const fills = cand.map(x => { try { x.fill.load('type'); return x.fill; } catch (e) { return null; } });
    await ctx.sync();
    cand.forEach((b, k) => {
      let solid = false; try { solid = fills[k] && (fills[k].type === 'Solid' || fills[k].type === 'Gradient'); } catch (e) {}
      if (!solid && !/Container$/.test(b.name || '')) return;
      // nothing above it (its own content sits inside it, below its top)
      const above = sh.items.some(o => o !== b && !fixed(o) && o.top + o.height <= b.top + 1 && o.height > 1);
      if (above) return;
      const bottom = b.top + b.height; b.top = 0; b.height = bottom;
    });
    await ctx.sync();
  }).catch(() => {});
}

// ---------- convert ----------
// The converted slide replaces the original in place (same number of slides).
// A slide longer than the page is NOT split on its own: the pane asks first.
const MAX_PARTS = 6;
const pendingSplit = new Map();      // slide id -> { pkg, rels, xml, plan }
async function convertOne(id, opts, rep) {
  const b64 = await exportSlide(id);
  const pkg = await UR.openPackage(b64);
  let r;
  try { r = UR.restyle(pkg, { watermark: opts.watermark }); }
  catch (e) { throw new Error(L('تحليل الشريحة: ', 'Reading the slide: ') + errText(e)); }
  let newId;
  try { newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), id); }
  catch (e) { throw new Error(L('إدراج الشريحة: ', 'Inserting the slide: ') + errText(e)); }
  const lay = await refine(newId, r.plan);
  await unmark(newId);
  await deleteSlide(id);
  if (lay.overflow >= 0 && UR.splitPoint(r.plan, lay) > 0) { pendingSplit.set(newId, { pkg, rels: r.rels, xml: r.xml, plan: r.plan }); rep.full.push(newId); }
  else if (lay.overflow >= 0) rep.tooLong.push(newId);
  rep.boxes += r.report.boxes; rep.tables += r.report.tables;
  r.report.links.forEach(u => rep.links.push(u)); r.report.linkFail.forEach(u => rep.linkFail.push(u));
  if (r.report.qr) rep.qr += r.report.qr;
  return newId;
}
// split a long slide into as many slides as needed (only after the user agrees)
async function splitSlide(id) {
  const st = pendingSplit.get(id); if (!st) return [id];
  const queue = [{ xml: st.xml, plan: st.plan }];
  let prev = id, made = [], parts = 0;
  while (queue.length) {
    const part = queue.shift();
    const newId = await insertAfter(await UR.buildPackage(st.pkg, part.xml, st.rels), prev);
    const lay = await refine(newId, part.plan);
    if (lay.overflow >= 0 && parts < MAX_PARTS) {
      const k = UR.splitPoint(part.plan, lay);
      if (k > 0) { await deleteSlide(newId); const [A, B] = UR.splitParts(part.xml, part.plan, k); queue.unshift(A, B); parts++; continue; }
    }
    await unmark(newId);
    made.push(newId); prev = newId;
  }
  await deleteSlide(id);
  pendingSplit.delete(id);
  return made;
}
async function askSplit(ids) {
  if (busy) return;
  busy = true; lock();
  try {
    let added = 0;
    for (let k = 0; k < ids.length; k++) {
      say(L('تقسيم الشريحة ', 'Splitting slide ') + (k + 1) + L(' من ', ' of ') + ids.length + '…');
      const made = await splitSlide(ids[k]); added += made.length - 1;
    }
    if ($('numChk').checked) await renumber();
    const box = res.querySelector('.splitask'); if (box) box.remove();
    say(added ? L('تم التقسيم، وأُضيفت ', 'Split; added ') + cnt(added) + '.' : L('تم.', 'Done.'));
  } catch (e) { say(L('تعذّر التقسيم: ', 'Could not split: ') + errText(e)); }
  finally { busy = false; lock(); }
}
async function doConvert() {
  const ids = await targetIds();
  if (!ids.length) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
  pendingSplit.clear();
  const rep = { done: [], full: [], tooLong: [], errs: [], boxes: 0, tables: 0, links: [], linkFail: [], qr: 0 };
  const opts = {};
  for (let k = 0; k < ids.length; k++) {
    prog(k / ids.length); say(L('تحويل الشريحة ', 'Restyling slide ') + (k + 1) + L(' من ', ' of ') + ids.length + '…');
    try { const nid = await convertOne(ids[k], opts, rep); rep.done.push(nid); if (scope === 'cur') await run(async ctx => { ctx.presentation.setSelectedSlides([nid]); await ctx.sync(); }).catch(() => {}); }
    catch (e) { rep.errs.push(errText(e)); }
  }
  prog(null);
  if ($('numChk').checked && rep.done.length) { say(L('ترقيم الصفحات…', 'Numbering pages…')); try { await renumber(); } catch (e) { rep.errs.push(L('الترقيم: ', 'Numbering: ') + errText(e)); } }
  const all = await allIds(), num = i => all.indexOf(i) + 1;
  res.innerHTML = '';
  if (rep.done.length) line('ok', L('✓ تم تحويل ', '✓ Restyled ') + cnt(rep.done.length) + ': ' + slideLinks(rep.done, num));
  if (rep.full.length) {
    const d = line('warn splitask', L('• أطول من الصفحة: ', '• Longer than the page: ') + slideLinks(rep.full, num)
      + '<br><span style="display:flex;gap:6px;margin-top:4px"><button type="button" class="mini" data-split="yes">' + L('قسّمها إلى شرائح جديدة', 'Split into new slides') + '</button><button type="button" class="mini" data-split="no">' + L('لا، اتركها', 'No, leave it') + '</button></span>');
    d.style.display = 'block';
    const list = rep.full.slice();
    d.querySelector('[data-split="yes"]').addEventListener('click', () => askSplit(list));
    d.querySelector('[data-split="no"]').addEventListener('click', () => { d.remove(); list.forEach(i => pendingSplit.delete(i)); });
  }
  if (rep.tooLong.length) line('err', L('• أطول من الصفحة ولا يمكن تقسيمها، راجعها: ', '• Too long and cannot be split, check: ') + slideLinks(rep.tooLong, num));
  if (rep.qr) line('ok', L('• استُبدل QR جديد: ', '• New QR placed: ') + rep.qr + (rep.links.length ? L(' (رُبط الرابط)', ' (link added)') : ''));
  rep.linkFail.forEach(u => line('warn', L('• أضف هذا الرابط للـQR يدويًا: ', '• Add this link to the QR yourself: ') + '<span dir="ltr">' + esc(u) + '</span>'));
  rep.errs.forEach(m => line('err', '✗ ' + esc(m)));
  if (rep.done.length) line('', '<span class="note">' + L('يلزم تثبيت خط UULA Sans على الجهاز.', 'The UULA Sans font must be installed.') + '</span>');
  say(rep.done.length ? L('تم التحويل.', 'Restyled.') : L('لم يتم تحويل أي شريحة.', 'No slide was restyled.'));
}

// ---------- page numbers ----------
function docSettings() { try { return Office.context.document.settings; } catch (e) { return null; } }
function numStart() { return Math.max(1, parseInt($('numStart').value, 10) || 1); }
let slideSize = null;
async function getSize() {
  if (slideSize) return slideSize;
  try { const id = (await allIds())[0]; const pkg = await UR.openPackage(await exportSlide(id)); slideSize = { W: pkg.slideW / PT, H: pkg.slideH / PT }; }
  catch (e) { slideSize = { W: 540, H: 779.5 }; }
  return slideSize;
}
async function renumber() {
  const start = numStart();
  const { W, H } = await getSize();
  await run(async ctx => {
    const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
    const cols = slides.items.map(s => { const sh = s.shapes; sh.load('items/name,items/id'); return sh; });
    await ctx.sync();
    slides.items.forEach((s, i) => {
      const n = i + 1 - start + 1;
      const boxes = cols[i].items.filter(x => /^CustomPageNumber_/.test(x.name || ''));
      if (n < 1) { boxes.forEach(b => b.delete()); return; }
      if (boxes.length) {
        boxes.slice(1).forEach(b => b.delete());
        const b = boxes[0]; b.name = 'CustomPageNumber_' + n; b.textFrame.textRange.text = String(n);
      } else {
        const tb = s.shapes.addTextBox(String(n), { left: W / 2 - 25, top: H - 34.5, width: 50, height: 20.6 });
        tb.name = 'CustomPageNumber_' + n;
        const tf = tb.textFrame; tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
        tf.verticalAlignment = 'Middle'; tf.autoSizeSetting = 'AutoSizeNone';
        const f = tf.textRange.font; f.name = 'UULA Sans'; f.size = 11; f.color = '#A6A6A6';
        tf.textRange.paragraphFormat.horizontalAlignment = 'Center';
      }
    });
    await ctx.sync();
  });
}
// delete the page numbers that are on the slides now: the tool's own, PowerPoint's slide-number boxes,
// and small boxes at the bottom of the page holding only a number
async function deleteNumbers() {
  const { H } = await getSize();
  let n = 0;
  await run(async ctx => {
    const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
    const cols = slides.items.map(s => { const sh = s.shapes; sh.load('items/name,items/type,items/top,items/height,items/width'); return sh; });
    await ctx.sync();
    const ask = [];
    cols.forEach(c => c.items.forEach(x => {
      const nm = x.name || '';
      if (/^CustomPageNumber_/.test(nm) || /^(Slide Number|رقم الشريحة)/i.test(nm)) { x.delete(); n++; return; }
      if (x.top >= H - 60 && x.width < 120 && x.height < 50 && /TextBox|Placeholder|GeometricShape/i.test(x.type || '')) { const tr = x.textFrame.textRange; tr.load('text'); ask.push([x, tr]); }
    }));
    await ctx.sync().catch(() => {});
    for (const [x, tr] of ask) { let t = ''; try { t = tr.text || ''; } catch (e) { continue; } if (/^\s*[0-9٠-٩]{1,4}\s*$/.test(t)) { x.delete(); n++; } }
    await ctx.sync();
  });
  return n;
}
$('numDel').addEventListener('click', async () => {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  try {
    say(L('حذف أرقام الصفحات…', 'Deleting page numbers…'));
    const n = await deleteNumbers();
    // numbering is turned off, so the next conversion does not put the numbers back
    $('numChk').checked = false; saveNum();
    say(n ? L('حُذف ', 'Deleted ') + n + L(' رقم صفحة، وأُوقف الترقيم التلقائي.', ' page numbers; automatic numbering is off.') : L('لا توجد أرقام صفحات.', 'No page numbers found.'));
  } catch (e) { say(L('تعذّر الحذف: ', 'Could not delete: ') + errText(e)); }
  finally { busy = false; lock(); }
});
$('numHere').addEventListener('click', async () => {
  if (!supported) return;
  try { const id = await currentId(); const all = await allIds(); const i = all.indexOf(id); if (i >= 0) { $('numStart').value = i + 1; saveNum(); } } catch (e) {}
});
function saveNum() { const s = docSettings(); if (s) { s.set('urNumStart', numStart()); s.set('urNum', $('numChk').checked); s.saveAsync(); } $('numRow').style.opacity = $('numChk').checked ? '' : '.5'; }
$('numStart').addEventListener('change', saveNum);
$('numChk').addEventListener('change', saveNum);
$('numAll').addEventListener('click', async () => {
  if (busy || !supported) return;
  busy = true; lock();
  try { say(L('ترقيم الصفحات…', 'Numbering pages…')); await renumber(); say(L('تم تحديث أرقام الصفحات.', 'Page numbers updated.')); }
  catch (e) { say(L('تعذّر الترقيم: ', 'Could not number pages: ') + errText(e)); }
  finally { busy = false; lock(); }
});

// ---------- templates ----------
const SW = { header: '1F71DD', topic: '000000', body: 'FFFFFF', bullets: 'FFFFFF', cols: 'FFFFFF', pill: '1F71DC', qhead: 'FF0000', question: 'FF0000', qa: '1F71DE', answer: '1F71DE', qqr: 'FF0000', mcq: 'FF0000', problem: '00B050', other: 'FBEBD3',
  note: 'FBEBD3', quote: 'E2F0D9', formula: 'D3E4F5', activity: 'F1F1F1', mars: 'F1F1F1', cancelled: 'FFC000', table: 'CAD0D8', tableMain: 'DAE0E9', solTable: 'CAD0D8', grid: 'D3D8E0', practice: 'EBE6FE' };
// frame buttons wear the frame's own colours, as it looks in the booklet
const TINT = { note: ['FBEBD3', 'E9C79A'], quote: ['E2F0D9', '2AB672'], def: ['E2F0D9', '2AB672'], formula: ['D3E4F5', '9DC3E6'], other: ['FBEBD3', 'E9C79A'],
  activity: ['F1F1F1', 'C9CDD2'], mars: ['F1F1F1', 'C9CDD2'], cancelled: ['FFFFFF', 'FFC000'], practice: ['F2F2F2', 'C9BFF5'] };
function tint(b, id) {
  const t = TINT[id]; if (!t) return;
  b.style.background = id === 'practice' ? 'radial-gradient(circle, #F2F2F2, #EBE6FE)' : '#' + t[0];
  b.style.borderColor = '#' + t[1];
}
// a small sign on every button, like the booklet's own: the red Q of a question, the bulb of a note …
function iconHtml(id, col) {
  const img = k => typeof UR_ASSETS !== 'undefined' && UR_ASSETS[k] ? '<i class="ic"><img alt="" src="data:image/png;base64,' + UR_ASSETS[k] + '"></i>' : null;
  const g = (t, c, extra) => '<i class="ic" style="color:#' + c + (extra || '') + '">' + t + '</i>';
  const box = (fill, line, t) => '<i class="ic"><b style="width:13px;height:11px;background:#' + fill + ';border:1px solid #' + line + ';font-size:8px;line-height:10px;color:#555">' + (t || '') + '</b></i>';
  const tbl = (head, col1) => '<i class="ic"><svg viewBox="0 0 14 12" width="14" height="12"><rect x=".5" y=".5" width="13" height="11" fill="#fff" stroke="#9aa3ad"/>' + (head ? '<rect x=".5" y=".5" width="13" height="3.5" fill="#' + head + '"/>' : '') + (col1 ? '<rect x="9.5" y="4" width="4" height="7.5" fill="#' + col1 + '"/>' : '') + '<path d="M.5 4h13M.5 7.7h13M5 .5v11M9.5 .5v11" stroke="#9aa3ad" stroke-width=".6"/></svg></i>';
  switch (id) {
    case 'chapter': return g('Aa', 'A6A6A6', ';font-size:9px');
    case 'lesson': case 'header': return g('Aa', '1F71DD', ';font-size:12px');
    case 'topic': return g('Aa', '000000', ';font-size:11px');
    case 'body': return g('≡', '8C8C8C', ';font-size:14px');
    case 'bullets': return g('•≡', '595959', ';font-size:11px');
    case 'cols': return g('••', '595959', ';font-size:11px;letter-spacing:2px');
    case 'pill': return '<i class="ic"><b style="width:14px;height:8px;border-radius:5px;background:#1F71DC"></b></i>';
    case 'qhead': return g('Q≡', 'FF0000', ';font-size:10px');
    case 'question': case 'qqr': return g('Q', 'FF0000', ';font-size:13px');
    case 'qa': return g('Q<span style="color:#1F71DE">✓</span>', 'FF0000', ';font-size:11px');
    case 'mcqV': return '<i class="ic"><svg viewBox="0 0 14 14" width="14" height="14" fill="none" stroke="#FF0000" stroke-width="1.1"><circle cx="11" cy="2.5" r="1.8"/><circle cx="11" cy="7" r="1.8"/><circle cx="11" cy="11.5" r="1.8"/><path d="M1.5 2.5h6M1.5 7h6M1.5 11.5h6" stroke="#595959"/></svg></i>';
    case 'mcq': case 'choice': return g('ⓐ', id === 'mcq' ? 'FF0000' : '595959', ';font-size:13px;font-weight:400');
    case 'answer': return g('✓', '1F71DE', ';font-size:13px');
    case 'correct': return '<i class="ic"><svg viewBox="0 0 14 14" width="15" height="15"><circle cx="7" cy="7" r="6.6" fill="#2070DF"/><text x="7" y="10" text-anchor="middle" font-family="-apple-system,Helvetica,Arial,sans-serif" font-size="10.5" font-weight="400" fill="#fff">a</text></svg></i>';
    case 'figure': return '<i class="ic"><svg viewBox="0 0 14 12" width="14" height="12"><rect x=".5" y=".5" width="13" height="11" rx="1.5" fill="#fff" stroke="#00B050"/><circle cx="4.5" cy="4" r="1.4" fill="#00B050"/><path d="M1 11l4-4 3 3 2-2 3 3" fill="none" stroke="#00B050"/></svg></i>';
    case 'given': return '<i class="ic"><b style="width:12px;height:10px;border:1.5px dashed #808080"></b></i>';
    case 'problem': return g('Q', 'FF0000', ';font-size:13px');
    case 'note': return img('bulb') || box('FBEBD3', 'E69138');
    case 'quote': case 'def': return img('book') || box('E2F0D9', '2AB672');
    case 'other': return img('star') || box('FBEBD3', 'C45A11');
    case 'activity': return img('palette') || box('F1F1F1', 'A6A6A6');
    case 'mars': return img('target') || box('F1F1F1', '595959');
    case 'formula': return '<i class="ic"><svg viewBox="0 0 14 14" width="15" height="15"><rect x=".3" y=".3" width="13.4" height="13.4" rx="3" fill="#2E75B6"/><path d="M10 3.4H4.2l3.2 3.6-3.2 3.6H10" fill="none" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></i>';
    case 'practice': return img('trophy') || box('EBE6FE', '7030A0');
    case 'cancelled': return '<i class="ic"><svg viewBox="0 0 16 14" width="16" height="14"><path d="M8 1.2L15 13H1z" fill="#FFC000" stroke="#E0A800" stroke-width=".8" stroke-linejoin="round"/><path d="M8 5.2v3.9" stroke="#1a1a1a" stroke-width="1.6" stroke-linecap="round"/><circle cx="8" cy="11.1" r=".95" fill="#1a1a1a"/></svg></i>';
    case 'table': return tbl('CAD0D8');
    case 'tableMain': return tbl('CAD0D8', 'DAE0E9');
    case 'tablePlain': return tbl(null);
    case 'solTable': return tbl('CAD0D8', null);
    case 'ab': return '';   // the name itself is the sign: «ⓐ ⓑ»
    case 'qrlink': return '<i class="ic"><svg viewBox="0 0 14 14" width="14" height="14"><rect x=".5" y=".5" width="13" height="13" rx="2.5" fill="#1F71DE"/><rect x="2.5" y="2.5" width="9" height="7.5" rx="1" fill="#fff"/><path d="M3.5 3.5h2v2h-2zM8.5 3.5h2v2h-2zM3.5 7h2v2h-2zM7 6.8h1v1h-1zM9 8h1.5v1H9z" fill="#222"/><path d="M3 11.6h8" stroke="#fff" stroke-width=".9"/></svg></i>';
    case 'graph': return '<i class="ic"><svg viewBox="0 0 14 12" width="14" height="12"><path d="M1 0v12M4.5 0v12M8 0v12M11.5 0v12M0 1.5h14M0 5h14M0 8.5h14M0 11.5h14" stroke="#c5ccd6" stroke-width=".6"/><path d="M1 10.5C4 10 5 2 8 2s3.5 5 5 6" fill="none" stroke="#1F71DE" stroke-width="1.4" stroke-linecap="round"/></svg></i>';
    case 'grid': return '<i class="ic"><svg viewBox="0 0 14 12" width="14" height="12"><path d="M1 0v12M4.5 0v12M8 0v12M11.5 0v12M0 1.5h14M0 5h14M0 8.5h14M0 11.5h14" stroke="#9aa3ad" stroke-width=".7"/><path d="M8 0v12M0 5h14" stroke="#333" stroke-width=".9"/></svg></i>';
  }
  return '<i style="background:#' + (col || 'FFFFFF') + '"></i>';
}
// one card per group: a coloured strip and the title; every group always open
const GROUP_COL = { 'عناوين': '1F71DD', 'نصوص': '8C8C8C', 'أسئلة': 'FF0000', 'الحلول': '1F71DE', 'إطارات': 'E69138', 'إطارات حول المحدد': 'E69138', 'جداول': '44546A', 'تنسيق جدول محدد': '44546A', 'نهاية الدرس': '7030A0', 'أخرى': '00B050' };
function groupCard(where, key, title) {
  const d = document.createElement('div'); d.className = 'tgroup';
  d.innerHTML = '<div class="gt"><span class="bar" style="background:#' + (GROUP_COL[key] || '9aa3ad') + '"></span><span>' + title + '</span></div>';
  const grid = document.createElement('div'); grid.className = 'tgrid'; d.appendChild(grid);
  return { d, grid };
}
try { localStorage.removeItem('ur-closed'); } catch (e) {}
function buildTemplateList() {
  const host = $('tpls'); host.innerHTML = '';
  const groups = [];
  for (const t of URT.LIST) { let g = groups.find(x => x.name === t.g); if (!g) groups.push(g = { name: t.g, items: [] }); g.items.push(t); }
  for (const g of groups) {
    const { d, grid } = groupCard('ins', g.name, L(g.name, URT.GROUP_EN[g.name] || g.name));
    for (const t of g.items) {
      const b = document.createElement('button'); b.type = 'button'; b.className = 'tbtn'; b.dataset.id = t.id;
      b.title = L(t.full || t.name, t.en); b.innerHTML = iconHtml(t.id, SW[t.id]) + '<span>' + L(t.name, t.en) + '</span>'; tint(b, t.id);
      grid.appendChild(b);
    }
    d.appendChild(grid); host.appendChild(d);
  }
}

// the graph grid chosen while pictures are selected: the grid takes each picture's place and size
// the empty grid drawn over one object (thin filled rectangles grouped), the object parked aside as its hidden original:
// the pair behaves like a solution («إظهار الحلول» → the object, «إخفاء الحلول» → the grid)
async function gridOverShape(ctx, sh, p, W, made) {
  const base = (p.name || '').replace(/ #ursel$/, '').replace(/^UULA Sol(Hidden|View) /, '').replace(/^UULA:[a-z]+\s*/, '') || 'Graph';
  const x0 = p.left, y0 = p.top, w = p.width, h = p.height, step = 13.39 * 1.75, T = 0.75;
  const lines = [];
  // each line a very thin filled rectangle (a zero-wide line is tilted by PowerPoint for Mac)
  const add = (l, t, ww, hh) => { const ln = sh.addGeometricShape('Rectangle', { left: l, top: t, width: ww, height: hh }); ln.fill.setSolidColor('#A0A8B4'); ln.lineFormat.visible = false; ln.name = 'UULA Graph Grid Line'; lines.push(ln); made.push(ln); };
  for (let x = (w / 2) % step; x <= w + 0.01; x += step) add(x0 + Math.min(x, w - T), y0, T, h);
  for (let y = (h / 2) % step; y <= h + 0.01; y += step) add(x0, y0 + Math.min(y, h - T), w, T);
  lines.forEach(l => l.load('id')); await ctx.sync();
  const g = sh.addGroup(lines.map(l => l.id)); g.name = 'UULA SolView UULA:graph ' + base; g.load('id');
  await ctx.sync();
  p.name = 'UULA SolHidden UULA:graph ' + base; p.left = x0 + 3 * W;
  await ctx.sync();
  return g;
}
async function gridSelected(id) {
  const n = await run(async ctx => {
    const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/name,items/type'); await ctx.sync();
    const pics = shapes.items.filter(s => !['TextBox', 'GeometricShape', 'Table', 'Line'].includes(s.type));
    if (!pics.length) return 0;
    pics.forEach(s => { s.name = (s.name || '').replace(/ #ursel$/, '') + ' #ursel'; });
    await ctx.sync();
    return pics.length;
  }).catch(() => 0);
  if (!n) return 0;
  // in place: the grid drawn from lines and grouped; the picture is kept, linked to the grid as a hidden solution
  // («إظهار الحلول» brings the picture back, «إخفاء الحلول» shows the grid again)
  const W = (await getSize()).W;
  const quick = await run(async ctx => {
    const sh = ctx.presentation.slides.getItem(id).shapes;
    sh.load('items/id,items/name,items/left,items/top,items/width,items/height'); await ctx.sync();
    const pics = sh.items.filter(s => / #ursel$/.test(s.name || ''));
    const made = [];
    try {
      const sel = [];
      for (const p of pics) { const g = await gridOverShape(ctx, sh, p, W, made); sel.push(g); }
      try { ctx.presentation.slides.getItem(id).setSelectedShapes(sel.map(g => g.id)); await ctx.sync(); } catch (e) {}
      return pics.length;
    } catch (e) {
      made.forEach(l => { try { l.delete(); } catch (x) {} }); await ctx.sync().catch(() => {});
      return null;
    }
  }).catch(() => null);
  if (quick) return quick;
  const pkg = await UR.openPackage(await exportSlide(id));
  const r = UR.gridOver(pkg);
  if (!r.count) { await clearSel(id); return 0; }
  const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, null), id);
  await deleteSlide(id);
  await run(async ctx => {
    ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
    const sh = ctx.presentation.slides.getItem(newId).shapes; sh.load('items/id,items/name'); await ctx.sync();
    const sel = sh.items.filter(s => / #ursel$/.test(s.name || ''));
    sel.forEach(s => { s.name = s.name.replace(/ #ursel$/, ''); });
    await ctx.sync();
    try { ctx.presentation.slides.getItem(newId).setSelectedShapes(sel.map(s => s.id)); await ctx.sync(); } catch (e) {}
  }).catch(() => {});
  return r.count;
}
// a container chosen while objects are selected: it goes around them (without the template's own text)
async function wrapSelected(id, tid) {
  const n = await run(async ctx => {
    const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/name'); await ctx.sync();
    if (!shapes.items.length) return 0;
    shapes.items.forEach(s => { s.name = (s.name || '').replace(/ #ursel$/, '') + ' #ursel'; });
    await ctx.sync();
    return shapes.items.length;
  }).catch(() => 0);
  if (!n) return false;
  const pkg = await UR.openPackage(await exportSlide(id));
  let r;
  try { r = UR.wrapBox(pkg, UR.WRAP[tid] || tid); }
  catch (e) { await clearSel(id); throw e; }
  const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), id);
  await deleteSlide(id);
  await run(async ctx => {
    const sh = ctx.presentation.slides.getItem(newId).shapes;
    sh.load('items/name,items/left,items/top,items/width,items/height'); await ctx.sync();
    const byIdx = {}; sh.items.forEach((s, k) => { const m = / #ur(\d+)$/.exec(s.name || ''); if (m) byIdx[+m[1]] = k; });
    const sel = r.sel.map(i => byIdx[i]).filter(k => k != null);
    // real sizes: the container fits around what it holds
    const y0 = Math.min(...sel.map(k => sh.items[k].top)), y1 = Math.max(...sel.map(k => sh.items[k].top + sh.items[k].height));
    const G = 13.39;
    const bk = byIdx[r.box];
    if (bk != null) { const b = sh.items[bk]; b.top = r.small ? y0 - 4 : y0 - G; b.height = Math.max(r.minH / PT, r.small ? (y1 - y0) + 8 : (y1 - y0) + 3 * G); }
    if (r.icon != null && byIdx[r.icon] != null && !r.small) sh.items[byIdx[r.icon]].top = y0 - G + G - 2;
    if (r.line != null && byIdx[r.line] != null) { const l = sh.items[byIdx[r.line]]; l.top = y0 + 17; l.height = Math.max(G, (y1 - y0) - 17); }
    await ctx.sync();
    sh.items.forEach(s => s.load('left,top,width,height')); await ctx.sync();
    // on the page grid, and what is under it moves with two grid steps
    const H = r.H / PT;
    const rects = sh.items.map((s, k) => ({ key: k, x: s.left * PT, y: s.top * PT, w: s.width * PT, h: s.height * PT,
      fixed: /^UULA (Footer|Watermark)|^CustomPageNumber_/.test((s.name || '').replace(/ #ur\d+$/, '')) || s.top > H - 54 || s.left + s.width <= 0 || s.left >= r.W / PT }));
    const group = [...sel, bk, r.icon != null ? byIdx[r.icon] : null, r.line != null ? byIdx[r.line] : null].filter(k => k != null);
    const mv = UR.reflow(rects, group, { keepX: true });
    for (const [k, m] of Object.entries(mv)) if (m.y != null) sh.items[+k].top = m.y / PT;
    await ctx.sync();
    ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
  });
  await unmark(newId);
  return true;
}
async function insertTemplate(tid) {
  if (busy || !supported) return;
  busy = true; lock(); tsay('');
  await snapOp([await currentId().catch(() => null)].filter(Boolean));
  try {
    const id = await currentId();
    if (!id) { tsay(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
    if (tid === 'cancelled') {
      const r = await suspendSelected().catch(() => ({ n: 0 }));
      if (r.n) { tsay(suspendSay(r)); return; }
    }
    if (tid === 'grid') { const n = await gridSelected(id); if (n) { tsay(n > 1 ? L('وُضعت الشبكة البيانية مكان ' + n + ' صور، كلٌّ بمقاسها.', 'The graph grid replaced ' + n + ' pictures, each at its size.') : L('وُضعت الشبكة البيانية مكان الصورة وبمقاسها.', 'The graph grid replaced the picture, at its size.')); return; } }
    if (UR.WRAP[tid] && await wrapSelected(id, tid)) { tsay(L('وُضعت الحاوية حول العناصر المحددة.', 'The container was put around the selected objects.')); return; }
    const pkg = await UR.openPackage(await exportSlide(id));
    const r = URT.apply(pkg, tid);
    const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), id);
    await deleteSlide(id);
    // real text heights, then place the template again
    await run(async ctx => {
      const { map } = await shapesByMark(ctx, newId);
      const idx = Array.from({ length: r.count }, (_, k) => r.first + k);
      const texts = idx.filter((i, k) => r.built.shapes[k].t === 'text' && map[i]);
      // (vertical options keep «no wrapping»: each box grows with its own answer)
      texts.forEach(i => { try { map[i].textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; if (tid !== 'mcqV') map[i].textFrame.wordWrap = true; } catch (e) {} });
      await ctx.sync();
      texts.forEach(i => map[i].load('height'));
      await ctx.sync();
      const heights = r.built.shapes.map((m, k) => { const s = map[r.first + k]; return m.t === 'text' && s ? Math.round(s.height * PT) : null; });
      const pl = URT.place(tid, pkg.slideW, r.top, r.built, heights);
      pl.pos.forEach((p, k) => {
        const s = map[r.first + k]; if (!s || !p) return;
        if (p.x != null) s.left = p.x / PT; if (p.y != null) s.top = p.y / PT;
        if (p.w != null && p.w > 0) s.width = p.w / PT;
        if (p.h != null && p.h > 0 && r.built.shapes[k].t !== 'text') s.height = p.h / PT;
      });
      await ctx.sync();
      // a template that has its own place at the top of the page (the lesson title): what is already on the slide
      // moves down as one piece, under the title with one page-grid line between them — nothing is restyled
      if (URT.spec(tid, pkg.slideW).fixed) {
        const mine = new Set(Object.keys(map).filter(i => +i >= r.first && +i < r.first + r.count).map(i => map[i].id));
        const sh = ctx.presentation.slides.getItem(newId).shapes;
        sh.load('items/id,items/name,items/left,items/top,items/width,items/height'); await ctx.sync();
        const W = pkg.slideW / PT, H = pkg.slideH / PT;
        const rest = sh.items.filter(x => !mine.has(x.id) && !/^UULA (Footer|Watermark)|^CustomPageNumber_/.test((x.name || '').replace(/ #ur\d+$/, '')) && !(x.width > W * 0.85 && x.height > H * 0.6) && x.top < H - 54);
        const onPage = rest.filter(x => x.left < W && x.left + x.width > 0);
        if (onPage.length) {
          const ROWE = 13.39, Y0 = 22.58, want = Y0 + Math.ceil((pl.bottom / PT + ROWE - Y0 - 0.3) / (ROWE / 2)) * (ROWE / 2);
          const dy = want - Math.min(...onPage.map(x => x.top));
          if (dy > 0.3) { rest.forEach(x => { x.top = x.top + dy; }); r.pushed = true; r.pushOver = Math.max(...onPage.map(x => x.top + x.height)) > H - 54 + 1; await ctx.sync(); }
        }
      }
      r.fitsReal = pl.bottom <= pkg.slideH - UR.FOOT;
      ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
      try { ctx.presentation.slides.getItem(newId).setSelectedShapes([]); await ctx.sync(); } catch (e) {}
    });
    await unmark(newId);
    const nm = tplName(tid); tsay(r.fitsReal === false ? L('تمت إضافة «' + nm[0] + '»، لكنه تجاوز أسفل الصفحة.', '«' + nm[1] + '» was added, but it goes past the bottom of the page.') : r.pushed ? (r.pushOver ? L('تمت إضافة «' + nm[0] + '» أعلى الصفحة ونزل المحتوى تحته، لكن المحتوى تجاوز أسفل الصفحة.', '«' + nm[1] + '» was added at the top and the content moved down, but it now goes past the bottom of the page.') : L('تمت إضافة «' + nm[0] + '» أعلى الصفحة، ونزل المحتوى تحته.', '«' + nm[1] + '» was added at the top; the content moved down under it.')) : L('تمت إضافة «' + nm[0] + '» في الأسفل.', '«' + nm[1] + '» was added at the bottom.'));
  } catch (e) { tsay(L('تعذّرت الإضافة: ', 'Could not insert: ') + errText(e)); }
  finally { busy = false; lock(); }
}

// ---------- one object at a time ----------
const ROLES = [
  { g: 'عناوين', ge: 'Headings', items: [['chapter', 'عنوان الوحدة', 'A6A6A6'], ['lesson', 'عنوان الدرس', '1F71DD'], ['topic', 'عنوان موضوع', '000000']] },
  { g: 'نصوص', ge: 'Text', items: [['body', 'فقرة شرح', 'FFFFFF'], ['bullets', 'نقاط مرقّمة', 'FFFFFF']] },
  { g: 'أسئلة', ge: 'Questions', items: [['qhead', 'عنوان الأسئلة', 'FF0000'], ['question', 'نص السؤال', 'FF0000'], ['choice', 'خيارات السؤال', 'FFFFFF']] },
  { g: 'الحلول', ge: 'Solutions', items: [['answer', 'حل السؤال', '1F71DE'], ['correct', 'الخيار الصحيح', '2070DF']] },
  { g: 'إطارات حول المحدد', ge: 'Frames around the selection', items: [['note', 'ملاحظة', 'FBEBD3'], ['quote', 'نص مقتطف', 'E2F0D9'], ['def', 'تعريف', 'E2F0D9'], ['formula', 'قوانين', 'D3E4F5'], ['other', 'صيغة أخرى', 'FBEBD3'], ['activity', 'نشاط', 'F1F1F1'], ['mars', 'من المريخ', 'F1F1F1'], ['cancelled', 'معلّق', 'FFFFFF']] },
  { g: 'تنسيق جدول محدد', ge: 'Selected table', items: [['table', 'جدول برأس', 'CAD0D8'], ['tableMain', 'رأس وعمود رئيسي', 'DAE0E9'], ['tablePlain', 'جدول بلا رأس', 'F1F1F1']] }
];
// the same buttons beside «Hide solutions» only mark: the colour stays as it is (colours are set in «Format»)
const SOL_TIP = {
  answer: ['● ← حل يُخفى مع «إخفاء الحلول» · الجدول يصير أزرق', '● → a solution hidden by «Hide solutions» · a table turns blue'],
  correct: ['● حدّد الخيار الصحيح ← حرفه في دائرة مملوءة', '● Select the correct choice → its letter in a filled circle'],
  given: ['● ← جزء من السؤال: يبقى ظاهرًا', '● → part of the question: stays visible'],
  graph: ['● ← حل بياني: شبكة فارغة مكانه عند الإخفاء', '● → graph solution: an empty grid in its place when hidden']
};
// what each button does (shown when the pointer rests on it)
const ROLE_TIP = {
  chapter: ['اسم الوحدة أعلى الصفحة (رمادي صغير)', 'Unit name at the top (small gray)'],
  lesson: ['اسم الدرس أعلى الصفحة (أزرق كبير)', 'Lesson name at the top (large blue)'],
  topic: ['عنوان موضوع داخل الدرس (أسود عريض)', 'A topic heading inside the lesson (bold black)'],
  body: ['فقرة شرح عادية بخط المذكرة', 'A normal paragraph in the booklet font'],
  bullets: ['قائمة نقاط، كل سطر نقطة', 'A bullet list, one point per line'],
  qhead: ['عنوان قسم الأسئلة (أحمر عريض)', 'Heading of a questions section (bold red)'],
  question: ['نص السؤال مع علامة السؤال الحمراء', 'Question text with the red question mark'],
  choice: ['يحوّل المحدد إلى خيارات سؤال اختيار من متعدد (ⓐ ⓑ ⓒ ⓓ)', 'Turns the selection into the options of a multiple-choice question (ⓐ ⓑ ⓒ ⓓ)'],
  answer: ['يعلّم المحدد كحل ويلوّن نصه بالأزرق، فيُخفى عند «إخفاء الحلول». لا يتحرك شيء في الشريحة.', 'Marks the selection as a solution and colours its text blue, so «Hide solutions» hides it. Nothing on the slide moves.'],
  correct: ['حرف الخيار الصحيح يصير أبيض على تظليل أزرق في مكانه (واضح حتى في الطباعة بالأبيض والأسود)، ويعود حرفًا عاديًا عند «إخفاء الحلول». مع تحديد مربع: هذا المربع فقط. بدون تحديد: كل الخيارات المميزة سابقًا كصحيحة (أو حرفها أزرق) في الشرائح المحددة دفعة واحدة.', 'The correct choice\'s letter turns white on a blue highlight in place (clear even in black-and-white print) and turns back to a plain letter when solutions are hidden. With a box selected: that box only. Nothing selected: every choice marked correct before (or with a blue letter) on the selected slides, at once.'],
  given: ['شكل أو نص من السؤال نفسه: يبقى ظاهرًا عند «إخفاء الحلول» حتى لو كان أزرق', 'A shape or text that belongs to the question: it stays when solutions are hidden, even if blue'],
  note: ['إطار ملاحظة حول ما حددته (مربع أو أكثر)', 'A note frame around the selection'],
  quote: ['إطار نص مقتطف حول ما حددته', 'A quote frame around the selection'],
  def: ['إطار تعريف حول ما حددته', 'A definition frame around the selection'],
  formula: ['إطار قوانين حول ما حددته', 'A formulas frame around the selection'],
  other: ['إطار صغير «صيغة أخرى» حول ما حددته', 'A small «other form» frame around the selection'],
  activity: ['إطار نشاط حول ما حددته', 'An activity frame around the selection'],
  mars: ['إطار «من المريخ» حول ما حددته', 'A «from Mars» frame around the selection'],
  cancelled: ['إطار «معلّق» حول ما حددته: إطار أصفر، والمحتوى تحته باهت، وفي وسطه «مـعـلـق» مع مثلث التحذير.', 'A «suspended» frame around the selection: yellow frame, the content faded under it, «suspended» with the warning triangle in the middle.'],
  table: ['حدّد جدولًا: صف رأس رمادي وصفوف متبادلة', 'Select a table: gray header row, alternating rows'],
  tableMain: ['حدّد جدولًا: صف رأس وعمود أول رئيسي', 'Select a table: header row and a main first column'],
  tablePlain: ['حدّد جدولًا: بلا صف رأس، صفوف متبادلة', 'Select a table: no header row, alternating rows']
};
const FRAME_ROLES = ['note', 'quote', 'def', 'formula', 'other', 'activity', 'mars'];
const ROLE_EN = { chapter: 'Unit title', lesson: 'Lesson title', topic: 'Topic heading', body: 'Paragraph', bullets: 'Bullet list', qhead: 'Questions heading', question: 'Question text', answer: 'Solution', choice: 'Choices',
  correct: 'Correct choice', given: 'Part of question', graph: 'Graph solution', note: 'Note', quote: 'Quote', def: 'Definition', formula: 'Formulas', other: 'Other form', activity: 'Activity', mars: 'From Mars', cancelled: 'Suspended', table: 'With header', tableMain: 'Header + main column', tablePlain: 'No header' };
function mergedBtn(it) {
  const b = document.createElement('button'); b.type = 'button'; b.className = 'tbtn'; b.dataset.item = it.id;
  if (it.role) b.dataset.role = it.role; if (it.tpl) b.dataset.id = it.tpl;
  b.title = mergedTip(it); b.innerHTML = iconHtml(it.tpl === 'header' ? 'lesson' : it.id, it.col) + '<span>' + L(it.ar, it.en) + '</span>';
  tint(b, it.id); return b;
}
// direction of multiple-choice options: auto / right-to-left / left-to-right (kept in the file)
let mcqDir = 'auto';
function dirSeg() {
  const d = document.createElement('div'); d.className = 'seg seg3 icoseg dirseg';
  const ic = { auto: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2.5l1.3 3.2 3.2 1.3-3.2 1.3L8 11.5 6.7 8.3 3.5 7l3.2-1.3z"/></svg>',
    rtl: '<svg viewBox="0 0 18 14" width="17" height="13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="14" cy="7" r="3"/><path d="M9 7H2M4.5 4.5L2 7l2.5 2.5"/></svg>',
    ltr: '<svg viewBox="0 0 18 14" width="17" height="13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="4" cy="7" r="3"/><path d="M9 7h7M13.5 4.5L16 7l-2.5 2.5"/></svg>' };
  const tip = { auto: ['اتجاه الخيارات تلقائي: سؤال عربي ← الخيارات من اليمين، سؤال إنجليزي → من اليسار', 'Options direction automatic: Arabic question → from the right, English question → from the left'],
    rtl: ['الخيارات من اليمين إلى اليسار: ⓐ في اليمين', 'Options right-to-left: ⓐ on the right'], ltr: ['الخيارات من اليسار إلى اليمين: ⓐ في اليسار (كالإنجليزي)', 'Options left-to-right: ⓐ on the left (as in English)'] };
  for (const v of ['auto', 'rtl', 'ltr']) {
    const b = document.createElement('button'); b.type = 'button'; b.dataset.v = v; b.title = L(tip[v][0], tip[v][1]);
    b.setAttribute('aria-pressed', String(v === mcqDir)); b.innerHTML = ic[v]; d.appendChild(b);
  }
  d.addEventListener('click', async e => {
    e.stopPropagation(); const b = e.target.closest('button'); if (!b) return;
    setMcqDir(b.dataset.v, true);
    // options selected on the slide: they turn to this direction now (the same button for inserting and for formatting)
    if (b.dataset.v !== 'auto') { await selCheck(); if (selNow) flipSelected(b.dataset.v); }
  });
  return d;
}
// «ⓐ ⓑ»: the answer circles as text — two text boxes «ⓐ» and «ⓑ», as far apart as four spaces. Nothing selected: inserted under the
// last object; drawn shapes selected: replaced by it in their place (one box per line of shapes)
const AB_TEXT = 'ⓐ    ⓑ';
async function abPress() {
  if (busy || !supported) return;
  busy = true; lock();
  const id = await currentId().catch(() => null);
  await snapOp(id ? [id] : []);
  try {
    if (!id) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
    const top = await (async () => { const pkg = await UR.openPackage(await exportSlide(id)); return URT.freeTop(pkg, 1).top / PT; })().catch(() => 60);
    const n = await run(async ctx => {
      const ss = ctx.presentation.getSelectedShapes(); ss.load('items/id,items/left,items/top,items/width,items/height'); await ctx.sync();
      const sh = ctx.presentation.slides.getItem(id).shapes;
      // the places: shapes that overlap in height make a line; in a line, shapes close to each other make one place
      // (two pairs side by side → two places); each place gets its own box where the shapes were
      const lines = [];
      for (const s of [...ss.items].sort((a, b) => a.top - b.top)) {
        const ln = lines.find(l => l.some(o => Math.min(o.top + o.height, s.top + s.height) - Math.max(o.top, s.top) > Math.min(o.height, s.height) * 0.3));
        if (ln) ln.push(s); else lines.push([s]);
      }
      const spots = [];
      for (const ln of lines) {
        ln.sort((a, b) => a.left - b.left);
        let cur = null;
        for (const s of ln) {
          const near = cur && s.left - cur.r < Math.max(24, 1.5 * Math.max(...cur.shapes.map(o => o.width), s.width)) && cur.shapes.length < 2;
          if (near) { cur.l = Math.min(cur.l, s.left); cur.r = Math.max(cur.r, s.left + s.width); cur.t = Math.min(cur.t, s.top); cur.b = Math.max(cur.b, s.top + s.height); cur.shapes.push(s); }
          else { cur = { l: s.left, r: s.left + s.width, t: s.top, b: s.top + s.height, shapes: [s] }; spots.push(cur); }
        }
      }
      if (!spots.length) spots.push({ l: 36, r: 110, t: top, b: top + 20, shapes: [] });
      const made = [];
      // two text boxes, «ⓐ» and «ⓑ», with the same gap as «ⓐ    ⓑ» (four spaces) in one box: the gap is measured once
      // from a hidden sample (the whole pair, minus the two letters alone)
      const style = t => {
        try {
          const tf = t.textFrame; tf.wordWrap = false; tf.autoSizeSetting = 'AutoSizeShapeToFitText';
          tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0; tf.verticalAlignment = 'Middle';
          tf.textRange.font.size = 12; tf.textRange.font.color = '#000000'; tf.textRange.paragraphFormat.horizontalAlignment = 'Left';
        } catch (e) {}
        return t;
      };
      const probe = [AB_TEXT, 'ⓐ', 'ⓑ'].map(x => style(sh.addTextBox(x, { left: -400, top: -400, width: 80, height: 20 })));
      await ctx.sync();
      probe.forEach(t => t.load('width')); await ctx.sync();
      let gap = probe[0].width - probe[1].width - probe[2].width;
      if (!(gap > 2 && gap < 60)) gap = 13;
      probe.forEach(t => t.delete());
      for (const p of spots) {
        const h = 20, y = (p.t + p.b) / 2 - h / 2;
        const a = style(sh.addTextBox('ⓐ', { left: p.l, top: y, width: 14, height: h })); a.name = 'UULA ab a';
        await ctx.sync(); a.load('left,width'); await ctx.sync();
        const b = style(sh.addTextBox('ⓑ', { left: a.left + a.width + gap, top: y, width: 14, height: h })); b.name = 'UULA ab b';
        made.push(a, b);
      }
      await ctx.sync();
      spots.forEach(p => p.shapes.forEach(s => s.delete()));
      made.forEach(t => t.load('id')); await ctx.sync();
      try { ctx.presentation.slides.getItem(id).setSelectedShapes(made.map(t => t.id)); await ctx.sync(); } catch (e) {}
      return { boxes: made.length / 2, replaced: spots.reduce((k, p) => k + p.shapes.length, 0) };
    });
    say(n.replaced ? L('استُبدلت الأشكال المحددة بمربعي «ⓐ» و«ⓑ»' + (n.boxes > 1 ? ' (' + n.boxes + ' أزواج)' : '') + '.', 'The selected shapes were replaced by the «ⓐ» and «ⓑ» boxes' + (n.boxes > 1 ? ' (' + n.boxes + ' pairs)' : '') + '.') : L('تمت إضافة مربعي «ⓐ» و«ⓑ» في الأسفل.', 'The «ⓐ» and «ⓑ» boxes were added at the bottom.'));
  } catch (e) { say(L('تعذّر: ', 'Could not: ') + errText(e)); }
  finally { busy = false; lock(); }
}
// «QR برابط»: the QR picture becomes the page's small QR and opens the address written beside it; the address goes
async function qrPress() {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  const ids = await selIds().catch(() => []);
  await snapOp(ids);
  try {
    if (!ids.length) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
    // one slide: the selected picture / address count; several slides: every address with its nearest picture
    if (ids.length === 1) await run(async ctx => {
      const ss = ctx.presentation.getSelectedShapes(); ss.load('items/name'); await ctx.sync();
      ss.items.forEach(x => { x.name = (x.name || '').replace(/ #ursel$/, '') + ' #ursel'; }); await ctx.sync();
    }).catch(() => {});
    let pics = 0, slides = 0, noUrl = 0, noPic = 0;
    const keep = [];
    for (let k = 0; k < ids.length; k++) {
      const id = ids[k];
      prog(k / ids.length); if (ids.length > 1) say(L('QR برابط ', 'Linked QR ') + (k + 1) + L(' من ', ' of ') + ids.length + '…');
      try {
        const pkg = await UR.openPackage(await exportSlide(id));
        const r = UR.qrLink(pkg);
        if (!r.count) { await clearSel(id); if (!r.urls) noUrl++; else noPic++; keep.push(id); continue; }
        const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), id);
        await deleteSlide(id);
        await clearSel(newId); await unmark(newId);
        keep.push(newId); pics += r.count; slides++;
      } catch (e) { keep.push(id); line('err', '✗ ' + esc(errText(e))); }
    }
    prog(null);
    await run(async ctx => { ctx.presentation.setSelectedSlides(keep); await ctx.sync(); }).catch(() => {});
    if (!pics) {
      say(noPic ? L('لم أجد صورة QR بجانب الرابط.', 'No QR picture beside the address.') : L('لم أجد رابطًا مكتوبًا ' + (ids.length > 1 ? 'في الشرائح المحددة' : 'في الشريحة') + ' (يبدأ بـ https:// أو www).', 'No written address on the ' + (ids.length > 1 ? 'selected slides' : 'slide') + ' (starting with https:// or www).'));
      return;
    }
    if (ids.length > 1 && slides < ids.length) line('', L('• شرائح بلا رابط مكتوب أو بلا صورة QR بقيت كما هي: ', '• Slides with no written address or QR picture stayed as they were: ') + (ids.length - slides) + '.');
    say(ids.length > 1 ? L('رُبطت ' + pics + ' صور QR بروابطها في ', pics + ' QR pictures were linked in ') + (lang !== 'en' && slides === 2 ? 'شريحتين' : cnt(slides)) + L('، وحُذفت الروابط المكتوبة.', '; the written addresses were removed.')
      : pics > 1 ? L('رُبطت ' + pics + ' صور QR بروابطها وحُذفت الروابط المكتوبة.', pics + ' QR pictures were linked and the written addresses removed.')
      : L('صُغّرت صورة الـQR ونُقلت إلى الهامش الأيسر وحافتها العلوية في مكانها، وصارت تفتح الرابط عند الضغط عليها، وحُذف الرابط المكتوب.', 'The QR picture was resized onto the left margin, its top edge kept; a click on it opens the address, and the written address was removed.'));
  } catch (e) { prog(null); say(L('تعذّر: ', 'Could not: ') + errText(e)); }
  finally { busy = false; lock(); }
}
// how many text boxes are selected now (for the options buttons)
async function selCount() {
  // an equation box can be reported as another type: every selected object counts except pictures, tables, lines and groups
  return run(async ctx => { const s = ctx.presentation.getSelectedShapes(); s.load('items/type'); await ctx.sync(); return s.items.filter(x => !['Image', 'Table', 'Line', 'Group', 'Chart', 'SmartArt', 'Media'].includes(x.type)).length; }).catch(() => 0);
}
// several selected text boxes become options: in one row, or one under the other; they stay selected afterwards
async function arrangeSelected(layout) {
  if (busy || !supported) return;
  busy = true; lock();
  await snapOp([await currentId().catch(() => null)].filter(Boolean));
  let slideId = null;
  try {
    slideId = await run(async ctx => {
      const slides = ctx.presentation.getSelectedSlides(); slides.load('items/id');
      const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/name'); await ctx.sync();
      if (!slides.items.length || !shapes.items.length) return null;
      shapes.items.forEach(x => { x.name = (x.name || '').replace(/ #ursel$/, '') + ' #ursel'; });
      await ctx.sync(); return slides.items[0].id;
    });
    if (!slideId) return;
    const pkg = await UR.openPackage(await exportSlide(slideId));
    const r = UR.arrangeChoices(pkg, layout, mcqDir);
    if (!r.count) { await clearSel(slideId); say(L('حدّد مربعَي نص أو أكثر.', 'Select two or more text boxes.')); return; }
    let newId = slideId;
    if (!(await UQ.apply(slideId, pkg.files.slide, r.xml, exportSlide).catch(() => false))) {
      newId = await insertAfter(await UR.buildPackage(pkg, r.xml, null), slideId);
      await deleteSlide(slideId);
    }
    await run(async ctx => {
      ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
      const sh = ctx.presentation.slides.getItem(newId).shapes; sh.load('items/id,items/name'); await ctx.sync();
      const sel = sh.items.filter(x => / #ursel$/.test(x.name || ''));
      // one under the other: by their real heights, each starting on the page grid right after the one above
      // (the UULA rule for options that follow each other: no empty line between them)
      if (layout !== 'row' && sel.length > 1) {
        // each box as wide as its text and a little more (no wrapping; it grows when more is typed), on the right
        // edge of the page in right-to-left, on the left edge in left-to-right
        // the real widths of the texts (no wrapping), then all the boxes as wide as the widest and a little more
        sel.forEach(x => { try { const tf = x.textFrame; tf.wordWrap = false; tf.autoSizeSetting = 'AutoSizeShapeToFitText'; } catch (e) {} });
        await ctx.sync();
        sel.forEach(x => x.load('left,top,width,height')); await ctx.sync();
        const { W: SW } = await getSize(), Mx = 36;
        const bw = Math.min(SW - 2 * Mx, Math.max(...sel.map(x => x.width)) + 8);
        sel.forEach(x => { try { x.textFrame.wordWrap = true; } catch (e) {} x.width = bw; x.left = r.dir === 'rtl' ? SW - Mx - bw : Mx; });
        await ctx.sync();
        sel.forEach(x => x.load('top,height')); await ctx.sync();
        const Y0 = 22.58, G = 6.695, snap = y => Y0 + Math.max(0, Math.ceil((y - Y0 - 0.3) / G)) * G;
        const list = [...sel].sort((a, b) => a.top - b.top);
        let y = list[0].top;
        for (const x of list) { x.top = y; y = snap(y + x.height); }
        await ctx.sync();
      }
      sel.forEach(x => { x.name = x.name.replace(/ #ursel$/, ''); }); await ctx.sync();
      try { ctx.presentation.slides.getItem(newId).setSelectedShapes(sel.map(x => x.id)); await ctx.sync(); } catch (e) {}
    }).catch(() => {});
    slideId = null;
    say((layout === 'row' ? L('رُتّبت ' + r.count + ' خيارات أفقيًا', r.count + ' options arranged in a row') : L('رُتّبت ' + r.count + ' خيارات عموديًا', r.count + ' options arranged in a column'))
      + (r.dir === 'rtl' ? L('، من اليمين لليسار.', ', right-to-left.') : L('، من اليسار لليمين.', ', left-to-right.')));
  } catch (e) { if (slideId) await clearSel(slideId); say(L('تعذّر ترتيب الخيارات: ', 'Could not arrange the options: ') + errText(e)); }
  finally { busy = false; lock(); }
}
async function flipSelected(dir) {
  if (busy || !supported) return;
  busy = true; lock();
  await snapOp([await currentId().catch(() => null)].filter(Boolean));
  let slideId = null;
  try {
    slideId = await run(async ctx => {
      const slides = ctx.presentation.getSelectedSlides(); slides.load('items/id');
      const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/name'); await ctx.sync();
      if (!slides.items.length || !shapes.items.length) return null;
      shapes.items.forEach(x => { x.name = (x.name || '').replace(/ #ursel$/, '') + ' #ursel'; });
      await ctx.sync(); return slides.items[0].id;
    });
    if (!slideId) return;
    const pkg = await UR.openPackage(await exportSlide(slideId));
    const r = UR.flipChoices(pkg, dir);
    let newId = slideId;
    if (!(await UQ.apply(slideId, pkg.files.slide, r.xml, exportSlide).catch(() => false))) {
      newId = await insertAfter(await UR.buildPackage(pkg, r.xml, null), slideId);
      await deleteSlide(slideId);
    }
    // the same options stay selected, so the other direction can be pressed again right away
    await run(async ctx => {
      ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
      const sh = ctx.presentation.slides.getItem(newId).shapes; sh.load('items/id,items/name'); await ctx.sync();
      const sel = sh.items.filter(x => / #ursel$/.test(x.name || ''));
      sel.forEach(x => { x.name = x.name.replace(/ #ursel$/, ''); }); await ctx.sync();
      try { ctx.presentation.slides.getItem(newId).setSelectedShapes(sel.map(x => x.id)); await ctx.sync(); } catch (e) {}
    }).catch(() => {});
    slideId = null;
    say(dir === 'rtl' ? L('حُوِّلت الخيارات المحددة إلى الاتجاه من اليمين لليسار.', 'The selected options now run right-to-left.') : L('حُوِّلت الخيارات المحددة إلى الاتجاه من اليسار لليمين.', 'The selected options now run left-to-right.'));
  } catch (e) { if (slideId) await clearSel(slideId); say(L('تعذّر تحويل الاتجاه: ', 'Could not change the direction: ') + errText(e)); }
  finally { busy = false; lock(); }
}
function setMcqDir(v, store) {
  mcqDir = v; URT.setOpt({ mcqDir: v });
  document.querySelectorAll('.dirseg button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === v)));
  if (store) { const st = docSettings(); if (st) { st.set('urMcqDir', v); st.saveAsync(); } }
}
function buildRoleList() {
  const host = $('roles'); host.innerHTML = '';
  for (const g of MERGED) {
    const { d, grid } = groupCard('obj', g.g, L(g.g, g.ge));
    if (g.items) g.items.forEach(it => grid.appendChild(mergedBtn(it)));
    else {
      grid.className = 'subrows';
      for (const r of g.rows) {
        const row = document.createElement('div'); row.className = 'subrow';
        const lab = document.createElement('span'); lab.className = 'sublab'; lab.textContent = r.label ? L(r.label, r.le) : '';
        const tg = document.createElement('div'); tg.className = 'tgrid';
        r.items.forEach(it => tg.appendChild(mergedBtn(it)));
        if (r.dir) tg.appendChild(dirSeg());
        row.appendChild(lab); row.appendChild(tg); grid.appendChild(row);
      }
    }
    d.appendChild(grid); host.appendChild(d);
  }
  selPaint();
}
// ---------- Format & Insert in one list ----------
// a box selected → it becomes the type pressed (role); nothing selected → a new one is added at the bottom (template)
const MERGED = [
  { g: 'عناوين', ge: 'Headings', items: [
    { id: 'lesson', ar: 'عنوان الدرس', en: 'Lesson title', col: '1F71DD', role: 'lesson', tpl: 'header' },
    { id: 'topic', ar: 'عنوان موضوع', en: 'Topic heading', col: '000000', role: 'topic', tpl: 'topic' }] },
  { g: 'نصوص', ge: 'Text', items: [
    { id: 'body', ar: 'فقرة شرح', en: 'Paragraph', col: 'FFFFFF', role: 'body', tpl: 'body' },
    { id: 'bullets', ar: 'نقاط مرقّمة', en: 'Bullet list', col: 'FFFFFF', role: 'bullets', tpl: 'bullets' },
    { id: 'cols', ar: 'نقاط بأعمدة', en: 'Bullet columns', col: 'FFFFFF', tpl: 'cols' },
    { id: 'pill', ar: 'كبسولة تعريف', en: 'Term pill', col: '1F71DC', tpl: 'pill' }] },
  // one card for questions: the general ones, then written (essay) questions with their solution,
  // then multiple choice with the correct choice and the direction of the options
  { g: 'أسئلة', ge: 'Questions', rows: [
    { items: [
      { id: 'qhead', ar: 'عنوان الأسئلة', en: 'Questions heading', col: 'FF0000', role: 'qhead', tpl: 'qhead' },
      { id: 'question', ar: 'سؤال', en: 'Question', col: 'FF0000', role: 'question', tpl: 'question' }] },
    { label: 'مقالي', le: 'Essay', items: [
      { id: 'problem', ar: 'مسألة', en: 'Problem', col: '00B050', tpl: 'problem' },
      { id: 'qqr', ar: 'سؤال مع QR', en: 'Question + QR', col: 'FF0000', tpl: 'qqr' },
      { id: 'answer', ar: 'حل السؤال', en: 'Solution', col: '1F71DE', role: 'answer', tpl: 'answer' },
      { id: 'graph', ar: 'حل بياني', en: 'Graph solution', col: 'A0A8B4', role: 'graph' }] },
    { label: 'اختياري', le: 'Choice', dir: true, items: [
      { id: 'mcq', ar: 'خيارات أفقية', en: 'Options in a row', col: 'FF0000', role: 'choice', tpl: 'mcq' },
      { id: 'mcqV', ar: 'خيارات عمودية', en: 'Options in a column', col: 'FF0000', tpl: 'mcqV' },
      { id: 'correct', ar: 'الخيار الصحيح', en: 'Correct choice', col: '2070DF', role: 'correct' },
      { id: 'ab', ar: 'ⓐ ⓑ', en: 'ⓐ ⓑ circles', col: 'FFFFFF' }] }] },
  { g: 'إطارات', ge: 'Frames', items: [
    { id: 'note', ar: 'ملاحظة', en: 'Note', role: 'note', tpl: 'note' },
    { id: 'quote', ar: 'نص مقتطف', en: 'Quote', role: 'quote', tpl: 'quote' },
    { id: 'def', ar: 'تعريف', en: 'Definition', role: 'def', tpl: 'def' },
    { id: 'formula', ar: 'قوانين', en: 'Formulas', role: 'formula', tpl: 'formula' },
    { id: 'other', ar: 'صيغة أخرى', en: 'Other form', role: 'other', tpl: 'other' },
    { id: 'activity', ar: 'نشاط', en: 'Activity', role: 'activity', tpl: 'activity' },
    { id: 'mars', ar: 'من المريخ', en: 'From Mars', role: 'mars', tpl: 'mars' },
    { id: 'cancelled', ar: 'معلّق', en: 'Suspended', role: 'cancelled', tpl: 'cancelled' }] },
  { g: 'جداول', ge: 'Tables', items: [
    { id: 'table', ar: 'جدول برأس', en: 'Table with header', col: 'CAD0D8', role: 'table', tpl: 'table' },
    { id: 'tableMain', ar: 'رأس وعمود رئيسي', en: 'Header + main column', col: 'DAE0E9', role: 'tableMain', tpl: 'tableMain' },
    { id: 'tablePlain', ar: 'جدول بلا رأس', en: 'No header', col: 'F1F1F1', role: 'tablePlain' },
    { id: 'solTable', ar: 'جدول الحل', en: 'Answer table', col: 'CAD0D8', tpl: 'solTable' },
    { id: 'grid', ar: 'شبكة بيانية', en: 'Graph grid', col: 'D3D8E0', tpl: 'grid' }] },
  { g: 'نهاية الدرس', ge: 'End of lesson', items: [
    { id: 'practice', ar: 'تدرب وتفوق', en: 'Practice & excel', col: 'EBE6FE', tpl: 'practice' },
    { id: 'qrlink', ar: 'QR برابط', en: 'Linked QR', col: '1F71DE' }] }
];
const ALL_ITEMS = MERGED.flatMap(g => g.items || g.rows.flatMap(r => r.items));
const ITEM = Object.fromEntries(ALL_ITEMS.map(x => [x.id, x]));
function tplName(tid) { const it = ALL_ITEMS.find(x => x.tpl === tid); return it ? [it.ar, it.en] : ['قالب', 'block']; }
// the tip of a button: one short line (● = something selected, ○ = nothing selected)
const SHORT_TIP = {
  qrlink: ['QR ← 1.42×1.65 سم عند الهامش الأيسر + رابط بالضغط · يُحذف النص', 'QR → 1.42×1.65 cm at the left margin + click link · text removed'],
  ab: ['○ ← مربعا ⓐ ⓑ · ● دوائر مرسومة ← تُستبدل في مكانها', '○ → ⓐ ⓑ boxes · ● drawn circles → replaced in place'],
  mcq: ['● مربعات ← خيارات في صف (حسب الاتجاه) · ○ ← قالب جديد', '● boxes → options in a row (by direction) · ○ → new block'],
  mcqV: ['● مربعات ← خيارات عمودية (حسب الاتجاه) · ○ ← قالب جديد', '● boxes → options in a column (by direction) · ○ → new block'],
  correct: SOL_TIP.correct, answer: SOL_TIP.answer, given: SOL_TIP.given, graph: SOL_TIP.graph,
  grid: ['● صور ← شبكة بمقاسها · ○ ← شبكة جديدة', '● pictures → grid at their size · ○ → new grid'],
  cancelled: ['● ← إطار معلّق فوقه · ○ ← إطار بعرض الصفحة', '● → suspended frame over it · ○ → page-wide frame'],
  practice: ['＋ «تدرب وتفوق» بعد آخر عنصر', '＋ «Practice & excel» after the last object']
};
function mergedTip(it) {
  const t = SHORT_TIP[it.id];
  if (t) return L(t[0], t[1]);
  if (it.role && it.tpl) return L('● ← يتحوّل إلى «' + it.ar + '» · ○ ← يُضاف جديد', '● → turned into «' + it.en + '» · ○ → a new one is added');
  if (it.role) return L('● حدّد ثم اضغط ← «' + it.ar + '»', '● Select, then press → «' + it.en + '»');
  return L('＋ «' + it.ar + '» تحت آخر عنصر', '＋ «' + it.en + '» under the last object');
}
// what is selected now, shown above the list: the next press converts it, or adds a new one
let selNow = 0;
function selPaint() {
  const h = $('selHint'); if (!h) return;
  h.dataset.on = selNow ? '1' : '';
  h.innerHTML = selNow
    ? '<b>●</b> ' + L('محدد ' + (selNow > 1 ? selNow + ' عناصر' : 'عنصر') + ' — الضغط سيحوّله', (selNow > 1 ? selNow + ' objects' : 'One object') + ' selected — pressing converts it')
    : '<b>○</b> ' + L('لا شيء محدد — الضغط سيضيف جديدًا في الأسفل', 'Nothing selected — pressing adds a new one at the bottom');
}
async function selCheck() {
  if (!supported || isWord) return;
  const n = await run(async ctx => { const s = ctx.presentation.getSelectedShapes(); s.load('items/name'); await ctx.sync(); return s.items.filter(x => !/^UULA (Footer|Watermark)/.test(x.name || '')).length; }).catch(() => 0);
  selNow = n; selPaint();
}
async function mergedPress(id) {
  const it = ITEM[id]; if (!it || busy || !supported) return;
  await selCheck();
  if (id === 'ab') return abPress();
  if (id === 'qrlink') return qrPress();
  if ((id === 'mcq' || id === 'mcqV') && (await selCount()) > 1) return arrangeSelected(id === 'mcq' ? 'row' : 'column');
  if (id === 'mcqV' && selNow === 1) return pinRole('choice');
  if (id === 'cancelled') {
    if (busy) return; busy = true; lock();
    await snapOp([await currentId().catch(() => null)].filter(Boolean));
    try { say(suspendSay(await suspendSelected())); } catch (e) { say(L('تعذّر: ', 'Could not: ') + errText(e)); }
    finally { busy = false; lock(); }
    return;
  }
  if (it.tpl === 'grid' || it.tpl === 'practice' || !it.role) {
    if (it.tpl) return insertTemplate(it.tpl);
  }
  if (id === 'correct' && !selNow) return needCorrect();
  if (selNow && it.role) return pinRole(it.role);
  if (it.tpl) return insertTemplate(it.tpl);
  needSel(it);
}
// a button that needs a selection, pressed with nothing selected: the line above the list says so, clearly, for a moment
let needTimer = 0;
function needSel(it) {
  const msg = L('حدّد أولًا الكائن الذي تريد تحويله إلى «' + it.ar + '»', 'First select the object to turn into «' + it.en + '»');
  say(msg + '.');
  const h = $('selHint'); if (!h) return;
  h.dataset.on = 'need'; h.innerHTML = msg + ' <b>!</b>';
  clearTimeout(needTimer); needTimer = setTimeout(selPaint, 3500);
}
$('roles').addEventListener('click', e => { const b = e.target.closest('.tbtn'); if (b) mergedPress(b.dataset.item); });
// the same solution buttons beside «Hide solutions»
function buildSolRoles() {
  const host = $('solRoles'); if (!host) return; host.innerHTML = '';
  for (const [id, nameAr, col] of [['answer', 'حل السؤال', '1F71DE'], ['graph', 'حل بياني', 'A0A8B4'], ['correct', 'الخيار الصحيح', '2070DF'], ['given', 'جزء من السؤال', 'A6A6A6']]) {
    const name = L(nameAr, ROLE_EN[id]);
    const b = document.createElement('button'); b.type = 'button'; b.className = 'tbtn'; b.dataset.role = id;
    b.title = L(SOL_TIP[id][0], SOL_TIP[id][1]); b.innerHTML = iconHtml(id, col) + '<span>' + name + '</span>'; tint(b, id);
    host.appendChild(b);
  }
}
async function undoRedo(dir) {
  if (busy || !supported || !(dir < 0 ? undoStack : redoStack).length) return;
  busy = true; lock();
  try {
    const n = await stepMark(dir);
    say(n ? (dir < 0 ? L('تم التراجع عن آخر تمييز.', 'The last mark was undone.') : L('تمت إعادة التمييز.', 'The mark was done again.')) : L('لم أجد العنصر.', 'The object was not found.'));
  } catch (e) { say(L('تعذّر: ', 'Could not: ') + errText(e)); }
  finally { busy = false; lock(); }
}
$('opUndo').addEventListener('click', undoOp);
$('undoBtn').addEventListener('click', () => undoRedo(-1));
$('redoBtn').addEventListener('click', () => undoRedo(1));
$('solRoles').addEventListener('click', e => { const b = e.target.closest('.tbtn'); if (b) pinRole(b.dataset.role, true); });
const rsay = m => say(m);
async function clearSel(slideId) {
  await run(async ctx => {
    const sh = ctx.presentation.slides.getItem(slideId).shapes; sh.load('items/name'); await ctx.sync();
    sh.items.forEach(s => { if (/ #ursel$/.test(s.name || '')) s.name = s.name.replace(/ #ursel$/, ''); });
    await ctx.sync();
  }).catch(() => {});
}
// solution marks change only the selected objects, in place: no slide is rebuilt, nothing moves
const QUIET = { answer: '1F71DE', correct: '2070DF', given: null, graph: null };
// ---------- quick undo: the whole last operation at once ----------
// before every operation of the tool the slides it may change are copied (by their place in the deck);
// «تراجع» puts those copies back in their places, so one press undoes the whole operation (not step by step)
const opStack = [];
async function snapOp(ids) {
  if (!ids || !ids.length || !supported || isWord || ids.length > 40) return;   // a very long deck: no copy (too heavy)
  try {
    const pos = await run(async ctx => { const sl = ctx.presentation.slides; sl.load('items/id'); await ctx.sync(); return ids.map(id => sl.items.findIndex(x => x.id === id)); });
    const copies = [];
    for (let k = 0; k < ids.length; k++) if (pos[k] >= 0) copies.push({ pos: pos[k], b64: await exportSlide(ids[k]) });
    if (!copies.length) return;
    opStack.push(copies); if (opStack.length > 10) opStack.shift();
    opPaint();
  } catch (e) {}
}
function opPaint() { const b = $('opUndo'); if (b) b.disabled = busy || !supported || !opStack.length; }
async function undoOp() {
  if (busy || !supported || !opStack.length) return;
  busy = true; lock();
  try {
    const copies = opStack.pop();
    let back = null;
    for (const c of copies) {
      const id = await run(async ctx => { const sl = ctx.presentation.slides; sl.load('items/id'); await ctx.sync(); return sl.items[c.pos] ? sl.items[c.pos].id : null; });
      if (!id) continue;
      back = await insertAfter(c.b64, id);
      await deleteSlide(id);
    }
    if (back) await run(async ctx => { ctx.presentation.setSelectedSlides([back]); await ctx.sync(); }).catch(() => {});
    solSlideId = null; await solLabel().catch(() => {});
    say(L('رجعت الشريحة كما كانت قبل آخر عملية.', 'The slide is back as it was before the last operation.'));
  } catch (e) { say(L('تعذّر التراجع: ', 'Could not undo: ') + errText(e)); }
  finally { busy = false; lock(); }
}
// «Undo» / «Redo» of the marks given with the solution buttons (name, text colour, place), in place
const undoStack = [], redoStack = [];
// the text colour of every cell of a table (null where PowerPoint can't tell, or the API is missing)
async function tableCells(ctx, shape) {
  try {
    const t = shape.getTable(); t.load('rowCount,columnCount'); await ctx.sync();
    const cells = [];
    for (let r = 0; r < t.rowCount; r++) for (let c = 0; c < t.columnCount; c++) {
      const cell = t.getCellOrNullObject(r, c); cell.load('isNullObject'); cell.font.load('color'); cells.push({ r, c, cell });
    }
    await ctx.sync();
    return cells.filter(x => !x.cell.isNullObject);
  } catch (e) { return null; }
}
async function tableColours(ctx, shape) {
  const cells = await tableCells(ctx, shape);
  return cells ? cells.map(x => ({ r: x.r, c: x.c, color: x.cell.font.color || null })) : null;
}
async function paintTable(ctx, shape, colour) {
  const cells = await tableCells(ctx, shape); if (!cells) return false;
  cells.forEach(x => { try { x.cell.font.color = colour; } catch (e) {} });
  await ctx.sync(); return true;
}
function undoPaint() {
  const u = $('undoBtn'), r = $('redoBtn');
  if (u) u.disabled = busy || !supported || !undoStack.length;
  if (r) r.disabled = busy || !supported || !redoStack.length;
}
// how the objects are now, to go back to it from the other side
async function snapState(ctx, slide, ids) {
  const sh = ctx.presentation.slides.getItem(slide).shapes; sh.load('items/id,items/name,items/left,items/type'); await ctx.sync();
  const list = ids.map(id => sh.items.find(x => x.id === id)).filter(Boolean);
  const fonts = list.map(x => { if (!['TextBox', 'GeometricShape'].includes(x.type)) return null; try { const f = x.textFrame.textRange.font; f.load('color,bold'); return f; } catch (e) { return null; } });
  await ctx.sync();
  const cells = []; for (const x of list) cells.push(x.type === 'Table' ? await tableColours(ctx, x) : null);
  return { slide, items: list.map((x, k) => ({ id: x.id, name: x.name, left: x.left, color: fonts[k] ? fonts[k].color : null, bold: fonts[k] ? fonts[k].bold : null, cells: cells[k] })) };
}
async function stepMark(dir) {
  const from = dir < 0 ? undoStack : redoStack, to = dir < 0 ? redoStack : undoStack;
  const u = from.pop(); undoPaint(); if (!u) return 0;
  const n = await run(async ctx => {
    const snap = await snapState(ctx, u.slide, u.items.map(i => i.id));
    const sh = ctx.presentation.slides.getItem(u.slide).shapes; sh.load('items/id'); await ctx.sync();
    let n = 0;
    // the filled circles of a correct choice: undo takes them away, redo draws them again
    if (u.added) { for (const a of u.added) { const c = sh.items.find(x => x.id === a.id); if (c) c.delete(); } snap.redraw = u.added.map(a => a.d); }
    to.push(snap);
    if (u.redraw) { const made = u.redraw.map(d => ({ c: makeDot(ctx, sh, d), d })); made.forEach(x => x.c.load('id')); await ctx.sync(); snap.added = made.map(x => ({ id: x.c.id, d: x.d })); }
    for (const it of u.items) {
      const s = sh.items.find(x => x.id === it.id); if (!s) continue;
      s.name = it.name; s.left = it.left; n++;
      if (it.color || it.bold != null) { try { const f = s.textFrame.textRange.font; if (it.color) f.color = it.color; if (it.bold != null) f.bold = it.bold; } catch (e) {} }
      if (it.cells) { try { const t = s.getTable(); for (const c of it.cells) if (c.color) t.getCellOrNullObject(c.r, c.c).font.color = c.color; } catch (e) {} }
    }
    await ctx.sync(); return n;
  });
  undoPaint(); return n;
}
// the correct choice, readable in black and white too: its circled letter (ⓐ…ⓩ) gets a filled blue circle over it with
// the plain letter in white (the circled letter under it turns white). On «Hide solutions» the circle goes (it is a
// solution) and the letter comes back black, like the other options.
const DOT_FILL = '#1F71DE';
function makeDot(ctx, shapes, d) {
  const c = shapes.addGeometricShape(((PowerPoint.GeometricShapeType || {}).ellipse) || 'Ellipse', { left: d.left, top: d.top, width: d.size, height: d.size });
  c.name = 'UULA:answer UULA CorrectDot';
  try { c.fill.setSolidColor(DOT_FILL); } catch (e) {}
  try { c.lineFormat.visible = false; } catch (e) {}
  try {
    const tf = c.textFrame; tf.wordWrap = false; tf.autoSizeSetting = 'AutoSizeNone';
    tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0; tf.verticalAlignment = 'Middle';
    tf.textRange.text = d.letter;
    const f = tf.textRange.font; f.color = '#FFFFFF'; f.size = d.font; f.bold = false; if (d.name) f.name = d.name;
    tf.textRange.paragraphFormat.horizontalAlignment = 'Center';
  } catch (e) {}
  return c;
}
// the place of the letter is measured, not guessed: a hidden sample box with the same letters, font and size is written
// off the slide and PowerPoint's own width of it gives where the letter ends (from the start side of the line).
// Returns null when the letter can't be placed for sure (then the caller highlights the letter in the text instead).
let dotMeasure = '';
async function dotCorrect(ctx, shapes, s, slideId) {
  dotMeasure = '';
  if (!['TextBox', 'GeometricShape'].includes(s.type)) return null;
  const tf = s.textFrame, tr = tf.textRange;
  tr.load('text'); try { tf.load('leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment'); } catch (e) {} try { tr.paragraphFormat.load('horizontalAlignment'); } catch (e) {}
  s.load('left,top,width,height'); await ctx.sync();
  const text = tr.text || '';
  const first = text.split(/[\r\n\v]/)[0];
  // the circled letters of the first line; several → the blue one
  const at = []; for (let i = 0; i < first.length; i++) if (/[ⓐ-ⓩ]/.test(first[i])) at.push(i);
  if (!at.length) { dotMeasure = 'no letter: ' + JSON.stringify(first.slice(0, 12)); return null; }
  let idx = at[0];
  if (at.length > 1) {
    const subs = at.map(i => { const x = tr.getSubstring(i, 1); try { x.font.load('color'); } catch (e) {} return x; });
    await ctx.sync();
    const blue = c => { const h = /^#?([0-9a-f]{6})$/i.exec(c || ''); if (!h) return false; const v = parseInt(h[1], 16), r = v >> 16, b = v & 255; return b > 150 && r < 120 && b - r > 60; };
    const k = subs.findIndex(x => { try { return blue(x.font.color); } catch (e) { return false; } });
    if (k < 0) { dotMeasure = 'no blue letter'; return null; }
    idx = at[k];
  }
  const sub = tr.getSubstring(idx, 1); try { sub.font.load('size,name'); } catch (e) {} await ctx.sync();
  const fs = sub.font.size || 12, fname = sub.font.name || null;
  // the two samples: the line up to the letter (included), and the letter alone
  const sample = str => {
    const t = shapes.addTextBox(str, { left: -3000, top: -3000, width: 10, height: 10 });
    try { const f = t.textFrame; f.wordWrap = false; f.autoSizeSetting = 'AutoSizeShapeToFitText'; f.leftMargin = 0; f.rightMargin = 0; f.topMargin = 0; f.bottomMargin = 0; const ft = f.textRange.font; ft.size = fs; if (fname) ft.name = fname; } catch (e) {}
    return t;
  };
  // PowerPoint keeps a sample box at a minimum size and adds room around it: every sample is made long (ten more
  // letters after it) and the sizes are taken by difference — ten letters vs eleven (one letter), two lines vs three
  const L1 = first[idx], TEN = L1.repeat(10);
  const pre = sample(first.slice(0, idx + 1) + TEN), ten = sample(TEN), eleven = sample(TEN + L1), r2 = sample(TEN + '\n' + TEN), r3 = sample(TEN + '\n' + TEN + '\n' + TEN);
  await ctx.sync();
  [pre, ten, eleven, r2, r3].forEach(x => x.load('width,height')); await ctx.sync();
  const g = eleven.width - ten.width, lh = r3.height - r2.height;
  const wPre = pre.width - ten.width;                  // the line up to the letter (included)
  [pre, ten, eleven, r2, r3].forEach(x => x.delete()); await ctx.sync();
  // the sizes did not follow the text: not sure → the other way
  dotMeasure = 'g=' + g.toFixed(1) + ' lh=' + lh.toFixed(1) + ' pre=' + wPre.toFixed(1) + ' fs=' + fs;
  // (the circled letter's advance is wide in the Mac fonts — about 1.65 × the size — so only absurd values are refused)
  if (!(g > fs * 0.4 && g < fs * 3 && lh > fs * 0.6 && lh < fs * 4 && wPre >= g - 0.5)) return null;
  const gw = g, pad = 0;
  // the start side of the line: right for Arabic text (or a right-to-left mark), left otherwise
  const rtl = /^[\s]*‏/.test(first) || (/[؀-ۿ]/.test(text) && !/^[\s]*‎/.test(first));
  const al = tr.paragraphFormat.horizontalAlignment;
  const lm = tf.leftMargin || 0, rm = tf.rightMargin || 0, tm = tf.topMargin || 0, bm = tf.bottomMargin || 0;
  // the paragraph's own margins (marL / indent) and alignment, and the box insets, read from the slide itself
  let geo = null;
  try {
    const nm0 = s.name; s.name = nm0 + ' #urdot'; await ctx.sync();
    const ex = ctx.presentation.slides.getItem(slideId).exportAsBase64(); await ctx.sync();
    s.name = nm0; await ctx.sync();
    const pkg = await UR.openPackage(ex.value), doc = UR.parseXml(pkg.files.slide), A = UR.NS.a, P = UR.NS.p;
    const nv = [...doc.getElementsByTagNameNS(P, 'cNvPr')].find(n => / #urdot$/.test(n.getAttribute('name') || ''));
    const sp = nv && nv.parentNode && nv.parentNode.parentNode;
    if (sp) {
      const bp = sp.getElementsByTagNameNS(A, 'bodyPr')[0], p0 = sp.getElementsByTagNameNS(A, 'p')[0], pr = p0 && [...p0.children].find(c => c.localName === 'pPr');
      const at = (el, a, d) => el && el.hasAttribute(a) ? +el.getAttribute(a) : d;
      geo = { lIns: at(bp, 'lIns', 91440) / 12700, rIns: at(bp, 'rIns', 91440) / 12700, tIns: at(bp, 'tIns', 45720) / 12700, bIns: at(bp, 'bIns', 45720) / 12700,
        marL: at(pr, 'marL', 0) / 12700, indent: at(pr, 'indent', 0) / 12700, algn: pr && pr.getAttribute('algn') || null, rtl: pr && pr.getAttribute('rtl') === '1' };
    }
  } catch (e) { geo = null; }
  const RTL = geo ? (geo.rtl || (geo.algn === 'r' && rtl)) : rtl;
  const L0 = geo ? geo.lIns : lm, R0 = geo ? geo.rIns : rm, T0 = geo ? geo.tIns : tm, B0 = geo ? geo.bIns : bm;
  // the line starts after the paragraph margin and the first-line indent (on the right for right-to-left)
  const inStart = geo ? geo.marL + geo.indent : 0;
  const atStart = geo ? (RTL ? geo.algn !== 'l' && geo.algn !== 'ctr' : geo.algn !== 'r' && geo.algn !== 'ctr') : al !== 'Center';
  let cx;
  if (atStart) cx = RTL ? s.left + s.width - R0 - inStart - (wPre - gw / 2) : s.left + L0 + inStart + (wPre - gw / 2);
  else {
    const all = sample(first + TEN), t2 = sample(TEN); await ctx.sync(); all.load('width'); t2.load('width'); await ctx.sync(); const wl = all.width - t2.width; all.delete(); t2.delete(); await ctx.sync();
    const inner0 = s.left + L0 + (RTL ? 0 : inStart), inner1 = s.left + s.width - R0 - (RTL ? inStart : 0);
    const lineL = (geo ? geo.algn === 'ctr' : al === 'Center') ? (inner0 + inner1) / 2 - wl / 2 : RTL ? inner0 : inner1 - wl;
    cx = RTL ? lineL + wl - (wPre - gw / 2) : lineL + (wPre - gw / 2);
  }
  const lines0 = text.split(/[\r\n\v]/).length;
  // one line in its box: the line is the box minus its insets (exact); otherwise the measured line height
  const lineH = lines0 === 1 && s.height - T0 - B0 > fs * 0.6 ? Math.min(lh, s.height - T0 - B0) : lh;
  const va0 = tf.verticalAlignment;
  const top0 = va0 === 'Middle' && lines0 === 1 ? s.top + (s.height - lineH) / 2 : va0 === 'Bottom' && lines0 === 1 ? s.top + s.height - B0 - lineH : s.top + T0;
  const cyM = top0 + lineH * 0.5;
  const cy = cyM;
  // the circle drawn in the glyph is about the size of the font (its advance has room on both sides)
  const dia = Math.round(Math.min(gw * 0.92, fs * 0.98) * 10) / 10;
  // already done (pressed twice): one circle is enough
  const all = shapes; all.load('items/name,items/left,items/top,items/width,items/height'); await ctx.sync();
  if (all.items.some(x => /UULA CorrectDot/.test(x.name || '') && Math.abs(x.left + x.width / 2 - cx) < dia / 2 && Math.abs(x.top + x.height / 2 - cy) < dia / 2)) return { id: null };
  sub.font.color = '#FFFFFF';
  const d = { left: Math.round((cx - dia / 2) * 10) / 10, top: Math.round((cy - dia / 2) * 10) / 10, size: dia, font: Math.round(fs * 0.8 * 2) / 2, letter: String.fromCharCode(first.charCodeAt(idx) - 0x24D0 + 97), name: null };
  const c = makeDot(ctx, shapes, d);
  c.load('id'); await ctx.sync();
  return { id: c.id, d };
}
let fillAfter = false;
async function quietMark(role, keepColour) {
  fillAfter = false;
  const id = await currentId();
  const hiddenNow = (role === 'answer' || role === 'graph') && await slideSolHidden(id);
  const W = hiddenNow ? (await getSize()).W : 0;
  return run(async ctx => {
    const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/name,items/type,items/left,items/top,items/width,items/height'); await ctx.sync();
    const sel = shapes.items;
    if (!sel.length) return 0;
    // what they were, for «Undo»
    const fonts = sel.map(x => { if (!['TextBox', 'GeometricShape'].includes(x.type)) return null; try { const f = x.textFrame.textRange.font; f.load('color,bold'); return f; } catch (e) { return null; } });
    sel.forEach(x => x.load('id')); await ctx.sync();
    // a table marked as (part of) a solution: its text turns solution blue at once
    const blueTable = role === 'answer';
    const cells = []; for (const x of sel) cells.push(blueTable && x.type === 'Table' ? await tableColours(ctx, x) : null);
    undoStack.push({ slide: id, items: sel.map((x, k) => ({ id: x.id, name: x.name, left: x.left, color: fonts[k] ? fonts[k].color : null, bold: fonts[k] ? fonts[k].bold : null, cells: cells[k] })) });
    if (undoStack.length > 30) undoStack.shift();
    redoStack.length = 0;
    undoPaint();
    let all = null;
    if (hiddenNow) { all = ctx.presentation.slides.getItem(id).shapes; all.load('items/name,items/left'); await ctx.sync(); }
    const base = n => (n || '').replace(/ #ursel$/, '').replace(/^UULA Sol(Hidden|View) /, '').replace(/^UULA:[a-z]+\s*/, '');
    for (const s of sel) {
      if (hiddenNow && /^UULA SolView /.test(s.name || '')) {
        // the student's copy of a question still showing part of a solution: the copy goes, its original (kept aside) is now a solution
        const b = base(s.name), orig = all.items.find(x => /^UULA SolHidden /.test(x.name || '') && base(x.name) === b);
        if (orig) orig.name = 'UULA SolHidden UULA:answer ' + b;
        s.delete(); continue;
      }
      s.name = 'UULA:' + role + ' ' + base(s.name);
      if (blueTable && s.type === 'Table') await paintTable(ctx, s, '#' + (QUIET.answer || '1F71DE'));
      if (QUIET[role] && !keepColour && ['TextBox', 'GeometricShape'].includes(s.type)) {
        try { const f = s.textFrame.textRange.font; f.color = '#' + QUIET[role]; if (role === 'correct') f.bold = true; } catch (e) {}
      }
      // the correct choice: its circled letter is filled in the text itself right after (see fillSlide)
      if (role === 'correct' && ['TextBox', 'GeometricShape'].includes(s.type)) {
        await ctx.sync();
        const dot = await dotCorrect(ctx, ctx.presentation.slides.getItem(id).shapes, s, id).catch(e => { dotMeasure = (dotMeasure ? dotMeasure + ' ' : '') + 'err: ' + errText(e); return null; });
        if (dot && dot.id) (undoStack[undoStack.length - 1].added = undoStack[undoStack.length - 1].added || []).push(dot);
        // the place could not be measured: the letter is highlighted in the text instead (see fillSlide)
        if (!dot) { s.name = s.name + ' #urfill'; fillAfter = true; }
      }
      // the solutions of this slide are hidden now: this one goes aside too, and «Show solutions» brings it back
      // (a graph leaves the empty grid in its place)
      if (hiddenNow && role === 'graph') { await ctx.sync(); await gridOverShape(ctx, ctx.presentation.slides.getItem(id).shapes, s, W, []); continue; }
      if (hiddenNow) { s.name = 'UULA SolHidden ' + s.name; s.left = s.left + 3 * W; }
    }
    await ctx.sync();
    return hiddenNow ? -sel.length : sel.length;
  });
}
// «Suspended»: a yellow frame over the selection, the content faded under it, and «مـعـلـق» with the warning
// triangle in the middle — drawn in place and grouped (the content itself stays as it is, under it)
function suspendSay(r) {
  if (!r.n) return L('تعذّر رسم إطار «معلّق».', 'Could not draw the «suspended» frame.');
  return (r.n < 0 ? L('تمت إضافة «معلّق» في الأسفل.', '«Suspended» was added at the bottom.') : L('وُضع إطار «معلّق» حول المحدد.', 'The «suspended» frame was put around the selection.')) + (r.failed && r.failed.length ? L(' لم يقبل الباوربوينت: ', ' PowerPoint refused: ') + r.failed.join('، ') : '');
}
async function suspendSelected() {
  return run(async ctx => {
    const slides = ctx.presentation.getSelectedSlides(); slides.load('items/id');
    const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/left,items/top,items/width,items/height');
    await ctx.sync();
    if (!slides.items.length) return { n: 0 };
    const sid = slides.items[0].id;
    const sh = ctx.presentation.slides.getItem(sid).shapes;
    let X, Y, Wd, Ht, n = shapes.items.length;
    if (n) {
      const x0 = Math.min(...shapes.items.map(s => s.left)), y0 = Math.min(...shapes.items.map(s => s.top));
      const x1 = Math.max(...shapes.items.map(s => s.left + s.width)), y1 = Math.max(...shapes.items.map(s => s.top + s.height));
      const pad = 6; X = x0 - pad; Y = y0 - pad; Wd = x1 - x0 + 2 * pad; Ht = y1 - y0 + 2 * pad;
    } else {
      // nothing selected: the same box on its own, under the last object, as big as the booklet's «suspended» box
      // (the full width of the page, about 74pt high)
      const pkg = await UR.openPackage(await exportSlide(sid));
      const ft = URT.freeTop(pkg, 1);
      X = 0; Wd = pkg.slideW / PT; Y = ft.top / PT; Ht = 5.5 * UR.ROW / PT; n = -1;
    }
    const failed = [];
    // each step on its own: one refusal does not stop the rest
    const step = async (what, fn) => { try { fn(); await ctx.sync(); return true; } catch (e) { failed.push(what); return false; } };
    let box = null;
    if (!(await step('frame', () => { box = sh.addGeometricShape('RoundRectangle', { left: X, top: Y, width: Wd, height: Ht }); })))
      await step('frame2', () => { box = sh.addGeometricShape('Rectangle', { left: X, top: Y, width: Wd, height: Ht }); });
    if (!box) return { n: 0, failed };
    await step('name', () => { box.name = 'UULA Suspended'; });
    // white, 40% see-through: the content shows faded under it; the yellow line is solid
    await step('fill', () => { box.fill.setSolidColor('#FFFFFF'); });
    await step('see-through', () => { box.fill.transparency = 0.4; });
    await step('line', () => { box.lineFormat.visible = true; box.lineFormat.color = '#F5C242'; box.lineFormat.weight = 1.5; });
    await step('corners', () => { box.adjustments.set(0, Math.min(0.5, 5 / Math.min(Wd, Ht))); });
    // the words written inside the rectangle itself, in its middle
    await step('words', () => {
      const tf = box.textFrame; tf.wordWrap = true; tf.autoSizeSetting = 'AutoSizeNone'; tf.verticalAlignment = 'Middle';
      tf.textRange.text = '⚠️ معلق';
      const f = tf.textRange.font; f.name = 'UULA Sans'; f.size = 20; f.bold = true; f.color = '#F5C242';
      tf.textRange.paragraphFormat.horizontalAlignment = 'Center';
    });
    return { n, failed: failed.filter(f => f !== 'frame' && f !== 'corners') };
  });
}
// fill the correct choices of one slide (in place when possible, otherwise the slide is rebuilt); returns the result
async function fillSlide(id, all) {
  const pkg = await UR.openPackage(await exportSlide(id));
  const r = UR.fillCorrect(pkg, all);
  if (!r.count && !r.dots && !/ #urfill/.test(pkg.files.slide)) return Object.assign(r, { id });
  // the highlight behind a letter can't be set in place by the PowerPoint API: the slide is rebuilt
  const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, null), id);
  await deleteSlide(id);
  await run(async ctx => { ctx.presentation.setSelectedSlides([newId]); await ctx.sync(); }).catch(() => {});
  return Object.assign(r, { id: newId });
}
// «الخيار الصحيح» with nothing selected: a clear message (the hint line above the list too)
function needCorrect() {
  const msg = L('حدّد الخيار الصحيح أولًا', 'Select the correct choice first');
  say(msg + '.');
  const h = $('selHint'); if (!h) return;
  h.dataset.on = 'need'; h.innerHTML = msg + ' <b>!</b>';
  clearTimeout(needTimer); needTimer = setTimeout(selPaint, 3500);
}
// (kept, not on a button now) every choice marked correct before (or whose circled letter is blue) on the
// selected slides gets its filled circle, all at once
async function correctAll() {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  const ids = await selIds().catch(() => []);
  await snapOp(ids);
  try {
    if (!ids.length) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
    let done = 0, slides = 0, skipped = 0;
    const keep = [];
    for (let k = 0; k < ids.length; k++) {
      prog(k / ids.length);
      const r = await fillSlide(ids[k], true).catch(e => { line('err', '✗ ' + esc(errText(e))); return { count: 0, skipped: 0, id: ids[k] }; });
      keep.push(r.id || ids[k]); skipped += r.skipped || 0;
      if (r.count) { done += r.count; slides++; }
    }
    prog(null);
    await run(async ctx => { ctx.presentation.setSelectedSlides(keep); await ctx.sync(); }).catch(() => {});
    if (skipped) line('warn', L('• خيارات مميزة كصحيحة لم أعرف أي حرف فيها هو الصحيح (أكثر من حرف بنفس اللون): ', '• Choices marked correct where the right letter is unclear (several letters of one colour): ') + skipped + '.');
    const q = done === 1 ? 'سؤال واحد' : done === 2 ? 'سؤالين' : done <= 10 ? done + ' أسئلة' : done + ' سؤالًا';
    say(done ? L('ظُلّل حرف الإجابة الصحيحة في ' + q + (ids.length > 1 ? '، في ' + (slides === 2 ? 'شريحتين' : cnt(slides)) : '') + '.', 'The correct letter was highlighted in ' + done + ' questions' + (ids.length > 1 ? ', in ' + cnt(slides) : '') + '.')
      : L('لم أجد خيارات مميزة كصحيحة (أو حرفها أزرق) بحرف غير مظلّل. لتمييز خيار: حدّده ثم اضغط الزر.', 'No choice marked correct (or with a blue letter) left unhighlighted. To mark one: select it, then press the button.'));
  } catch (e) { say(L('تعذّر: ', 'Could not: ') + errText(e)); }
  finally { busy = false; lock(); }
}
async function pinRole(role, keepColour) {
  if (busy || !supported) return;
  if (role === 'correct') {
    const n = await run(async ctx => { const ss = ctx.presentation.getSelectedShapes(); ss.load('items'); await ctx.sync(); return ss.items.length; }).catch(() => 0);
    if (!n) return needCorrect();
  }
  busy = true; lock(); rsay('');
  await snapOp([await currentId().catch(() => null)].filter(Boolean));
  if (role === 'cancelled') {
    try { rsay(suspendSay(await suspendSelected())); }
    catch (e) { rsay(L('تعذّر: ', 'Could not: ') + errText(e)); }
    finally { busy = false; lock(); }
    return;
  }
  if (role in QUIET) {
    try {
      let n = await quietMark(role, keepColour);
      if (fillAfter) { const id0 = await currentId().catch(() => null); if (id0) { await fillSlide(id0, false).catch(() => null); undoStack.length = 0; redoStack.length = 0; undoPaint(); } }
      if (fillAfter) { rsay(L('تعذّر قياس مكان الحرف، فظُلّل بدل الدائرة.', 'The letter could not be measured, so it was highlighted instead.') + (dotMeasure ? ' [' + dotMeasure + ']' : '')); return; }
      if (n < 0 && role === 'graph') { rsay(L('عُلِّم كحل بياني ووُضعت مكانه الشبكة الفارغة، لأن حلول الشريحة مخفية الآن. يرجع مع «إظهار الحلول».', 'Marked as a graph solution and replaced by the empty grid, since the slide\'s solutions are hidden now. «Show solutions» brings it back.')); return; }
      if (n < 0) { rsay(L('عُلِّم كحل وأُخفي، لأن حلول الشريحة مخفية الآن. يظهر مع «إظهار الحلول».', 'Marked as a solution and hidden, since the slide\'s solutions are hidden now. «Show solutions» brings it back.')); return; }
      rsay(!n ? L('حدّد أولًا الكائن الذي تريد تمييزه.', 'First select the object to mark.')
        : role === 'given' ? L('عُلِّم كجزء من السؤال: يبقى ظاهرًا عند إخفاء الحلول.', 'Marked as part of the question: it stays when solutions are hidden.')
        : role === 'correct' ? L('عُلِّم كخيار صحيح: صار حرفه دائرة مملوءة تظهر حتى في الطباعة بالأبيض والأسود، ويعود حرفًا عاديًا عند إخفاء الحلول.', 'Marked as the correct choice: its letter is now white on a blue highlight, clear even in black-and-white print; it turns back to a plain letter when solutions are hidden.')
        : role === 'graph' ? L('عُلِّم كحل بياني: تحل مكانه شبكة فارغة عند إخفاء الحلول.', 'Marked as a graph solution: an empty grid takes its place when solutions are hidden.')
        : keepColour ? L('عُلِّم كحل: يُخفى عند إخفاء الحلول.', 'Marked as a solution: it is hidden when solutions are hidden.') : L('عُلِّم كحل (أزرق): يُخفى عند إخفاء الحلول.', 'Marked as a solution (blue): it is hidden when solutions are hidden.'));
    } catch (e) { rsay(L('تعذّر التمييز: ', 'Could not mark: ') + errText(e)); }
    finally { busy = false; lock(); }
    return;
  }
  let slideId = null;
  try {
    const got = await run(async ctx => {
      const slides = ctx.presentation.getSelectedSlides(); slides.load('items/id');
      const shapes = ctx.presentation.getSelectedShapes(); shapes.load('items/name,items/type');
      await ctx.sync();
      if (!slides.items.length || !shapes.items.length) return null;
      shapes.items.forEach(s => { s.name = (s.name || '').replace(/ #ursel$/, '') + ' #ursel'; });
      await ctx.sync();
      // pictures (an equation from UULA Arabic Math, a graph …), groups and other objects are not text boxes
      const other = shapes.items.some(s => !['TextBox', 'GeometricShape'].includes(s.type));
      return { id: slides.items[0].id, n: shapes.items.length, other };
    });
    if (!got) { rsay(L('حدّد مربعًا في الشريحة أولًا.', 'Select a box on the slide first.')); return; }
    // several objects selected and a frame chosen: one frame around all of them
    // a picture or other object in the selection: the frame goes around it (as for several objects)
    if ((got.n > 1 || got.other) && FRAME_ROLES.includes(role)) {
      if (await wrapSelected(got.id, role)) rsay(got.n > 1 ? L('وُضع إطار واحد حول العناصر المحددة.', 'One frame was put around the selected objects.') : L('وُضع الإطار حول العنصر المحدد.', 'The frame was put around the selected object.'));
      return;
    }
    slideId = got.id;
    const pkg = await UR.openPackage(await exportSlide(slideId));
    let r;
    try { r = UR.applyRole(pkg, role, { dir: mcqDir }); if (['given', 'correct'].includes(role)) r.targets = []; }
    catch (e) { await clearSel(slideId); rsay(e.message === 'nosel' ? L('حدّد مربعًا في الشريحة أولًا.', 'Select a box on the slide first.') : L('تعذّر التحويل: ', 'Could not restyle: ') + errText(e)); return; }
    // in place when it can be done exactly (nothing moves on screen); otherwise the slide is rebuilt
    let newId = slideId;
    if (!(await UQ.apply(slideId, pkg.files.slide, r.xml, exportSlide).catch(() => false))) {
      newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), slideId);
      await deleteSlide(slideId);
    }
    slideId = null;
    // real text heights: fit the container around the text
    await run(async ctx => {
      const { map } = await shapesByMark(ctx, newId);
      const texts = r.targets.filter(t => t.kind === 'text' && map[t.idx]);
      texts.forEach(t => { const s = map[t.idx]; try { s.textFrame.wordWrap = true; s.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; } catch (e) {} });
      await ctx.sync();
      texts.forEach(t => map[t.idx].load('height,top'));
      await ctx.sync();
      for (const t of texts) {
        if (!t.box || !map[t.box]) continue;
        const h = map[t.idx].height, other = t.boxType === 'other';
        const b = map[t.box]; b.top = t.top / PT; b.height = other ? h + 8 : h + 3 * 13.3;
        if (t.line != null && map[t.line]) map[t.line].height = Math.max(13.3, h - 17);
      }
      await ctx.sync();
      // on the page grid: two grid steps under what is above, and what is under keeps two grid steps
      const sh = ctx.presentation.slides.getItem(newId).shapes;
      sh.load('items/name,items/left,items/top,items/width,items/height'); await ctx.sync();
      const H = r.H / PT;
      const rects = sh.items.map((s, k) => ({ key: k, x: s.left * PT, y: s.top * PT, w: s.width * PT, h: s.height * PT,
        fixed: /^UULA (Footer|Watermark)|^CustomPageNumber_/.test((s.name || '').replace(/ #ur\d+$/, '')) || s.top > H - 54 || s.left + s.width <= 0 || s.left >= r.W / PT }));
      const keyOf = idx => { const s = map[idx]; return s ? sh.items.findIndex(x => x.name === s.name) : -1; };
      const order = r.targets.slice().sort((a, c) => (rects[keyOf(a.idx)] || {}).y - (rects[keyOf(c.idx)] || {}).y);
      for (const t of order) {
        const tk = keyOf(t.idx); if (tk < 0) continue;
        const group = [tk, t.box != null ? keyOf(t.box) : -1, t.icon != null ? keyOf(t.icon) : -1, t.line != null ? keyOf(t.line) : -1].filter(k => k >= 0);
        const mv = UR.reflow(rects, group, { textKey: tk, keepX: !!t.box || t.role === 'chapter' || t.role === 'lesson' });
        for (const [k, m] of Object.entries(mv)) {
          const R = rects[+k], s = sh.items[+k];
          if (m.y != null) { R.y = m.y; s.top = m.y / PT; }
          if (m.x != null) { R.x = m.x; s.left = m.x / PT; }
          if (m.w != null) { R.w = m.w; s.width = m.w / PT; }
        }
      }
      await ctx.sync();
      ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
      try { const ids = r.targets.map(t => map[t.idx]).filter(Boolean).map(s => s.id); ctx.presentation.slides.getItem(newId).setSelectedShapes(ids); await ctx.sync(); } catch (e) {}
    });
    await unmark(newId);
    rsay(role === 'given' ? L('عُلِّم كشكل من السؤال: يبقى ظاهرًا عند إخفاء الحلول.', 'Marked as part of the question: it stays when solutions are hidden.')
      : role === 'correct' ? L('عُلِّم كخيار صحيح (أزرق عريض): يتحول للأسود عند إخفاء الحلول.', 'Marked as the correct choice (bold blue): it turns black when solutions are hidden.')
      : role === 'choice' ? L('حُوِّل إلى خيارات.', 'Turned into options.')
      : ITEM[role] ? L('حُوِّل إلى «' + ITEM[role].ar + '».', 'Turned into «' + ITEM[role].en + '».') : L('تم.', 'Done.'));
  } catch (e) { if (slideId) await clearSel(slideId); rsay(L('تعذّر التحويل: ', 'Could not restyle: ') + errText(e)); }
  finally { busy = false; lock(); }
}

// ---------- vertical spacing only ----------
// only: the ids of the selected objects (two or more) — then the gaps are set between those only, the rest stays
async function spaceSlide(id, only) {
  const { W, H } = await getSize();
  return run(async ctx => {
    const sh = ctx.presentation.slides.getItem(id).shapes;
    sh.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height'); await ctx.sync();
    // text boxes take the height of their text (so the gaps are real)
    for (const s of sh.items) if (s.type === 'TextBox' && (!only || only.includes(s.id))) { try { s.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; } catch (e) {} }
    await ctx.sync();
    const fills = sh.items.map(s => { try { s.fill.load('type'); return s.fill; } catch (e) { return null; } });
    sh.items.forEach(s => s.load('top,height'));
    await ctx.sync();
    const rects = sh.items.map((s, k) => {
      const nm = s.name || '';
      const fixed = /^UULA (Footer|Watermark)|^CustomPageNumber_/.test(nm) || s.top >= H - 54 || s.left + s.width <= 0 || s.left >= W || s.top + s.height <= 0 || (s.width > W * 0.85 && s.height > H * 0.6) || / #ur\d+$/.test(nm) && false;
      let solid = false; try { solid = fills[k] && fills[k].type === 'Solid' || fills[k] && fills[k].type === 'Gradient'; } catch (e) {}
      const box = (s.type === 'GeometricShape' && solid && s.width > W * 0.5) || /Container$/.test(nm);
      return { key: k, x: s.left * PT, y: s.top * PT, w: s.width * PT, h: s.height * PT, fixed, box };
    });
    const mine = only ? rects.filter(r => only.includes(sh.items[r.key].id)) : rects;
    const mv = UR.spacing(mine, W * PT, H * PT, !only);
    // with a selection, the others stay where they are but still count: a selected object that would land on one of
    // them goes under it (one grid row lower, on the page grid), and the selected ones after it follow by as much
    if (only) {
      const others = rects.filter(r => !mine.includes(r) && !r.fixed);
      const ROWE = 2 * 6.695 * PT, Y0 = 22.58 * PT, snap = y => Y0 + Math.ceil((y - Y0 - 0.3 * PT) / ROWE * 2) * ROWE / 2;
      const placed = mine.map(r => ({ r, y: mv[r.key] && mv[r.key].y != null ? mv[r.key].y : r.y, h: mv[r.key] && mv[r.key].h != null ? mv[r.key].h : r.h })).sort((a, b) => a.y - b.y);
      let push = 0;
      for (const p of placed) {
        p.y += push;
        for (let guard = 0; guard < 20; guard++) {
          const hit = others.find(o => o.x < p.r.x + p.r.w - PT && o.x + o.w > p.r.x + PT && o.y < p.y + p.h - PT && o.y + o.h > p.y + PT);
          if (!hit) break;
          const ny = snap(hit.y + hit.h + ROWE); push += ny - p.y; p.y = ny;
        }
        mv[p.r.key] = Object.assign(mv[p.r.key] || {}, { y: p.y });
      }
    }
    // the rule: the content must fit between the top of the page and its bottom limit — if it can't, nothing moves
    let low = 0;
    for (const r of mine) if (!r.fixed) { const m = mv[r.key] || {}; low = Math.max(low, (m.y != null ? m.y : r.y) + (m.h != null ? m.h : r.h)); }
    if (low > (H - 54) * PT + PT) return { status: 'big', moved: 0 };
    let moved = 0;
    for (const [k, m] of Object.entries(mv)) {
      const s = sh.items[+k];
      if (m.y != null && Math.abs(m.y / PT - s.top) > 0.2) { s.top = m.y / PT; moved++; }
      if (m.h != null) s.height = m.h / PT;
    }
    await ctx.sync();
    return { status: 'ok', moved };
  }).then(async r => { await frameTop(id); return r; });
}
// distribute: after the gaps are fixed, the whole page is used — the content is cut into bands (objects that overlap in
// height stay together, so a box keeps its content and a figure stays beside its question), the first band goes up to the
// top margin and the free space (top, middle, bottom) is shared out equally under every band, in whole page-grid lines.
// Nothing is resized or restyled: whole bands only move down. With a selection only the selected objects move, inside the
// room between the first of them and the next unselected object under them.
async function distributeSlide(id, only) {
  if ((await spaceSlide(id, only)).status === 'big') return { status: 'big' };
  const { W, H } = await getSize();
  return run(async ctx => {
    const sh = ctx.presentation.slides.getItem(id).shapes;
    sh.load('items/id,items/name,items/left,items/top,items/width,items/height'); await ctx.sync();
    const lim = H - 54, G2 = 13.39, Y0 = 22.58;
    const isFixed = s => /^UULA (Footer|Watermark)|^CustomPageNumber_/.test(s.name || '') || s.top >= lim || s.left + s.width <= 0 || s.left >= W || s.top + s.height <= 0 || (s.width > W * 0.85 && s.height > H * 0.6);
    const all = sh.items.filter(s => !isFixed(s));
    const free = all.filter(s => !only || only.includes(s.id));
    if (!free.length) return { status: 'none' };
    // bands
    const bands = [];
    for (const s of free.slice().sort((a, b) => a.top - b.top)) {
      const b = bands[bands.length - 1];
      if (b && s.top < b.bottom - 0.5) { b.items.push(s); b.bottom = Math.max(b.bottom, s.top + s.height); }
      else bands.push({ top: s.top, bottom: s.top + s.height, items: [s] });
    }
    // the room: from the top margin (or, with a selection, from the first selected object) down to the footer,
    // or to the first unselected object that sits under the selection
    let top = only ? bands[0].top : Math.min(bands[0].top, Y0), floor = lim - G2 / 2;
    // something fixed at the top of the page (a header) keeps its place: the content starts under it
    if (!only) for (const s of sh.items) if (isFixed(s) && s.top < bands[0].top && s.top + s.height > top && s.top + s.height <= bands[0].top && s.left < W && s.left + s.width > 0 && s.top < lim) top = s.top + s.height + G2;
    if (only) {
      const last = bands[bands.length - 1].bottom;
      const under = all.filter(s => !only.includes(s.id) && s.top >= last - 0.5).map(s => s.top);
      if (under.length) floor = Math.min(...under) - G2;
    }
    const lift = Math.max(0, bands[0].top - top);
    const room = floor - (bands[bands.length - 1].bottom - lift);
    // the free space goes into the gaps between the parts (the last part ends at the bottom of the room), in half
    // page-grid steps; only real gaps get it (a question and the next one, a question and its options…), not lines
    // that sit tight together (the options of one question), so the look of each part stays as it is
    const G = G2 / 2, units = Math.max(0, Math.floor((room + 0.3) / G));
    const gaps = bands.slice(0, -1).map((b, k) => ({ k, g: bands[k + 1].top - b.bottom }));
    let cand = gaps.filter(x => x.g >= G2 * 0.75);
    if (!cand.length) cand = gaps;
    if (lift < 0.5 && (!cand.length || units < 1)) return { status: 'full' };
    const add = new Array(bands.length).fill(0);
    if (cand.length) {
      const per = Math.floor(units / cand.length); let rem = units - per * cand.length;
      cand.forEach(x => { add[x.k] = per; });
      // what is left over goes to the widest gaps first (between questions rather than inside one)
      for (const x of cand.slice().sort((p, q) => q.g - p.g || p.k - q.k)) { if (rem <= 0) break; add[x.k]++; rem--; }
    }
    // too little room for a half step in every gap: every gap gets the same small share instead (fair before on-grid)
    const step = cand.length && units < cand.length ? Math.max(0, room) / cand.length / G : 1;
    if (step !== 1) cand.forEach(x => { add[x.k] = 1; });
    let dy = -lift;
    bands.forEach((b, k) => { if (Math.abs(dy) > 0.2) b.items.forEach(s => { s.top = s.top + dy; }); dy += add[k] * G * step; });
    await ctx.sync();
    return { status: 'ok', pts: Math.round(units * G), gaps: cand.length };
  }).then(async r => { await frameTop(id); return r; });
}
// horizontal: to the right / left margin (selected objects only, if any are selected on the current slide)
let hMode = load('ur-halign', 'auto');
function setH(v) { hMode = v; save('ur-halign', v); $('hAlign').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === v))); }
$('hAlign').addEventListener('click', e => { const b = e.target.closest('button'); if (b) setH(b.dataset.v); });
setH(hMode);
async function alignSlide(id, mode, useSel) {
  const { W, H } = await getSize();
  return run(async ctx => {
    const sh = ctx.presentation.slides.getItem(id).shapes;
    sh.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
    let selIds = [];
    if (useSel) { const ss = ctx.presentation.getSelectedShapes(); ss.load('items/id'); await ctx.sync(); selIds = ss.items.map(s => s.id); }
    await ctx.sync();
    const fills = sh.items.map(s => { try { s.fill.load('type'); return s.fill; } catch (e) { return null; } });
    await ctx.sync();
    const rects = sh.items.map((s, k) => {
      const nm = s.name || '';
      const fixed = /^UULA (Footer|Watermark)|^CustomPageNumber_/.test(nm) || s.top >= H - 54 || s.left + s.width <= 0 || s.left >= W || (s.width > W * 0.85 && s.height > H * 0.6);
      let solid = false; try { solid = fills[k] && (fills[k].type === 'Solid' || fills[k].type === 'Gradient'); } catch (e) {}
      return { key: k, x: s.left * PT, y: s.top * PT, w: s.width * PT, h: s.height * PT, fixed, box: (s.type === 'GeometricShape' && solid && s.width > W * 0.5) || /Container$/.test(nm), kind: s.type === 'Line' ? 'line' : '' };
    });
    const only = selIds.length ? sh.items.map((s, k) => selIds.includes(s.id) ? k : -1).filter(k => k >= 0) : null;
    const mv = UR.align(rects, W * PT, mode, only);
    for (const [k, m] of Object.entries(mv)) { if (m.w != null) sh.items[+k].width = m.w / PT; if (m.x != null) sh.items[+k].left = m.x / PT; }
    await ctx.sync();
    return Object.keys(mv).length;
  }).then(async r => { await frameTop(id); return r; });
}
async function spacingRun(kind) {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  await snapOp(await targetIds().catch(() => []));
  try {
    const ids = await targetIds();
    if (!ids.length) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
    const over = [], skipped = [], spread = [];
    // two or more objects selected: only those are spaced (vertical, distribute and horizontal alike)
    const only = await run(async ctx => { const ss = ctx.presentation.getSelectedShapes(); ss.load('items/id'); await ctx.sync(); return ss.items.map(x => x.id); }).catch(() => []);
    const sel = only.length > 1 ? only : null;
    for (let k = 0; k < ids.length; k++) {
      prog(k / ids.length); say(L('تنسيق المسافات ', 'Spacing ') + (k + 1) + L(' من ', ' of ') + ids.length + '…');
      try {
        if (kind === 'v') { const r = await spaceSlide(ids[k], sel); if (r.status === 'big') over.push(ids[k]); }
        else if (kind === 'd') { const r = await distributeSlide(ids[k], sel); if (r.status === 'big') over.push(ids[k]); else if (r.status !== 'ok') skipped.push({ id: ids[k], why: r.status }); else spread.push(r.pts); }
        else await alignSlide(ids[k], hMode, scope === 'cur');
      } catch (e) { line('err', '✗ ' + esc(errText(e))); }
    }
    prog(null);
    const BIG = L('كمية المحتوى أكبر من حجم الشريحة، فلم يتحرك شيء', 'There is more content than the slide can hold, so nothing moved');
    if (over.length && ids.length > 1) { const all = await allIds(); line('warn', '• ' + BIG + ': ' + slideLinks(over, i => all.indexOf(i) + 1)); }
    if (over.length && over.length === ids.length) { say(BIG + '.'); return; }
    if (kind === 'd') {
      const all = await allIds(), num = i => all.indexOf(i) + 1;
      if (spread.length) line('ok', L('✓ وُزّعت المساحة الفارغة بالتساوي على المسافات بين أجزاء الشريحة.', '✓ The free space was shared out evenly between the parts of the slide.'));
      const why = { full: L('لا توجد مساحة فارغة', 'no free space'), none: L('فارغة', 'empty') };
      for (const w of ['full', 'none']) { const list = skipped.filter(x => x.why === w).map(x => x.id); if (list.length) line('', '• ' + why[w] + ': ' + slideLinks(list, num)); }
    }
    say(sel ? L('تم تنسيق المسافات بين ' + sel.length + ' عناصر محددة فقط.', 'Spacing done between the ' + sel.length + ' selected objects only.') : L('تم تنسيق المسافات.', 'Spacing done.'));
  } catch (e) { prog(null); say(L('حدث خطأ: ', 'Error: ') + errText(e)); }
  finally { busy = false; lock(); }
}
$('spV').addEventListener('click', () => spacingRun('v'));
$('spH').addEventListener('click', () => spacingRun('h'));
$('spD').addEventListener('click', () => spacingRun('d'));

// ---------- hide / show the solutions ----------
// an empty grid from a graph picture: the coloured drawing is removed and the grid lines / axes it covered are drawn again
// the UULA empty grid (lines A0A8B4, clear in print, on a clear background, cells of two page-grid steps), cut to w×h points
function gridPicture(wPt, hPt) {
  const S = 4, cell = 13.39 * 1.75, cv = document.createElement('canvas');
  cv.width = Math.max(4, Math.round(wPt * S)); cv.height = Math.max(4, Math.round(hPt * S));
  const g = cv.getContext('2d');
  g.strokeStyle = '#A0A8B4'; g.lineWidth = 0.75 * S;
  const step = cell * S;
  // the lines run off the edges, as a grid cut out of a bigger sheet; the cut is centred on the picture
  const ox = ((cv.width / 2) % step), oy = ((cv.height / 2) % step);
  g.beginPath();
  for (let x = ox; x <= cv.width; x += step) { g.moveTo(Math.round(x) + 0.5, 0); g.lineTo(Math.round(x) + 0.5, cv.height); }
  for (let y = oy; y <= cv.height; y += step) { g.moveTo(0, Math.round(y) + 0.5); g.lineTo(cv.width, Math.round(y) + 0.5); }
  g.stroke();
  return cv.toDataURL('image/png').split(',')[1];
}
async function emptyGrid(b64, ext) {
  const mime = /jpe?g/i.test(ext) ? 'image/jpeg' : 'image/png';
  const img = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = no; i.src = 'data:' + mime + ';base64,' + b64; });
  const w = img.naturalWidth, h = img.naturalHeight;
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
  const g = cv.getContext('2d'); g.drawImage(img, 0, 0);
  const im = g.getImageData(0, 0, w, h), d = im.data;
  const colGray = new Uint32Array(w), rowGray = new Uint32Array(h), colDark = new Uint32Array(w), rowDark = new Uint32Array(h);
  let colored = 0, transparent = 0, grayR = 0, grayN = 0;
  const kind = new Uint8Array(w * h);           // 1 coloured, 2 gray, 3 dark
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const k = (y * w + x) * 4, a = d[k + 3]; if (a < 30) { transparent++; continue; }
    const r = d[k], gg = d[k + 1], b = d[k + 2], mx = Math.max(r, gg, b), mn = Math.min(r, gg, b), av = (r + gg + b) / 3;
    if (mx - mn > 55) { kind[y * w + x] = 1; colored++; }
    else if (mx - mn < 25 && av >= 130 && av <= 238) { kind[y * w + x] = 2; colGray[x]++; rowGray[y]++; grayR += av; grayN++; }
    else if (mx - mn < 45 && av < 110) { kind[y * w + x] = 3; colDark[x]++; rowDark[y]++; }
  }
  if (colored < w * h * 0.002) return null;     // nothing drawn on it
  const bgClear = transparent > w * h * 0.3;
  const gv = grayN ? Math.round(grayR / grayN) : 190;
  const gx = new Uint8Array(w), gy = new Uint8Array(h), ax = new Uint8Array(w), ay = new Uint8Array(h);
  for (let x = 0; x < w; x++) { if (colGray[x] > h * 0.3) gx[x] = 1; if (colDark[x] > h * 0.45) ax[x] = 1; }
  for (let y = 0; y < h; y++) { if (rowGray[y] > w * 0.3) gy[y] = 1; if (rowDark[y] > w * 0.45) ay[y] = 1; }
  // a coloured pixel is erased (with a 1px halo for antialiasing), then the line under it comes back
  const erase = new Uint8Array(w * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (kind[y * w + x] === 1)
    for (let yy = Math.max(0, y - 1); yy <= Math.min(h - 1, y + 1); yy++) for (let xx = Math.max(0, x - 1); xx <= Math.min(w - 1, x + 1); xx++) if (kind[yy * w + xx] !== 3 || kind[y * w + x] === 1) erase[yy * w + xx] = 1;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    if (!erase[y * w + x]) continue;
    const k = (y * w + x) * 4;
    if (ax[x] || ay[y]) { d[k] = d[k + 1] = d[k + 2] = 0; d[k + 3] = 255; }
    else if (gx[x] || gy[y]) { d[k] = d[k + 1] = d[k + 2] = gv; d[k + 3] = 255; }
    else if (bgClear) d[k + 3] = 0;
    else { d[k] = d[k + 1] = d[k + 2] = 255; d[k + 3] = 255; }
  }
  g.putImageData(im, 0, 0);
  return cv.toDataURL('image/png').split(',')[1];
}
// empty lines for the student in the place of a hidden solution, or nothing (saved on this device)
let solLines = !!load('ur-sollines', false);
function paintSolLines() { document.querySelectorAll('#solLines button').forEach(b => b.setAttribute('aria-pressed', String((b.dataset.v === '1') === solLines))); }
document.getElementById('solLines').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; solLines = b.dataset.v === '1'; save('ur-sollines', solLines); paintSolLines(); });
paintSolLines();
// the button follows the slide on screen: its solutions hidden → «Show solutions», shown → «Hide solutions»
let solState = false;
async function slideSolHidden(id) {
  if (!id) return false;
  return run(async ctx => { const sh = ctx.presentation.slides.getItem(id).shapes; sh.load('items/name'); await ctx.sync(); return sh.items.some(x => /^UULA Sol(Hidden|View) /.test(x.name || '')); }).catch(() => false);
}
function solHidden() { return solState; }
function solPaint() { const h = solState; $('solLbl').textContent = h ? L('إظهار الحلول', 'Show solutions') : L('إخفاء الحلول', 'Hide solutions'); $('solSlash').style.display = h ? 'none' : ''; }
async function solLabel() {
  solPaint();
  if (!supported) return;
  try { solState = await slideSolHidden(await currentId()); } catch (e) {}
  solPaint();
}
// another slide chosen: the button is checked again
let solTimer = 0, solSlideId = null;
let solWatching = false;
function solWatch() {
  if (solWatching) return; solWatching = true;
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      clearTimeout(solTimer); solTimer = setTimeout(async () => {
        if (busy || isWord) return;
        const id = await currentId().catch(() => null);
        if (id && id !== solSlideId) { solSlideId = id; solLabel(); }
        selCheck();
      }, 300);
    });
  } catch (e) {}
  window.addEventListener('focus', () => { if (!busy && !isWord) solLabel(); });
}
// hide / show without rebuilding the slide (nothing moves on screen) when only whole objects go aside or come back;
// returns null when the slide needs the full way (copies, empty grids, writing lines, or objects PowerPoint keeps hidden)
async function solInPlace(id, mode, pkg) {
  const U = UR.util, W = pkg.slideW / PT;
  const doc = UR.parseXml(pkg.files.slide);
  const tops = U.topItems(doc).map(el => { const c = U.core(el), n = U.cnvOf(c); return { el, n, name: n ? n.getAttribute('name') || '' : '', g: U.geomOf(el, pkg) }; });
  if (mode === 'show') {
    if (tops.some(t => /^UULA SolHidden /.test(t.name) && [...t.el.getElementsByTagNameNS(UR.NS.p, 'cNvPr')].some(n => n.getAttribute('hidden') === '1'))) return null;
    return run(async ctx => {
      const sh = ctx.presentation.slides.getItem(id).shapes; sh.load('items/name,items/left'); await ctx.sync();
      let n = 0;
      for (const s of sh.items) {
        if (/^UULA SolView /.test(s.name || '')) s.delete();
        else if (/^UULA SolHidden /.test(s.name || '')) { s.name = s.name.replace(/^UULA SolHidden /, ''); if (s.left >= 2.5 * W) s.left = s.left - 3 * W; n++; }
      }
      await ctx.sync(); return n;
    });
  }
  if (solLines) return null;
  const found = UR.findSolutions(pkg);
  if (!found.length || found.some(f => f.how !== 'hide')) return null;
  const want = found.map(f => tops[f.i]).filter(Boolean);
  if (want.length !== found.length || want.some(t => !t.g)) return null;
  return run(async ctx => {
    const sh = ctx.presentation.slides.getItem(id).shapes; sh.load('items/name,items/left,items/top'); await ctx.sync();
    const used = new Set(), match = [];
    for (const t of want) {
      const s = sh.items.find(x => !used.has(x) && (x.name || '') === t.name && Math.abs(x.left - t.g.x / PT) < 1.5 && Math.abs(x.top - t.g.y / PT) < 1.5);
      if (!s) return null;                                  // not sure which object it is: the full way
      used.add(s); match.push(s);
    }
    match.forEach(s => { s.name = 'UULA SolHidden ' + s.name; s.left = s.left + 3 * W; });
    await ctx.sync(); return match.length;
  });
}
async function solSlide(id, mode) {
  const pkg = await UR.openPackage(await exportSlide(id));
  if (mode === 'show' ? UR.hasSolutions(pkg.files.slide) : !UR.hasSolutions(pkg.files.slide)) {
    const quick = await solInPlace(id, mode, pkg).catch(() => null);
    if (quick != null) return quick;
  }
  let r;
  if (mode === 'show') { if (!UR.hasSolutions(pkg.files.slide)) return 0; r = UR.showSolutions(pkg); r.rels = null; }
  else {
    if (UR.hasSolutions(pkg.files.slide)) return 0;
    const found = UR.findSolutions(pkg);
    if (!found.length) return 0;
    const grids = {};
    for (const f of found.filter(x => x.how === 'grid')) {
      // marked «حل بياني»: the empty grid at its size, whatever the object is
      if (f.pinned && f.w && f.h) { grids[f.i] = gridPicture(f.w, f.h); continue; }
      const file = f.media && pkg.zip.file(f.media); if (!file || !/\.(png|jpe?g)$/i.test(f.media)) continue;
      try {
        // a graph that is part of a solution: in its place the UULA empty grid, cut to the picture's size
        const drawn = await emptyGrid(await file.async('base64'), f.media);
        if (drawn) grids[f.i] = f.w && f.h ? gridPicture(f.w, f.h) : drawn;
      } catch (e) {}
    }
    r = UR.hideSolutions(pkg, grids, { lines: solLines });
  }
  if (!r.count) return 0;
  const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), id);
  await deleteSlide(id);
  solNew[id] = newId;
  return r.count;
}
let solNew = {};
$('solBtn').addEventListener('click', async () => {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  const ids0 = await selIds().catch(() => []);
  await snapOp(ids0);
  solState = await slideSolHidden(await currentId().catch(() => null));
  const mode = solState ? 'show' : 'hide';
  solNew = {};
  try {
    const ids = ids0;
    if (!ids.length) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
    let total = 0, slides = 0;
    for (let k = 0; k < ids.length; k++) {
      prog(k / ids.length); say((mode === 'hide' ? L('إخفاء الحلول ', 'Hiding solutions ') : L('إظهار الحلول ', 'Showing solutions ')) + (k + 1) + L(' من ', ' of ') + ids.length + '…');
      try { const n = await solSlide(ids[k], mode); if (n) { total += n; slides++; } }
      catch (e) { line('err', '✗ ' + esc(errText(e))); }
    }
    prog(null);
    // the same slides stay selected (a rebuilt slide in place of the old one)
    const keep = ids.map(i => solNew[i] || i);
    await run(async ctx => { ctx.presentation.setSelectedSlides(keep); await ctx.sync(); }).catch(() => {});
    await solLabel();
    say(total ? (mode === 'hide' ? L('أُخفيت الحلول في ', 'Solutions hidden in ') : L('ظهرت الحلول في ', 'Solutions shown in ')) + cnt(slides) + '.' : (mode === 'hide' ? L('لم أجد حلولًا لإخفائها.', 'No solutions found to hide.') : L('لا توجد حلول مخفية.', 'No hidden solutions.')));
  } catch (e) { prog(null); say(L('حدث خطأ: ', 'Error: ') + errText(e)); }
  finally { busy = false; lock(); }
});

// ---------- Word file → slides ----------
// every tagged part of the Word file becomes its own text box (named "UULA:<type>"), then the slide is restyled
function slideEmpty(pkg) {
  const doc = UR.parseXml(pkg.files.slide), U = UR.util;
  return !U.topItems(doc).some(el => { const c = U.core(el), n = U.cnvOf(c), nm = n ? n.getAttribute('name') || '' : ''; if (/^CustomPageNumber_/.test(nm)) return false; return c.localName !== 'sp' || /\S/.test([...c.getElementsByTagNameNS(UR.NS.a, 't')].map(t => t.textContent).join('')); });
}
let docxPending = null;
async function importDocx(file) {
  res.innerHTML = '';
  say(L('قراءة ملف الوورد…', 'Reading the Word file…'));
  const parsed = await UW.read(await file.arrayBuffer());
  const units = parsed.units;
  const rp0 = parsed.report;
  if (!units.length) {
    say(rp0.skippedParts ? L('لا يوجد في الملف أي جزء مميز. ميّز الأجزاء في الوورد أولًا.', 'Nothing in the file is tagged. Tag the parts in Word first.') : L('لم أجد محتوى في الملف.', 'No content found in the file.'));
    return;
  }
  const id = await currentId();
  if (!id) { say(L('حدّد شريحة أولًا.', 'Select a slide first.')); return; }
  const b64 = await exportSlide(id);
  const pkg0 = await UR.openPackage(b64);
  const parts = UW.batches(units, pkg0.slideW, pkg0.slideH);
  const empty = slideEmpty(pkg0);
  const adds = parts.length - (empty ? 1 : 0);
  const rp = parsed.report;
  line('', esc(L('عدد الأجزاء المميزة التي ستُنقل: ', 'Tagged parts to bring over: ') + units.length + '.'));
  if (rp.skippedParts) line('warn', L('• في ملف الوورد أجزاء غير مميزة (عددها ' + rp.skippedParts + ')، تجاهلتها ولن تُنقل. لرؤيتها: في الوورد افتح «مراجعة» واضغط «لوّن غير المميز».', '• ' + rp.skippedParts + ' untagged parts of the Word file are skipped. To see them: in Word open Review and press Colour untagged.')
    + (rp.skippedText.length ? '<br><span class="note">' + rp.skippedText.slice(0, 4).map(x => '«' + esc(x) + '…»').join('، ') + '</span>' : ''));
  if (adds > 0) {
    docxPending = { units, parts, b64, id, empty, report: rp };
    const d = line('warn splitask', L('سيُضاف تقريبًا ', 'About ') + cnt(adds) + L(' بعد الشريحة الحالية' + (empty ? ' (والشريحة الحالية الفارغة تُستخدم أولًا)' : '') + '.', ' will be added after the current slide' + (empty ? ' (the empty current slide is used first)' : '') + '.')
      + '<br><span style="display:flex;gap:6px;margin-top:4px"><button type="button" class="mini" data-dx="yes">' + L('موافق، أضفها', 'OK, add them') + '</button><button type="button" class="mini" data-dx="no">' + L('إلغاء', 'Cancel') + '</button></span>');
    d.style.display = 'block';
    d.querySelector('[data-dx="yes"]').addEventListener('click', () => { d.remove(); runDocx(); });
    d.querySelector('[data-dx="no"]').addEventListener('click', () => { d.remove(); docxPending = null; say(L('أُلغي الإدراج.', 'Cancelled.')); });
    say(L('بانتظار موافقتك.', 'Waiting for your OK.'));
    return;
  }
  docxPending = { units, parts, b64, id, empty, report: rp };
  await runDocx(true);
}
async function runDocx(already) {
  const st = docxPending; if (!st) return;
  if (!already) { if (busy) return; busy = true; lock(); }
  try {
    let prev = st.id, made = [], mediaKey = 1, first = true;
    for (let k = 0; k < st.parts.length; k++) {
      prog(k / st.parts.length); say(L('إنشاء الشريحة ', 'Building slide ') + (k + 1) + L(' من ', ' of ') + st.parts.length + '…');
      const pkg = await UR.openPackage(st.b64);
      UW.fillSlide(pkg, st.parts[k], { clear: true, mediaKey });
      mediaKey += st.parts[k].length + 1;
      const r = UR.restyle(pkg, {});
      r.xml = UW.syncFallback(r.xml);
      const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), prev);
      if (first && st.empty) await deleteSlide(st.id);
      first = false;
      const lay = await refine(newId, r.plan);
      await unmark(newId);
      // still longer than the page: split it (already agreed)
      if (lay.overflow >= 0 && UR.splitPoint(r.plan, lay) > 0) {
        pendingSplit.set(newId, { pkg, rels: r.rels, xml: r.xml, plan: r.plan });
        const ids = await splitSlide(newId);
        made.push(...ids); prev = ids[ids.length - 1];
      } else { made.push(newId); prev = newId; }
    }
    prog(null);
    if ($('numChk').checked) { try { await renumber(); } catch (e) {} }
    await run(async ctx => { ctx.presentation.setSelectedSlides([made[0]]); await ctx.sync(); }).catch(() => {});
    const all = await allIds(), num = i => all.indexOf(i) + 1;
    line('ok', L('✓ أُدرج ملف الوورد في ', '✓ The Word file was inserted in ') + cnt(made.length) + ': ' + slideLinks(made, num));
    if (st.report.inlineObj) line('warn', L('• ' + st.report.inlineObj + ' معادلة كصورة داخل السطر وُضعت كصورة مستقلة، راجعها.', '• ' + st.report.inlineObj + ' inline equation pictures were placed as separate pictures; check them.'));
    line('', '<span class="note">' + L('يلزم تثبيت خط UULA Sans على الجهاز.', 'The UULA Sans font must be installed.') + '</span>');
    say(L('تم إدراج ملف الوورد.', 'The Word file was inserted.'));
  } catch (e) { prog(null); say(L('تعذّر الإدراج: ', 'Could not insert: ') + errText(e)); }
  finally { docxPending = null; busy = false; lock(); }
}
$('docxBtn').addEventListener('click', () => { if (!busy && supported) $('docxFile').click(); });
$('docxFile').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = '';
  if (!f || busy) return;
  busy = true; lock();
  try { await importDocx(f); }
  catch (e2) { prog(null); say(L('تعذّر قراءة الملف: ', 'Could not read the file: ') + errText(e2)); }
  finally { busy = false; lock(); }
});

// ---------- Word: tagging ----------
let isWord = false;
const wtype = id => UW.TYPE[id];
const wname = id => { const t = wtype(id); return t ? L(t.ar, t.en) : id; };
function buildWordRoles() {
  const host = $('wroles'); host.innerHTML = '';
  for (const g of UW.TYPES) {
    const { d, grid } = groupCard('word', g.g, L(g.g, g.ge));
    for (const [id, ar, en, col] of g.items) {
      const b = document.createElement('button'); b.type = 'button'; b.className = 'tbtn'; b.dataset.wrole = id;
      b.title = L(ar, en); b.innerHTML = iconHtml(id, col) + '<span>' + L(ar, en) + '</span>'; tint(b, id);
      grid.appendChild(b);
    }
    d.appendChild(grid); host.appendChild(d);
  }
}
function styleCC(cc, type) {
  const t = wtype(type);
  cc.tag = 'UULA:' + type; cc.title = L(t.ar, t.en); cc.color = '#' + t.col; cc.appearance = 'BoundingBox';
}
const isRedHex = h => { const m = /^#?([0-9A-F]{2})([0-9A-F]{2})([0-9A-F]{2})$/i.exec(h || ''); if (!m) return false; const [r, g, b] = m.slice(1).map(x => parseInt(x, 16)); return r > 170 && g < 90 && b < 90; };
const isUula = cc => cc && !cc.isNullObject && /^UULA:/.test(cc.tag || '');
// an empty line between two tagged parts that touch (so their frames do not stick together)
function spacer(par, where) {
  const s = par.insertParagraph('', where);
  try { s.spaceAfter = 0; s.spaceBefore = 0; s.font.size = 6; } catch (e) {}
  return s;
}
async function tagSelection(type) {
  return Word.run(async ctx => {
    const sel = ctx.document.getSelection();
    sel.load('text,isEmpty');
    const pcc = sel.parentContentControlOrNullObject; pcc.load('isNullObject,tag,text');
    await ctx.sync();
    // inside a tagged part with nothing (or all of it) selected: change its type
    if (isUula(pcc) && (sel.isEmpty || sel.text.trim() === (pcc.text || '').trim())) { styleCC(pcc, type); await ctx.sync(); return 'retag'; }
    // the cursor (or a selection) inside a text box / shape: the whole box is tagged
    let pb = null;
    try { pb = sel.parentBody; pb.load('type'); await ctx.sync(); } catch (e) { pb = null; }
    if (pb && pb.type && !['MainDoc', 'TableCell', 'Header', 'Footer', 'Section', 'Footnote', 'Endnote'].includes(pb.type)) {
      const all = pb.getRange('Content'); all.load('text'); await ctx.sync();
      const whole = sel.isEmpty || sel.text.trim().length >= all.text.trim().length - 1;
      if (whole && all.text.trim()) {
        const cc = all.insertContentControl(); styleCC(cc, type); await ctx.sync();
        return 'box';
      }
    }
    if (sel.isEmpty || !sel.text.trim()) {
      // the cursor in a paragraph: that whole paragraph
      const ps = sel.paragraphs; ps.load('items/text'); await ctx.sync();
      if (!ps.items.length || !ps.items[0].text.trim()) return 'empty';
    }
    const ps = sel.paragraphs; ps.load('items/text,items/tableNestingLevel'); await ctx.sync();
    const first = ps.items[0], last = ps.items[ps.items.length - 1];
    const partial = !sel.isEmpty && ps.items.length === 1 && sel.text.trim().length < first.text.trim().length - 1;
    let range;
    if (partial) range = sel;                       // part of one paragraph: an inline frame
    else if (first.tableNestingLevel > 0 && ps.items.every(p => p.tableNestingLevel > 0)) { const tb = first.parentTable; range = tb.getRange('Whole'); }
    else {
      const prev = first.getPreviousOrNullObject(), next = last.getNextOrNullObject();
      prev.load('isNullObject,text'); next.load('isNullObject,text');
      await ctx.sync();
      const pc = prev.isNullObject ? null : prev.parentContentControlOrNullObject, nc = next.isNullObject ? null : next.parentContentControlOrNullObject;
      if (pc) pc.load('isNullObject,tag'); if (nc) nc.load('isNullObject,tag');
      await ctx.sync();
      const inPrev = isUula(pc) && !(isUula(pcc) && pc.tag === pcc.tag), inNext = isUula(nc) && !(isUula(pcc) && nc.tag === pcc.tag);
      if (inPrev && prev.text.trim()) spacer(first, 'Before');
      if (inNext && next.text.trim()) spacer(last, 'After');
      range = first.getRange('Whole').expandTo(last.getRange('Whole'));
    }
    const cc = range.insertContentControl();
    styleCC(cc, type);
    cc.font.load('highlightColor');
    await ctx.sync();
    // the "not tagged" colour goes once the part is tagged
    if ((cc.font.highlightColor || '').toUpperCase() === UW.UNTAGGED) { cc.font.highlightColor = null; await ctx.sync(); }
    return partial ? 'inline' : 'new';
  });
}
async function untagSelection() {
  return Word.run(async ctx => {
    const sel = ctx.document.getSelection();
    const pcc = sel.parentContentControlOrNullObject; pcc.load('isNullObject,tag');
    const inside = sel.contentControls; inside.load('items/tag');
    await ctx.sync();
    let n = 0;
    for (const c of inside.items) if (/^UULA:/.test(c.tag || '')) { c.delete(true); n++; }
    if (isUula(pcc)) { pcc.delete(true); n++; }
    await ctx.sync();
    return n;
  });
}
async function autoTag() {
  return Word.run(async ctx => {
    const ps = ctx.document.body.paragraphs;
    ps.load('items/text,items/style,items/tableNestingLevel');
    await ctx.sync();
    const extra = ps.items.map(p => {
      const cc = p.parentContentControlOrNullObject; cc.load('isNullObject,tag');
      const pics = p.inlinePictures; pics.load('items/width');
      p.font.load('color,bold,size');
      return { cc, pics };
    });
    await ctx.sync();
    const paras = ps.items.map((p, i) => {
      const e = extra[i];
      const tagged = !e.cc.isNullObject;
      const table = p.tableNestingLevel > 0;
      const pic = e.pics.items.length > 0;
      const o = { text: p.text, style: p.style, color: p.font.color || '', bold: p.font.bold === true, size: p.font.size || 0, table, pic, big: pic && Math.max(...e.pics.items.map(x => x.width)) > 230 };
      o.type = tagged ? null : UW.guess(o);
      o.skip = tagged;
      return o;
    });
    // table paragraphs: one group per table
    const groups = [];
    const plain = paras.map((p, i) => p.table ? Object.assign({}, p, { skip: true }) : p);
    UW.group(plain).forEach(g => groups.push(g));
    for (let i = 0; i < paras.length; i++) {
      if (!paras[i].table || paras[i].skip) continue;
      let j = i; while (j + 1 < paras.length && paras[j + 1].table) j++;
      groups.push({ type: 'table', from: i, to: j, table: true });
      i = j;
    }
    groups.sort((a, b) => a.from - b.from);
    // ranges first (before anything moves), then spacers, then the frames
    const plans = groups.map(g => ({ g, range: g.table ? ps.items[g.from].parentTable.getRange('Whole') : ps.items[g.from].getRange('Whole').expandTo(ps.items[g.to].getRange('Whole')) }));
    let prevEnd = -10, prevTagged = false;
    for (let k = 0; k < plans.length; k++) {
      const g = plans[k].g;
      const touch = g.from === prevEnd + 1 || (g.from > 0 && paras[g.from - 1].skip && paras[g.from - 1].text.trim() && !paras[g.from - 1].table);
      if (touch && !g.table) spacer(ps.items[g.from], 'Before');
      // a part tagged before, right after this one
      const nx = paras[g.to + 1];
      if (nx && nx.skip && (nx.text.trim() || nx.table) && !g.table && !(plans[k + 1] && plans[k + 1].g.from === g.to + 1)) spacer(ps.items[g.to], 'After');
      prevEnd = g.to;
    }
    let n = 0;
    for (const pl of plans) { const cc = pl.range.insertContentControl(); styleCC(cc, pl.g.type); n++; if (n % 40 === 0) await ctx.sync(); }
    await ctx.sync();
    // answers of a word or two written inside a question in another colour (red / blue): an inline «حل» frame
    const qpars = [];
    for (const pl of plans) if (['question', 'choice'].includes(pl.g.type)) for (let i = pl.g.from; i <= pl.g.to; i++) if (paras[i].text.trim()) qpars.push(ps.items[i]);
    const words = qpars.map(p => { const w = p.getTextRanges([' '], false); w.load('items/text,items/font/color'); return w; });
    await ctx.sync();
    for (const w of words) {
      const it = w.items; if (it.length < 2) continue;
      const base = (it[0].font.color || '').toUpperCase();
      const odd = x => { const c = (x.font.color || '').toUpperCase(); return c && c !== base && (UW.isBlue(c) || isRedHex(c)); };
      for (let a = 0; a < it.length; a++) {
        if (!odd(it[a])) continue;
        let b = a; while (b + 1 < it.length && odd(it[b + 1])) b++;
        const cc = it[a].expandTo(it[b]).insertContentControl(); styleCC(cc, 'answer'); n++;
        a = b;
      }
    }
    await ctx.sync();
    return n;
  });
}
async function clearAllTags() {
  return Word.run(async ctx => {
    const ccs = ctx.document.contentControls; ccs.load('items/tag'); await ctx.sync();
    let n = 0;
    for (const c of ccs.items) if (/^UULA:/.test(c.tag || '')) { c.delete(true); n++; }
    await ctx.sync();
    return n;
  });
}
// content that is not tagged: counted, and coloured on request so it stands out
async function untaggedParas(ctx) {
  const ps = ctx.document.body.paragraphs; ps.load('items/text,items/tableNestingLevel'); await ctx.sync();
  const ex = ps.items.map(p => { const cc = p.parentContentControlOrNullObject; cc.load('isNullObject,tag'); const pics = p.inlinePictures; pics.load('items/width'); p.font.load('highlightColor'); return { cc, pics }; });
  await ctx.sync();
  // tables count as tagged (their content is text; solutions inside are tagged by hand)
  return ps.items.filter((p, i) => (p.text.trim() || ex[i].pics.items.length) && !(isUula(ex[i].cc)) && !(p.tableNestingLevel > 0));
}
async function markUntagged(on) {
  return Word.run(async ctx => {
    if (on) { const list = await untaggedParas(ctx); list.forEach(p => { p.font.highlightColor = UW.UNTAGGED; }); await ctx.sync(); return list.length; }
    const ps = ctx.document.body.paragraphs; ps.load('items/font/highlightColor'); await ctx.sync();
    let n = 0; for (const p of ps.items) if ((p.font.highlightColor || '').toUpperCase() === UW.UNTAGGED) { p.font.highlightColor = null; n++; }
    await ctx.sync(); return n;
  });
}
async function reviewList() {
  const items = await Word.run(async ctx => {
    const ccs = ctx.document.contentControls; ccs.load('items/id,items/tag,items/text'); await ctx.sync();
    return ccs.items.filter(c => /^UULA:[a-z]+$/.test(c.tag || '')).map(c => ({ id: c.id, type: c.tag.slice(5), text: (c.text || '').replace(/\s+/g, ' ').trim().slice(0, 70) }));
  });
  const list = $('wlist'), sum = $('wsum');
  list.innerHTML = ''; sum.innerHTML = '';
  const counts = {};
  for (const it of items) {
    counts[it.type] = (counts[it.type] || 0) + 1;
    const t = wtype(it.type) || { col: '999999' };
    const d = document.createElement('div'); d.className = 'witem'; d.dataset.id = it.id;
    d.innerHTML = '<i style="background:#' + t.col + '"></i><b>' + esc(wname(it.type)) + '</b><span>' + esc(it.text || '—') + '</span><button type="button" class="x" title="' + L('إزالة التمييز', 'Remove tag') + '">×</button>';
    list.appendChild(d);
  }
  for (const [k, v] of Object.entries(counts)) { const s = document.createElement('span'); s.textContent = wname(k) + ' ' + v; sum.appendChild(s); }
  const un = await Word.run(async ctx => (await untaggedParas(ctx)).length).catch(() => 0);
  $('wwarn').hidden = !un;
  $('wwarnT').textContent = L('⚠ فقرات غير مميزة (عددها ' + un + ')، ولن تُنقل إلى الباوربوينت.', '⚠ ' + un + ' paragraphs are not tagged and will not be brought into PowerPoint.');
  if (!items.length) list.innerHTML = '<span class="note">' + L('لا توجد أجزاء مميزة بعد.', 'No tagged parts yet.') + '</span>';
}
$('wlist').addEventListener('click', e => {
  const it = e.target.closest('.witem'); if (!it) return;
  const id = +it.dataset.id;
  if (e.target.closest('.x')) {
    Word.run(async ctx => { ctx.document.contentControls.getById(id).delete(true); await ctx.sync(); }).then(reviewList).catch(er => say(errText(er)));
    return;
  }
  Word.run(async ctx => { ctx.document.contentControls.getById(id).select('Select'); await ctx.sync(); }).catch(() => {});
});
$('wroles').addEventListener('click', async e => {
  const b = e.target.closest('.tbtn'); if (!b || busy || !supported) return;
  busy = true; lock();
  try {
    const r = await tagSelection(b.dataset.wrole);
    if (r === 'empty') say(L('حدّد المحتوى أولًا.', 'Select the content first.'));
    else say((r === 'retag' ? L('تغيّر النوع إلى: ', 'Type changed to: ') : r === 'box' ? L('مُيّز مربع النص كاملًا كـ: ', 'The whole text box is tagged as: ') : L('مُيّز كـ: ', 'Tagged as: ')) + wname(b.dataset.wrole));
  } catch (er) { say(L('تعذّر التمييز: ', 'Could not tag: ') + errText(er)); }
  finally { busy = false; lock(); }
});
$('wUntag').addEventListener('click', async () => {
  if (busy || !supported) return; busy = true; lock();
  try { const n = await untagSelection(); say(n ? L('أُزيل التمييز.', 'Tag removed.') : L('لا يوجد تمييز في المحدد.', 'Nothing tagged in the selection.')); }
  catch (er) { say(errText(er)); } finally { busy = false; lock(); }
});
$('wAuto').addEventListener('click', async () => {
  if (busy || !supported) return; busy = true; lock();
  try { say(L('تمييز المستند…', 'Tagging the document…')); const n = await autoTag(); say(n ? L('مُيّز ', 'Tagged ') + n + L(' جزءًا. راجعها في «مراجعة».', ' parts. Check them in Review.') : L('كل المحتوى مميز.', 'Everything is already tagged.')); }
  catch (er) { say(L('تعذّر التمييز التلقائي: ', 'Auto-tag failed: ') + errText(er)); } finally { busy = false; lock(); }
});
$('wMark').addEventListener('click', async () => {
  if (busy || !supported) return; busy = true; lock();
  try { const n = await markUntagged(true); say(n ? L('لُوّنت الفقرات غير المميزة (' + n + ') بالأزرق الفاتح.', n + ' untagged paragraphs are coloured light blue.') : L('كل المحتوى مميز.', 'Everything is tagged.')); }
  catch (er) { say(errText(er)); } finally { busy = false; lock(); }
});
$('wUnmark').addEventListener('click', async () => {
  if (busy || !supported) return; busy = true; lock();
  try { await markUntagged(false); say(L('أُزيل التلوين.', 'Colouring removed.')); }
  catch (er) { say(errText(er)); } finally { busy = false; lock(); }
});
$('wRefresh').addEventListener('click', () => { if (supported) reviewList().catch(er => say(errText(er))); });
$('wClear').addEventListener('click', () => {
  if (busy || !supported) return;
  res.innerHTML = '';
  const d = line('warn', L('إزالة كل التمييز من المستند؟ (يبقى النص كما هو)', 'Remove all tags from the document? (the text stays)')
    + '<br><span style="display:flex;gap:6px;margin-top:4px"><button type="button" class="mini" data-c="yes">' + L('نعم، أزلها', 'Yes, remove') + '</button><button type="button" class="mini" data-c="no">' + L('لا', 'No') + '</button></span>');
  d.style.display = 'block';
  d.querySelector('[data-c="no"]').addEventListener('click', () => d.remove());
  d.querySelector('[data-c="yes"]').addEventListener('click', async () => {
    d.remove(); busy = true; lock();
    try { const n = await clearAllTags(); say(L('أُزيل ', 'Removed ') + n + L(' تمييز.', ' tags.')); await reviewList(); }
    catch (er) { say(errText(er)); } finally { busy = false; lock(); }
  });
});

// ---------- main button ----------
function lock() {
  $('goBtn').disabled = busy || !supported;
  $('solBtn').disabled = busy || !supported;
  $('spV').disabled = busy || !supported; $('spH').disabled = busy || !supported;
  document.querySelectorAll('.tbtn').forEach(b => { b.disabled = busy || !supported; });
  undoPaint(); opPaint();
  ['docxBtn', 'numDel', 'wAuto', 'wUntag', 'wClear', 'wMark', 'wUnmark'].forEach(k => { if ($(k)) $(k).disabled = busy || !supported; });
}
$('goBtn').addEventListener('click', async () => {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  await snapOp(await targetIds().catch(() => []));
  try { await doConvert(); }
  catch (e) { prog(null); say(L('حدث خطأ: ', 'Error: ') + errText(e)); }
  finally { busy = false; lock(); }
});

// tabs
function setTab(t) {
  document.querySelectorAll('#tabs button, #wtabs button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.tab === t)));
  document.querySelectorAll('[data-pane]').forEach(s => { s.hidden = s.dataset.pane !== t; });
  save(isWord ? 'ur-wtab' : 'ur-tab', t);
  if (t === 'wrev' && supported) reviewList().catch(() => {});
}
$('tabs').addEventListener('click', e => { const b = e.target.closest('button[data-tab]'); if (b) setTab(b.dataset.tab); });
$('wtabs').addEventListener('click', e => { const b = e.target.closest('button[data-tab]'); if (b) setTab(b.dataset.tab); });
setTab((t => t === 'conv' || t === 'spc' ? 'obj' : t)(load('ur-tab', 'obj')));  // the old تحويل and تنسيق المسافات tabs are now inside تنسيق
// language
function applyLang() {
  document.documentElement.lang = lang; document.documentElement.dir = lang === 'en' ? 'ltr' : 'rtl';
  document.querySelectorAll('[data-en]').forEach(el => { if (el.dataset.ar == null) el.dataset.ar = el.textContent; el.textContent = lang === 'en' ? el.dataset.en : el.dataset.ar; });
  document.querySelectorAll('[data-en-title]').forEach(el => { el.title = lang === 'en' ? el.dataset.enTitle : el.dataset.arTitle; });
  $('langBtn').textContent = lang === 'en' ? 'ع' : 'EN';
  buildRoleList(); buildSolRoles(); buildWordRoles(); solLabel(); lock(); if (!isWord) solWatch();
  if (isWord && supported && !$('wlist').closest('[data-pane]').hidden) reviewList().catch(() => {});
}
$('langBtn').addEventListener('click', () => { lang = lang === 'en' ? 'ar' : 'en'; save('ur-lang', lang); applyLang(); say(''); });

// ---------- menu / about ----------
function menu(o) { $('appMenu').hidden = !o; $('menuBtn').setAttribute('aria-expanded', String(o)); }
$('menuBtn').addEventListener('click', e => { e.stopPropagation(); menu($('appMenu').hidden); });
document.addEventListener('click', e => { if (!$('appMenu').hidden && !e.target.closest('.appmenu')) menu(false); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') { menu(false); $('modal').hidden = true; } });
function modal(t, html) { $('modalT').textContent = t; $('modalB').innerHTML = html; $('modal').hidden = false; }
$('modalX').addEventListener('click', () => { $('modal').hidden = true; });
$('modal').addEventListener('click', e => { if (e.target === $('modal')) $('modal').hidden = true; });
$('aboutBtn').addEventListener('click', () => {
  menu(false);
  modal(L('حول الأداة', 'About'), (lang === 'en'
    ? '<p>Restyles A4 slides into the UULA study-guide style with one click, and inserts the ready study-guide objects.</p><ul><li>UULA Sans font, 11pt text on the page grid</li><li>Headings, questions and answers in UULA colors</li><li>Full-width containers, tables and pictures in one style</li><li>Hide / show the solutions</li><li>In Word: tag the parts (question, solution, definition…), then insert the Word file into PowerPoint</li></ul><p>No content is deleted; only old decoration is removed.</p>'
    : '<p>تحوّل شرائح A4 إلى ستايل مذكرات علا بضغطة زر، وتضيف قوالب المذكرة الجاهزة.</p><ul><li>خط UULA Sans، والنص 11 على شبكة الصفحة</li><li>العناوين والأسئلة والإجابات بألوان علا</li><li>إطارات بعرض الصفحة، وجداول وصور بالشكل الموحّد</li><li>إخفاء الحلول وإظهارها</li><li>في الوورد: تمييز الأجزاء (سؤال، حل، تعريف…) ثم إدراج ملف الوورد في الباوربوينت</li></ul><p>لا يُحذف أي محتوى؛ تُحذف الزخارف القديمة فقط.</p>')
    + '<p class="foot" dir="ltr">Developed by <b>Zaher Alkablawi</b> · Version ' + VER + '<br>© 2026 UULA. All rights reserved.</p>');
});


// ---------- updates (same method as UULA Arabic Math) ----------
(function () {
  const head = () => fetch('app.html?chk=' + Date.now(), { method: 'HEAD', cache: 'no-store' })
    .then(r => r.ok ? (r.headers.get('etag') || r.headers.get('last-modified')) : null).catch(() => null);
  let base = null, last = 0;
  const note = m => { if (busy) return; say(m); setTimeout(() => { if (status.textContent === m) say(''); }, 5000); };
  async function check(manual) {
    last = Date.now(); const t = await head();
    if (!base) { base = t; if (manual) note(t ? L('لديك آخر إصدار.', 'You have the latest version.') : L('تعذّر الاتصال.', 'Could not connect.')); return; }
    if (t && t !== base) { $('updBtn').hidden = false; if (manual) note(L('يوجد تحديث.', 'An update is available.')); }
    else if (manual) note(t ? L('لديك آخر إصدار.', 'You have the latest version.') : L('تعذّر الاتصال.', 'Could not connect.'));
  }
  head().then(t => { base = t; });
  setInterval(() => check(false), 60 * 60 * 1000);
  const soon = () => { if (Date.now() - last > 60 * 60 * 1000) check(false); };
  window.addEventListener('focus', soon); document.addEventListener('visibilitychange', () => { if (!document.hidden) soon(); });
  $('updBtn').addEventListener('click', () => { if (!busy) location.replace('app.html?v=' + Date.now()); });
  $('chkUpd').addEventListener('click', () => { menu(false); check(true); });
})();

// ---------- start ----------
$('appVer').textContent = VER;
applyLang();
const ready = (window.Office && Office.onReady) ? Office.onReady() : Promise.resolve({});
ready.then(info => {
  if (info && info.host === Office.HostType.Word) {
    isWord = true;
    $('tabs').hidden = true; $('wtabs').hidden = false;
    let wt = load('ur-wtab', 'wtag'); if (!['wtag', 'wrev'].includes(wt)) wt = 'wtag';
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) { setTab(wt); say(L('حدّث برنامج Word لاستخدام الأداة.', 'Update Word to use this tool.')); return; }
    supported = true; lock(); setTab(wt);
    return;
  }
  if (!window.PowerPoint || (info && info.host && info.host !== Office.HostType.PowerPoint)) { say(L('هذه الأداة تعمل داخل PowerPoint وWord فقط.', 'This tool works inside PowerPoint and Word only.')); return; }
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) { say(L('حدّث برنامج PowerPoint لاستخدام الأداة.', 'Update PowerPoint to use this tool.')); return; }
  supported = true;
  const s = docSettings();
  if (s) {
    const st = s.get('urNumStart'); if (st) $('numStart').value = st;
    const md = s.get('urMcqDir'); if (md) setMcqDir(md);
    const on = s.get('urNum'); if (on === false) $('numChk').checked = false;
  }
  $('numRow').style.opacity = $('numChk').checked ? '' : '.5';
  solLabel();
  lock();
});
})();

(() => {
'use strict';
const $ = id => document.getElementById(id);
const VER = '1.0';
const PT = UR.PT;
const save = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
const load = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } };
let scope = load('ur-scope', 'cur'); if (scope !== 'all') scope = 'cur';
let busy = false, supported = false;

// ---------- status / results ----------
const status = $('status'), res = $('res');
const say = (m, cls) => { status.textContent = m || ''; status.className = 'status' + (cls ? ' ' + cls : ''); };
const tsay = m => { $('tstatus').textContent = m || ''; };
const prog = f => { $('prog').hidden = f == null; if (f != null) $('prog').firstElementChild.style.width = Math.round(f * 100) + '%'; };
function line(cls, html) { const d = document.createElement('div'); d.className = 'ln ' + cls; d.innerHTML = html; res.appendChild(d); return d; }
const slideLinks = (ids, num) => ids.map(id => '<a data-id="' + id + '">' + num(id) + '</a>').join('، ');
res.addEventListener('click', e => {
  const a = e.target.closest('a[data-id]'); if (!a) return;
  PowerPoint.run(async ctx => { ctx.presentation.setSelectedSlides([a.dataset.id]); await ctx.sync(); }).catch(() => {});
});
const errText = e => { const loc = e && e.debugInfo && e.debugInfo.errorLocation; return (e && (e.code || e.message) || String(e)) + (loc ? ' (' + loc + ')' : ''); };
const cnt = n => n === 1 ? 'شريحة واحدة' : n === 2 ? 'شريحتان' : n <= 10 ? n + ' شرائح' : n + ' شريحة';
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
  return run(async ctx => {
    const shapes = ctx.presentation.slides.getItem(slideId).shapes;
    shapes.load('items/name'); await ctx.sync();
    for (const s of shapes.items) if (/ #ur\d+$/.test(s.name || '')) s.name = s.name.replace(/ #ur\d+$/, '');
    await ctx.sync();
  }).catch(() => {});
}

// ---------- convert ----------
const MAX_PARTS = 6;
async function convertOne(id, opts, rep) {
  const b64 = await exportSlide(id);
  const pkg = await UR.openPackage(b64);
  let r;
  try { r = UR.restyle(pkg, { watermark: opts.watermark }); }
  catch (e) { throw new Error('تحليل الشريحة: ' + errText(e)); }
  const queue = [{ xml: r.xml, plan: r.plan }];
  let prev = id, made = [], parts = 0;
  while (queue.length) {
    const part = queue.shift();
    const pb = await UR.buildPackage(pkg, part.xml, r.rels);
    let newId;
    try { newId = await insertAfter(pb, prev); }
    catch (e) { throw new Error('إدراج الشريحة: ' + errText(e)); }
    const lay = await refine(newId, part.plan);
    if (lay.overflow >= 0 && parts < MAX_PARTS) {
      const k = UR.splitPoint(part.plan, lay);
      if (k > 0) {
        await deleteSlide(newId);
        const [A, B] = UR.splitParts(part.xml, part.plan, k);
        queue.unshift(A, B); parts++;
        continue;
      }
      rep.full.push(newId);
    }
    await unmark(newId);
    made.push(newId); prev = newId;
  }
  await deleteSlide(id);
  if (made.length > 1) rep.split.push(made[0]);
  rep.boxes += r.report.boxes; rep.tables += r.report.tables;
  r.report.links.forEach(u => rep.links.push(u)); r.report.linkFail.forEach(u => rep.linkFail.push(u));
  if (r.report.qr) rep.qr += r.report.qr;
  return made;
}
async function doConvert() {
  const ids = await targetIds();
  if (!ids.length) { say('حدّد شريحة أولًا.'); return; }
  const rep = { done: [], split: [], full: [], errs: [], boxes: 0, tables: 0, links: [], linkFail: [], qr: 0 };
  const opts = { watermark: $('wmChk').checked };
  for (let k = 0; k < ids.length; k++) {
    prog(k / ids.length); say('تحويل الشريحة ' + (k + 1) + ' من ' + ids.length + '…');
    try { const made = await convertOne(ids[k], opts, rep); rep.done.push(...made); }
    catch (e) { rep.errs.push(errText(e)); }
  }
  prog(null);
  if ($('numChk').checked && rep.done.length) { say('ترقيم الصفحات…'); try { await renumber(); } catch (e) { rep.errs.push('الترقيم: ' + errText(e)); } }
  const all = await allIds(), num = i => all.indexOf(i) + 1;
  res.innerHTML = '';
  if (rep.done.length) line('ok', '✓ ' + cnt(ids.length - rep.errs.length) + ' ← ' + cnt(rep.done.length) + ': ' + slideLinks(rep.done, num));
  if (rep.split.length) line('warn', '• قُسّمت لأن المحتوى أطول من الصفحة: ' + slideLinks(rep.split, num));
  if (rep.full.length) line('err', '• ما زالت ممتلئة، راجعها: ' + slideLinks(rep.full, num));
  if (rep.qr) line('ok', '• استُبدل QR جديد: ' + rep.qr + (rep.links.length ? ' (رُبط الرابط)' : ''));
  rep.linkFail.forEach(u => line('warn', '• أضف هذا الرابط للـQR يدويًا: <span dir="ltr">' + esc(u) + '</span>'));
  rep.errs.forEach(m => line('err', '✗ ' + esc(m)));
  if (rep.done.length) line('', '<span class="note">يلزم تثبيت خط UULA Sans على الجهاز.</span>');
  say(rep.done.length ? 'تم التحويل.' : 'لم يتم تحويل أي شريحة.');
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
  try { say('ترقيم الصفحات…'); await renumber(); say('تم تحديث أرقام الصفحات.'); }
  catch (e) { say('تعذّر الترقيم: ' + errText(e)); }
  finally { busy = false; lock(); }
});
$('wmChk').addEventListener('change', () => save('ur-wm', $('wmChk').checked));
$('wmChk').checked = load('ur-wm', true);

// ---------- templates ----------
const SW = { header: '1F71DD', topic: '000000', body: 'FFFFFF', bullets: 'FFFFFF', cols: 'FFFFFF', pill: '1F71DC', qhead: 'FF0000', question: 'FF0000', qa: '1F71DE', answer: '1F71DE', qqr: 'FF0000', mcq: 'FF0000', problem: '00B050', other: 'FBEBD3',
  note: 'FBEBD3', quote: 'E2F0D9', formula: 'D3E4F5', activity: 'F1F1F1', mars: 'F1F1F1', cancelled: 'FFC000', table: 'CAD0D8', tableMain: 'DAE0E9', solTable: 'CAD0D8', practice: 'EBE6FE' };
function buildTemplateList() {
  const host = $('tpls'); host.innerHTML = '';
  const groups = [];
  for (const t of URT.LIST) { let g = groups.find(x => x.name === t.g); if (!g) groups.push(g = { name: t.g, items: [] }); g.items.push(t); }
  for (const g of groups) {
    const d = document.createElement('div'); d.className = 'tgroup';
    d.innerHTML = '<div class="gt">' + g.name + '</div>';
    const grid = document.createElement('div'); grid.className = 'tgrid';
    for (const t of g.items) {
      const b = document.createElement('button'); b.type = 'button'; b.className = 'tbtn'; b.dataset.id = t.id;
      b.innerHTML = '<i style="background:#' + (SW[t.id] || 'FFFFFF') + '"></i><span>' + t.name + '</span>';
      grid.appendChild(b);
    }
    d.appendChild(grid); host.appendChild(d);
  }
}
$('tpls').addEventListener('click', e => { const b = e.target.closest('.tbtn'); if (b) insertTemplate(b.dataset.id); });
async function insertTemplate(tid) {
  if (busy || !supported) return;
  busy = true; lock(); tsay('');
  try {
    const id = await currentId();
    if (!id) { tsay('حدّد شريحة أولًا.'); return; }
    const pkg = await UR.openPackage(await exportSlide(id));
    const r = URT.apply(pkg, tid);
    const newId = await insertAfter(await UR.buildPackage(pkg, r.xml, r.rels), id);
    await deleteSlide(id);
    // real text heights, then place the template again
    await run(async ctx => {
      const { map } = await shapesByMark(ctx, newId);
      const idx = Array.from({ length: r.count }, (_, k) => r.first + k);
      const texts = idx.filter((i, k) => r.built.shapes[k].t === 'text' && map[i]);
      texts.forEach(i => { try { map[i].textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; map[i].textFrame.wordWrap = true; } catch (e) {} });
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
      r.fitsReal = pl.bottom <= pkg.slideH - UR.FOOT;
      ctx.presentation.setSelectedSlides([newId]); await ctx.sync();
      try {
        const sel = idx.map(i => map[i]).filter(Boolean).map(s => s.id);
        ctx.presentation.slides.getItem(newId).setSelectedShapes(sel); await ctx.sync();
      } catch (e) {}
    });
    await unmark(newId);
    tsay(r.fitsReal === false ? 'أُضيف القالب، لكنه تجاوز أسفل الصفحة.' : 'أُضيف القالب.');
  } catch (e) { tsay('تعذّرت الإضافة: ' + errText(e)); }
  finally { busy = false; lock(); }
}

// ---------- main button ----------
function lock() {
  $('goBtn').disabled = busy || !supported;
  document.querySelectorAll('.tbtn').forEach(b => { b.disabled = busy || !supported; });
}
$('goBtn').addEventListener('click', async () => {
  if (busy || !supported) return;
  busy = true; lock(); res.innerHTML = '';
  try { await doConvert(); }
  catch (e) { prog(null); say('حدث خطأ: ' + errText(e)); }
  finally { busy = false; lock(); }
});
function setScope(v) { scope = v; save('ur-scope', v); $('scope').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === v))); }
$('scope').addEventListener('click', e => { const b = e.target.closest('button'); if (b) setScope(b.dataset.v); });
setScope(scope);
$('setTog').addEventListener('click', () => { const o = $('setTog').getAttribute('aria-expanded') !== 'true'; $('setTog').setAttribute('aria-expanded', String(o)); $('setBody').hidden = !o; });

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
  modal('حول الأداة', '<p>تحوّل شرائح A4 إلى ستايل مذكرات علا بضغطة زر، وتضيف قوالب المذكرة الجاهزة.</p>'
    + '<ul><li>خط UULA Sans، والنص 11 على شبكة سطرها 13.3</li><li>العناوين والأسئلة والإجابات بألوان علا</li><li>حاويات بعرض الصفحة، وجداول وصور بالشكل الموحّد</li><li>تقسيم الشريحة تلقائيًا إذا طال المحتوى</li></ul>'
    + '<p>لا يُحذف أي محتوى؛ تُحذف الزخارف القديمة فقط.</p>'
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
    if (!base) { base = t; if (manual) note(t ? 'لديك آخر إصدار.' : 'تعذّر الاتصال.'); return; }
    if (t && t !== base) { $('updBtn').hidden = false; if (manual) note('يوجد تحديث.'); }
    else if (manual) note(t ? 'لديك آخر إصدار.' : 'تعذّر الاتصال.');
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
buildTemplateList();
lock();
const ready = (window.Office && Office.onReady) ? Office.onReady() : Promise.resolve({});
ready.then(info => {
  if (!window.PowerPoint || (info && info.host && info.host !== Office.HostType.PowerPoint)) { say('هذه الأداة تعمل داخل PowerPoint فقط.'); return; }
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) { say('حدّث برنامج PowerPoint لاستخدام الأداة.'); return; }
  supported = true;
  const s = docSettings();
  if (s) {
    const st = s.get('urNumStart'); if (st) $('numStart').value = st;
    const on = s.get('urNum'); if (on === false) $('numChk').checked = false;
  }
  $('numRow').style.opacity = $('numChk').checked ? '' : '.5';
  lock();
});
})();

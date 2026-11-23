// UQ — apply a computed slide change to the objects in place (no new slide, nothing moves on screen).
// The engine computes the new slide XML as before; here only the differences are applied with the PowerPoint API,
// then the slide is read back and compared with what was computed. Any difference → the caller takes the full way.
const UQ = (() => {
let why = '';
const no = r => { why = r; return null; };
const NS = UR.NS, PT = UR.PT, U = UR.util;
const A = NS.a, P = NS.p;
const kid = (el, n, ns) => el ? [...el.children].find(c => c.localName === n && (!ns || c.namespaceURI === ns)) || null : null;
const kids = (el, n) => el ? [...el.children].filter(c => c.localName === n) : [];
const desc = (el, ns, n) => el ? [...el.getElementsByTagNameNS(ns, n)] : [];

// ---------- a comparable picture of one object ----------
function colorOf(fillParent) {
  if (!fillParent) return 'inherit';
  if (kid(fillParent, 'noFill', A)) return 'none';
  const sf = kid(fillParent, 'solidFill', A);
  if (sf) { const c = kid(sf, 'srgbClr', A); return c && !c.children.length ? c.getAttribute('val').toUpperCase() : 'other'; }
  if ([...fillParent.children].some(c => /Fill$/.test(c.localName))) return 'other';
  return 'inherit';
}
function lineOf(spPr) {
  const ln = kid(spPr, 'ln', A); if (!ln) return 'inherit';
  const f = colorOf(ln); if (f === 'none') return 'none';
  return f + '/' + (ln.getAttribute('w') || '') + '/' + ((kid(ln, 'prstDash', A) || { getAttribute: () => 'solid' }).getAttribute('val') || 'solid');
}
const pct = (el, n, def) => { const x = kid(el, n, A); if (!x) return def; const p = kid(x, 'spcPct', A), s = kid(x, 'spcPts', A); return p ? 'p' + p.getAttribute('val') : s ? 's' + s.getAttribute('val') : def; };
function paraSig(p, txBox) {
  const pr = kid(p, 'pPr', A);
  const g = (a, d) => (pr && pr.getAttribute(a)) || d;
  const bu = pr ? (kid(pr, 'buNone', A) ? 'none' : kid(pr, 'buChar', A) ? 'char:' + kid(pr, 'buChar', A).getAttribute('char') : kid(pr, 'buAutoNum', A) ? 'num' : txBox ? 'none' : 'inherit') : (txBox ? 'none' : 'inherit');
  return [g('algn', 'l'), g('rtl', '0'), g('marL', '0'), g('indent', '0'), g('lvl', '0'), pct(pr, 'lnSpc', 'p100000'), pct(pr, 'spcBef', 's0'), pct(pr, 'spcAft', 's0'), bu].join(',');
}
// the text of a paragraph, character by character with its look
function runChars(p) {
  const out = [];
  for (const c of p.children) {
    if (c.localName === 'r') {
      const pr = kid(c, 'rPr', A), t = (kid(c, 't', A) || {}).textContent || '';
      const look = pr ? [pr.getAttribute('sz') || '', pr.getAttribute('b') === '1' ? 'b' : '', pr.getAttribute('i') === '1' ? 'i' : '', colorOf(pr), (kid(pr, 'latin', A) || { getAttribute: () => '' }).getAttribute('typeface'), (kid(pr, 'cs', A) || { getAttribute: () => '' }).getAttribute('typeface')].join('|') : '||||';
      for (const ch of t) out.push([ch, look]);
    } else if (c.localName === 'br') out.push(['\v', '']);
    else if (c.localName === 'fld' || c.localName === 'AlternateContent') return null;   // fields and equations: not here
  }
  return out;
}
function textSig(txBody, txBox) {
  if (!txBody) return null;
  const bp = kid(txBody, 'bodyPr', A);
  const g = (a, d) => (bp && bp.getAttribute(a)) || d;
  const fit = bp ? (kid(bp, 'spAutoFit', A) ? 'shape' : kid(bp, 'normAutofit', A) ? 'text' : 'none') : 'none';
  const body = [g('lIns', '91440'), g('tIns', '45720'), g('rIns', '91440'), g('bIns', '45720'), g('wrap', 'square'), g('anchor', 't'), fit].join(',');
  const paras = kids(txBody, 'p').map(p => ({ pr: paraSig(p, txBox), chars: runChars(p) }));
  if (paras.some(p => !p.chars)) return null;
  return { body, paras };
}
function geomOf(el) {
  const c = U.core(el);
  const x = c.localName === 'graphicFrame' ? kid(c, 'xfrm', P) : kid(kid(c, 'spPr', P) || kid(c, 'grpSpPr', P), 'xfrm', A);
  const off = kid(x, 'off', A), ext = kid(x, 'ext', A);
  return off && ext ? { x: +off.getAttribute('x'), y: +off.getAttribute('y'), w: +ext.getAttribute('cx'), h: +ext.getAttribute('cy'), rot: x.getAttribute('rot') || '0', flip: (x.getAttribute('flipH') || '') + (x.getAttribute('flipV') || '') } : null;
}
// everything of an object that matters for its look, without its name and id
function sigOf(el) {
  const c = U.core(el), kind = el.localName === 'AlternateContent' ? 'mc' : c.localName;
  const s = { kind, g: geomOf(el) };
  if (kind === 'sp') {
    const spPr = kid(c, 'spPr', P), nv = kid(c, 'nvSpPr', P);
    s.ph = !!desc(nv, P, 'ph').length;
    s.geo = (kid(spPr, 'prstGeom', A) || { getAttribute: () => 'cust' }).getAttribute('prst');
    s.fill = colorOf(spPr); s.line = lineOf(spPr);
    // no theme style on the shape: an unset fill or line is no fill / no line
    if (!kid(c, 'style', P)) { if (s.fill === 'inherit') s.fill = 'none'; if (s.line === 'inherit') s.line = 'none'; }
    s.fx = !!kid(spPr, 'effectLst', A) && !!kid(spPr, 'effectLst', A).children.length;
    const txBox = (kid(nv, 'cNvSpPr', P) || { getAttribute: () => '' }).getAttribute('txBox') === '1';
    s.text = textSig(kid(c, 'txBody', P), txBox);
    s.badText = !!kid(c, 'txBody', P) && !s.text;
  } else if (kind === 'graphicFrame') {
    const tbl = desc(c, A, 'tbl')[0];
    if (tbl) {
      s.style = (kid(kid(tbl, 'tblPr', A), 'tableStyleId', A) || {}).textContent || '';
      s.cols = kids(kid(tbl, 'tblGrid', A), 'gridCol').map(x => x.getAttribute('w'));
      s.rows = kids(tbl, 'tr').map(tr => ({ h: tr.getAttribute('h'), cells: kids(tr, 'tc').map(tc => {
        const pr = kid(tc, 'tcPr', A);
        const bd = ['lnL', 'lnR', 'lnT', 'lnB'].map(n => { const l = kid(pr, n, A); return l ? colorOf(l) + '/' + (l.getAttribute('w') || '') : 'inherit'; }).join(';');
        return { fill: colorOf(pr), bd, text: textSig(kid(tc, 'txBody', A), true), span: (tc.getAttribute('gridSpan') || '') + (tc.getAttribute('rowSpan') || '') + (tc.getAttribute('hMerge') || '') + (tc.getAttribute('vMerge') || '') };
      }) }));
    } else s.other = true;
  } else s.raw = new XMLSerializer().serializeToString(c).replace(/ (id|name)="[^"]*"/g, '');
  return s;
}
const near = (a, b) => Math.abs(a - b) <= 1.2 * PT;
function sameGeom(a, b, fitH) { return !!a && !!b && near(a.x, b.x) && near(a.y, b.y) && near(a.w, b.w) && (fitH || near(a.h, b.h)) && a.rot === b.rot && a.flip === b.flip; }
function sameText(a, b) {
  if (!a || !b) return a === b;
  if (a.body !== b.body || a.paras.length !== b.paras.length) return false;
  return a.paras.every((p, k) => p.pr === b.paras[k].pr && p.chars.length === b.paras[k].chars.length && p.chars.every((c, j) => c[0] === b.paras[k].chars[j][0] && c[1] === b.paras[k].chars[j][1]));
}
// looks the same (geometry with a little tolerance: points are rounded on the way)
function sameLook(a, b, withGeom) {
  if (a.kind !== b.kind) return false;
  // a text that sizes its box to the text: PowerPoint sets the height itself
  const fitH = a.kind === 'sp' && b.text && /,shape$/.test(b.text.body);
  if (withGeom && !sameGeom(a.g, b.g, fitH)) return false;
  if (a.kind === 'sp') return a.ph === b.ph && a.geo === b.geo && a.fill === b.fill && a.line === b.line && a.fx === b.fx && sameText(a.text, b.text);
  if (a.kind === 'graphicFrame') {
    if (a.other || b.other || a.style !== b.style || a.cols.join() !== b.cols.join() || a.rows.length !== b.rows.length) return false;
    return a.rows.every((r, i) => near(+r.h, +b.rows[i].h) && r.cells.length === b.rows[i].cells.length && r.cells.every((c, j) => { const d = b.rows[i].cells[j]; return c.fill === d.fill && c.bd === d.bd && c.span === d.span && sameText(c.text, d.text); }));
  }
  return a.raw === b.raw;
}

// ---------- plan: what changes, and can it be done in place ----------
function itemsOf(xml) {
  const doc = UR.parseXml(xml);
  return U.topItems(doc).map(el => { const n = U.cnvOf(U.core(el)); return { el, id: n ? n.getAttribute('id') : '', name: n ? n.getAttribute('name') || '' : '', sig: sigOf(el) }; });
}
const plainChars = t => t ? t.paras.map(p => p.chars.map(c => c[0]).join('')).join('\n') : '';
function plan(beforeXml, afterXml) {
  const B = itemsOf(beforeXml), Af = itemsOf(afterXml);
  if (B.length !== Af.length || B.some((b, k) => b.id !== Af[k].id)) return no('plan: if B length Af length B some b k b id Af k id retu');    // something added, removed or reordered
  const changes = [];
  for (let k = 0; k < B.length; k++) {
    const b = B[k], a = Af[k];
    if (b.name === a.name && sameLook(b.sig, a.sig, true)) continue;
    const s0 = b.sig, s1 = a.sig;
    if (s0.kind !== s1.kind) return no('plan: if s0 kind s1 kind return null ');
    if (s0.kind === 'sp') {
      if (s0.ph || s0.badText || s1.badText || s0.geo !== s1.geo || s0.fx !== s1.fx) return no('plan: if s0 ph s0 badText s1 badText s0 geo s1 geo s0 fx');
      if (/other/.test(s1.fill + s1.line)) return no('plan: if other test s1 fill s1 line return null ');
      if (plainChars(s0.text) !== plainChars(s1.text)) return no('plan: if plainChars s0 text plainChars s1 text return nu');                    // the text itself changed
      if (s1.text && s1.text.paras.some((p, i) => { const q = s0.text.paras[i]; return p.pr.split(',').slice(1, 8).join() !== q.pr.split(',').slice(1, 8).join() || (p.pr.split(',')[8] !== q.pr.split(',')[8] && !/^(none|char:.*)$/.test(p.pr.split(',')[8])); })) return no('plan: if s1 text s1 text paras some p i const q s0 text ');
    } else if (s0.kind === 'graphicFrame') {
      if (s0.other || s0.style || s1.style || s0.cols.join() !== s1.cols.join() || s0.rows.length !== s1.rows.length) return no('plan: if s0 other s0 style s1 style s0 cols join s1 cols');
      if (s0.rows.some((r, i) => r.cells.length !== s1.rows[i].cells.length || r.cells.some((c, j) => c.span !== s1.rows[i].cells[j].span || plainChars(c.text) !== plainChars(s1.rows[i].cells[j].text) || /other/.test(s1.rows[i].cells[j].fill)))) return no('plan: if s0 rows some r i r cells length s1 rows i cells');
    } else if (s0.kind === 'pic' || s0.kind === 'grpSp') {
      if (s0.raw.replace(/<a:off[^>]*>|<a:ext [^>]*>/g, '') !== s1.raw.replace(/<a:off[^>]*>|<a:ext [^>]*>/g, '')) return no('plan: if s0 raw replace a off a ext g s1 raw replace a o');   // only moved or sized
    } else return no('plan:  else return null ');
    changes.push({ k, before: b, after: a });
  }
  return { changes, after: Af };
}

// ---------- applying ----------
const ALIGN = { l: 'Left', r: 'Right', ctr: 'Center', just: 'Justify', dist: 'Distributed' };
const ANCHOR = { t: 'Top', ctr: 'Middle', b: 'Bottom' };
function applyFont(f, look) {
  const [sz, b, i, col, latin] = look.split('|');
  if (sz) f.size = +sz / 100;
  f.bold = b === 'b'; f.italic = i === 'i';
  if (/^[0-9A-F]{6}$/.test(col)) f.color = '#' + col;
  if (latin) f.name = latin;
}
// text: runs of the same look get the look together
function applyText(tf, tr, sig) {
  const b = sig.body.split(',');
  tf.leftMargin = +b[0] / PT; tf.topMargin = +b[1] / PT; tf.rightMargin = +b[2] / PT; tf.bottomMargin = +b[3] / PT;
  tf.wordWrap = b[4] !== 'none';
  if (ANCHOR[b[5]]) tf.verticalAlignment = ANCHOR[b[5]];
  tf.autoSizeSetting = b[6] === 'shape' ? 'AutoSizeShapeToFitText' : b[6] === 'text' ? 'AutoSizeTextToFitShape' : 'AutoSizeNone';
  let pos = 0;
  sig.paras.forEach((p, pi) => {
    const pr = p.pr.split(','), len = p.chars.length;
    if (len) {
      const pr0 = tr.getSubstring(pos, len).paragraphFormat;
      if (ALIGN[pr[0]]) pr0.horizontalAlignment = ALIGN[pr[0]];
      pr0.bulletFormat.visible = /^char:/.test(pr[8]);
    }
    let j = 0;
    while (j < len) {
      let e = j; while (e < len && p.chars[e][1] === p.chars[j][1]) e++;
      if (p.chars[j][1]) applyFont(tr.getSubstring(pos + j, e - j).font, p.chars[j][1]);
      j = e;
    }
    pos += len + (pi < sig.paras.length - 1 ? 1 : 0);
  });
}
function applyFill(fill, col) { if (col === 'none') fill.clear(); else if (/^[0-9A-F]{6}$/.test(col)) fill.setSolidColor('#' + col); }
function applyLine(lf, line) {
  if (line === 'none') { lf.visible = false; return; }
  const [col, w, dash] = line.split('/');
  if (/^[0-9A-F]{6}$/.test(col)) { lf.visible = true; lf.color = '#' + col; }
  if (w) lf.weight = +w / PT;
  if (dash && dash !== 'solid') lf.dashStyle = { dash: 'Dash', sysDash: 'RoundDot', sysDot: 'RoundDot', lgDash: 'LongDash', dashDot: 'DashDot' }[dash] || 'Solid';
}

// returns true when done in place and checked; false = take the full way (nothing is left half-done that matters:
// the full way starts again from the slide as it was computed)
async function apply(slideId, beforeXml, afterXml, exportSlide) {
  const pl = plan(beforeXml, afterXml);
  if (!pl) return false;
  why = '';
  if (!pl.changes.length) return true;
  const totalText = t => t.paras.reduce((n, p) => n + p.chars.length, 0) + Math.max(0, t.paras.length - 1);
  const ok = await PowerPoint.run(async ctx => {
    const sh = ctx.presentation.slides.getItem(slideId).shapes;
    sh.load('items/name,items/left,items/top,items/width,items/height,items/type'); await ctx.sync();
    const used = new Set(), jobs = [];
    for (const ch of pl.changes) {
      const g = ch.before.sig.g; if (!g) return false;
      const s = sh.items.find(x => !used.has(x) && (x.name || '') === ch.before.name && Math.abs(x.left - g.x / PT) < 1.5 && Math.abs(x.top - g.y / PT) < 1.5);
      if (!s) { why = 'no match ' + ch.before.name; return false; }
      used.add(s); jobs.push({ s, ch });
    }
    // the text has to be the same as PowerPoint sees it (so the character positions are right)
    const trs = jobs.map(j => j.ch.after.sig.kind === 'sp' && j.ch.after.sig.text ? j.s.textFrame.textRange : null);
    trs.forEach(t => t && t.load('text')); await ctx.sync();
    for (let k = 0; k < jobs.length; k++) if (trs[k] && trs[k].text.length !== totalText(jobs[k].ch.after.sig.text)) { why = 'text length ' + trs[k].text.length + ' vs ' + totalText(jobs[k].ch.after.sig.text); return false; }
    for (let k = 0; k < jobs.length; k++) {
      const { s, ch } = jobs[k], a = ch.after.sig;
      s.name = ch.after.name;
      if (a.kind === 'sp') {
        if (a.fill !== ch.before.sig.fill && a.fill !== 'inherit') applyFill(s.fill, a.fill);
        if (a.line !== ch.before.sig.line && a.line !== 'inherit') applyLine(s.lineFormat, a.line);
        if (a.text) applyText(s.textFrame, trs[k], a.text);
      }
      if (a.g) { s.left = a.g.x / PT; s.top = a.g.y / PT; s.width = a.g.w / PT; if (!(a.kind === 'sp' && a.text && /,shape$/.test(a.text.body))) s.height = a.g.h / PT; } else if (a.kind === 'graphicFrame') {
        const tbl = s.getTable();
        a.rows.forEach((r, i) => r.cells.forEach((c, j) => {
          const b0 = ch.before.sig.rows[i].cells[j];
          const cell = tbl.getCellOrNullObject(i, j);
          if (c.fill !== b0.fill) { if (c.fill === 'none') cell.fill.clear(); else if (/^[0-9A-F]{6}$/.test(c.fill)) cell.fill.setSolidColor('#' + c.fill); }
          if (c.bd !== b0.bd) c.bd.split(';').forEach((v, q) => {
            if (v === 'inherit') return;
            const side = cell.borders[['left', 'right', 'top', 'bottom'][q]], [col, w] = v.split('/');
            if (col === 'none') side.transparency = 1; else if (/^[0-9A-F]{6}$/.test(col)) { side.color = '#' + col; side.transparency = 0; if (w) side.weight = +w / PT; }
          });
          const chars = c.text ? c.text.paras.flatMap(p => p.chars) : [];
          if (chars.length && chars.every(x => x[1] === chars[0][1])) applyFont(cell.font, chars[0][1]);
          else if (chars.length) throw new Error('mixed');
          const al = c.text && c.text.paras[0] && c.text.paras[0].pr.split(',')[0];
          if (ALIGN[al]) cell.horizontalAlignment = ALIGN[al];
        }));
        a.rows.forEach((r, i) => { if (!near(+r.h, +ch.before.sig.rows[i].h)) tbl.rows.getItemAt(i).height = +r.h / PT; });
      }
    }
    await ctx.sync();
    return true;
  }).catch(e => { why = 'api: ' + (e && e.message); return false; });
  if (!ok) return false;
  // read it back: it must look exactly as computed
  const pkg = await UR.openPackage(await exportSlide(slideId));
  const now = itemsOf(pkg.files.slide);
  for (const ch of pl.changes) {
    const it = now.find(x => x.id === ch.after.id) || now.find(x => x.name === ch.after.name);
    if (!it || !sameLook(it.sig, ch.after.sig, true)) { why = 'check ' + ch.after.name + ' ' + JSON.stringify([it && it.sig, ch.after.sig]).slice(0, 3000); return false; }
  }
  return true;
}
return { apply, plan, get why() { return why; } };
})();

/* ==========================================================================
   СТУ-Нормоконтроль – автоисправление документа и выгрузка .docx.
   Правит word/document.xml прямо в архиве, сохраняя всю остальную структуру,
   поэтому Word открывает файл штатно. Правки включаются опциями.
   ========================================================================== */
(function (global) {
  'use strict';

  const WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
  const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];

  function child(parent, name) { if (!parent) return null; for (const el of parent.children) if (el.localName === name) return el; return null; }
  function insertOrdered(parent, el, order) {
    const idx = order.indexOf(el.localName);
    for (const c of parent.children) {
      const ci = order.indexOf(c.localName);
      if (ci >= 0 && ci > idx) { parent.insertBefore(el, c); return; }
    }
    parent.appendChild(el);
  }
  function ensure(parent, name, order, first) {
    let el = child(parent, name);
    if (!el) {
      el = parent.ownerDocument.createElementNS(WNS, 'w:' + name);
      if (first) parent.insertBefore(el, parent.firstChild); else if (order) insertOrdered(parent, el, order); else parent.appendChild(el);
    }
    return el;
  }
  function setW(el, attr, val) { el.setAttribute('w:' + attr, val); }

  function fixRun(run, o) {
    const rPr = ensure(run, 'rPr', null, true);
    if (o.font) {
      const rf = ensure(rPr, 'rFonts', RPR_ORDER);
      setW(rf, 'ascii', 'Times New Roman'); setW(rf, 'hAnsi', 'Times New Roman'); setW(rf, 'cs', 'Times New Roman');
      if (!rf.getAttribute('w:eastAsia')) setW(rf, 'eastAsia', 'Times New Roman');
    }
    if (o.size) {
      setW(ensure(rPr, 'sz', RPR_ORDER), 'val', '28');
      setW(ensure(rPr, 'szCs', RPR_ORDER), 'val', '28');
    }
  }

  function roleOptions(role, o) {
    const a = Object.assign({}, o);
    if (role === 'heading' || role === 'structural' || role === 'tableCaption' || role === 'figureCaption') { a.align = false; a.indent = false; a.spacing = false; a.size = o.size; }
    else if (role === 'list' || role === 'manualList') { a.align = false; a.indent = false; }
    else if (role === 'toc' || role === 'title' || role === 'image' || role === 'table' || role === 'empty') a.size = o.size && role !== 'toc' && role !== 'title';
    return a;
  }

  function fixParagraph(p, role, o) {
    const a = roleOptions(role, o);
    for (const r of p.getElementsByTagName('w:r')) fixRun(r, a);
    const pPr = ensure(p, 'pPr', null, true);
    if (a.align) setW(ensure(pPr, 'jc', PPR_ORDER), 'val', 'both');
    if (a.indent) {
      const ind = ensure(pPr, 'ind', PPR_ORDER);
      if (!ind.getAttribute('w:firstLine') && !ind.getAttribute('w:hanging')) setW(ind, 'firstLine', '708');
    }
    if (a.spacing) {
      const sp = ensure(pPr, 'spacing', PPR_ORDER);
      const line = sp.getAttribute('w:line');
      const rule = sp.getAttribute('w:lineRule') || 'auto';
      if (!line || (rule === 'auto' && line !== '240' && line !== '360')) { setW(sp, 'line', '240'); setW(sp, 'lineRule', 'auto'); }
    }
  }

  function fixTable(tbl, o) {
    for (const r of tbl.getElementsByTagName('w:r')) fixRun(r, { font: o.font, size: false });
  }

  function fixSectPr(sectPr, o) {
    if (!o.page) return;
    const pgSz = ensure(sectPr, 'pgSz');
    const w = +(pgSz.getAttribute('w:w') || 11906), h = +(pgSz.getAttribute('w:h') || 16838);
    const landscape = (pgSz.getAttribute('w:orient') === 'landscape') || (w > h);
    if (landscape) { setW(pgSz, 'w', '16838'); setW(pgSz, 'h', '11906'); setW(pgSz, 'orient', 'landscape'); }
    else { setW(pgSz, 'w', '11906'); setW(pgSz, 'h', '16838'); pgSz.removeAttribute('w:orient'); }
    const mar = ensure(sectPr, 'pgMar');
    if (landscape) { setW(mar, 'top', '1701'); setW(mar, 'bottom', '567'); setW(mar, 'left', '1134'); setW(mar, 'right', '1134'); }
    else { setW(mar, 'top', '1134'); setW(mar, 'bottom', '1134'); setW(mar, 'left', '1701'); setW(mar, 'right', '567'); }
  }

  function fixDashes(dom) {
    const ts = dom.getElementsByTagName('w:t');
    const BOUND = /[0-9A-Za-zА-Яа-я_/.-]/;
    for (const t of ts) {
      const before = t.textContent;
      if (!/\d\s*[-–—]\s*\d/.test(before)) continue;
      const after = before.replace(/(\d+(?:[.,]\d+)?)\s*([-–—])\s*(\d+(?:[.,]\d+)?)/g, (m, a, s, b, idx) => {
        const prev = before[idx - 1] || '', next = before[idx + m.length] || '';
        if (BOUND.test(prev) || BOUND.test(next)) return m;
        return a + '\u2013' + b;
      });
      if (after !== before) t.textContent = after;
    }
  }

  // meta: [{kind, role}] в порядке блоков (как при разборе)
  function buildFixed(buffer, meta, opts) {
    const chk = (s) => { if (global.STUFIX_CHECK) global.STUFIX_CHECK(s); };
    chk('start');
    const files = global.fflate.unzipSync(new Uint8Array(buffer));
    chk('unzip');
    const dec = new TextDecoder('utf-8');
    const enc = new TextEncoder();
    const xmlText = dec.decode(files['word/document.xml']);
    const dom = new DOMParser().parseFromString(xmlText, 'application/xml');
    chk('parse');
    const body = dom.documentElement.getElementsByTagName('w:body')[0];
    const skipBefore = opts.titleEnd || 0;   // титульный лист не трогаем
    let idx = 0;
    function walk(container) {
      for (const el of Array.from(container.children)) {
        if (el.localName === 'p') {
          const m = meta[idx]; const bi = idx; idx++;
          if (bi >= skipBefore) fixParagraph(el, m ? m.role : 'body', opts);
        } else if (el.localName === 'tbl') {
          const bi = idx; idx++;
          if (bi >= skipBefore) fixTable(el, opts);
        } else if (el.localName === 'sdt') { const c = child(el, 'sdtContent'); if (c) walk(c); }
      }
    }
    walk(body);
    chk('walk');
    for (const s of body.getElementsByTagName('w:sectPr')) fixSectPr(s, opts);
    chk('sect');
    if (opts.dash) fixDashes(dom);
    chk('dash');
    const out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + new XMLSerializer().serializeToString(dom.documentElement);
    chk('serialize');
    files['word/document.xml'] = enc.encode(out);
    const zipped = global.fflate.zipSync(files, { level: 6 });
    chk('zip');
    return zipped;
  }

  global.STUFix = { buildFixed };
})(window);

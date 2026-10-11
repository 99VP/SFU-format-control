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
    a.clean = (role === 'body' || role === 'where' || role === 'list' || role === 'manualList');
    if (role === 'heading' || role === 'structural' || role === 'tableCaption' || role === 'figureCaption') { a.align = false; a.indent = false; a.spacing = false; a.size = o.size; }
    else if (role === 'list' || role === 'manualList') { a.align = false; a.indent = false; }
    else if (role === 'toc' || role === 'title' || role === 'image' || role === 'table' || role === 'empty') a.size = o.size && role !== 'toc' && role !== 'title';
    return a;
  }

  // Убирает подчёркивание и приводит цвет текста к чёрному (СТУ, п. 7.1.1).
  function clearEmphasis(p) {
    for (const run of p.getElementsByTagName('w:r')) {
      const rPr = child(run, 'rPr'); if (!rPr) continue;
      const u = child(rPr, 'u'); if (u) rPr.removeChild(u);
      const col = child(rPr, 'color'); if (col) setW(col, 'val', '000000');
    }
  }

  function fixParagraph(p, role, o) {
    const a = roleOptions(role, o);
    for (const r of p.getElementsByTagName('w:r')) fixRun(r, a);
    if (o.emph && a.clean) clearEmphasis(p);
    const pPr = ensure(p, 'pPr', null, true);
    if (a.align) setW(ensure(pPr, 'jc', PPR_ORDER), 'val', 'both');
    if (a.indent) {
      const ind = ensure(pPr, 'ind', PPR_ORDER);
      const fl = ind.getAttribute('w:firstLine');
      const ha = ind.getAttribute('w:hanging');
      const num = (v) => (v == null ? null : (parseInt(v, 10) || 0));
      // нулевой или отсутствующий отступ первой строки приводим к 12,5 мм (708 twips)
      if (!(ha != null && num(ha) > 0) && (fl == null || num(fl) <= 0)) {
        setW(ind, 'firstLine', '708');
        if (ha != null) ind.removeAttribute('w:hanging');
      }
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
    const URL_RE = /(?:https?:\/\/|ftp:\/\/|www\.)[^\s<>"']+/gi;
    for (const t of ts) {
      const before = t.textContent;
      if (!/\d\s*[-–—]\s*\d/.test(before) && !/ \u2014 /.test(before)) continue;
      const urls = []; let um;
      while ((um = URL_RE.exec(before))) urls.push([um.index, um.index + um[0].length]);
      URL_RE.lastIndex = 0;
      // длинное тире между словами → короткое тире (внутри ссылок не трогаем)
      let after = before.replace(/ \u2014 /g, (m, idx) => (urls.some((u) => idx >= u[0] && idx < u[1]) ? m : ' \u2013 '));
      // диапазоны чисел → тире «–» без пробелов
      after = after.replace(/(\d+(?:[.,]\d+)?)\s*([-–—])\s*(\d+(?:[.,]\d+)?)/g, (m, a, s, b, idx) => {
        if (urls.some((u) => idx >= u[0] && idx < u[1])) return m; // внутри ссылки не трогаем
        const prev = after[idx - 1] || '', next = after[idx + m.length] || '';
        if (BOUND.test(prev) || BOUND.test(next)) return m;
        return a + '\u2013' + b;
      });
      if (after !== before) t.textContent = after;
    }
  }

  // Чистка текста: двойные пробелы и пробел перед знаком препинания.
  function fixSpaces(dom) {
    const ts = dom.getElementsByTagName('w:t');
    for (const t of ts) {
      const before = t.textContent;
      const after = before.replace(/ {2,}/g, ' ').replace(/ +([,.;!?])(\s|$)/g, '$1$2');
      if (after !== before) t.textContent = after;
    }
  }

  // Отсылка к источнику: переносим точку до квадратных скобок («текст.[7]»).
  function fixCitations(dom) {
    const ts = dom.getElementsByTagName('w:t');
    const re = /[ \t]*\[(\d{1,4}(?:\s*[;,][^\]]*)?)\][ \t]*\./g;
    for (const t of ts) {
      const before = t.textContent;
      if (!/\[\d/.test(before)) continue;
      const after = before.replace(re, '.[$1]');
      if (after !== before) t.textContent = after;
    }
  }

  // Обновление номера страницы в строке содержания (последний числовой прогон после вкладки).
  function fixTocPage(p, newPage) {
    const runs = p.getElementsByTagName('w:r');
    let tabRun = -1;
    for (let i = 0; i < runs.length; i++) if (child(runs[i], 'tab')) tabRun = i;
    const numT = [];
    for (let i = tabRun + 1; i < runs.length; i++) {
      for (const t of runs[i].getElementsByTagName('w:t')) if (/^\d{1,4}$/.test(t.textContent)) numT.push(t);
    }
    if (!numT.length) {
      const all = p.getElementsByTagName('w:t');
      for (let i = all.length - 1; i >= 0; i--) if (/^\d{1,4}$/.test(all[i].textContent)) { numT.push(all[i]); break; }
    }
    if (!numT.length) return false;
    numT[0].textContent = String(newPage);
    for (let i = 1; i < numT.length; i++) numT[i].textContent = '';
    return true;
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
          if (bi >= skipBefore && opts.toc && opts.tocPages && opts.tocPages[bi] != null) fixTocPage(el, opts.tocPages[bi]);
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
    if (opts.text) fixSpaces(dom);
    chk('text');
    if (opts.text) fixCitations(dom);
    chk('cite');
    const out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + new XMLSerializer().serializeToString(dom.documentElement);
    chk('serialize');
    files['word/document.xml'] = enc.encode(out);
    const zipped = global.fflate.zipSync(files, { level: 6 });
    chk('zip');
    return zipped;
  }

  global.STUFix = { buildFixed };
})(window);

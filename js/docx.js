/* ==========================================================================
   СТУ-Нормоконтроль – разбор файла .docx (OOXML) без серверной части.
   Читает zip через fflate, разбирает XML штатным DOMParser браузера.
   На выходе – модель документа, достаточная для проверки по СТУ 7.5–07–2021.

   Единицы OOXML:
     • twips (двадцатые доли пункта), 1 pt = 20 twips, 1 мм = 56.6929 twips,
       1 дюйм = 1440 twips;
     • w:sz – половинные пункты (sz=28 → 14 pt);
     • w:line при lineRule="auto" – 240-е доли строки (240 → одинарный).
   ========================================================================== */
(function (global) {
  'use strict';

  const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  };

  const TWIP_PER_MM = 1440 / 25.4; // 56.6929…
  const mm = (tw) => (tw == null ? null : tw / TWIP_PER_MM);
  const pt = (tw) => (tw == null ? null : tw / 20);
  const halfPt = (v) => (v == null ? null : v / 2);

  // ------------------------------------------------------------------ XML utils
  function parseXML(bytes) {
    const text = new TextDecoder('utf-8').decode(bytes);
    return new DOMParser().parseFromString(text, 'application/xml');
  }
  function kids(node, name) {
    const out = [];
    if (!node) return out;
    for (let c = node.firstElementChild; c; c = c.nextElementSibling) {
      if (!name || c.localName === name) out.push(c);
    }
    return out;
  }
  function kid(node, name) { return kids(node, name)[0] || null; }
  function attr(el, local) {
    if (!el) return null;
    let v = el.getAttributeNS(NS.w, local);
    if (v == null) v = el.getAttribute('w:' + local);
    if (v == null) v = el.getAttribute(local);
    return v;
  }
  // w:b, w:i, w:caps – «переключатели»: отсутствие val = true, "0"/"false" = false.
  function toggle(el) {
    if (!el) return undefined;
    const v = attr(el, 'val');
    if (v == null) return true;
    return !(v === '0' || v === 'false' || v === 'off');
  }
  function onOff(el) { return !!(el && (attr(el, 'val') == null || !(attr(el, 'val') === '0' || attr(el, 'val') === 'false'))); }
  function relId(el) { return el ? (el.getAttributeNS(NS.r, 'id') || el.getAttribute('r:id') || attr(el, 'id')) : null; }

  // ------------------------------------------------------------------ rPr / pPr
  function parseRPr(rPr) {
    const o = {};
    if (!rPr) return o;
    const rf = kid(rPr, 'rFonts');
    if (rf) {
      o.font = attr(rf, 'hAnsi') || attr(rf, 'ascii') || attr(rf, 'cs') || null;
      o.fontAscii = attr(rf, 'ascii');
      o.fontHAnsi = attr(rf, 'hAnsi');
      o.fontEastAsia = attr(rf, 'eastAsia');
    }
    const sz = kid(rPr, 'sz'); if (sz) o.size = halfPt(+attr(sz, 'val'));
    const szCs = kid(rPr, 'szCs'); if (szCs) o.sizeCs = halfPt(+attr(szCs, 'val'));
    if (kid(rPr, 'b')) o.bold = toggle(kid(rPr, 'b'));
    if (kid(rPr, 'i')) o.italic = toggle(kid(rPr, 'i'));
    const u = kid(rPr, 'u'); if (u) o.underline = (attr(u, 'val') || 'single') !== 'none';
    if (kid(rPr, 'strike')) o.strike = toggle(kid(rPr, 'strike'));
    const c = kid(rPr, 'color'); if (c) { const val = attr(c, 'val'); if (val && val !== 'auto') o.color = val; const th = attr(c, 'themeColor'); if (th) o.themeColor = th; }
    const hl = kid(rPr, 'highlight'); if (hl) { const val = attr(hl, 'val'); if (val && val !== 'none') o.highlight = val; }
    if (kid(rPr, 'caps')) o.caps = toggle(kid(rPr, 'caps'));
    if (kid(rPr, 'smallCaps')) o.smallCaps = toggle(kid(rPr, 'smallCaps'));
    const rs = kid(rPr, 'rStyle'); if (rs) o.rStyle = attr(rs, 'val');
    return o;
  }
  function mergeRPr(base, over) {
    const o = Object.assign({}, base || {});
    if (over) for (const k in over) if (over[k] !== undefined) o[k] = over[k];
    return o;
  }
  function parsePPr(pPr) {
    const o = {};
    if (!pPr) return o;
    const st = kid(pPr, 'pStyle'); if (st) o.styleId = attr(st, 'val');
    const jc = kid(pPr, 'jc'); if (jc) o.jc = attr(jc, 'val');
    const ind = kid(pPr, 'ind');
    if (ind) {
      const io = {};
      const L = attr(ind, 'left') != null ? attr(ind, 'left') : attr(ind, 'start');
      if (L != null) io.left = +L;
      const R = attr(ind, 'right') != null ? attr(ind, 'right') : attr(ind, 'end');
      if (R != null) io.right = +R;
      const FL = attr(ind, 'firstLine'); if (FL != null) io.firstLine = +FL;
      const HG = attr(ind, 'hanging'); if (HG != null) io.hanging = +HG;
      o.ind = io;
    }
    const sp = kid(pPr, 'spacing');
    if (sp) {
      const so = {};
      const b = attr(sp, 'before'); if (b != null) so.before = +b;
      const a = attr(sp, 'after'); if (a != null) so.after = +a;
      const l = attr(sp, 'line'); if (l != null) so.line = +l;
      const lr = attr(sp, 'lineRule'); if (lr != null) so.lineRule = lr;
      o.spacing = so;
    }
    const tabs = kid(pPr, 'tabs');
    if (tabs) {
      o.tabs = kids(tabs, 'tab').map((t) => ({
        pos: attr(t, 'pos') != null ? +attr(t, 'pos') : null,
        val: attr(t, 'val') || 'left',
        leader: attr(t, 'leader') || 'none',
      })).filter((t) => t.pos != null).sort((a, b) => a.pos - b.pos);
    }
    const numPr = kid(pPr, 'numPr');
    if (numPr) {
      o.numPr = {
        numId: attr(kid(numPr, 'numId'), 'val'),
        ilvl: attr(kid(numPr, 'ilvl'), 'val') || '0',
      };
    }
    if (kid(pPr, 'pageBreakBefore')) o.pageBreakBefore = onOff(kid(pPr, 'pageBreakBefore'));
    const ol = kid(pPr, 'outlineLvl'); if (ol) o.outlineLvl = +attr(ol, 'val');
    const rPr = kid(pPr, 'rPr'); if (rPr) o.markRPr = parseRPr(rPr);
    const shd = kid(pPr, 'shd');
    if (shd && attr(shd, 'fill') && attr(shd, 'fill') !== 'auto') o.shading = attr(shd, 'fill');
    const sectPr = kid(pPr, 'sectPr'); if (sectPr) o.sectPr = sectPr;
    if (kid(pPr, 'keepNext')) o.keepNext = onOff(kid(pPr, 'keepNext'));
    if (kid(pPr, 'suppressLineNumbers')) o.suppressLineNumbers = onOff(kid(pPr, 'suppressLineNumbers'));
    return o;
  }

  // ------------------------------------------------------------------ styles
  function buildStyles(stylesXml, themeFonts) {
    const styles = {};
    const docDefaults = { rPr: {}, pPr: {} };
    if (!stylesXml) return { styles, docDefaults, defaultParaStyle: null };
    const doc = stylesXml.documentElement;
    const dd = kid(doc, 'docDefaults');
    if (dd) {
      const rPrDef = kid(dd, 'rPrDefault');
      if (rPrDef) docDefaults.rPr = parseRPr(kid(rPrDef, 'rPr'));
      const pPrDef = kid(dd, 'pPrDefault');
      if (pPrDef) docDefaults.pPr = parsePPr(kid(pPrDef, 'pPr'));
    }
    let defaultParaStyle = null;
    for (const st of kids(doc, 'style')) {
      const id = attr(st, 'styleId');
      const type = attr(st, 'type');
      const name = attr(kid(st, 'name'), 'val');
      const basedOn = attr(kid(st, 'basedOn'), 'val');
      const isDefault = attr(st, 'default') === '1' || attr(st, 'default') === 'true';
      const rPr = parseRPr(kid(st, 'rPr'));
      const pPr = parsePPr(kid(st, 'pPr'));
      styles[id] = { id, type, name, basedOn, isDefault, rPr, pPr };
      if (type === 'paragraph' && isDefault && !defaultParaStyle) defaultParaStyle = id;
    }
    // Дополняем шрифт из темы, если он не задан явно в стиле.
    if (themeFonts && themeFonts.minor) {
      const minor = { font: themeFonts.minor };
      for (const id in styles) {
        if (styles[id].rPr.font == null) styles[id].rPr.font = themeFonts.minor;
      }
      if (docDefaults.rPr.font == null) docDefaults.rPr.font = themeFonts.minor;
    }
    return { styles, docDefaults, defaultParaStyle };
  }
  function resolveStyleChain(styleId, styles) {
    const chain = [];
    let id = styleId, guard = 0;
    while (id && styles[id] && guard++ < 20) { chain.push(styles[id]); id = styles[id].basedOn; }
    return chain;
  }
  function styleRPr(styleId, styles, cache) {
    if (cache && cache[styleId]) return cache[styleId];
    let r = {};
    for (const s of resolveStyleChain(styleId, styles).reverse()) r = mergeRPr(r, s.rPr);
    if (cache) cache[styleId] = r;
    return r;
  }
  function stylePPr(styleId, styles, cache) {
    if (cache && cache[styleId]) return cache[styleId];
    let p = {};
    for (const s of resolveStyleChain(styleId, styles).reverse()) p = mergeRPr(p, s.pPr);
    if (cache) cache[styleId] = p;
    return p;
  }

  // ------------------------------------------------------------------ numbering
  function buildNumbering(numXml) {
    const nums = {}, abstractNums = {};
    if (!numXml) return { nums, abstractNums, level: () => null };
    const doc = numXml.documentElement;
    for (const an of kids(doc, 'abstractNum')) {
      const id = attr(an, 'abstractNumId');
      const levels = {};
      for (const lvl of kids(an, 'lvl')) {
        const ilvl = attr(lvl, 'ilvl');
        const fmt = attr(kid(lvl, 'numFmt'), 'val');
        const lvlText = attr(kid(lvl, 'lvlText'), 'val');
        const start = attr(kid(lvl, 'start'), 'val');
        const pStyle = attr(kid(lvl, 'pStyle'), 'val');
        const ind = parsePPr(kid(lvl, 'pPr')).ind || null;
        levels[ilvl] = { ilvl: +ilvl, fmt, lvlText, start: start ? +start : 1, pStyle, ind };
      }
      abstractNums[id] = { id, levels };
    }
    for (const n of kids(doc, 'num')) {
      const numId = attr(n, 'numId');
      const aId = attr(kid(n, 'abstractNumId'), 'val');
      nums[numId] = { numId, abstractNumId: aId };
    }
    function level(numId, ilvl) {
      const n = nums[String(numId)];
      if (!n) return null;
      const an = abstractNums[n.abstractNumId];
      if (!an) return null;
      return an.levels[String(ilvl || 0)] || null;
    }
    return { nums, abstractNums, level };
  }

  // ------------------------------------------------------------------ sectPr
  function parseSectPr(sectPr) {
    const o = {};
    if (!sectPr) { o.w = 11906; o.h = 16838; o.orient = 'portrait'; o.margin = { top: 1134, right: 850, bottom: 1134, left: 1701, header: 708, footer: 708, gutter: 0 }; return o; }
    const pgSz = kid(sectPr, 'pgSz');
    if (pgSz) { o.w = +attr(pgSz, 'w'); o.h = +attr(pgSz, 'h'); o.orient = attr(pgSz, 'orient') || (o.w > o.h ? 'landscape' : 'portrait'); }
    const mar = kid(sectPr, 'pgMar');
    o.margin = {
      top: mar ? +attr(mar, 'top') : null,
      right: mar ? +attr(mar, 'right') : null,
      bottom: mar ? +attr(mar, 'bottom') : null,
      left: mar ? +attr(mar, 'left') : null,
      header: mar ? +attr(mar, 'header') : null,
      footer: mar ? +attr(mar, 'footer') : null,
      gutter: mar ? +attr(mar, 'gutter') : null,
    };
    o.titlePg = !!kid(sectPr, 'titlePg');
    const cols = kid(sectPr, 'cols'); if (cols) { o.cols = +attr(cols, 'num') || 1; o.colsSpace = attr(cols, 'space'); }
    const type = kid(sectPr, 'type'); if (type) o.type = attr(type, 'val');
    // ссылки на колонтитулы
    o.headerRefs = []; o.footerRefs = [];
    for (const h of kids(sectPr, 'headerReference')) o.headerRefs.push({ type: attr(h, 'type'), id: relId(h) });
    for (const f of kids(sectPr, 'footerReference')) o.footerRefs.push({ type: attr(f, 'type'), id: relId(f) });
    return o;
  }

  // ------------------------------------------------------------------ runs
  function parseRun(run) {
    const rPr = parseRPr(kid(run, 'rPr'));
    const parts = [];
    for (const c of kids(run)) {
      if (c.localName === 't') parts.push({ type: 'text', text: c.textContent });
      else if (c.localName === 'tab') parts.push({ type: 'tab' });
      else if (c.localName === 'br') parts.push({ type: attr(c, 'type') === 'page' ? 'pagebreak' : 'break' });
      else if (c.localName === 'cr') parts.push({ type: 'break' });
      else if (c.localName === 'fldChar') parts.push({ type: 'fldChar', ftype: attr(c, 'fldCharType') });
      else if (c.localName === 'instrText') parts.push({ type: 'instrText', text: c.textContent });
      else if (c.localName === 'noBreakHyphen') parts.push({ type: 'text', text: '\u2011' });
    }
    // Графика может лежать внутри mc:AlternateContent (Choice – DrawingML,
    // Fallback – VML), поэтому ищем её в любом месте прогона.
    const g = run.getElementsByTagName('w:drawing')[0] || run.getElementsByTagName('w:pict')[0] || run.getElementsByTagName('w:object')[0];
    if (g) {
      const blip = g.getElementsByTagName('a:blip')[0];
      let embed = blip ? (blip.getAttributeNS(NS.r, 'embed') || blip.getAttribute('r:embed')) : null;
      if (!embed) {
        const vi = g.getElementsByTagName('v:imagedata')[0];
        if (vi) embed = vi.getAttributeNS(NS.r, 'id') || vi.getAttribute('r:id');
      }
      let cx = null, cy = null;
      const ext = g.getElementsByTagName('wp:extent')[0];
      if (ext) { cx = +ext.getAttribute('cx'); cy = +ext.getAttribute('cy'); }
      // Плавающий рисунок (anchor) не должен раздвигать строку, как в Word.
      let anchor = false, offH = null, offV = null, relH = null, relV = null;
      const anc = g.getElementsByTagName('wp:anchor')[0];
      if (anc) {
        anchor = true;
        const ph = anc.getElementsByTagName('wp:positionH')[0];
        const pv = anc.getElementsByTagName('wp:positionV')[0];
        if (ph) { relH = ph.getAttribute('relativeFrom'); const oh = ph.getElementsByTagName('wp:posOffset')[0]; if (oh) offH = (+oh.textContent) / 914400 * 96; }
        if (pv) { relV = pv.getAttribute('relativeFrom'); const ov = pv.getElementsByTagName('wp:posOffset')[0]; if (ov) offV = (+ov.textContent) / 914400 * 96; }
      }
      // Текст внутри фигуры-надписи (если это не картинка).
      let txbx = null;
      const tb = g.getElementsByTagName('w:txbxContent')[0];
      if (tb) {
        let s = '';
        const walkT = (n) => {
          for (const c of n.childNodes) {
            if (c.nodeType !== 1) continue;
            if (c.localName === 't') s += c.textContent;
            else if (c.localName === 'br') s += '\n';
            else if (c.localName === 'p') { if (s) s += '\n'; walkT(c); }
            else walkT(c);
          }
        };
        walkT(tb);
        txbx = s.replace(/\n{2,}/g, '\n').trim();
      }
      // размер шрифта внутри надписи (полуточки → пункты)
      let txbxSize = null;
      if (tb) for (const s of tb.getElementsByTagName('w:sz')) { const v = +s.getAttribute('w:val'); if (v && (txbxSize == null || v > txbxSize)) txbxSize = v; }
      if (txbxSize) txbxSize = txbxSize / 2;
      parts.push({ type: 'drawing', embed, cx, cy, anchor, offH, offV, relH, relV, txbx, txbxSize });
    }
    let text = '';
    for (const p of parts) {
      if (p.type === 'text') text += p.text;
      else if (p.type === 'tab') text += '\t';
      else if (p.type === 'break' || p.type === 'pagebreak') text += '\n';
      else if (p.type === 'drawing') text += '\uFFFC';
    }
    let fieldInstr = '';
    for (const p of parts) if (p.type === 'instrText') fieldInstr += p.text;
    return { rPr, parts, text, fieldInstr };
  }

  function parseParagraph(p, ctx) {
    const pPr = parsePPr(kid(p, 'pPr'));
    // Прогоны нужно брать строго в порядке документа, в том числе внутри
    // гиперссылок и полей (w:hyperlink, w:sdt, w:fldSimple). Иначе, например,
    // номер страницы в содержании «перескакивает» вперёд текста.
    const allRaw = [];
    (function collect(node) {
      for (const c of Array.from(node.childNodes)) {
        if (c.nodeType !== 1) continue;
        if (c.localName === 'r') allRaw.push(parseRun(c));
        else if (c.localName === 'pPr') continue;
        else collect(c);
      }
    })(p);
    const styleId = pPr.styleId || ctx.defaultParaStyle || null;
    const stRPr = styleRPr(styleId, ctx.styles, ctx.styleRPrCache);
    const stPPr = stylePPr(styleId, ctx.styles, ctx.stylePPrCache);
    const stName = (ctx.styles[styleId] && ctx.styles[styleId].name) || '';
    const base = mergeRPr(ctx.docDefaults.rPr, stRPr);
    const runs = [];
    for (const r of allRaw) {
      let eff = mergeRPr(base, r.rPr);
      if (r.rPr.rStyle) {
        const cs = styleRPr(r.rPr.rStyle, ctx.styles, ctx.styleRPrCache);
        eff = mergeRPr(mergeRPr(base, cs), r.rPr);
      }
      runs.push({
        text: r.text, parts: r.parts, rPr: r.rPr, eff, fieldInstr: r.fieldInstr || '',
        font: eff.font || ctx.baseFont,
        size: eff.size != null ? eff.size : 14,
        bold: !!eff.bold, italic: !!eff.italic, underline: !!eff.underline, caps: !!eff.caps,
        highlight: eff.highlight || null, colorVal: eff.color || null,
        directColor: r.rPr && r.rPr.color ? r.rPr.color : null,
        themeColor: (r.rPr && r.rPr.themeColor) || null,
        hasDrawing: r.parts.some((x) => x.type === 'drawing'),
        hasPageBreak: r.parts.some((x) => x.type === 'pagebreak'),
      });
    }
    const markEff = mergeRPr(base, pPr.markRPr || {});
    // Разрыв страницы в начале абзаца трактуем как «начать с новой страницы».
    let leadingBreak = false;
    const firstRunWithParts = runs.find((r) => r.parts && r.parts.length);
    if (firstRunWithParts && firstRunWithParts.parts[0] && firstRunWithParts.parts[0].type === 'pagebreak') {
      const other = runs.map((r) => (r.text || '')).join('').replace(/\n/g, '').replace(/\uFFFC/g, '').trim();
      if (other !== '') {
        leadingBreak = true;
        firstRunWithParts.parts = firstRunWithParts.parts.slice(1);
        if (firstRunWithParts.text.charAt(0) === '\n') firstRunWithParts.text = firstRunWithParts.text.slice(1);
      }
    }
    let text = '';
    for (const r of runs) text += r.text;
    const fieldText = runs.map((r) => r.fieldInstr).join(' ');
    const markRPr = {
      font: markEff.font || ctx.baseFont,
      size: markEff.size != null ? markEff.size : 14,
      bold: !!markEff.bold,
    };
    const sp = mergeRPr(stPPr.spacing || {}, pPr.spacing || {});
    const ind = mergeRPr(stPPr.ind || {}, pPr.ind || {});
    const effPPr = {
      jc: pPr.jc || stPPr.jc || null,
      ind,
      spacing: sp,
      numPr: pPr.numPr || null,
      outlineLvl: pPr.outlineLvl != null ? pPr.outlineLvl : (stPPr.outlineLvl != null ? stPPr.outlineLvl : null),
      pageBreakBefore: !!(pPr.pageBreakBefore || stPPr.pageBreakBefore || leadingBreak),
      sectPr: pPr.sectPr || null,
      keepNext: !!(pPr.keepNext || stPPr.keepNext),
      tabs: pPr.tabs || stPPr.tabs || [],
    };
    // нормализуем undефайны
    for (const k in effPPr.ind) if (effPPr.ind[k] === undefined) effPPr.ind[k] = null;
    for (const k in effPPr.spacing) if (effPPr.spacing[k] === undefined) effPPr.spacing[k] = null;
    return {
      kind: 'p',
      text,
      runs,
      styleId,
      styleName: stName,
      pPr: effPPr,
      jc: effPPr.jc,
      ind: effPPr.ind,
      spacing: effPPr.spacing,
      numPr: effPPr.numPr,
      outlineLvl: effPPr.outlineLvl,
      pageBreakBefore: effPPr.pageBreakBefore,
      sectPr: effPPr.sectPr,
      tabs: effPPr.tabs,
      fieldText,
      markRPr,
      hasDrawing: runs.some((r) => r.hasDrawing),
      isEmpty: text.trim() === '',
      raw: p,
    };
  }

  function parseTable(tbl, ctx) {
    const rows = [];
    for (const tr of kids(tbl, 'tr')) {
      const cells = [];
      for (const tc of kids(tr, 'tc')) {
        const tcPr = kid(tc, 'tcPr');
        const gridSpan = tcPr ? +(attr(kid(tcPr, 'gridSpan'), 'val') || 1) : 1;
        const paras = kids(tc, 'p').map((p) => parseParagraph(p, ctx));
        cells.push({ paras, gridSpan, tcPr });
      }
      rows.push({ cells });
    }
    // поиск объединений по вертикали пропускаем – для проверок не критично
    const tblPr = kid(tbl, 'tblPr');
    const borders = kid(tblPr, 'tblBorders');
    const hasBorders = !!(borders && ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].some((side) => {
      const b = kid(borders, side); return b && attr(b, 'val') !== 'none' && attr(b, 'val') !== 'nil';
    }));
    const tblW = kid(tblPr, 'tblW');
    const tblGrid = kid(tbl, 'tblGrid');
    const grid = tblGrid ? kids(tblGrid, 'gridCol').map((c) => +attr(c, 'w') || 0) : [];
    return {
      kind: 'table',
      rows,
      hasBorders,
      grid,
      width: tblW ? { w: attr(tblW, 'w'), type: attr(tblW, 'type') } : null,
      text: rows.map((r) => r.cells.map((c) => c.paras.map((p) => p.text).join(' ')).join('\t')).join('\n'),
    };
  }

  // ------------------------------------------------------------------ theme fonts
  function parseThemeFonts(xml) {
    if (!xml) return {};
    try {
      const els = xml.getElementsByTagName('a:latin');
      const major = els[0] && els[0].getAttribute('typeface');
      const minor = els[1] && els[1].getAttribute('typeface');
      return { major, minor };
    } catch (e) { return {}; }
  }

  // ------------------------------------------------------------------ footers/headers
  function parseHeaderFooter(xml, idMap) {
    if (!xml) return null;
    const root = xml.documentElement;
    const paras = kids(root, 'p').map((p) => {
      const pPr = parsePPr(kid(p, 'pPr'));
      const runs = kids(p, 'r').map(parseRun);
      let text = '';
      let hasPageField = false;
      for (const r of runs) text += r.text;
      // поля PAGE / NUMPAGES
      for (const fc of root.getElementsByTagName('w:fldChar')) {
        // nothing
      }
      const instrNodes = Array.prototype.slice.call(root.getElementsByTagName('w:instrText'));
      for (const it of instrNodes) if (/PAGE/i.test(it.textContent)) hasPageField = true;
      const fldSimple = Array.prototype.slice.call(root.getElementsByTagName('w:fldSimple'));
      for (const fs of fldSimple) {
        const instr = fs.getAttributeNS(NS.w, 'instr') || fs.getAttribute('w:instr') || '';
        if (/PAGE/i.test(instr)) hasPageField = true;
      }
      const rPr = runs[0] ? runs[0].eff : {};
      return { jc: pPr.jc || null, ind: pPr.ind || {}, text, hasPageField, font: (runs[0] && runs[0].font) || null, size: (runs[0] && runs[0].size) || null };
    });
    return { paras, text: paras.map((p) => p.text).join('\n'), hasPageField: paras.some((p) => p.hasPageField) };
  }

  // ------------------------------------------------------------------ main parse
  function parse(arrayBuffer) {
    const u8 = new Uint8Array(arrayBuffer);
    const files = global.fflate.unzipSync(u8, { filter: (f) => /^word\/|^docProps\//.test(f.name) });
    const get = (name) => files[name] ? parseXML(files[name]) : null;
    const getBytes = (name) => files[name] || null;

    const documentXml = get('word/document.xml');
    if (!documentXml) throw new Error('В файле нет word/document.xml – это не документ Word (.docx).');

    const themeXml = get('word/theme/theme1.xml');
    const themeFonts = parseThemeFonts(themeXml);
    const stylesInfo = buildStyles(get('word/styles.xml'), themeFonts);
    const numbering = buildNumbering(get('word/numbering.xml'));
    const settingsXml = get('word/settings.xml');

    const baseFont = stylesInfo.docDefaults.rPr.font || themeFonts.minor || 'Times New Roman';

    const ctx = {
      styles: stylesInfo.styles,
      docDefaults: stylesInfo.docDefaults,
      defaultParaStyle: stylesInfo.defaultParaStyle,
      baseFont,
      styleRPrCache: {},
      stylePPrCache: {},
    };

    // Ссылки на части (rels) для колонтитулов
    const relsMap = {};
    const relsXml = get('word/_rels/document.xml.rels');
    if (relsXml) {
      for (const rel of kids(relsXml.documentElement, 'Relationship')) {
        relsMap[rel.getAttribute('Id')] = rel.getAttribute('Target');
      }
    }

    const body = kid(documentXml.documentElement, 'body');
    const blocks = [];
    let globalSection = null;
    // Последний sectPr в body – настройки последней секции.
    for (const s of kids(body, 'sectPr')) globalSection = s;

    function pushBlock(b) { b.index = blocks.length; blocks.push(b); }

    function walk(container) {
      for (const el of kids(container)) {
        if (el.localName === 'p') pushBlock(parseParagraph(el, ctx));
        else if (el.localName === 'tbl') pushBlock(parseTable(el, ctx));
        else if (el.localName === 'sdt') {
          const content = kid(el, 'sdtContent');
          if (content) walk(content);
        }
      }
    }
    walk(body);

    // Секции: агрегируем по sectPr. Обычно одна секция.
    const sections = [];
    let currentSectionStart = 0;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === 'p' && b.sectPr) {
        const sect = parseSectPr(b.sectPr);
        sect.startIndex = currentSectionStart;
        sect.endIndex = i;
        sections.push(sect);
        currentSectionStart = i + 1;
        delete b.sectPr;
      }
    }
    const lastSect = parseSectPr(globalSection);
    lastSect.startIndex = currentSectionStart;
    lastSect.endIndex = blocks.length - 1;
    sections.push(lastSect);

    // Колонтитулы
    function loadHF(refs) {
      const out = {};
      for (const ref of refs || []) {
        const target = relsMap[ref.id];
        if (!target) continue;
        const path = 'word/' + target.replace(/^\/?word\//, '');
        const x = get(path);
        if (x) out[ref.type || 'default'] = parseHeaderFooter(x, relsMap);
      }
      return out;
    }

    // привязываем колонтитулы к секциям (по ссылкам в каждой секции);
    // сохраняем и «сырые» ссылки.
    for (const sect of sections) {
      sect.headers = loadHF(sect.headerRefs);
      sect.footers = loadHF(sect.footerRefs);
    }

    // Настройки
    const settings = {};
    if (settingsXml) {
      const root = settingsXml.documentElement;
      settings.evenAndOddHeaders = !!kid(root, 'evenAndOddHeaders');
      settings.mirrorMargins = !!kid(root, 'mirrorMargins');
      const df = kid(root, 'defaultTabStop');
      settings.defaultTabStop = df ? +attr(df, 'val') : 708;
    }

    const props = {};
    const coreXml = get('docProps/core.xml');
    if (coreXml) {
      const t = kid(coreXml.documentElement, 'title');
      if (t) props.title = t.textContent;
    }

    // Медиафайлы (картинки)
    const media = {};
    for (const name in files) {
      if (/^word\/media\//.test(name)) {
        const rel = name.replace(/^word\//, '');
        try { media[rel] = 'data:' + mimeOf(name) + ';base64,' + bytesToBase64(files[name]); } catch (e) { }
      }
    }

    return {
      baseFont,
      blocks,
      sections,
      numbering,
      settings,
      props,
      styles: stylesInfo.styles,
      docDefaults: stylesInfo.docDefaults,
      themeFonts,
      rels: relsMap,
      media,
    };
  }

  function mimeOf(name) {
    if (/\.png$/i.test(name)) return 'image/png';
    if (/\.jpe?g$/i.test(name)) return 'image/jpeg';
    if (/\.gif$/i.test(name)) return 'image/gif';
    if (/\.bmp$/i.test(name)) return 'image/bmp';
    if (/\.svg$/i.test(name)) return 'image/svg+xml';
    if (/\.emf$/i.test(name)) return 'image/emf';
    if (/\.wmf$/i.test(name)) return 'image/wmf';
    return 'application/octet-stream';
  }
  function bytesToBase64(u8) {
    let bin = '';
    const chunk = 0x8000;
    for (let i = 0; i < u8.length; i += chunk) {
      bin += String.fromCharCode.apply(null, u8.subarray(i, i + chunk));
    }
    return btoa(bin);
  }

  global.STUDocx = { parse, helpers: { mm, pt, TWIP_PER_MM, kids, kid, attr } };
})(window);

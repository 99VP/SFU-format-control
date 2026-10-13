/* ==========================================================================
   СТУ-Нормоконтроль – построение постраничной развёртки документа.
   Браузер сам переносит слова внутри абзаца; мы раскладываем абзацы по
   листам А4 с реальными полями, шрифтами, отступами и интервалами, рисуем
   табуляцию (в том числе с отточием в содержании), маркеры списков и
   колонтитул с номером страницы, затем подсвечиваем проблемы.
   ========================================================================== */
(function (global) {
  'use strict';

  const PX_PER_MM = 96 / 25.4;
  const MM = (v) => v * PX_PER_MM;         // мм → px
  const TW = (v) => (v / 1440) * 96;       // twips → px
  const PT = (v) => (v * 96) / 72;         // pt → px
  const twMM = (v) => v / (1440 / 25.4);

  const LEVEL_CLASS = { error: 'stu-mk-err', warning: 'stu-mk-warn', info: 'stu-mk-info' };
  const LEVEL_RANK = { error: 3, warning: 2, info: 1 };
  const FONT_STACK = '"Times New Roman", "Liberation Serif", "Tinos", "DejaVu Serif", serif';

  let _mctx = null;
  function measure(text, run) {
    if (!_mctx) { const c = document.createElement('canvas'); _mctx = c.getContext('2d'); }
    const px = PT((run && run.size) || 14);
    _mctx.font = (run && run.italic ? 'italic ' : '') + (run && run.bold ? '700 ' : '') + px + 'px ' + FONT_STACK;
    return _mctx.measureText(text).width;
  }

  function geomOf(section) {
    const s = section || {};
    const w = s.w || 11906, h = s.h || 16838;
    const portrait = w <= h;
    const m = s.margin || {};
    const dflt = portrait
      ? { left: 1701, right: 567, top: 1134, bottom: 1134 }
      : { left: 1134, right: 1134, top: 1701, bottom: 567 };
    return {
      wmm: twMM(w), hmm: twMM(h), portrait,
      ml: twMM(m.left != null ? m.left : dflt.left),
      mr: twMM(m.right != null ? m.right : dflt.right),
      mt: twMM(m.top != null ? m.top : dflt.top),
      mb: twMM(m.bottom != null ? m.bottom : dflt.bottom),
    };
  }

  function lineHeightPx(size, spacing) {
    const sizePx = PT(size || 14);
    if (spacing) {
      if (spacing.line && (spacing.lineRule === 'exact' || spacing.lineRule === 'atLeast')) return PT(spacing.line / 20);
      const f = spacing.line ? spacing.line / 240 : 1;
      return sizePx * 1.15 * f;
    }
    return sizePx * 1.15;
  }

  function styleParagraphText(el, p) {
    const jc = p.jc;
    const centered = (jc === 'center' || jc === 'right' || jc === 'end');
    if (jc === 'both' || jc === 'justify' || jc === 'distribute') el.style.textAlign = 'justify';
    else if (jc === 'center') el.style.textAlign = 'center';
    else if (jc === 'right' || jc === 'end') el.style.textAlign = 'right';
    else el.style.textAlign = 'left';
    const ind = p.ind || {};
    // у центрированных и правосторонних строк отступы не применяем, иначе текст уезжает
    if (!centered) {
      if (ind.left) el.style.marginLeft = TW(ind.left) + 'px';
      if (ind.right) el.style.marginRight = TW(ind.right) + 'px';
      if (ind.firstLine) el.style.textIndent = TW(ind.firstLine) + 'px';
      else if (ind.hanging) el.style.textIndent = (-TW(ind.hanging)) + 'px';
    }
    const sp = p.spacing || {};
    const firstRun = (p.runs || []).find((r) => r.size) || (p.runs || [])[0] || { size: 14 };
    const lh = lineHeightPx(firstRun.size, sp);
    el.style.lineHeight = lh + 'px';
    el.style.fontSize = PT(firstRun.size || 14) + 'px';
    el.style.fontFamily = FONT_STACK;
    el.style.marginTop = (sp.before ? PT(sp.before / 20) : 0) + 'px';
    el.style.marginBottom = (sp.after ? PT(sp.after / 20) : 0) + 'px';
  }

  // ------------------------------------------------------------- маркеры списков
  function bulletChar(lvlText) {
    if (!lvlText) return '\u2013';
    let ch = lvlText.replace(/[\uF0A7\uF0B7\uF0FC\uF0D8\uF0A8]/g, '\u2022');
    if (ch === 'o' || ch === '-') return '\u2013';
    if (ch === '\u2013' || ch === '\u2014') return '\u2013';
    if (/[\u2022\u25CF\u25AA\u25E6\u00B7\u25A0\u25AB\u25C6]/.test(ch)) return '\u2022';
    if (ch === '\uF02D' || ch === '\uF0FC') return '\u{1F4D6}';   // значок «книжка»
    if (/[\uE000-\uF8FF]/.test(ch)) return '\u2022';
    return ch;
  }
  function formatLevel(lvlText, numId, counters) {
    if (!lvlText) return '';
    return lvlText.replace(/%(\d)/g, (m, d) => {
      const key = numId + '/' + (+d - 1);
      return counters[key] != null ? String(counters[key]) : '1';
    });
  }
  function computeMarkers(doc) {
    const markers = {};
    const num = doc.numbering; if (!num) return markers;
    const counters = {};
    const blocks = doc.blocks;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind !== 'p' || !b.numPr || !b.numPr.numId) continue;
      const lvl = num.level(b.numPr.numId, b.numPr.ilvl);
      if (!lvl) continue;
      const key = b.numPr.numId + '/' + (b.numPr.ilvl || '0');
      if (lvl.fmt === 'bullet') { markers[i] = bulletChar(lvl.lvlText); continue; }
      counters[key] = (counters[key] != null ? counters[key] : ((lvl.start || 1) - 1)) + 1;
      // сбрасываем более глубокие уровни
      for (const k in counters) {
        if (k.indexOf(b.numPr.numId + '/') === 0 && k !== key) {
          const lv = +k.split('/')[1];
          if (lv > (+b.numPr.ilvl || 0)) delete counters[k];
        }
      }
      markers[i] = formatLevel(lvl.lvlText, b.numPr.numId, counters) || (counters[key] + ')');
    }
    return markers;
  }

  // ------------------------------------------------------------- табуляция
  function appendTab(container, run, cursor, p, ctx) {
    const tabs = (p.tabs || []).map((t) => ({ pos: TW(t.pos), val: t.val, leader: t.leader })).sort((a, b) => a.pos - b.pos);
    let stop = null, leader = 'none';
    for (const t of tabs) { if (t.pos > cursor.x + 0.5) { stop = t.pos; leader = t.leader || 'none'; break; } }
    if (stop == null) { const step = ctx.tabStep || TW(708); stop = Math.ceil((cursor.x + 1) / step) * step; }
    const ind = p.ind || {};
    const startOffset = TW((ind.left || 0) + Math.max(0, ind.firstLine || 0));
    const maxStop = (ctx.textWidthPx || 1e9) - startOffset;
    if (stop > maxStop) stop = maxStop;
    const w = Math.max(0, stop - cursor.x);
    const span = document.createElement('span');
    span.className = 'stu-tab';
    span.style.display = 'inline-block';
    span.style.width = w + 'px';
    if (leader === 'dot' || leader === 'middleDot' || leader === 'hyphen' || leader === 'underscore') span.classList.add('stu-tab-' + (leader === 'middleDot' ? 'dot' : leader));
    if (run.underline) span.style.borderBottom = '1px solid currentColor';
    container.appendChild(span);
    cursor.x = stop;
  }

  function appendDrawing(container, part, ctx, blockIndex) {
    let src = null;
    if (part.embed && ctx.doc.rels && ctx.doc.media) {
      const rel = ctx.doc.rels[part.embed];
      if (rel) { src = ctx.doc.media['media/' + rel.split('/').pop()] || ctx.doc.media[rel]; }
    }
    // Плавающий рисунок на титульном листе выводим вне потока, иначе он
    // раздвигает строку и последняя строка титула уезжает на другой лист.
    const role = ctx.roles && ctx.roles[blockIndex] ? ctx.roles[blockIndex].role : null;
    const floating = part.anchor && (role === 'title' || !!part.flowless);
    if (src) {
      const img = document.createElement('img');
      img.src = src; img.className = 'stu-img';
      img.style.verticalAlign = 'bottom';
      // Размер задаём сразу из OOXML (cx/cy), иначе картинка получит высоту
      // только после декодирования и сдвинет текст вниз (текст обрежется).
      let wpx = null, hpx = null;
      if (part.cx) wpx = (part.cx / 914400) * 96;
      if (part.cy) hpx = (part.cy / 914400) * 96;
      const limit = Math.min(MM(160), ctx.textWidthPx || MM(160));
      if (wpx && wpx > limit) { if (hpx) hpx = hpx * (limit / wpx); wpx = limit; }
      if (wpx) img.style.width = wpx + 'px';
      else img.style.maxWidth = limit + 'px';
      if (hpx) img.style.height = hpx + 'px';
      // выравнивание по якорю: по центру / вправо; иначе в строке (рисунки могут идти в ряд)
      if (part.hAlign === 'center') { img.style.display = 'block'; img.style.margin = '0 auto'; }
      else if (part._block) img.style.display = 'block';
      // якорь: Word размещает плавающий рисунок ниже — добавляем отступ сверху
      if (part.anchor && !floating && part.offV != null && part.offV > 6) img.style.marginTop = Math.round(part.offV) + 'px';
      if (floating) {
        img.style.position = 'absolute';
        img.style.left = Math.max(0, part.offH != null ? part.offH : 0) + 'px';
        img.style.top = Math.max(-8, part.offV != null ? part.offV : 0) + 'px';
        if (container.style.position !== 'relative') container.style.position = 'relative';
      }
      container.appendChild(img);
      return img;
    } else if (part.txbx) {
      // Фигура-надпись: показываем её текст тем же размером, что в Word.
      const sp = document.createElement('span');
      sp.className = 'stu-txbx';
      sp.textContent = part.txbx;
      if (part.txbxSize) sp.style.fontSize = PT(part.txbxSize) + 'px';
      sp.style.lineHeight = '1.15';
      sp.title = 'Надпись (фигура) Word. Пунктирная рамка условна: точное положение и рамка могут не совпадать с оригиналом.';
      if (floating) {
        sp.style.position = 'absolute';
        sp.style.left = Math.max(0, part.offH != null ? part.offH : 0) + 'px';
        sp.style.top = Math.max(-8, part.offV != null ? part.offV : 0) + 'px';
        if (container.style.position !== 'relative') container.style.position = 'relative';
      }
      container.appendChild(sp);
      return sp;
    } else {
      const ph = document.createElement('span'); ph.className = 'stu-img-ph'; ph.textContent = '\u25A3 рисунок';
      if (floating) {
        ph.style.position = 'absolute';
        ph.style.left = Math.max(0, part.offH != null ? part.offH : 0) + 'px';
        ph.style.top = Math.max(-8, part.offV != null ? part.offV : 0) + 'px';
        if (container.style.position !== 'relative') container.style.position = 'relative';
      }
      container.appendChild(ph);
      return ph;
    }
  }

  // Позиционирование рисунка по якорю Word (абсолютно), с резервированием высоты.
  function placeAnchored(container, part, ctx, blockIndex, hasText) {
    const node = appendDrawing(container, Object.assign({}, part, { hAlign: null, _block: false }), ctx, blockIndex);
    if (!node) return;
    const wpx = part.cx ? (part.cx / 914400) * 96 : null;
    const hpx = part.cy ? (part.cy / 914400) * 96 : null;
    const tw = ctx.textWidthPx || MM(170);
    const ml = ctx.marginLeftPx || 0;
    let left = 0;
    if (part.hAlign === 'center') left = Math.max(0, (tw - (wpx || 0)) / 2);
    else if (part.hAlign === 'right') left = Math.max(0, tw - (wpx || 0));
    else if (part.relH === 'column' || part.relH === 'character') left = part.offH || 0;
    else if (part.relH === 'margin' || part.relH === 'page') left = (part.offH || 0) - ml;
    // если в абзаце есть текст — рисунок ставим ниже текста, чтобы не накладывался
    let top = Math.max(-8, part.offV || 0);
    if (hasText) top = Math.max(top, container.offsetHeight || 0);
    node.style.position = 'absolute';
    node.style.left = Math.max(0, left) + 'px';
    node.style.top = top + 'px';
    node.style.margin = '0';
    if (container.style.position !== 'relative') container.style.position = 'relative';
    const need = top + (hpx || 0);
    const cur = parseFloat(container.style.minHeight) || 0;
    if (need > cur) container.style.minHeight = need + 'px';
  }

  function appendMarkedText(container, text, marks, run, blockIndex, runIndex, ctx, cursor) {
    const sorted = marks.slice().sort((a, b) => a.start - b.start);
    let pos = 0;
    for (const mk of sorted) {
      const s = Math.max(pos, Math.min(mk.start, text.length));
      const e = Math.max(s, Math.min(mk.end, text.length));
      if (s > pos) { const sp = makeSpan(text.slice(pos, s), run, blockIndex, runIndex, ctx); container.appendChild(sp); }
      if (e > s) {
        const slice = text.slice(s, e);
        // в оборотах подсвечиваем только слова, пробелы внутри не трогаем
        // (для замечаний именно про пробелы — подсвечиваем и пробелы)
        const pieces = mk.space ? [slice] : slice.split(/(\s+)/).filter((x) => x !== '');
        for (const piece of pieces) {
          const sp = makeSpan(piece, run, blockIndex, runIndex, ctx);
          const isSpace = /^\s+$/.test(piece);
          if (!isSpace || mk.space) {
            sp.classList.add(LEVEL_CLASS[mk.level] || 'stu-mk-info');
            if (mk.punct) sp.classList.add('stu-mk-punct');
            if (mk.src) sp.classList.add('stu-mk-src');
            if (mk.ai) sp.classList.add('stu-ai');
            if (mk.id) sp.setAttribute('data-find', mk.id);
            if (/^[\u200B-\u200D\u2060\uFEFF\u00AD]+$/.test(piece)) sp.classList.add('stu-char-marker');
            if (isSpace) { sp.classList.add('stu-space-mark'); sp.style.whiteSpace = 'pre'; }
          }
          container.appendChild(sp);
        }
      }
      pos = e;
    }
    if (pos < text.length) { const sp = makeSpan(text.slice(pos), run, blockIndex, runIndex, ctx); container.appendChild(sp); }
    cursor.x += measure(text, run);
  }

  function makeSpan(text, run, blockIndex, runIndex, ctx) {
    const span = document.createElement('span');
    span.textContent = text;
    styleSpan(span, run, blockIndex, runIndex, ctx);
    return span;
  }

  function styleSpan(span, run, blockIndex, runIndex, ctx) {
    if (run.font) span.style.fontFamily = '"' + run.font + '", ' + FONT_STACK;
    if (run.size) span.style.fontSize = PT(run.size) + 'px';
    if (run.bold) span.style.fontWeight = '700';
    if (run.italic) span.style.fontStyle = 'italic';
    if (run.underline && !ctx._tocMode) span.style.textDecoration = 'underline';
    if (run.color && /^[0-9a-fA-F]{6}$/.test(run.color)) span.style.color = '#' + run.color;
    if (run.caps) span.style.textTransform = 'uppercase';
    const marks = ctx.runMarks[blockIndex];
    if (marks && marks[runIndex] != null) span.classList.add(LEVEL_CLASS[marks[runIndex]]);
  }

  function renderRuns(container, p, blockIndex, ctx) {
    ctx._curPara = p;
    const role = ctx.roles && ctx.roles[blockIndex] ? ctx.roles[blockIndex].role : null;
    // Подпись рисунка/таблицы, идущая в одном абзаце с изображением: сначала рисунок, потом подпись.
    if ((role === 'figureCaption' || role === 'tableCaption') && p.hasDrawing) {
      for (const r of p.runs || []) for (const pp of (r.parts || [])) {
        if (pp.type === 'drawing') { pp._block = true; appendDrawing(container, pp, ctx, blockIndex); }
      }
      ctx._skipDraw = true;
    }
    const cursor = { x: 0 };
    const marker = ctx.markers[blockIndex];
    let listHang = 0;
    if (marker && role === 'list' && p.numPr && ctx.doc.numbering) {
      const lvl = ctx.doc.numbering.level(p.numPr.numId, p.numPr.ilvl);
      if (lvl && lvl.ind && lvl.ind.hanging) listHang = TW(lvl.ind.hanging);
    }
    if (marker) {
      const firstRun = (p.runs || [])[0] || { size: 14 };
      const m = document.createElement('span');
      m.className = 'stu-listmark';
      const mk = String(marker).replace(/\s+$/, '');
      m.style.whiteSpace = 'pre';
      m.textContent = mk + '\u00a0';
      styleSpan(m, firstRun, blockIndex, -1, ctx);
      m.style.fontFamily = FONT_STACK;   // маркер рисуем обычным шрифтом, без «квадратиков»
      const mkLvl = ctx.markerMarks[blockIndex];
      if (mkLvl) { m.classList.add(LEVEL_CLASS[mkLvl.level]); if (mkLvl.id) m.setAttribute('data-find', mkLvl.id); }
      container.appendChild(m);
      const mW = measure(mk + '\u00a0', firstRun);
      cursor.x += mW;
      // выравниваем текст на позиции отступа (как в Word: маркер в начале, текст на отступе)
      if (listHang > 0) {
        const gap = Math.max(0, listHang - mW);
        if (gap > 0.5) { const sp = document.createElement('span'); sp.className = 'stu-tab'; sp.style.display = 'inline-block'; sp.style.width = gap + 'px'; container.appendChild(sp); cursor.x += gap; }
      }
    }
    // Сегменты по табуляторам; позиции считаем абсолютно (как Word), затем
    // сдвигаем всю строку в соответствии с выравниванием абзаца.
    const segs = []; let seg = [];
    (p.runs || []).forEach((run, ri) => {
      const parts = run.parts && run.parts.length ? run.parts : [{ type: 'text', text: run.text }];
      for (const part of parts) {
        if (part.type === 'tab') { segs.push(seg); seg = []; }
        else if (part.type === 'instrText') { /* код поля */ }
        else seg.push({ run, ri, part });
      }
    });
    segs.push(seg);

    const ind = p.ind || {};
    const hasTabsP = (p.tabs && p.tabs.length) || /\t/.test(p.text || '');
    const centered = hasTabsP && (p.jc === 'center' || p.jc === 'right' || p.jc === 'end');
    const startOffset = centered ? 0 : TW((ind.left || 0) + Math.max(0, ind.firstLine || 0));
    const textW = Math.max(0, (ctx.textWidthPx || 0) - startOffset);
    const tabs = (p.tabs || []).map((t) => ({ pos: TW(t.pos), val: t.val, leader: t.leader })).sort((a, b) => a.pos - b.pos);
    const widths = segs.map(segWidth);
    const starts = [];
    let x = marker ? cursor.x : 0;
    for (let s = 0; s < segs.length; s++) {
      if (s > 0) {
        let stop = null, val = 'left';
        for (const t of tabs) { if (t.pos > x + 0.5) { stop = t.pos; val = t.val || 'left'; break; } }
        if (stop == null) { const step = ctx.tabStep || TW(708); stop = Math.ceil((x + 1) / step) * step; }
        if (stop > textW) stop = textW;
        if (val === 'right' || val === 'end') stop = Math.max(x, stop - widths[s]);
        else if (val === 'center') stop = Math.max(x, stop - widths[s] / 2);
        x = stop;
      }
      starts.push(x); x += widths[s];
    }
    const total = x;
    let offset = 0;
    if (!marker && hasTabsP) {
      if (p.jc === 'center') offset = Math.max(0, (textW - total) / 2);
      else if (p.jc === 'right' || p.jc === 'end') offset = Math.max(0, textW - total);
    }
    ctx._defer = [];
    for (let s = 0; s < segs.length; s++) {
      const desired = offset + starts[s];
      const gap = desired - cursor.x;
      if (gap > 0.5) {
        const sp = document.createElement('span'); sp.className = 'stu-tab';
        sp.style.display = 'inline-block'; sp.style.width = gap + 'px';
        container.appendChild(sp); cursor.x += gap;
      }
      for (const it of segs[s]) renderPart(container, it.run, it.part, blockIndex, it.ri, ctx, cursor);
    }
    // плавающие рисунки выводим после текста абзаца (как в Word: текст, затем рисунок)
    const hasText = /\S/.test((p.text || '').replace(/\uFFFC/g, ''));
    for (const part of ctx._defer) {
      if (part.anchor && part.wrap === 'wrapTopAndBottom') placeAnchored(container, part, ctx, blockIndex, hasText);
      else { if (hasText) part._block = true; appendDrawing(container, part, ctx, blockIndex); }
    }
    ctx._defer = null;
    ctx._skipDraw = false;
  }

  function segWidth(seg) {
    let w = 0;
    for (const x of seg) if (x.part.type === 'text') w += measure(x.part.text, x.run);
    return w;
  }

  function appendTabMeasured(container, p, ctx, cursor, nextWidth) {
    const tabs = p.tabs || [];
    let stop = null, val = 'left', leader = 'none';
    for (const t of tabs) { if (t.pos > cursor.x + 0.5) { stop = t.pos; val = t.val || 'left'; leader = t.leader || 'none'; break; } }
    if (stop == null) { const step = ctx.tabStep || TW(708); stop = Math.ceil((cursor.x + 1) / step) * step; val = 'left'; }
    const ind = p.ind || {};
    const startOffset = TW((ind.left || 0) + Math.max(0, ind.firstLine || 0));
    const maxStop = (ctx.textWidthPx || 1e9) - startOffset;
    if (stop > maxStop) stop = maxStop;
    let w = stop - cursor.x;
    if (val === 'right' || val === 'end') w -= (nextWidth || 0);
    else if (val === 'center') w -= (nextWidth || 0) / 2;
    w = Math.max(0, w);
    const span = document.createElement('span');
    span.className = 'stu-tab';
    span.style.display = 'inline-block';
    span.style.width = w + 'px';
    if (leader === 'dot' || leader === 'middleDot' || leader === 'hyphen' || leader === 'underscore') span.classList.add('stu-tab-' + (leader === 'middleDot' ? 'dot' : leader));
    container.appendChild(span);
    cursor.x += w;
  }

  function renderPart(container, run, part, blockIndex, ri, ctx, cursor) {
    if (part.type === 'text') {
      const marks = (run.parts && run.parts.length === 1) ? ((ctx.charMarks[blockIndex] || {})[ri]) : null;
      if (marks && marks.length) appendMarkedText(container, part.text, marks, run, blockIndex, ri, ctx, cursor);
      else { container.appendChild(makeSpan(part.text, run, blockIndex, ri, ctx)); cursor.x += measure(part.text, run); }
    } else if (part.type === 'tab') {
      appendTab(container, run, cursor, ctx._curPara || {}, ctx);
    } else if (part.type === 'break') {
      container.appendChild(document.createElement('br')); cursor.x = 0;
    } else if (part.type === 'drawing') {
      if (ctx._skipDraw) return;
      if (ctx._defer) ctx._defer.push(part);
      else appendDrawing(container, part, ctx, blockIndex);
    }
  }

  // Содержание: «название … номер» – гибкая строка с отточием.
  function renderTocParagraph(el, p, blockIndex, ctx) {
    el.classList.add('stu-toc');
    ctx._curPara = p;
    const prevTocMode = ctx._tocMode; ctx._tocMode = true;   // в содержании не подчёркиваем
    const ind = p.ind || {};
    el.style.marginLeft = TW((ind.left || 0) + (ind.firstLine || 0)) + 'px';
    if (ind.right) el.style.marginRight = TW(ind.right) + 'px';
    el.style.textIndent = '0';
    const segments = [];
    let seg = [];
    (p.runs || []).forEach((run, ri) => {
      const parts = run.parts && run.parts.length ? run.parts : [{ type: 'text', text: run.text }];
      for (const part of parts) {
        if (part.type === 'tab') { segments.push(seg); seg = []; }
        else if (part.type === 'instrText') { /* пропускаем */ }
        else seg.push({ run, ri, part });
      }
    });
    segments.push(seg);
    const nonEmpty = segments.filter((s) => s.some((x) => (x.part.type === 'text' && x.part.text.trim()) || x.part.type === 'drawing'));
    nonEmpty.forEach((s, si) => {
      const cell = document.createElement('span');
      cell.className = 'stu-toc-cell';
      const cursor = { x: 0 };
      for (const x of s) renderPart(cell, x.run, x.part, blockIndex, x.ri, ctx, cursor);
      el.appendChild(cell);
      if (si < nonEmpty.length - 1) {
        const lead = document.createElement('span');
        lead.className = 'stu-toc-leader';
        el.appendChild(lead);
      }
    });
    ctx._tocMode = prevTocMode;
  }

  function buildParagraph(p, blockIndex, ctx) {
    const el = document.createElement('p');
    el.className = 'stu-p';
    el.setAttribute('data-block', blockIndex);
    el.id = 'stu-block-' + blockIndex;
    const role = ctx.roles && ctx.roles[blockIndex] ? ctx.roles[blockIndex].role : null;
    // pre-wrap оставляем только там, где важны пробелы (титул, ведущие пробелы).
    // У обычного текста он не схлопывает концевые пробелы и строка вылезает
    // за правое поле.
    const needsPre = role === 'title' || /^\s/.test(p.text || '');
    el.style.whiteSpace = needsPre ? 'pre-wrap' : 'normal';
    styleParagraphText(el, p);
    // Перечисления: маркер и текст ставим по отступам из нумерации (как в Word).
    if (role === 'list' && p.numPr && ctx.doc.numbering) {
      const lvl = ctx.doc.numbering.level(p.numPr.numId, p.numPr.ilvl);
      if (lvl && lvl.ind) {
        const li = lvl.ind;
        if (li.left != null) el.style.marginLeft = TW(li.left) + 'px';
        if (li.hanging) el.style.textIndent = (-TW(li.hanging)) + 'px';
        else if (li.firstLine != null) el.style.textIndent = TW(li.firstLine) + 'px';
      }
    }
    const hasTabs = (p.tabs && p.tabs.length) || /\t/.test(p.text || '');
    // Центрирование/право без табуляторов отдаём браузеру (как Word), поэтому
    // text-align трогаем только при табуляторах, где позиции считаем сами.
    if (hasTabs && role !== 'toc') el.style.textAlign = 'left';
    if (role === 'image') { el.style.fontSize = '0'; el.style.lineHeight = 'normal'; }  // убираем лишний интервал вокруг картинки
    if (role === 'toc' && /\t/.test(p.text || '')) renderTocParagraph(el, p, blockIndex, ctx);
    else renderRuns(el, p, blockIndex, ctx);
    if (p.isEmpty && !ctx.markers[blockIndex]) { el.style.minHeight = el.style.lineHeight; }
    applyBlockMark(el, blockIndex, ctx);
    return el;
  }

  function applyBlockMark(el, blockIndex, ctx) {
    const marks = ctx.blockMarks[blockIndex];
    const ai = ctx.aiBlock[blockIndex];
    const src = ctx.blockSrc && ctx.blockSrc[blockIndex];
    const ind = ctx.indentMark && ctx.indentMark[blockIndex];
    if ((!marks || !marks.length) && !ai && !src && !ind) return;
    if (marks && marks.length) {
      let worst = 'info';
      for (const lv of marks) if (LEVEL_RANK[lv] > LEVEL_RANK[worst]) worst = lv;
      el.classList.add(LEVEL_CLASS[worst]);
    }
    if (ai) el.classList.add('stu-ai');
    if (src) el.classList.add('stu-src-block');
    if (ind) el.classList.add('stu-indent-line');
    el.classList.add('stu-marked');
    el.setAttribute('data-findings', (marks ? marks.length : 0) + (ai ? 1 : 0));
    if (ctx.blockFind[blockIndex] != null) el.setAttribute('data-find', ctx.blockFind[blockIndex]);
  }

  function buildTableCellParagraph(p, blockIndex, ctx) {
    const el = document.createElement('div');
    el.className = 'stu-p stu-cell-p';
    styleParagraphText(el, p);
    el.style.margin = '0';
    renderRuns(el, p, blockIndex, ctx);
    return el;
  }

  function buildTable(tbl, blockIndex, ctx) {
    const table = document.createElement('table');
    table.className = 'stu-table';
    table.setAttribute('data-block', blockIndex);
    table.id = 'stu-block-' + blockIndex;
    if (tbl.grid && tbl.grid.length) {
      const total = tbl.grid.reduce((a, b) => a + b, 0) || 1;
      const cg = document.createElement('colgroup');
      tbl.grid.forEach((w) => { const col = document.createElement('col'); col.style.width = ((w / total) * 100).toFixed(2) + '%'; cg.appendChild(col); });
      table.appendChild(cg);
      table.style.tableLayout = 'fixed';
    }
    tbl.rows.forEach((row) => {
      const tr = document.createElement('tr');
      row.cells.forEach((cell) => {
        const td = document.createElement('td');
        if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
        cell.paras.forEach((p) => td.appendChild(buildTableCellParagraph(p, blockIndex, ctx)));
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    applyBlockMark(table, blockIndex, ctx);
    return table;
  }

  // --------------------------------------------------------------- grouping
  function makeItems(doc, forcedPageStarts, roles) {
    const items = [];
    const blocks = doc.blocks;
    const roleOf = (i) => (roles && roles[i] ? roles[i].role : null);
    const sectionStarts = new Set(doc.sections.map((s, i) => (i > 0 ? s.startIndex : -1)));
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === 'p' && /^Таблица\s+/i.test(b.text || '') && blocks[i + 1] && blocks[i + 1].kind === 'table') {
        items.push({ indices: [i, i + 1], forcedBefore: forcedPageStarts.has(i) || !!b.pageBreakBefore, sectionStart: sectionStarts.has(i) });
        i++; continue;
      }
      // (рисунок и подпись не группируем — как в Word, поток сохраняем)
      const all = [];
      if (b.kind === 'p' && b.runs) for (const r of b.runs) for (const pp of (r.parts || [])) all.push(pp);
      while (all.length && all[all.length - 1].type === 'text' && all[all.length - 1].text === '') all.pop();
      const trailing = all.some((x) => x.type === 'pagebreak');
      items.push({ indices: [i], forcedBefore: forcedPageStarts.has(i) || !!b.pageBreakBefore, forcedAfter: !!trailing, sectionStart: sectionStarts.has(i) });
    }
    return items;
  }

  function sectionForBlock(doc, blockIndex) {
    let res = doc.sections[0];
    for (const s of doc.sections) if (blockIndex >= s.startIndex) res = s;
    return res;
  }

  // ---------------------------------------------------------------- rulers
  function buildHRuler(g) {
    const r = document.createElement('div');
    r.className = 'stu-hrule';
    r.style.width = MM(g.wmm) + 'px';
    const mL = document.createElement('div'); mL.className = 'stu-rule-margin'; mL.style.left = '0'; mL.style.width = MM(g.ml) + 'px';
    const mR = document.createElement('div'); mR.className = 'stu-rule-margin'; mR.style.right = '0'; mR.style.width = MM(g.mr) + 'px';
    r.appendChild(mL); r.appendChild(mR);
    // Ноль линейки - у границы текста (после левого поля).
    const textLeft = MM(g.ml);
    const textRight = MM(g.wmm - g.mr);
    const textW = Math.max(0, textRight - textLeft);
    for (let cm = 0; cm <= Math.floor(textW / 10) + 1; cm++) {
      const x = textLeft + MM(cm * 10);
      if (x > textRight) break;
      const t = document.createElement('div'); t.className = 'stu-rule-tick'; t.style.left = x + 'px'; r.appendChild(t);
      const l = document.createElement('div'); l.className = 'stu-rule-lab'; l.style.left = x + 'px'; l.textContent = cm; r.appendChild(l);
    }
    // граница текста (начало координат)
    const zero = document.createElement('div'); zero.className = 'stu-rule-zero'; zero.style.left = textLeft + 'px'; r.appendChild(zero);
    // ориентир абзацного отступа 12,5 мм (на линейке это 1,25)
    const ref = document.createElement('div'); ref.className = 'stu-rule-indent-ref'; ref.style.left = (textLeft + MM(12.5)) + 'px'; ref.title = 'Абзацный отступ 12,5 мм'; r.appendChild(ref);
    return r;
  }
  function buildVRuler(g) {
    const r = document.createElement('div');
    r.className = 'stu-vrule'; r.style.height = MM(g.hmm) + 'px';
    const mT = document.createElement('div'); mT.className = 'stu-rule-margin-v'; mT.style.top = '0'; mT.style.height = MM(g.mt) + 'px';
    const mB = document.createElement('div'); mB.className = 'stu-rule-margin-v'; mB.style.bottom = '0'; mB.style.height = MM(g.mb) + 'px';
    r.appendChild(mT); r.appendChild(mB);
    for (let cm = 0; cm <= Math.ceil(g.hmm / 10); cm++) {
      const y = MM(cm * 10); if (y > MM(g.hmm)) break;
      const t = document.createElement('div'); t.className = 'stu-rule-tick-v'; t.style.top = y + 'px'; r.appendChild(t);
      if (cm > 0 && cm <= Math.floor(g.hmm / 10)) { const l = document.createElement('div'); l.className = 'stu-rule-lab-v'; l.style.top = y + 'px'; l.textContent = cm; r.appendChild(l); }
    }
    return r;
  }

  // ---------------------------------------------------------------- footer
  function footerForPage(section, pageIndex) {
    const f = (section && section.footers) || {};
    if (pageIndex === 0 && f.first) return f.first;
    return f.default || f.first || null;
  }

  // ---------------------------------------------------------------- split paragraphs
  function textLeaves(root) {
    const arr = [];
    (function w(n) {
      for (const c of Array.from(n.childNodes)) {
        if (c.nodeType === 3) arr.push(c);
        else if (c.nodeType === 1 && c.tagName !== 'BR') w(c);
      }
    })(root);
    return arr;
  }
  // Делит абзац на две части так, чтобы первая поместилась в availPx (как перенос в Word).
  function splitElement(el, availPx, widthPx) {
    if (el.tagName !== 'P') return null;
    const leaves = textLeaves(el);
    const total = leaves.reduce((a, n) => a + (n.nodeValue || '').length, 0);
    if (total < 40) return null;
    const host = document.createElement('div');
    host.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;';
    host.style.width = widthPx + 'px';
    // Важно: клон должен полностью повторять абзац (в т.ч. абзацный отступ),
    // иначе замер врёт и голова не влезает (обрезается последняя строка).
    const clone = el.cloneNode(true);
    host.appendChild(clone); document.body.appendChild(host);
    const cleaves = textLeaves(clone);
    for (const n of cleaves) n.__full = n.nodeValue;
    const totalClone = cleaves.reduce((a, n) => a + n.__full.length, 0);
    const setKeep = (K) => { let left = K; for (const n of cleaves) { const take = Math.max(0, Math.min(n.__full.length, left)); n.nodeValue = n.__full.slice(0, take); left -= take; } };
    let lo = 0, hi = totalClone, best = 0;
    for (let it = 0; it < 18 && lo <= hi; it++) { const mid = (lo + hi) >> 1; setKeep(mid); if (clone.scrollHeight <= availPx) { best = mid; lo = mid + 1; } else hi = mid - 1; }
    host.remove();
    if (best <= 0 || best >= totalClone) return null;
    const charAt = (i) => { let left = i; for (const n of leaves) { const t = n.nodeValue || ''; if (left < t.length) return t[left]; left -= t.length; } return ''; };
    let cut = best;
    while (cut > 0 && charAt(cut - 1) !== ' ') cut--;
    if (cut <= 0) return null;
    const head = el.cloneNode(true), tail = el.cloneNode(true);
    const hl = textLeaves(head), tl = textLeaves(tail);
    let left = cut;
    for (const n of hl) { const t = n.nodeValue || ''; if (left >= t.length) { left -= t.length; } else { n.nodeValue = t.slice(0, left); left = 0; } }
    left = cut; let firstTail = true;
    for (const n of tl) { const t = n.nodeValue || ''; if (left >= t.length) { n.nodeValue = ''; left -= t.length; } else { const rest = t.slice(left); n.nodeValue = firstTail ? rest.replace(/^\s+/, '') : rest; left = 0; firstTail = false; } }
    if (!textLeaves(tail).some((n) => (n.nodeValue || '').trim())) return null;
    tail.removeAttribute('id');
    return { head, tail };
  }

  // ---------------------------------------------------------------- render
  function renderDocument(doc, roles, findings, viewer) {
    viewer.innerHTML = '';
    const ctx = { doc, runMarks: {}, blockMarks: {}, charMarks: {}, aiBlock: {}, markerMarks: {}, blockFind: {}, blockSrc: {}, indentMark: {}, roles, markers: computeMarkers(doc), tabStep: TW((doc.settings && doc.settings.defaultTabStop) || 708) };
    for (const f of findings) {
      if (f.blockIndex == null || f.pageLevel) continue;
      if (ctx.blockFind[f.blockIndex] == null) ctx.blockFind[f.blockIndex] = f.id;
      if (f.title && /отступ/i.test(f.title)) { ctx.indentMark[f.blockIndex] = true; continue; }
      if (f.chars && f.chars.length) {
        const m = ctx.charMarks[f.blockIndex] = ctx.charMarks[f.blockIndex] || {};
        for (const c of f.chars) for (const rng of c.ranges) (m[c.runIndex] = m[c.runIndex] || []).push({ start: rng[0], end: rng[1], level: f.level, ai: !!f.ai, punct: f.category === 'punct', src: f.category === 'sources', space: !!(f.title && /пробел/i.test(f.title)), id: f.id });
      }
      if (f.isMarker) {
        const cur = ctx.markerMarks[f.blockIndex];
        if (!cur || LEVEL_RANK[f.level] > LEVEL_RANK[cur.level]) ctx.markerMarks[f.blockIndex] = { level: f.level, id: f.id };
        continue;
      }
      if (f.category === 'sources' && !(f.chars && f.chars.length)) ctx.blockSrc[f.blockIndex] = true;
      if (f.blockIndices && f.blockIndices.length) {
        for (const bi of f.blockIndices) (ctx.blockMarks[bi] = ctx.blockMarks[bi] || []).push(f.level);
        continue;
      }
      if (f.runIndices && f.runIndices.length) {
        ctx.runMarks[f.blockIndex] = ctx.runMarks[f.blockIndex] || {};
        for (const ri of f.runIndices) {
          const cur = ctx.runMarks[f.blockIndex][ri];
          if (!cur || LEVEL_RANK[f.level] > LEVEL_RANK[cur]) ctx.runMarks[f.blockIndex][ri] = f.level;
        }
      } else if (!(f.chars && f.chars.length)) {
        (ctx.blockMarks[f.blockIndex] = ctx.blockMarks[f.blockIndex] || []).push(f.level);
        if (f.ai) ctx.aiBlock[f.blockIndex] = true;
      }
    }

    const forcedPageStarts = global.STURules.computeForcedPageStarts(doc);
    const items = makeItems(doc, forcedPageStarts, roles);
    // С новой страницы начинаем только структурные элементы (ВВЕДЕНИЕ, СОДЕРЖАНИЕ…),
    // но только если в файле уже нет разрыва рядом (иначе как в Word).
    for (const it of items) {
      const r = roles[it.indices[0]];
      if (r && r.role === 'structural') {
        const prev = it.indices[0] > 0 ? doc.blocks[it.indices[0] - 1] : null;
        const hasBreakNear = (doc.blocks[it.indices[0]] && doc.blocks[it.indices[0]].pageBreakBefore) || !!it.forcedBefore ||
          (prev && prev.kind === 'p' && prev.runs && prev.runs.some((run) => (run.parts || []).some((p) => p.type === 'pagebreak')));
        if (!hasBreakNear) it.forcedBefore = true;
      }
    }
    const pages = [];
    const pageForBlock = {};
    const blockEls = {};
    let cur = null;

    function newPage(section) {
      const g = geomOf(section);
      ctx.textWidthPx = MM(g.wmm - g.ml - g.mr);
      ctx.marginLeftPx = MM(g.ml);
      const wrap = document.createElement('div');
      wrap.className = 'stu-page-wrap';
      wrap.setAttribute('data-page', pages.length);

      const hruleRow = document.createElement('div');
      hruleRow.className = 'stu-hrule-row';
      hruleRow.appendChild(buildHRuler(g));

      const sheetRow = document.createElement('div');
      sheetRow.className = 'stu-sheet-row';
      const vrule = buildVRuler(g);
      const sheet = document.createElement('div');
      sheet.className = 'stu-sheet';
      sheet.style.width = MM(g.wmm) + 'px';
      sheet.style.height = MM(g.hmm) + 'px';
      const flow = document.createElement('div');
      flow.className = 'stu-flow';
      flow.style.left = MM(g.ml) + 'px';
      flow.style.top = MM(g.mt) + 'px';
      flow.style.width = MM(g.wmm - g.ml - g.mr) + 'px';
      flow.style.height = MM(g.hmm - g.mt - g.mb) + 'px';
      sheet.appendChild(flow);

      // колонтитул
      const pageIndex = pages.length;
      const foot = footerForPage(section, pageIndex);
      let fdiv = null;
      if (foot && (foot.hasPageField || (foot.text || '').trim())) {
        fdiv = document.createElement('div');
        fdiv.className = 'stu-footer';
        fdiv.style.left = MM(g.ml) + 'px';
        fdiv.style.width = MM(g.wmm - g.ml - g.mr) + 'px';
        fdiv.style.bottom = MM(Math.max(4, (g.mb - 8))) + 'px';
        if (foot.hasPageField) {
          fdiv.textContent = String(pageIndex + 1);
          fdiv.style.fontFamily = FONT_STACK;
          fdiv.style.fontSize = PT(14) + 'px';
          const p0 = (foot.paras || [])[0];
          fdiv.style.textAlign = (p0 && p0.jc) ? 'center' : 'center';
        } else {
          fdiv.textContent = (foot.text || '').trim();
        }
        sheet.appendChild(fdiv);
      }

      sheetRow.appendChild(vrule);
      sheetRow.appendChild(sheet);
      wrap.appendChild(hruleRow);
      wrap.appendChild(sheetRow);
      viewer.appendChild(wrap);

      const page = { index: pages.length, wrap, sheet, flow, geom: g, section, blocks: [], textH: MM(g.hmm - g.mt - g.mb), footerEl: fdiv };
      pages.push(page);
      return page;
    }

    for (const item of items) {
      const firstIdx = item.indices[0];
      const section = sectionForBlock(doc, firstIdx);
      if (!cur) cur = newPage(section);
      if (cur.flow.childElementCount > 0 && (item.forcedBefore || (item.sectionStart && cur.section !== section))) cur = newPage(section);
      const nodes = [];
      for (const idx of item.indices) {
        const b = doc.blocks[idx];
        nodes.push({ idx, node: b.kind === 'table' ? buildTable(b, idx, ctx) : buildParagraph(b, idx, ctx) });
      }
      const hadContent = cur.flow.childElementCount > 0;
      const group = nodes.length > 1;
      for (const n of nodes) cur.flow.appendChild(n.node);
      // Титульный лист не разрываем: допускаем небольшой перелив (иначе
      // последняя строка «уезжает» на следующий лист из-за расхождения в метриках).
      const firstRole = roles && roles[firstIdx] ? roles[firstIdx].role : null;
      const tol = firstRole === 'title' ? 20 : 1;
      if (cur.flow.scrollHeight > cur.textH + tol) {
        // одиночный абзац: переносим часть текста на следующую страницу, как Word
        if (!group && nodes[0].node.tagName === 'P') {
          const idx = nodes[0].idx;
          let node = nodes[0].node; let guard = 0;
          const headPage = cur.index; let headEl = null;
          while (cur.flow.scrollHeight > cur.textH + 1 && guard++ < 30) {
            const avail = cur.textH - node.offsetTop;
            let ok = false;
            if (avail > 24) {
              const sp = splitElement(node, avail, cur.flow.clientWidth);
              if (sp) {
                cur.flow.removeChild(node);
                cur.flow.appendChild(sp.head);
                if (!headEl) headEl = sp.head;
                cur = newPage(section);
                node = sp.tail; node.style.marginTop = '0'; node.style.textIndent = '0';
                cur.flow.appendChild(node); ok = true;
              }
            }
            if (!ok) {
              if (cur.flow.childElementCount > 1) { cur.flow.removeChild(node); cur = newPage(section); cur.flow.appendChild(node); ok = true; }
              else break;
            }
          }
          pageForBlock[idx] = headEl ? headPage : cur.index;
          blockEls[idx] = headEl || node; cur.blocks.push(idx);
          if (item.forcedAfter) cur = newPage(sectionForBlock(doc, Math.min(firstIdx + 1, doc.blocks.length - 1)));
          continue;
        }
        if (hadContent && cur.flow.childElementCount > nodes.length) {
          for (const n of nodes) cur.flow.removeChild(n.node);
          cur = newPage(section);
          for (const n of nodes) cur.flow.appendChild(n.node);
        }
      }
      // Заголовок не оставляем один внизу листа: переносим его вместе со следующим текстом.
      if (nodes.length === 1 && (roles[firstIdx] && (roles[firstIdx].role === 'heading' || roles[firstIdx].role === 'structural'))) {
        const nd = nodes[0].node;
        const nxt = item.indices[0] + 1;
        const nrole = roles[nxt] ? roles[nxt].role : null;
        const bottom = nd.offsetTop + nd.offsetHeight;
        if (cur.flow.childElementCount > 1 && nrole && nrole !== 'heading' && nrole !== 'structural' && (cur.textH - bottom) < 60) {
          cur.flow.removeChild(nd);
          cur = newPage(section);
          cur.flow.appendChild(nd);
        }
      }
      for (const n of nodes) { pageForBlock[n.idx] = cur.index; blockEls[n.idx] = n.node; cur.blocks.push(n.idx); }
      if (item.forcedAfter) cur = newPage(sectionForBlock(doc, Math.min(firstIdx + 1, doc.blocks.length - 1)));
    }

    const findingsByPage = new Map();
    for (const f of findings) {
      if (f.pageLevel) continue;   // замечания уровня листа обработаем ниже
      const pi = f.blockIndex != null ? pageForBlock[f.blockIndex] : null;
      f.page = pi != null ? pi : null;
      if (pi != null) {
        f.element = blockEls[f.blockIndex] || null;
        if (!findingsByPage.has(pi)) findingsByPage.set(pi, []);
        findingsByPage.get(pi).push(f);
      }
      if (pi != null && f.category === 'font' && blockEls[f.blockIndex]) {
        const page = pages[pi]; const b = doc.blocks[f.blockIndex]; const g = page.geom;
        const leftMM = g.ml + (b && b.ind && b.ind.left ? twMM(b.ind.left) : 0) + (b && b.ind && b.ind.firstLine ? twMM(b.ind.firstLine) : 0);
        const mk = document.createElement('div');
        mk.className = 'stu-rule-err' + (f.level === 'error' ? ' err' : ' warn');
        mk.style.left = MM(leftMM) + 'px'; mk.title = f.title;
        const hr = page.wrap.querySelector('.stu-hrule');
        if (hr) hr.appendChild(mk);
      }
    }

    // Замечания уровня листа (поля, формат, номер страницы): обводим лист
    // или колонтитул, чтобы было понятно, к какому месту относится замечание.
    for (const f of findings) {
      if (!f.pageLevel) continue;
      let targets = [];
      if (f.category === 'page' && f.blockIndex != null) {
        const sect = doc.sections.find((s) => s.startIndex === f.blockIndex);
        targets = pages.filter((p) => p.section === sect);
      }
      if (f.category === 'pagenum') {
        if (f.blockIndex === 0) targets = [pages[0]];
        else { const wp = pages.find((p) => p.footerEl); targets = wp ? [wp] : [pages[0]]; }
      }
      if (!targets.length) targets = [pages[0]];
      f.page = targets[0].index;
      let el = targets[0].sheet;
      if (f.category === 'pagenum') {
        const withFooter = targets[0].footerEl;
        if (withFooter) el = withFooter;
        else {
          // на титульном листе колонтитула нет: показываем место, где стоял бы номер
          const mark = document.createElement('div');
          const g0 = targets[0].geom;
          mark.className = 'stu-footer stu-num-mark';
          mark.style.left = MM(g0.ml) + 'px';
          mark.style.width = MM(g0.wmm - g0.ml - g0.mr) + 'px';
          mark.style.bottom = MM(Math.max(4, (g0.mb - 8))) + 'px';
          mark.textContent = String(targets[0].index + 1);
          targets[0].sheet.appendChild(mark);
          el = mark;
        }
      }
      f.element = el;
      const cls = f.level === 'error' ? 'stu-page-err' : (f.level === 'warning' ? 'stu-page-warn' : 'stu-page-info');
      for (const p of targets) {
        p.sheet.classList.add(cls);
        p.sheet.setAttribute('data-page-finding', '1');
        if (!findingsByPage.has(p.index)) findingsByPage.set(p.index, []);
        findingsByPage.get(p.index).push(f);
      }
      if (el !== targets[0].sheet) el.classList.add(cls);
      el.setAttribute('data-find-id', f.id);
    }

    // Делаем ссылки в тексте кликабельными (с подтверждением на стороне интерфейса).
    linkify(viewer);

    return { pages, pageForBlock, blockEls, findingsByPage };
  }

  function linkify(root) {
    const nodes = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let n; while ((n = walker.nextNode())) nodes.push(n);
    for (const node of nodes) {
      const text = node.nodeValue || '';
      const re = /((?:https?:\/\/|ftp:\/\/|www\.)[^\s<>"']+)/gi;
      if (!re.test(text)) continue;
      re.lastIndex = 0;
      const frag = document.createDocumentFragment();
      let last = 0, m;
      while ((m = re.exec(text))) {
        if (m.index > last) frag.appendChild(document.createTextNode(text.slice(last, m.index)));
        const a = document.createElement('a');
        a.className = 'stu-link';
        a.href = m[0];
        a.setAttribute('data-url', m[0]);
        a.textContent = m[0];
        a.addEventListener('click', (e) => {
          e.preventDefault(); e.stopPropagation();
          if (global.STULink) global.STULink(m[0]); else window.open(m[0], '_blank', 'noopener');
        });
        frag.appendChild(a);
        last = m.index + m[0].length;
      }
      if (last < text.length) frag.appendChild(document.createTextNode(text.slice(last)));
      if (node.parentNode) node.parentNode.replaceChild(frag, node);
    }
  }

  global.STURender = { renderDocument, geomOf, PX_PER_MM };
})(window);

/* ==========================================================================
   СТУ-Нормоконтроль – интерфейс: загрузка файла, экран обработки,
   панель замечаний, список листов и развёртка документа.
   ========================================================================== */
(function () {
  'use strict';

  const $ = (sel) => document.querySelector(sel);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ---- сохранение последнего файла между обновлениями страницы (IndexedDB)
  const IDB_NAME = 'stu-normokontrol', IDB_STORE = 'files';
  function idbOpen() {
    return new Promise((res, rej) => {
      if (!('indexedDB' in window)) return rej(new Error('no idb'));
      const r = indexedDB.open(IDB_NAME, 1);
      r.onupgradeneeded = () => { try { r.result.createObjectStore(IDB_STORE); } catch (e) { } };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbPut(key, val) {
    try { const db = await idbOpen(); await new Promise((res, rej) => { const tx = db.transaction(IDB_STORE, 'readwrite'); tx.objectStore(IDB_STORE).put(val, key); tx.oncomplete = res; tx.onerror = () => rej(tx.error); }); } catch (e) { }
  }
  async function idbGet(key) {
    try { const db = await idbOpen(); return await new Promise((res, rej) => { const tx = db.transaction(IDB_STORE, 'readonly'); const rq = tx.objectStore(IDB_STORE).get(key); rq.onsuccess = () => res(rq.result || null); rq.onerror = () => rej(rq.error); }); } catch (e) { return null; }
  }
  async function idbDel(key) {
    try { const db = await idbOpen(); await new Promise((res) => { const tx = db.transaction(IDB_STORE, 'readwrite'); tx.objectStore(IDB_STORE).delete(key); tx.oncomplete = res; tx.onerror = res; }); } catch (e) { }
  }
  function saveLast(name, buf, docType) { try { return idbPut('last', { name: name, buf: buf, docType: docType, at: Date.now() }); } catch (e) { } }

  // История проверок (локально): имя, дата, счётчики, тип.
  async function saveHistory(name, buf, docType) {
    try {
      const counts = {};
      for (const f of state.findings) counts[f.level] = (counts[f.level] || 0) + 1;
      const rec = { name: name, at: Date.now(), err: counts.error || 0, warn: counts.warning || 0, info: counts.info || 0, type: docType };
      let h = (await idbGet('history')) || [];
      h = h.filter((x) => x.name !== name);
      h.unshift(rec); h = h.slice(0, 20);
      await idbPut('history', h);
      await idbPut('blob:' + name, buf);
    } catch (e) { }
  }
  async function renderHistory() {
    const card = $('#history-card'), list = $('#history-list');
    if (!card || !list) return;
    let h = [];
    try { h = (await idbGet('history')) || []; } catch (e) { }
    if (!h.length) { card.style.display = 'none'; list.innerHTML = ''; return; }
    card.style.display = '';
    list.innerHTML = '';
    const shown = state.historyExpanded ? h : h.slice(0, 5);
    shown.forEach((r) => {
      const el = document.createElement('div'); el.className = 'hist-item';
      const nm = document.createElement('span'); nm.className = 'hist-name'; nm.textContent = r.name; nm.title = r.name;
      const meta = document.createElement('span'); meta.className = 'hist-meta';
      meta.innerHTML = '<span class="badge err">' + (r.err | 0) + '</span><span class="badge warn">' + (r.warn | 0) + '</span><span class="badge info">' + (r.info | 0) + '</span>';
      const d = document.createElement('span'); d.textContent = new Date(r.at).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
      meta.appendChild(d);
      el.appendChild(nm); el.appendChild(meta);
      el.addEventListener('click', async () => {
        let buf = null; try { buf = await idbGet('blob:' + r.name); } catch (e) { }
        if (!buf) { alert('Файл «' + r.name + '» больше не хранится – загрузите его снова.'); return; }
        if (r.type && DOC_TYPES.some((x) => x[0] === r.type)) state.docType = r.type;
        showProcessing(r.name);
        processBuffer(r.name, buf);
      });
      list.appendChild(el);
    });
    if (h.length > 5) {
      const more = document.createElement('button');
      more.className = 'btn';
      more.style.marginTop = '8px';
      more.textContent = state.historyExpanded ? 'Свернуть' : ('Показать все (' + h.length + ')');
      more.addEventListener('click', () => { state.historyExpanded = !state.historyExpanded; renderHistory(); });
      list.appendChild(more);
    }
  }
  async function clearHistory() {
    try {
      const h = (await idbGet('history')) || [];
      for (const r of h) { try { await idbDel('blob:' + r.name); } catch (e) { } }
      await idbDel('history');
    } catch (e) { }
    renderHistory();
  }

  const state = {
    doc: null, roles: null, findings: [], render: null,
    fileName: '', docType: 'vkr', filter: 'all', buffer: null, autoType: null, manualType: false, titleEnd: 0, historyExpanded: false,
  };

  const DOC_TYPES = [
    ['praktika', 'Отчёт о практике'], ['kursovaya', 'Курсовая'],
    ['otchet', 'Отчёт (лаб. / практ.)'], ['referat', 'Реферат'],
    ['vkr', 'ВКР (бакалаврская)'], ['diplom', 'Дипломная (ДР / ДП)'],
  ];

  // ------------------------------------------------------------------ theme
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    localStorage.setItem('stu-theme', t);
    const sun = $('#ico-sun'), moon = $('#ico-moon');
    if (sun && moon) { sun.style.display = t === 'dark' ? 'none' : 'block'; moon.style.display = t === 'dark' ? 'block' : 'none'; }
  }
  function toggleTheme(e) {
    const cur = document.documentElement.getAttribute('data-theme') || 'light';
    const next = cur === 'dark' ? 'light' : 'dark';
    const x = e && e.clientX ? (e.clientX / window.innerWidth * 100) : 90;
    const y = e && e.clientY ? (e.clientY / window.innerHeight * 100) : 4;
    document.documentElement.style.setProperty('--vt-x', x + '%');
    document.documentElement.style.setProperty('--vt-y', y + '%');
    if (document.startViewTransition) {
      document.documentElement.classList.add('vt-wave');
      const vt = document.startViewTransition(() => applyTheme(next));
      vt.finished.finally(() => document.documentElement.classList.remove('vt-wave'));
    } else {
      applyTheme(next);
    }
  }

  // ------------------------------------------------------------------ screens
  function showStart() {
    $('#screen-start').style.display = 'flex';
    $('#screen-processing').style.display = 'none';
    $('#workspace').classList.remove('on');
    $('#topfile').textContent = '';
    $('#fix-btn').style.display = 'none';
    const bb = $('#back-btn'); if (bb) bb.style.display = 'none';
    const rb = $('#report-btn'); if (rb) rb.style.display = 'none';
    const pb0 = $('#pdf-btn'); if (pb0) pb0.style.display = 'none';
    const tw = $('#toptype-wrap'); if (tw) tw.style.display = 'none';
    renderHistory();
  }
  function showProcessing(name) {
    $('#screen-start').style.display = 'none';
    $('#workspace').classList.remove('on');
    $('#screen-processing').style.display = 'flex';
    $('#proc-name').textContent = name;
    setSteps(-1, 0);
  }
  function showWorkspace() {
    $('#screen-start').style.display = 'none';
    $('#screen-processing').style.display = 'none';
    $('#workspace').classList.add('on');
    $('#fix-btn').style.display = '';
    const bb = $('#back-btn'); if (bb) bb.style.display = '';
    const rb = $('#report-btn'); if (rb) rb.style.display = '';
    const pb = $('#pdf-btn'); if (pb) pb.style.display = '';
  }

  const STEPS = ['Чтение файла', 'Разбор структуры документа', 'Проверка по СТУ 7.5–07–2021', 'Построение развёртки', 'Готово'];
  function setSteps(active, pct) {
    const ul = $('#proc-steps'); ul.innerHTML = '';
    STEPS.forEach((s, i) => {
      const li = document.createElement('li');
      li.className = i < active ? 'done' : (i === active ? 'active' : '');
      li.innerHTML = '<span class="dot"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="M20 6 9 17l-5-5"/></svg></span>' + s;
      ul.appendChild(li);
    });
    $('#proc-bar').style.width = (pct || 0) + '%';
  }

  // ------------------------------------------------------------------ file
  async function handleFile(file) {
    if (!file) return;
    if (/\.pdf$/i.test(file.name)) {
      state.fileName = file.name;
      showProcessing(file.name);
      const buf = await file.arrayBuffer();
      await handlePdf(file.name, buf);
      return;
    }
    if (/\.doc$/i.test(file.name)) {
      if (await tryConvertDoc(file)) return;
      alert('Формат .doc не читается в браузере.\n' +
        'Откройте файл в Word и сохраните как .docx («Сохранить как» → «Документ Word (*.docx)»).\n\n' +
        'Либо запустите локальный сервер (python3 server/server.py) с установленным LibreOffice – тогда .doc преобразуется автоматически.');
      return;
    }
    if (!/\.docx$/i.test(file.name)) {
      alert('Принимаются только файлы Word в форматах .docx и .doc.');
      return;
    }
    state.fileName = file.name;
    showProcessing(file.name);
    const buf = await file.arrayBuffer();
    await processBuffer(file.name, buf);
  }

  // Попытка преобразовать .doc в .docx через локальный сервер.
  async function tryConvertDoc(file) {
    if (!/^https?:$/.test(location.protocol)) return false;
    try {
      const r = await fetch('api/convert-doc', {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: file,
      });
      if (!r.ok) {
        const j = await r.json().catch(() => ({}));
        if (j && j.hint) alert('Сервер не смог преобразовать .doc.\n' + j.hint);
        return true;
      }
      const buf = await r.arrayBuffer();
      await processBuffer(file.name.replace(/\.doc$/i, '.docx'), buf);
      return true;
    } catch (e) {
      return false;
    }
  }

  // Просмотр PDF-файла: страницы как картинки (pdf.js).
  async function handlePdf(name, buf) {
    try {
      if (!window.pdfjsLib) { alert('PDF-модуль не загружен.'); showStart(); return; }
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
      const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
      const viewer = $('#viewer'); viewer.innerHTML = '';
      const pages = [];
      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const vp = page.getViewport({ scale: 2 });
        const canvas = document.createElement('canvas'); canvas.width = vp.width; canvas.height = vp.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
        const vp1 = page.getViewport({ scale: 1 });
        const wrap = document.createElement('div'); wrap.className = 'stu-page-wrap'; wrap.setAttribute('data-page', p - 1);
        const sheet = document.createElement('div'); sheet.className = 'stu-sheet stu-pdf-page';
        sheet.style.width = (vp1.width * 96 / 72) + 'px';
        sheet.style.height = (vp1.height * 96 / 72) + 'px';
        const img = document.createElement('img'); img.src = canvas.toDataURL('image/jpeg', 0.9);
        img.style.width = '100%'; img.style.height = '100%'; img.style.display = 'block';
        sheet.appendChild(img); wrap.appendChild(sheet); viewer.appendChild(wrap);
        pages.push({ index: p - 1, wrap, sheet, blocks: [] });
      }
      state.doc = null; state.roles = null; state.findings = []; state.buffer = null; state.titleEnd = 0;
      state.render = { pages: pages, pageForBlock: {}, blockEls: {}, findingsByPage: new Map() };
      showWorkspace();
      $('#topfile').textContent = name;
      const tw = $('#toptype-wrap'); if (tw) tw.style.display = 'none';
      const fb = $('#fix-btn'); if (fb) fb.style.display = 'none';
      const fm = $('#file-meta');
      const g0 = pages[0] ? { w: pages[0].sheet.style.width, h: pages[0].sheet.style.height } : null;
      if (fm) fm.innerHTML = '<div class="fm-name">' + escapeHtml(name) + '</div>' +
        '<div class="fm-sub">PDF · листов: ' + pdf.numPages + '</div>' +
        '<div class="fm-sub">Просмотр PDF. Проверки по СТУ выполняются для файлов Word.</div>';
      buildFindings();
      buildPages();
      wireViewerScroll();
    } catch (e) {
      alert('Не удалось прочитать PDF.\n' + (e && e.message ? e.message : e));
      showStart();
    }
  }

  async function processBuffer(name, buf) {
    state.fileName = name;
    state.buffer = buf;
    try {
      setSteps(0, 8); await sleep(200);
      setSteps(1, 28); await sleep(150);
      const doc = STUDocx.parse(buf);
      state.doc = doc;
      const det = STURules.detectDocType(doc);
      state.autoType = det;
      if (det && det.type && !state.manualType) state.docType = det.type;
      syncTypeChips();
      setSteps(2, 58); await sleep(150);
      const res = STURules.runChecks(doc, state.docType);
      setSteps(3, 82); await sleep(150);
      showWorkspace();
      $('#topfile').textContent = name;
      finishRender(doc, res, name, buf);
      setSteps(4, 100); await sleep(200);
    } catch (err) {
      console.error(err);
      alert('Не удалось обработать файл.\n' + (err && err.message ? err.message : err));
      showStart();
    }
  }

  function buildFileMeta(doc) {
    const m = $('#file-meta');
    const s = doc.sections[0];
    const g = STURender.geomOf(s);
    const req = STURules.DOC_REQUIREMENTS ? STURules.DOC_REQUIREMENTS[state.docType] : null;
    const typeLabel = req ? req.label : state.docType;
    const auto = state.autoType && state.autoType.confidence === 'high' ? ' (определён автоматически)' : '';
    m.innerHTML =
      '<div class="fm-name">' + escapeHtml(state.fileName) + '</div>' +
      '<div class="fm-sub">' + g.wmm.toFixed(0) + '×' + g.hmm.toFixed(0) + ' мм · ' +
      (g.portrait ? 'книжная' : 'альбомная') + ' · листов: ' + state.render.pages.length + '</div>' +
      '<div class="fm-sub">Тип: ' + escapeHtml(typeLabel) + auto + '</div>';
  }

  function syncTypeChips() {
    document.querySelectorAll('#type-chips .type-chip').forEach((x) => x.classList.toggle('active', x.getAttribute('data-type') === state.docType));
    const sel = $('#toptype'); if (sel) sel.value = state.docType;
    const wrap = $('#toptype-wrap'); if (wrap) wrap.style.display = state.doc ? '' : 'none';
  }

  // Построить развёртку, досчитать проверки содержания по страницам, собрать панели.
  function finishRender(doc, res, name, buf) {
    state.roles = res.roles; state.titleEnd = res.titleEnd; state.tocEntries = res.tocEntries || [];
    let findings = res.findings.slice();
    let rnd = STURender.renderDocument(doc, res.roles, findings, $('#viewer'));
    const extra = [];
    if (STURules.crossCheckToc && res.tocEntries && res.tocEntries.length) {
      STURules.crossCheckToc(doc, res.roles, res.tocEntries, rnd.pageForBlock, (f) => extra.push(f));
    }
    if (extra.length) findings = findings.concat(extra);
    // Убираем замечания про пустую строку, если объект и текст оказались на разных листах.
    const beforeN = findings.length;
    findings = findings.filter((f) => !f.pageRel || rnd.pageForBlock[f.pageRel[0]] === rnd.pageForBlock[f.pageRel[1]]);
    if (extra.length || findings.length !== beforeN) rnd = STURender.renderDocument(doc, res.roles, findings, $('#viewer'));
    state.findings = findings; state.render = rnd;
    buildFileMeta(doc); buildFindings(); buildPages(); wireViewerScroll(); wireDocClicks();
    if (name && buf) { saveLast(name, buf, state.docType); saveHistory(name, buf, state.docType); }
  }

  // Перепроверить текущий документ (например, после смены вида вручную).
  function rerun() {
    if (!state.doc || !state.buffer) return;
    try {
      const res = STURules.runChecks(state.doc, state.docType);
      finishRender(state.doc, res, state.fileName, state.buffer);
    } catch (e) { console.error(e); }
  }

  // ------------------------------------------------------------------ findings
  const ORDER = { error: 0, warning: 1, info: 2 };
  function buildFindings() {
    const list = $('#find-list');
    state.findEls = new Map();
    const counts = { error: 0, warning: 0, info: 0 };
    for (const f of state.findings) counts[f.level] = (counts[f.level] || 0) + 1;
    $('#sum-err').textContent = counts.error; $('#sum-warn').textContent = counts.warning; $('#sum-info').textContent = counts.info;

    const items = state.findings
      .filter((f) => state.filter === 'all' || (state.filter === 'ai' ? f.ai : f.level === state.filter))
      .slice()
      .sort((a, b) => {
        const ai = a.blockIndex == null ? 1e9 : a.blockIndex;
        const bi = b.blockIndex == null ? 1e9 : b.blockIndex;
        if (ai !== bi) return ai - bi;
        return ORDER[a.level] - ORDER[b.level];
      });

    if (!items.length) {
      list.innerHTML = '<div class="no-find">' + (state.findings.length ? 'Нет замечаний выбранного уровня.' : 'Замечаний не найдено – документ соответствует проверяемым требованиям СТУ.') + '</div>';
      return;
    }
    const frag = document.createDocumentFragment();
    for (const f of items) {
      const el = document.createElement('div');
      el.className = 'find-item ' + f.level + (f.ai ? ' ai' : '');
      el.innerHTML =
        '<div class="fi-top"><span class="fi-title">' + escapeHtml(f.title) + '</span>' +
        (f.ai ? '<span class="fi-ai">ИИ</span>' : '') +
        '<span class="fi-badge">' + levelLabel(f.level) + (f.page != null ? ' · лист ' + (f.page + 1) : '') + '</span></div>' +
        '<div class="fi-detail">' + escapeHtml(f.detail) + '</div>' +
        (f.recommend ? '<div class="fi-rec"><b>Как исправить:</b> ' + escapeHtml(f.recommend) + '</div>' : '') +
        '<div class="fi-ref">' + escapeHtml(f.ref || '') + '</div>';
      el.addEventListener('click', () => focusFinding(f));
      el.__finding = f;
      state.findEls.set(f.id, el);
      frag.appendChild(el);
    }
    list.innerHTML = '';
    list.appendChild(frag);
  }

  // Клик по подсвеченному месту в документе открывает замечание слева.
  function wireDocClicks() {
    const v = $('#viewer');
    if (!v || v.__clickBound) return; v.__clickBound = true;
    v.addEventListener('click', (e) => {
      // открываем только то, на что навели (по конкретной подсветке)
      const fs = e.target.closest('[data-find]');
      if (fs) {
        const f = state.findings.find((x) => x.id === fs.getAttribute('data-find'));
        if (f) { flashFinding(f); return; }
      }
      const pg = e.target.closest('[data-find-id]');
      if (pg) {
        const f = state.findings.find((x) => x.id === pg.getAttribute('data-find-id'));
        if (f) flashFinding(f);
      }
    });
  }

  function flashFinding(f) {
    const item = state.findEls && state.findEls.get(f.id);
    if (item) {
      item.scrollIntoView({ behavior: 'smooth', block: 'center' });
      item.classList.remove('stu-flash'); void item.offsetWidth; item.classList.add('stu-flash');
    }
  }

  function focusFinding(f) {
    if (f.element) {
      f.element.scrollIntoView({ behavior: 'smooth', block: 'center' });
      f.element.classList.remove('stu-flash');
      void f.element.offsetWidth;
      f.element.classList.add('stu-flash');
      setActivePage(f.page);
    }
  }

  function levelLabel(l) { return l === 'error' ? 'Ошибка' : l === 'warning' ? 'Предупреждение' : 'Инфо'; }

  // ------------------------------------------------------------------ pages
  function buildPages() {
    const wrap = $('#page-list');
    const byPage = state.render.findingsByPage;
    const frag = document.createDocumentFragment();
    state.render.pages.forEach((p, i) => {
      const fs = byPage.get(i) || [];
      const err = fs.filter((f) => f.level === 'error').length;
      const warn = fs.filter((f) => f.level === 'warning').length;
      const info = fs.filter((f) => f.level === 'info').length;
      const el = document.createElement('div');
      el.className = 'page-item';
      el.setAttribute('data-page', i);
      el.innerHTML = '<span class="pg-num">Лист ' + (i + 1) + '</span><span class="pg-meta">' +
        (err ? '<span class="badge err">' + err + '</span>' : '') +
        (warn ? '<span class="badge warn">' + warn + '</span>' : '') +
        (info ? '<span class="badge info">' + info + '</span>' : '') +
        (!fs.length ? '<span class="badge ok">✓</span>' : '') + '</span>';
      el.addEventListener('click', () => {
        setActivePage(i);
        if (fs.length) { focusFinding(fs[0]); }
        else { p.wrap.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
      });
      frag.appendChild(el);
    });
    wrap.innerHTML = '';
    wrap.appendChild(frag);
  }
  function setActivePage(i) {
    document.querySelectorAll('.page-item').forEach((el) => el.classList.toggle('active', +el.getAttribute('data-page') === i));
  }
  let scrollBound = false;
  function wireViewerScroll() {
    if (scrollBound) return; scrollBound = true;
    const v = $('#viewer');
    let raf = null;
    v.addEventListener('scroll', () => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = null;
        const vTop = v.getBoundingClientRect().top;
        let best = 0, bestD = Infinity;
        state.render.pages.forEach((p, i) => {
          const d = Math.abs(p.wrap.getBoundingClientRect().top - vTop - 10);
          if (d < bestD) { bestD = d; best = i; }
        });
        setActivePage(best);
      });
    }, { passive: true });
  }

  // --------------- исправление и выгрузка .docx
  function exportFixed() {
    if (!state.buffer || !state.doc) return;
    const opts = {
      font: $('#opt-font').checked, align: $('#opt-align').checked, indent: $('#opt-indent').checked,
      spacing: $('#opt-spacing').checked, page: $('#opt-page').checked, dash: $('#opt-dash').checked,
      text: $('#opt-text') ? $('#opt-text').checked : false,
      emph: $('#opt-emph') ? $('#opt-emph').checked : false,
      marker: $('#opt-marker') ? $('#opt-marker').checked : false,
      listEnd: $('#opt-marker') ? $('#opt-marker').checked : false,
      sectionBreak: $('#opt-section') ? $('#opt-section').checked : false,
      cleanEmpties: true,
      titleEnd: state.titleEnd || 0,
    };
    // Новые номера страниц для содержания (по построенной развёртке)
    const tocEl = $('#opt-toc');
    opts.toc = tocEl ? tocEl.checked : false;
    if (opts.toc && state.tocEntries && state.tocEntries.length && state.render) {
      const roles = state.roles, blocks = state.doc.blocks;
      const norm = (s) => (s || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
      const byNum = {}, byTitle = {};
      for (let i = 0; i < blocks.length; i++) {
        const r = roles[i]; if (!r) continue;
        if (r.role === 'heading' && r.number) { if (byNum[r.number] == null) byNum[r.number] = i; }
        if (r.role === 'structural') byTitle[norm(blocks[i].text)] = i;
      }
      const map = {};
      for (const e of state.tocEntries) {
        let bi = null;
        if (e.number != null && byNum[e.number] != null) bi = byNum[e.number];
        else if (e.title && byTitle[norm(e.title)] != null) bi = byTitle[norm(e.title)];
        if (bi == null) continue;
        const p = state.render.pageForBlock[bi];
        if (p == null) continue;
        map[e.index] = p + 1;
      }
      opts.tocPages = map;
    }
    // Пустые строки вокруг рисунков/таблиц (только для тех, что на одном листе).
    opts.inserts = state.findings.filter((f) => f.blankInsert).map((f) => f.blankInsert);
    // Индексы записей списка источников (для точки в конце).
    const src = new Set();
    {
      const roles = state.roles || [];
      let inSrc = false;
      for (let i = 0; i < state.doc.blocks.length; i++) {
        const r = roles[i]; if (!r) continue;
        if (r.role === 'structural' && r.structKey === 'istochniki') { inSrc = true; continue; }
        if (inSrc && r.role === 'structural') inSrc = false;
        if (inSrc) src.add(i);
      }
    }
    opts.sources = src;
    // единый межстрочный интервал по большинству
    let n240 = 0, n360 = 0;
    for (const b of state.doc.blocks) {
      if (b.kind !== 'p') continue;
      const sp = b.spacing || {};
      if (sp.lineRule && sp.lineRule !== 'auto') continue;
      const l = sp.line || 240;
      if (l >= 330) n360++; else n240++;
    }
    opts.lineTarget = n360 > n240 ? 360 : 240;
    try {
      const meta = state.doc.blocks.map((b, i) => ({ kind: b.kind, role: state.roles[i] ? state.roles[i].role : 'body', level: state.roles[i] ? state.roles[i].level : 0 }));
      const bytes = STUFix.buildFixed(state.buffer, meta, opts);
      const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = state.fileName.replace(/\.docx$/i, '') + ' (исправлено).docx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
      $('#fix-dialog').style.display = 'none';
    } catch (e) {
      alert('Не удалось собрать файл.\n' + (e && e.message ? e.message : e));
    }
  }

  // --------------- отчёт о замечаниях (txt)
  function buildReport() {
    if (!state.doc) return;
    const L = { error: 'ОШИБКА', warning: 'ПРЕДУПРЕЖДЕНИЕ', info: 'ИНФО' };
    const c = { error: 0, warning: 0, info: 0 };
    for (const f of state.findings) c[f.level] = (c[f.level] || 0) + 1;
    const req = STURules.DOC_REQUIREMENTS[state.docType];
    const lines = [];
    lines.push('СТУ-Нормоконтроль: отчёт о проверке');
    lines.push('Файл: ' + state.fileName);
    lines.push('Вид документа: ' + ((req && req.label) || state.docType));
    lines.push('Дата: ' + new Date().toLocaleString('ru-RU'));
    lines.push('Замечаний: ошибок ' + c.error + ', предупреждений ' + c.warning + ', инфо ' + c.info);
    lines.push('');
    state.findings.forEach((f, i) => {
      lines.push((i + 1) + '. [' + L[f.level] + (f.ai ? ' / ИИ' : '') + '] ' + (f.page != null ? ('лист ' + (f.page + 1)) : 'общее') + ': ' + f.title);
      lines.push('   ' + f.detail);
      if (f.recommend) lines.push('   Как исправить: ' + f.recommend);
      if (f.ref) lines.push('   ' + f.ref);
      lines.push('');
    });
    const blob = new Blob([lines.join('\n')], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = state.fileName.replace(/\.docx$/i, '') + ' (отчёт).txt';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  // Сохранение развёртки в pdf-файл (готовый файл через html2canvas + jsPDF).
  async function exportPdf() {
    if (!(window.html2canvas && window.jspdf)) { alert('PDF-модуль не загружен.'); return; }
    const sheets = Array.from(document.querySelectorAll('.stu-page-wrap .stu-sheet'));
    if (!sheets.length) return;
    const btn = $('#pdf-btn'); const old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = 'PDF…'; }
    try {
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'portrait' });
      const pw = pdf.internal.pageSize.getWidth();
      const ph = pdf.internal.pageSize.getHeight();
      for (let i = 0; i < sheets.length; i++) {
        const canvas = await window.html2canvas(sheets[i], { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false });
        const img = canvas.toDataURL('image/jpeg', 0.92);
        const ratio = Math.min(pw / canvas.width, ph / canvas.height);
        const w = canvas.width * ratio, h = canvas.height * ratio;
        if (i > 0) pdf.addPage();
        pdf.addImage(img, 'JPEG', (pw - w) / 2, (ph - h) / 2, w, h);
      }
      pdf.save((state.fileName || 'документ').replace(/\.(docx|pdf)$/i, '') + ' (развёртка).pdf');
    } catch (e) {
      alert('Не удалось собрать PDF.\n' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old; }
    }
  }

  // ------------------------------------------------------------------ utils
  function escapeHtml(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

  // Свернуть / развернуть боковую панель.
  function setCollapsed(side, on) {
    const ws = $('#workspace');
    ws.classList.toggle(side + '-collapsed', !!on);
    const panel = side === 'left' ? $('#panel-left') : $('#panel-right');
    if (panel) panel.classList.toggle('collapsed', !!on);
    const btn = document.querySelector('.panel-toggle[data-side="' + side + '"]');
    if (btn) btn.textContent = side === 'left' ? (on ? '›' : '‹') : (on ? '‹' : '›');
    try { localStorage.setItem('stu-collapse-' + side, on ? '1' : '0'); } catch (e) { }
  }

  // Синхронный прогон (для автотеста: результат готов к моменту загрузки страницы).
  function processSync(name, buf) {
    state.fileName = name;
    state.buffer = buf;
    try {
      const doc = STUDocx.parse(buf); state.doc = doc;
      const det2 = STURules.detectDocType(doc); state.autoType = det2;
      if (det2 && det2.type && !state.manualType) state.docType = det2.type;
      syncTypeChips();
      const res = STURules.runChecks(doc, state.docType);
      showWorkspace();
      $('#topfile').textContent = name;
      finishRender(doc, res, name, buf);
    } catch (e) { console.error(e); document.title = 'ОШИБКА: ' + e.message; }
  }

  // ------------------------------------------------------------------ init UI
  function init() {
    applyTheme(localStorage.getItem('stu-theme') || 'light');
    $('#theme-btn').addEventListener('click', toggleTheme);
    const legendBtn = $('#legend-btn'), legendBox = $('#legend');
    if (legendBtn && legendBox) legendBtn.addEventListener('click', () => { legendBox.style.display = legendBox.style.display === 'none' ? '' : 'none'; });

    // чипы типа документа
    const tc = $('#type-chips');
    DOC_TYPES.forEach(([id, label]) => {
      const b = document.createElement('button');
      b.className = 'type-chip' + (id === state.docType ? ' active' : '');
      b.textContent = label; b.setAttribute('data-type', id);
      b.addEventListener('click', () => {
        state.docType = id;
        state.manualType = true;
        tc.querySelectorAll('.type-chip').forEach((x) => x.classList.toggle('active', x === b));
        if (state.doc) rerun();
      });
      tc.appendChild(b);
    });
    // селектор вида документа в верхней панели
    const topSel = $('#toptype');
    if (topSel) {
      DOC_TYPES.forEach(([id, label]) => { const o = document.createElement('option'); o.value = id; o.textContent = label; topSel.appendChild(o); });
      topSel.value = state.docType;
      topSel.addEventListener('change', () => {
        state.docType = topSel.value; state.manualType = true;
        syncTypeChips();
        if (state.doc) rerun();
      });
    }

    // dropzone
    const drop = $('#drop-card');
    const input = $('#file-input');
    $('#pick-btn').addEventListener('click', () => input.click());
    drop.addEventListener('click', (e) => { if (e.target === drop || e.target.classList.contains('drop-icon')) input.click(); });
    input.addEventListener('change', () => { if (input.files[0]) handleFile(input.files[0]); input.value = ''; });
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('dragover'); }));

    // окно: подсветка и приём файла в любом месте (в т.ч. на экране результата)
    const overlay = $('#drop-overlay');
    let dragDepth = 0;
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('dragenter', (e) => {
      if (e.dataTransfer && Array.from(e.dataTransfer.types || []).indexOf('Files') >= 0) { dragDepth++; overlay.style.display = 'flex'; }
    });
    window.addEventListener('dragleave', () => { dragDepth--; if (dragDepth <= 0) { dragDepth = 0; overlay.style.display = 'none'; } });
    window.addEventListener('drop', (e) => {
      e.preventDefault(); dragDepth = 0; overlay.style.display = 'none';
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });

    // фильтры
    document.querySelectorAll('.filter-chip').forEach((c) => c.addEventListener('click', () => {
      state.filter = c.getAttribute('data-level');
      document.querySelectorAll('.filter-chip').forEach((x) => x.classList.toggle('active', x === c));
      buildFindings();
    }));

    $('#new-btn').addEventListener('click', () => { state.doc = null; state.render = null; state.buffer = null; $('#viewer').innerHTML = ''; showStart(); $('#file-input').click(); });
    const pdfBtn = $('#pdf-btn'); if (pdfBtn) pdfBtn.addEventListener('click', exportPdf);

    // исправление и выгрузка .docx
    $('#fix-btn').addEventListener('click', () => { if (state.buffer) $('#fix-dialog').style.display = 'flex'; });
    $('#fix-cancel').addEventListener('click', () => { $('#fix-dialog').style.display = 'none'; });
    $('#fix-ok').addEventListener('click', exportFixed);
    $('#fix-dialog').addEventListener('click', (e) => { if (e.target === $('#fix-dialog')) $('#fix-dialog').style.display = 'none'; });

    // Переход по ссылке с подтверждением
    const linkDlg = $('#link-dialog'); let pendingUrl = '';
    window.STULink = (url) => { pendingUrl = url; $('#link-url').textContent = url; linkDlg.style.display = 'flex'; };
    $('#link-no').addEventListener('click', () => { linkDlg.style.display = 'none'; pendingUrl = ''; });
    $('#link-yes').addEventListener('click', () => { if (pendingUrl) window.open(pendingUrl, '_blank', 'noopener'); linkDlg.style.display = 'none'; pendingUrl = ''; });
    linkDlg.addEventListener('click', (e) => { if (e.target === linkDlg) { linkDlg.style.display = 'none'; pendingUrl = ''; } });
    const rbtn = $('#report-btn'); if (rbtn) rbtn.addEventListener('click', buildReport);

    // история проверок: очистка
    const hclear = $('#history-clear');
    if (hclear) hclear.addEventListener('click', (e) => { e.stopPropagation(); clearHistory(); });

    // сворачивание / разворачивание панелей
    ['left', 'right'].forEach((side) => {
      let on = false;
      try { on = localStorage.getItem('stu-collapse-' + side) === '1'; } catch (e) { }
      setCollapsed(side, on);
      const head = side === 'left' ? $('#head-left') : $('#head-right');
      if (head) head.addEventListener('click', () => setCollapsed(side, !$('#workspace').classList.contains(side + '-collapsed')));
    });
    try {
      const c = new URLSearchParams(location.search).get('collapse');
      if (c) c.split(',').forEach((s) => { if (s === 'left' || s === 'right') setCollapsed(s, true); });
    } catch (e) { }

    // «Новый документ» – вернуться к началу и забыть сохранённый файл
    $('#back-btn').addEventListener('click', () => {
      state.doc = null; state.render = null; state.buffer = null; $('#viewer').innerHTML = '';
      showStart(); idbDel('last');
    });

    // восстановление последнего проверенного файла после обновления страницы
    const hasFileParam = new URLSearchParams(location.search).get('f');
    if (!hasFileParam) {
      idbGet('last').then((rec) => {
        if (!rec || !rec.buf) return;
        if (rec.docType && DOC_TYPES.some((d) => d[0] === rec.docType)) {
          state.docType = rec.docType;
          document.querySelectorAll('.type-chip').forEach((x) => x.classList.toggle('active', x.getAttribute('data-type') === rec.docType));
        }
        showProcessing(rec.name || 'документ');
        processBuffer(rec.name || 'документ.docx', rec.buf);
      }).catch(() => { });
    }

    showStart();

    // Открытие файла по ссылке (?f=путь) – удобно для проверки и встраивания.
    try {
      const f = new URLSearchParams(location.search).get('f');
      if (f) {
        const t = new URLSearchParams(location.search).get('type');
        if (t && DOC_TYPES.some((d) => d[0] === t)) state.docType = t;
        const sync = new URLSearchParams(location.search).get('sync') === '1';
        if (sync) {
          const xhr = new XMLHttpRequest();
          xhr.open('GET', f, false);
          xhr.overrideMimeType('text/plain; charset=x-user-defined');
          xhr.send(null);
          let bin = xhr.responseText, u8 = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i) & 0xff;
          processSync(f.split('/').pop(), u8.buffer);
        } else {
          showProcessing(f);
          fetch(f, { cache: 'no-cache' })
            .then((r) => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
            .then((b) => processBuffer(f.split('/').pop(), b))
            .catch((e) => { alert('Не удалось открыть файл: ' + e.message); showStart(); });
        }
      }
    } catch (e) { }
  }

  document.addEventListener('DOMContentLoaded', init);
})();

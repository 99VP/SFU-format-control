/* ==========================================================================
   СТУ-Нормоконтроль – правила проверки документа по СТУ 7.5–07–2021.
   Здесь описаны классификация фрагментов и все проверки с указанием пункта
   стандарта. Каждая находка: уровень, объяснение, рекомендация и ссылка.

   Намеренно НЕ проверяется содержание (актуальность, новизна и т.п.) –
   только структура и оформление, как требует нормоконтроль.
   ========================================================================== */
(function (global) {
  'use strict';

  const H = global.STUDocx.helpers;
  const toMM = H.mm;              // twips → мм
  const TOL = 30;                 // допуск в twips ≈ 0,53 мм

  const MAX_APPENDIX_LETTERS = 'АБВГДЕЖИКЛМНОПРСТУФХЦШЩЭЮЯ'.split(''); // без Ё З Й О Ч Ь Ы Ъ

  const STRUCTURAL = [
    { key: 'referat', names: ['РЕФЕРАТ'], label: 'Реферат' },
    { key: 'annotaciya', names: ['АННОТАЦИЯ'], label: 'Аннотация' },
    { key: 'soderzhanie', names: ['СОДЕРЖАНИЕ'], label: 'Содержание' },
    { key: 'vvedenie', names: ['ВВЕДЕНИЕ'], label: 'Введение' },
    { key: 'zaklyuchenie', names: ['ЗАКЛЮЧЕНИЕ'], label: 'Заключение' },
    { key: 'sokrashcheniya', names: ['СПИСОК СОКРАЩЕНИЙ', 'ПЕРЕЧЕНЬ СОКРАЩЕНИЙ', 'СПИСОК УСЛОВНЫХ ОБОЗНАЧЕНИЙ'], label: 'Список сокращений' },
    { key: 'istochniki', names: ['СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ И ЛИТЕРАТУРЫ', 'СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ', 'СПИСОК ЛИТЕРАТУРЫ', 'БИБЛИОГРАФИЧЕСКИЙ СПИСОК'], label: 'Список использованных источников' },
    { key: 'prilozheniya', names: ['ПРИЛОЖЕНИЯ', 'ПРИЛОЖЕНИЕ'], label: 'Приложения' },
  ];
  const STRUCT_ORDER = STRUCTURAL.map((s) => s.key);

  // Требования по видам документов: required – ошибка при отсутствии,
  // recommended – информационно.
  const DOC_REQUIREMENTS = {
    vkr: { label: 'ВКР (бакалаврская работа)', required: ['referat'], expected: ['vvedenie', 'zaklyuchenie', 'istochniki'], recommended: ['soderzhanie'], needMain: true },
    diplom: { label: 'Дипломная работа / дипломный проект', required: [], expected: ['vvedenie', 'zaklyuchenie', 'istochniki'], recommended: ['soderzhanie'], needMain: true },
    magister: { label: 'Магистерская диссертация', required: ['referat'], expected: ['vvedenie', 'zaklyuchenie', 'istochniki'], recommended: ['soderzhanie'], needMain: true },
    kursovaya: { label: 'Курсовая работа (проект)', required: [], expected: [], recommended: ['soderzhanie', 'vvedenie', 'zaklyuchenie', 'istochniki'], needMain: true },
    referat: { label: 'Реферат', required: [], expected: [], recommended: [], needMain: true },
    esse: { label: 'Эссе', required: [], expected: [], recommended: [], needMain: true },
    kontrolnaya: { label: 'Контрольная работа', required: [], expected: [], recommended: [], needMain: true },
    rgz: { label: 'РГЗ / РГР / расчётное задание', required: [], expected: [], recommended: [], needMain: true },
    praktika: { label: 'Отчёт о практике', required: [], expected: [], recommended: ['soderzhanie', 'vvedenie', 'zaklyuchenie', 'istochniki'], needMain: true },
    praktika_letnyaya: { label: 'Отчёт о летней практике', required: [], expected: [], recommended: ['soderzhanie', 'vvedenie', 'zaklyuchenie', 'istochniki'], needMain: true },
    otchet: { label: 'Отчёт (практика, лабораторная)', required: [], expected: [], recommended: ['soderzhanie'], needMain: true },
  };

  const ADJ = '[«"„\u2018\u2019]';  // открывающие кавычки
  // ------------------------------------------------------------------ helpers
  function isTOCLine(text) {
    return /\u2026{2,}/.test(text) || /\.{4,}\s*\d+\s*$/.test(text) || /\t\d+\s*$/.test(text) && /\.{3,}/.test(text);
  }
  function normText(s) { return (s || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim(); }
  function firstRunText(p) {
    for (const r of p.runs || []) if (normText(r.text)) return r.text;
    return '';
  }
  function isBoldPara(p) {
    const runs = (p.runs || []).filter((r) => normText(r.text));
    if (!runs.length) return p.markRPr && p.markRPr.bold;
    return runs.some((r) => r.bold);
  }
  function allBoldPara(p) {
    const runs = (p.runs || []).filter((r) => normText(r.text));
    if (!runs.length) return false;
    return runs.every((r) => r.bold);
  }
  function isUnderlined(p) { return (p.runs || []).some((r) => r.underline && normText(r.text)); }
  function fontKey(f) { return (f || '').replace(/["']/g, '').trim().toLowerCase(); }
  function isTimesNewRoman(f) { return /times new roman/i.test(f || ''); }
  function plural(n, forms) {
    const m10 = n % 10, m100 = n % 100;
    if (m10 === 1 && m100 !== 11) return forms[0];
    if (m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return forms[1];
    return forms[2];
  }
  function lvlFromNumber(num) { return num.split('.').length; } // "1.2.3" → 3
  // Читаемое имя маркера перечисления (символы Symbol показываем словами).
  function markerName(t) {
    if (!t) return 'без символа';
    const m = t.replace(/[\uF0A7\uF0B7]/g, '\u2022');
    if (/\u2022/.test(m)) return 'кружок';
    if (/\u25CF/.test(m)) return 'жирный круг';
    if (/\u25AA/.test(m)) return 'квадрат';
    if (/\u25E6/.test(m)) return 'маленький кружок';
    if (/[*]/.test(m)) return 'звёздочка';
    if (/\u00b7/.test(m)) return 'средняя точка';
    if (/[\d]/.test(m)) return 'цифра';
    if (/[–—\-]/.test(m)) return 'дефис';
    if (/[A-Za-zА-Яа-я]/.test(m)) return 'буква «' + m + '»';
    return 'символ «' + t + '»';
  }

  // Система координат страниц (без пагинации движком браузера):
  function computeForcedPageStarts(doc) {
    const starts = new Set([0]);
    // секции
    for (let s = 1; s < doc.sections.length; s++) {
      const sect = doc.sections[s];
      if (sect.type !== 'continuous') starts.add(sect.startIndex);
    }
    // явные разрывы и pageBreakBefore
    let prevTrailingBreak = false;
    doc.blocks.forEach((b, i) => {
      if (i > 0) {
        if (b.pageBreakBefore || prevTrailingBreak) starts.add(i);
      }
      prevTrailingBreak = false;
      if (b.kind === 'p' && b.runs && b.runs.length) {
        const all = [];
        for (const r of b.runs) for (const p of (r.parts || [])) all.push(p);
        while (all.length && all[all.length - 1].type === 'text' && all[all.length - 1].text === '') all.pop();
        if (all.length && all[all.length - 1].type === 'pagebreak') prevTrailingBreak = true;
      }
    });
    return starts;
  }

  // ------------------------------------------------------------------ classify
  function computeTitleEnd(doc, forcedPageStarts) {
    const blocks = doc.blocks || [];
    // Титульный лист это одна страница: ориентируемся на первый разрыв страницы.
    const starts = [...forcedPageStarts].filter((i) => i > 0).sort((a, b) => a - b);
    let end = starts.length ? starts[0] : null;
    // Первый структурный заголовок основной части означает конец титула,
    // но только если он не на самом титульном листе (там бывают строки «РЕФЕРАТ» и т.п.).
    let structStart = null;
    for (let i = 0; i < blocks.length; i++) {
      const t = (blocks[i].text || '').slice(0, 80).trim().toUpperCase().replace(/[.:]+$/, '');
      if (/^(РЕФЕРАТ|АННОТАЦИЯ|СОДЕРЖАНИЕ|ВВЕДЕНИЕ)$/.test(t)) { structStart = i; break; }
    }
    if (structStart != null && structStart >= 8) end = (end == null ? structStart : Math.min(end, structStart));
    if (end == null) {
      // разрыва нет: определяем титул по содержимому первых строк
      const MARK = /СИБИРСКИЙ ФЕДЕРАЛЬНЫЙ|МИНИСТЕРСТВ|ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ|ИНСТИТУТ|КАФЕДРА|УТВЕРЖДАЮ|КРАСНОЯРСК|РУКОВОДИТЕЛЬ|СТУДЕНТ|ПОДПИСЬ|ГРУПП|20\d\d/;
      let last = -1;
      for (let i = 0; i < Math.min(blocks.length, 40); i++) if (MARK.test((blocks[i].text || '').toUpperCase())) last = i;
      if (last >= 0) end = last + 1;
    }
    if (end == null) {
      // разрыва нет: ищем первый структурный заголовок основной части
      for (let i = 0; i < blocks.length; i++) {
        const t = (blocks[i].text || '').slice(0, 80).trim().toUpperCase().replace(/[.:]+$/, '');
        if (/^(РЕФЕРАТ|АННОТАЦИЯ|СОДЕРЖАНИЕ|ВВЕДЕНИЕ)$/.test(t)) { end = i; break; }
      }
      if (end == null) {
        for (let i = 0; i < blocks.length; i++) {
          const raw = blocks[i].text || ''; const t = normText(raw);
          if (/\t/.test(raw) || /\d\s*$/.test(t)) continue;
          if (/^\d+(?:\.\d+)*\s+\S/.test(t) && t.length < 200) { end = i; break; }
        }
      }
    }
    if (end == null || end > 80) end = 0;
    return Math.max(0, end);
  }

  function classify(doc, forcedPageStarts) {
    const blocks = doc.blocks;
    const roles = new Array(blocks.length);
    const numberedHeadings = [];
    const titleEnd = computeTitleEnd(doc, forcedPageStarts || new Set([0]));
    let inSources = false;
    // сначала разметим структурные заголовки и заголовки разделов
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (i < titleEnd) { roles[i] = { role: 'title' }; continue; }   // титульный лист не проверяем
      if (b.kind !== 'p') { roles[i] = { role: 'table' }; continue; }
      const text = b.text || '';
      const t = normText(text);
      const upper = t.toUpperCase();
      if (!t) { roles[i] = { role: 'empty' }; continue; }
      // структурный заголовок – высший приоритет
      let structKey = null;
      for (const s of STRUCTURAL) {
        if (s.names.includes(upper) || (s.key === 'prilozheniya' && /^ПРИЛОЖЕНИЕ\s+[А-ЯЁ]/.test(upper))) { structKey = s.key; break; }
      }
      if (structKey) { roles[i] = { role: 'structural', structKey }; inSources = (structKey === 'istochniki'); continue; }
      // строки оглавления: поле PAGEREF/_Toc, отточие, вкладка с номером или стиль TOC
      const fldText = b.fieldText || '';
      if (/PAGEREF|_Toc\d/i.test(fldText) || /PAGEREF|_Toc\d/i.test(text) || isTOCLine(text) ||
        (/\t/.test(text) && /\d\s*$/.test(t)) || /toc|оглавлен/i.test(b.styleName || '')) {
        roles[i] = { role: 'toc' }; continue;
      }
      // подпись к таблице/рисунку (может идти в одном абзаце с картинкой)
      const clean = normText(text.replace(/\uFFFC/g, ''));
      if (/^Таблица\s+[А-ЯA-Z0-9]/i.test(clean)) { roles[i] = { role: 'tableCaption' }; continue; }
      if (/^Рисунок\s+[А-ЯA-Z0-9]/i.test(clean)) { roles[i] = { role: 'figureCaption' }; continue; }
      // отдельный абзац только с изображением (картинка/формула-картинка)
      if (b.hasDrawing && clean === '') { roles[i] = { role: 'image' }; continue; }
      if (/^(Продолжение|Окончание)\s+(таблицы|приложения)/i.test(t)) { roles[i] = { role: 'continuation' }; continue; }
      // заголовок раздела/подраздела/пункта
      const m = t.match(/^(\d+(?:\.\d+){0,4})\.?\s+(.*)$/);
      const styleName = (b.styleName || '');
      const looksHeading = outlineHeading(b, styleName);
      if (m && (looksHeading || isBoldPara(b)) && t.length < 220) {
        const num = m[1];
        const rest = m[2];
        roles[i] = { role: 'heading', number: num, level: lvlFromNumber(num), title: rest };
        numberedHeadings.push({ index: i, number: num, level: lvlFromNumber(num), title: rest });
        continue;
      }
      if (looksHeading && !m) { roles[i] = { role: 'heading', number: null, level: 0, title: t }; continue; }
      // формула
      if (looksLikeFormula(t)) { roles[i] = { role: 'formula' }; continue; }
      if (/^где\s+/i.test(t)) { roles[i] = { role: 'where' }; continue; }
      // перечисление
      if (inSources && b.numPr && b.numPr.numId) { roles[i] = { role: 'body' }; continue; } // нумерованный список литературы
      if (b.numPr && b.numPr.numId) { roles[i] = { role: 'list' }; continue; }
      // в списке источников нумерация «1. …» – это не перечисление
      if (inSources && /^\s*\d{1,3}[.)]\s/.test(text)) { roles[i] = { role: 'body' }; continue; }
      if (/^\s*(?:[-–—•·◦○▪∙*]|\d{1,3}[.)]|[а-яёa-z][.)])\s+\S/.test(text)) { roles[i] = { role: 'manualList' }; continue; }
      roles[i] = { role: 'body' };
    }
    // Оглавление: всё между «СОДЕРЖАНИЕ» и следующим структурным элементом – это TOC.
    let soder = -1;
    for (let i = 0; i < blocks.length; i++) if (roles[i] && roles[i].role === 'structural' && roles[i].structKey === 'soderzhanie') { soder = i; break; }
    if (soder >= 0) {
      for (let i = soder + 1; i < blocks.length; i++) {
        const r = roles[i];
        if (r && r.role === 'structural') break;
        if (r && (r.role === 'body' || r.role === 'heading' || r.role === 'list' || r.role === 'manualList')) roles[i] = { role: 'toc' };
      }
    }
    return { roles, numberedHeadings, titleEnd };
  }
  function outlineHeading(b, styleName) {
    if (b.outlineLvl != null && b.outlineLvl >= 0) return true;
    if (/heading|заголовок/i.test(styleName || '')) return true;
    return false;
  }
  function looksLikeFormula(t) {
    if (!t || t.length > 130) return false;
    if (!/[=]/.test(t)) return false;
    if (/[А-Яа-я]{4,}/.test(t)) return false; // длинные слова – скорее текст
    return /[A-Za-zА-Яа-я0-9][^=]*=[^=]+/.test(t);
  }

  // ------------------------------------------------------------------ collect
  function collectText(blocks) { return blocks.map((b) => b.text || '').join('\n'); }

  // ------------------------------------------------------------------ checks
  function runChecks(doc, docType) {
    const findings = [];
    const add = (f) => findings.push(Object.assign({ id: 'f' + findings.length }, f));
    const blocks = doc.blocks;
    const forcedPageStarts = computeForcedPageStarts(doc);
    const { roles, titleEnd } = classify(doc, forcedPageStarts);
    const req = DOC_REQUIREMENTS[docType] || DOC_REQUIREMENTS.vkr;

    // ============================================================= Лист и поля
    doc.sections.forEach((sect, si) => {
      const portrait = (sect.w || 11906) <= (sect.h || 16838);
      const ref = { top: 1134, bottom: 1134, left: portrait ? 1701 : 1134, right: portrait ? 567 : 1134 };
      const target = portrait
        ? { left: 1701, right: 567, top: 1134, bottom: 1134, name: 'вертикальной' }
        : { left: 1134, right: 1134, top: 1701, bottom: 567, name: 'горизонтальной' };
      const wMM = Math.round(toMM(sect.w)), hMM = Math.round(toMM(sect.h));
      if (Math.abs(wMM - 210) > 2 || Math.abs(hMM - 297) > 2) {
        add({
          category: 'page', level: 'error', blockIndex: sect.startIndex, pageLevel: true,
          title: 'Формат листа не А4',
          detail: `Размер страницы ${wMM}×${hMM} мм.`,
          recommend: 'Установите формат листа А4 (210×297 мм).',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      }
      const m = sect.margin || {};
      const checks = [['left', 'левое', 30], ['right', 'правое', 10], ['top', 'верхнее', 20], ['bottom', 'нижнее', 20]];
      const targetPortrait = { left: 1701, right: 567, top: 1134, bottom: 1134 };
      const targetLandscape = { left: 1134, right: 1134, top: 1701, bottom: 567 };
      const tgt = portrait ? targetPortrait : targetLandscape;
      const bad = [];
      for (const [k, name] of [['left', 'левое'], ['right', 'правое'], ['top', 'верхнее'], ['bottom', 'нижнее']]) {
        const val = m[k];
        if (val == null) continue;
        if (Math.abs(val - tgt[k]) > TOL) {
          bad.push(`${name} ${toMM(val).toFixed(1).replace('.', ',')} мм (требуется ${toMM(tgt[k]).toFixed(0)} мм)`);
        }
      }
      if (bad.length) {
        add({
          category: 'page', level: 'error', blockIndex: sect.startIndex, pageLevel: true,
          title: 'Неверные размеры полей',
          detail: `При ${target.name} ориентации: ${bad.join('; ')}.`,
          recommend: portrait
            ? 'Задайте поля: левое 30 мм, верхнее и нижнее 20 мм, правое 10 мм.'
            : 'Задайте поля: верхнее 30 мм, нижнее 10 мм, левое и правое 20 мм.',
          ref: 'СТУ 7.5–07–2021, п. 7.1.2',
        });
      }
    });

    // ============================================================ Шрифт и абзац
    const indentValues = [];
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind !== 'p') continue;
      const r = roles[i];
      if (r.role !== 'body' && r.role !== 'where') continue;
      if (b.isEmpty) continue;
      // абзацный отступ
      const fl = b.ind ? b.ind.firstLine : null;
      const hang = b.ind ? b.ind.hanging : null;
      const effectiveFirst = fl != null ? fl : (hang != null ? -hang : null);
      const lineRule = b.spacing ? b.spacing.lineRule : null;
      const line = b.spacing ? b.spacing.line : null;
      const runs = (b.runs || []).filter((x) => normText(x.text));
      if (runs.length) {
        const wrongFont = runs.filter((x) => !isTimesNewRoman(x.font));
        if (wrongFont.length) {
          const fonts = [...new Set(wrongFont.map((x) => x.font))].join(', ');
          add({
            category: 'font', level: 'error', blockIndex: i,
            runIndices: runs.map((x) => b.runs.indexOf(x)).filter((idx) => !isTimesNewRoman(b.runs[idx].font)),
            title: 'Неверный шрифт',
            detail: `Основной текст набран шрифтом ${fonts || 'без явного шрифта'}, требуется Times New Roman.`,
            recommend: 'Установите для текста шрифт Times New Roman.',
            ref: 'СТУ 7.5–07–2021, п. 7.1.1',
          });
        }
        const wrongSize = runs.filter((x) => Math.abs((x.size || 14) - 14) > 0.01);
        if (wrongSize.length) {
          const sizes = [...new Set(wrongSize.map((x) => (x.size || 14)))].map((s) => String(s).replace('.', ',')).join(', ');
          add({
            category: 'font', level: 'error', blockIndex: i,
            runIndices: runs.map((x) => b.runs.indexOf(x)).filter((idx) => Math.abs((b.runs[idx].size || 14) - 14) > 0.01),
            title: 'Неверный размер шрифта',
            detail: `Размер текста ${sizes} пт, требуется 14 пт.`,
            recommend: 'Установите для основного текста размер 14 пт.',
            ref: 'СТУ 7.5–07–2021, п. 7.1.1',
          });
        }
      }
      if (b.jc !== 'both' && b.jc !== 'justify' && b.jc !== 'distribute') {
        add({
          category: 'font', level: 'error', blockIndex: i,
          title: 'Неверное выравнивание',
          detail: `Абзац выровнен не по ширине (${jcName(b.jc)}).`,
          recommend: 'Установите выравнивание «по ширине».',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      }
      // абзацный отступ 12,5 мм (учитываем и отступ всего абзаца слева)
      const leftInd = b.ind ? (b.ind.left || 0) : 0;
      if (effectiveFirst == null || effectiveFirst === 0) {
        if (leftInd < 300) add({
          category: 'font', level: 'error', blockIndex: i,
          title: 'Нет абзацного отступа',
          detail: 'Абзац без отступа первой строки, требуется 12,5 мм (5 знаков).',
          recommend: 'Задайте абзацный отступ первой строки 12,5 мм.',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      } else if (Math.abs(toMM(effectiveFirst) - 12.5) > 0.6) {
        add({
          category: 'font', level: 'error', blockIndex: i,
          title: 'Неверный абзацный отступ',
          detail: `Отступ первой строки ${toMM(effectiveFirst).toFixed(1).replace('.', ',')} мм, требуется 12,5 мм (5 знаков).`,
          recommend: 'Задайте абзацный отступ первой строки 12,5 мм.',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      } else {
        indentValues.push(effectiveFirst);
      }
      // межстрочный интервал
      let factor = null;
      if (line == null) factor = 1;
      else if (lineRule === 'auto' || lineRule == null) factor = line / 240;
      else factor = null; // exact/atLeast
      if (factor == null) {
        add({
          category: 'font', level: 'warning', blockIndex: i,
          title: 'Межстрочный интервал задан фиксированно',
          detail: 'Интервал задан как «точно»/«минимум».',
          recommend: 'Задайте одинарный или полуторный межстрочный интервал.',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      } else if (Math.abs(factor - 1) > 0.02 && Math.abs(factor - 1.5) > 0.02) {
        add({
          category: 'font', level: 'error', blockIndex: i,
          title: 'Неверный межстрочный интервал',
          detail: `Интервал ${String(Math.round(factor * 100) / 100).replace('.', ',')}, требуется одинарный или полуторный.`,
          recommend: 'Задайте одинарный или полуторный межстрочный интервал.',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      }
    }
    // одинаковость абзацного отступа
    if (indentValues.length) {
      const uniq = [...new Set(indentValues.map((v) => Math.round(v)))];
      if (uniq.length > 1) {
        add({
          category: 'font', level: 'warning', blockIndex: null,
          title: 'Абзацный отступ не одинаков',
          detail: `В документе встречаются разные отступы первой строки: ${uniq.map((v) => toMM(v).toFixed(1).replace('.', ',')).join(', ')} мм.`,
          recommend: 'Приведите абзацный отступ к одному значению – 12,5 мм по всему тексту.',
          ref: 'СТУ 7.5–07–2021, п. 7.1.1',
        });
      }
    }

    checkEmphasis(blocks, roles, add);

    // =============================================================== Структура
    const foundStruct = [];
    for (let i = 0; i < blocks.length; i++) {
      const r = roles[i];
      if (r.role === 'structural') foundStruct.push({ key: r.structKey, index: i });
    }
    const presentKeys = new Set(foundStruct.map((f) => f.key));
    // «Основная часть» – наличие нумерованных разделов
    const hasMain = blocks.some((b, i) => roles[i].role === 'heading' && roles[i].level === 1);
    if (req.needMain && !hasMain) {
      add({
        category: 'structure', level: 'error', blockIndex: null,
        title: 'Не найдена основная часть',
        detail: 'В документе нет нумерованных разделов основной части.',
        recommend: 'Оформите основную часть с нумерованными разделами (1, 2, 3 …).',
        ref: 'СТУ 7.5–07–2021, п. 6.1, 6.6',
      });
    }
    for (const key of req.required) {
      if (!presentKeys.has(key)) {
        const info = STRUCTURAL.find((s) => s.key === key);
        add({
          category: 'structure', level: 'error', blockIndex: null,
          title: `Отсутствует элемент «${info.label}»`,
          detail: `Для документа «${req.label}» обязателен структурный элемент «${info.label}».`,
          recommend: `Добавьте структурный элемент «${info.label.toUpperCase()}».`,
          ref: 'СТУ 7.5–07–2021, п. 6.1, 6.3',
        });
      }
    }
    for (const key of (req.expected || [])) {
      if (!presentKeys.has(key)) {
        const info = STRUCTURAL.find((s) => s.key === key);
        add({
          category: 'structure', level: 'warning', blockIndex: null,
          title: `Нет элемента «${info.label}»`,
          detail: `Для документа «${req.label}» обычно требуется структурный элемент «${info.label}» (есть в подавляющем большинстве работ).`,
          recommend: `Добавьте структурный элемент «${info.label.toUpperCase()}».`,
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
      }
    }
    for (const key of req.recommended) {
      if (!presentKeys.has(key)) {
        const info = STRUCTURAL.find((s) => s.key === key);
        // «Содержание» рекомендуется только при объёме > 24 стр. – оставим информацией
        add({
          category: 'structure', level: 'info', blockIndex: null,
          title: `Нет элемента «${info.label}»`,
          detail: `Структурный элемент «${info.label}» отсутствует.`,
          recommend: `При необходимости добавьте «${info.label.toUpperCase()}».`,
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
      }
    }
    // порядок структурных элементов
    let prevOrder = -1;
    for (const f of foundStruct) {
      const ord = STRUCT_ORDER.indexOf(f.key);
      if (ord < prevOrder) {
        add({
          category: 'structure', level: 'warning', blockIndex: f.index,
          title: 'Нарушен порядок структурных элементов',
          detail: `Элемент «${labelOf(f.key)}» расположен не в порядке, установленном стандартом.`,
          recommend: 'Расположите элементы в порядке: титульный лист, реферат, содержание, введение, основная часть, заключение, список сокращений, список источников, приложения.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
      }
      prevOrder = Math.max(prevOrder, ord);
    }
    // каждый структурный элемент – с новой страницы
    for (const f of foundStruct) {
      if (f.index === 0) continue;
      if (!forcedPageStarts.has(f.index)) {
        add({
          category: 'structure', level: 'warning', blockIndex: f.index,
          title: 'Структурный элемент без новой страницы',
          detail: `Элемент «${labelOf(f.key)}» не начинается с новой страницы.`,
          recommend: `Начните «${labelOf(f.key).toUpperCase()}» с новой страницы.`,
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
      }
    }
    // каждый раздел (уровень 1) – с новой страницы
    for (let i = 1; i < blocks.length; i++) {
      const r = roles[i];
      if (r && r.role === 'heading' && r.level === 1 && !forcedPageStarts.has(i)) {
        add({
          category: 'structure', level: 'warning', blockIndex: i,
          title: 'Раздел без новой страницы',
          detail: `Заголовок «${r.number ? r.number + ' ' : ''}${r.title || ''}» не начинается с новой страницы.`,
          recommend: 'Начинайте каждый раздел с новой страницы.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
      }
    }

    // =============================================================== Заголовки
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const r = roles[i];
      if (r.role === 'structural') {
        const t = normText(b.text);
        if (t !== t.toUpperCase()) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Заголовок не прописными буквами',
          detail: `Заголовок «${t}» оформлен не полностью прописными буквами.`,
          recommend: 'Печатайте заголовки структурных элементов прописными буквами.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
        if (b.jc !== 'center') add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Заголовок не по центру',
          detail: `Заголовок «${t}» выровнен не по центру (${jcName(b.jc)}).`,
          recommend: 'Расположите заголовок посередине строки.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
        const fl = b.ind ? b.ind.firstLine : null;
        const hang = b.ind ? b.ind.hanging : null;
        const eff = fl != null ? fl : (hang != null ? -hang : null);
        if (eff) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'У заголовка есть абзацный отступ',
          detail: `Заголовок «${t}» имеет абзацный отступ ${toMM(eff).toFixed(1).replace('.', ',')} мм.`,
          recommend: 'Уберите абзацный отступ у заголовка.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
        if (!isBoldPara(b)) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Заголовок не полужирный',
          detail: `Заголовок «${t}» не выделен полужирным шрифтом.`,
          recommend: 'Выделите заголовок полужирным шрифтом.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
        if (isUnderlined(b)) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Заголовок подчёркнут',
          detail: `Заголовок «${t}» подчёркнут.`,
          recommend: 'Уберите подчёркивание заголовка.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
        continue;
      }
      if (r.role === 'heading') {
        const t = normText(b.text);
        if (/\.$/.test(t)) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Точка в конце заголовка',
          detail: `Заголовок «${t}» заканчивается точкой.`,
          recommend: 'Уберите точку в конце заголовка.',
          ref: 'СТУ 7.5–07–2021, п. 7.5.3',
        });
        if (!isBoldPara(b)) add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Заголовок раздела не полужирный',
          detail: `Заголовок «${t}» не выделен полужирным.`,
          recommend: 'Выделите заголовок раздела полужирным шрифтом.',
          ref: 'СТУ 7.5–07–2021, п. 7.5.3',
        });
        if (isUnderlined(b)) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Заголовок подчёркнут',
          detail: `Заголовок «${t}» подчёркнут.`,
          recommend: 'Уберите подчёркивание заголовка.',
          ref: 'СТУ 7.5–07–2021, п. 7.5.3',
        });
        const fl = b.ind ? b.ind.firstLine : null;
        const hang = b.ind ? b.ind.hanging : null;
        const eff = fl != null ? fl : (hang != null ? -hang : null);
        if (!eff) add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Заголовок без абзацного отступа',
          detail: `Заголовок «${t}» напечатан без абзацного отступа.`,
          recommend: 'Печатайте заголовки разделов и подразделов с абзацного отступа (12,5 мм).',
          ref: 'СТУ 7.5–07–2021, п. 7.4.2',
        });
      }
    }
    // Стиль и уровень заголовков, отделение от текста, последовательность нумерации
    checkHeadingStructure(blocks, roles, add);
    checkHeadingSpacing(blocks, roles, add);
    checkPunctuation(blocks, roles, add);
    checkCitations(blocks, roles, add);
    checkNumberingSequence(roles, blocks, add);

    // Структурные заголовки печатают прописными буквами (СТУ, п. 6.1).
    for (let i = 0; i < blocks.length; i++) {
      const r = roles[i];
      if (!r || r.role !== 'structural') continue;
      if (blocks[i].kind !== 'p') continue;
      const t = normText(blocks[i].text);
      if (!t) continue;
      if (/[а-яё]/.test(t)) add({
        category: 'heading', level: 'warning', blockIndex: i,
        title: 'Заголовок не прописными буквами',
        detail: `«${t}» набран строчными буквами.`,
        recommend: 'Структурные заголовки (СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и другие) печатают прописными буквами полужирным шрифтом по центру.',
        ref: 'СТУ 7.5–07–2021, п. 6.1',
      });
    }

    // ======================================================= Нумерация страниц
    checkPageNumbers(doc, add);
    checkHeaders(doc, add);

    // ============================================================ Перечисления
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const r = roles[i];
      if (r.role === 'list' && doc.numbering) {
        const lvl = doc.numbering.level(b.numPr.numId, b.numPr.ilvl);
        if (lvl) {
          const isTop = (b.numPr.ilvl || '0') === '0';
          if (isTop && lvl.fmt !== 'bullet') {
            add({
              category: 'list', level: 'warning', blockIndex: i,
              title: 'Перечисление не дефисом',
              detail: `Позиция перечисления обозначена не дефисом (формат «${lvl.fmt}»).`,
              recommend: 'Перед каждой позицией перечисления ставьте дефис. Буквы со скобкой допустимы только при ссылках в тексте.',
              ref: 'СТУ 7.5–07–2021, п. 7.4.3',
            });
          } else if (lvl.fmt === 'bullet' && lvl.lvlText && /[•◦o\u25aa\u25cf\u25e6\u00b7\uf0a7\uf0b7]/.test(lvl.lvlText)) {
            add({
              category: 'list', level: 'warning', blockIndex: i,
              title: 'Маркер перечисления не дефис',
              detail: `Маркер перечисления: ${markerName(lvl.lvlText)}.`,
              recommend: 'Перед каждой позицией перечисления ставьте дефис (тире) «–».',
              ref: 'СТУ 7.5–07–2021, п. 7.4.3',
            });
          }
        }
      }
      if (r.role === 'manualList') {
        const raw = b.text || '';
        const m = raw.match(/^(\s*)([-–—•·◦○▪∙*]|\d{1,3}[.)]|[а-яёa-zA-Z][.)])(\s+)/);
        if (m) {
          const marker = m[2];
          if (marker !== '-' && marker !== '–') {
            const start = m[1].length;
            const kind = /^\d/.test(marker) ? 'номер' : (/^[а-яёa-z]/i.test(marker) && /[.)]$/.test(marker) ? 'буква' : 'знак');
            add({
              category: 'list', level: 'warning', blockIndex: i,
              chars: rangeAt(b, start, start + marker.length),
              title: 'Маркер перечисления не по стандарту',
              detail: `Позиция перечисления начинается с «${marker}» (${kind}).`,
              recommend: 'Перед каждой позицией перечисления ставьте дефис (тире) «–». Букву со скобкой или цифру со скобкой используют только при ссылках в тексте.',
              ref: 'СТУ 7.5–07–2021, п. 7.4.3',
            });
          }
        }
      }
      // Позиция перечисления продолжает фразу и пишется со строчной буквы
      // (кроме имён собственных и аббревиатур).
      if (r.role === 'list' || r.role === 'manualList') {
        const t = normText(b.text);
        const noM = t.replace(/^\s*(?:[-–—•·◦○▪∙*]|\d{1,3}[.)]|[а-яёa-z][.)])\s+/, '');
        const w = noM.match(/^([А-ЯЁ][а-яё\-]+)/);
        if (w) {
          const idx = (b.text || '').indexOf(w[1]);
          add({
            category: 'list', level: 'warning', blockIndex: i,
            chars: idx >= 0 ? rangeAt(b, idx, idx + w[1].length) : undefined,
            title: 'Позиция перечисления с прописной буквы',
            detail: `Позиция начинается со слова «${w[1]}».`,
            recommend: 'Позиции перечисления продолжают фразу и пишутся со строчной буквы (кроме имён собственных и аббревиатур).',
            ref: 'СТУ 7.5–07–2021, п. 7.4.3',
          });
        }
        // Знак в конце позиции: не последняя – «;», последняя – «.».
        const nextR = roles[i + 1];
        const isLast = !(nextR && (nextR.role === 'list' || nextR.role === 'manualList'));
        const want = isLast ? '.' : ';';
        const lastCh = t.slice(-1);
        if (t && lastCh !== want && !/[!?]$/.test(t)) {
          add({
            category: 'list', level: 'warning', blockIndex: i,
            title: 'Знак в конце позиции перечисления',
            detail: `Позиция заканчивается на «${lastCh || 'пробел'}», а ${isLast ? 'последняя позиция' : 'не последняя позиция'} должна заканчиваться на «${want}».`,
            recommend: isLast ? 'Последнюю позицию перечисления заканчивайте точкой.' : 'Не последнюю позицию перечисления заканчивайте точкой с запятой.',
            ref: 'СТУ 7.5–07–2021, п. 7.4.3',
          });
        }
      }
    }

    // ================================================================ Формулы
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      if (roles[i].role === 'formula') {
        if (!endsWithNumber(b.text)) {
          add({
            category: 'formula', level: 'info', blockIndex: i,
            title: 'У формулы, возможно, нет номера',
            detail: 'Формула не заканчивается номером в круглых скобках.',
            recommend: 'Номер формулы указывайте в круглых скобках в крайнем правом положении на строке.',
            ref: 'СТУ 7.5–07–2021, п. 7.6.2',
          });
        }
        const before = blocks[i - 1], after = blocks[i + 1];
        if (before && before.kind === 'p' && !before.isEmpty && !isBlankish(before)) {
          add({ category: 'formula', level: 'warning', blockIndex: i, title: 'Формула не отделена свободной строкой сверху', detail: 'Перед формулой нет свободной строки.', recommend: 'Оставьте одну свободную строку выше формулы.', ref: 'СТУ 7.5–07–2021, п. 7.6.1' });
        }
        if (after && after.kind === 'p' && !after.isEmpty && roles[i + 1] && roles[i + 1].role !== 'where') {
          add({ category: 'formula', level: 'warning', blockIndex: i, title: 'Формула не отделена свободной строкой снизу', detail: 'Под формулой нет свободной строки.', recommend: 'Оставьте одну свободную строку ниже формулы.', ref: 'СТУ 7.5–07–2021, п. 7.6.1' });
        }
      }
      if (roles[i].role === 'where') {
        const t = normText(b.text);
        if (/^где\s*:/.test(t)) add({
          category: 'formula', level: 'error', blockIndex: i,
          title: 'Двоеточие после «где»',
          detail: 'Пояснения символов начинаются с «где:».',
          recommend: 'После слова «где» двоеточие не ставят.',
          ref: 'СТУ 7.5–07–2021, п. 7.6.5',
        });
        const fl = b.ind ? b.ind.firstLine : null;
        if (fl) add({
          category: 'formula', level: 'warning', blockIndex: i,
          title: 'Пояснение символов с абзацным отступом',
          detail: 'Строка пояснений начинается с абзацного отступа.',
          recommend: 'Первую строку пояснения начинайте со слова «где» без абзацного отступа.',
          ref: 'СТУ 7.5–07–2021, п. 7.6.5',
        });
      }
    }

    // ================================================================ Таблицы
    checkTables(doc, blocks, roles, add);
    checkBlankLines(blocks, roles, add);

    // =========================================================== Иллюстрации
    checkFigures(blocks, roles, add);

    // ======================================== Ссылки, источники, приложения
    checkTypography(blocks, roles, add);
    checkAiMarkers(blocks, roles, add);
    checkConsistency(blocks, roles, add);
    checkReferencesAndSources(doc, blocks, roles, presentKeys, add);
    checkAbbreviations(blocks, roles, add);
    checkAbbrevUsage(blocks, roles, add);
    checkFormulas(blocks, roles, add);
    checkAppendices(blocks, roles, forcedPageStarts, add);
    const tocEntries = parseToc(blocks, roles, add);

    // ==================================== Рамка и основная надпись (ДП/КП)
    if (docType === 'diplom' || docType === 'kursovaya') {
      add({
        category: 'page', level: 'warning', blockIndex: null, pageLevel: true,
        title: 'Рамка и основная надпись (ДП / КП)',
        detail: 'Текст пояснительной записки дипломного/курсового проекта печатают на листах с рамкой и основной надписью (формы 1–4 приложения С).',
        recommend: 'Оформите листы с рамкой и основной надписью; расстояние от рамки до границ текста – 5 мм, от строки до рамки – 15 мм.',
        ref: 'СТУ 7.5–07–2021, п. 7.1.3',
      });
    }

    return { findings, roles, forcedPageStarts, titleEnd, tocEntries };
  }

  function jcName(jc) {
    switch (jc) {
      case 'both': case 'justify': return 'по ширине';
      case 'center': return 'по центру';
      case 'right': case 'end': return 'по правому краю';
      case 'left': case 'start': return 'по левому краю';
      case null: case undefined: return 'по левому краю (по умолчанию)';
      default: return jc;
    }
  }
  function labelOf(key) { const s = STRUCTURAL.find((x) => x.key === key); return s ? s.label : key; }
  function endsWithNumber(t) { return /\(\s*[\dА-ЯA-Z.]+\s*\)\s*$/.test(normText(t)); }
  function isBlankish(p) { return !normText(p.text); }

  // --------------------------------------------------------------- numbering seq
  function actualHeadingLevel(b) {
    if (b.outlineLvl != null && b.outlineLvl >= 0) return b.outlineLvl + 1;
    const m = (b.styleName || '').match(/(?:heading|заголовок)\s*([1-9])/i);
    if (m) return +m[1];
    return null;
  }

  // Стиль и уровень заголовков. 1 уровень – разделы и структурные элементы,
  // 2 – подразделы, 3 – пункты и т. д.
  function checkHeadingStructure(blocks, roles, add) {
    const levelStyles = {};
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const r = roles[i];
      const t = normText(b.text);
      if (r.role === 'structural') {
        if (/^\d/.test(t)) add({
          category: 'heading', level: 'error', blockIndex: i,
          title: 'Структурный элемент пронумерован',
          detail: `Заголовок «${t}» начинается с номера.`,
          recommend: 'Заголовки структурных элементов не нумеруют.',
          ref: 'СТУ 7.5–07–2021, п. 6.1',
        });
        const lvl = actualHeadingLevel(b);
        if (lvl !== 1) add({
          category: 'heading', level: 'info', blockIndex: i,
          title: 'Уровень заголовка структурного элемента',
          detail: `«${t}» имеет уровень ${lvl == null ? 'не задан' : lvl}. В стандарте уровень заголовков не оговорён; для единообразия обычно 1-й уровень.`,
          recommend: 'При желании назначьте уровень 1 (стиль «Заголовок 1»).',
          ref: '',
        }); else (levelStyles[1] = levelStyles[1] || []).push(b.styleName || '');
        continue;
      }
      if (r.role === 'heading') {
        const expected = r.level || 1;
        const lvl = actualHeadingLevel(b);
        if (r.title && /^[а-яёa-z]/.test(r.title)) add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Заголовок со строчной буквы',
          detail: `Заголовок «${t}» начинается со строчной буквы.`,
          recommend: 'Заголовок печатайте с прописной буквы.',
          ref: 'СТУ 7.5–07–2021, п. 7.5.3',
        });
        if (lvl == null) add({
          category: 'heading', level: 'info', blockIndex: i,
          title: 'Заголовок без стиля',
          detail: `Заголовок «${t}» не оформлен стилем заголовка.`,
          recommend: `Назначьте заголовку уровень ${expected} (стиль «Заголовок ${expected}»).`,
          ref: 'СТУ 7.5–07–2021, п. 7.4.2, 7.5',
        });
        else if (lvl !== expected) add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Неверный уровень заголовка',
          detail: `Заголовок «${t}» имеет уровень ${lvl}, а по нумерации должен быть уровня ${expected}.`,
          recommend: `Установите заголовку уровень ${expected} (стиль «Заголовок ${expected}»).`,
          ref: 'СТУ 7.5–07–2021, п. 7.4.2, 7.5',
        });
        else (levelStyles[lvl] = levelStyles[lvl] || []).push(b.styleName || '');
        continue;
      }
      if (r.role === 'body') {
        const lvl = actualHeadingLevel(b);
        if (lvl != null) add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Обычный текст оформлен стилем заголовка',
          detail: `Текст «${t.slice(0, 50)}» имеет уровень ${lvl}.`,
          recommend: 'Для основного текста используйте стиль «Обычный», а не стиль заголовка.',
          ref: 'СТУ 7.5–07–2021, п. 7.5',
        });
      }
    }
  }

  // --------------------------------------------------------------- numbering seq
  // Заголовки отделяют от текста интервалом в одну строку (п. 7.5.3, 6.1).
  function checkHeadingSpacing(blocks, roles, add) {
    for (let i = 0; i < blocks.length - 1; i++) {
      const r = roles[i]; if (r.role !== 'heading' && r.role !== 'structural') continue;
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const next = blocks[i + 1];
      if (!next || next.kind !== 'p') continue;
      const nr = roles[i + 1];
      if (next.isEmpty) continue;
      if (nr.role === 'heading' || nr.role === 'structural' || nr.role === 'image' || nr.role === 'tableCaption' || nr.role === 'figureCaption') continue;
      const after = (b.spacing && b.spacing.after) || 0;
      const before = (next.spacing && next.spacing.before) || 0;
      if (Math.max(after, before) < 200) {
        add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Заголовок не отделён от текста',
          detail: `После заголовка «${normText(b.text)}» сразу идёт текст без пустой строки.`,
          recommend: 'Заголовки отделяют от текста интервалом в одну строку.',
          ref: 'СТУ 7.5–07–2021, п. 7.5.3',
        });
      }
    }
  }

  // Пунктуация: точка в конце абзаца и двоеточие перед перечислением.
  // Отсылка к источнику: точка должна стоять до квадратных скобок («текст.[7]»).
  function checkCitations(blocks, roles, add) {
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const r = roles[i]; if (!r) continue;
      if (r.role === 'title' || r.role === 'toc') continue;
      const text = b.text || '';
      const urls = urlSpans(text);
      const re = /\[(\d{1,4}(?:\s*[;,][^\]]*)?)\]\s*\./g;
      let m; const ranges = [];
      while ((m = re.exec(text))) {
        if (inSpans(m.index, urls)) continue;
        ranges.push([m.index, m.index + m[0].length]);
      }
      if (!ranges.length) continue;
      let chars = []; for (const [s, e] of ranges) chars = chars.concat(rangeAt(b, s, e));
      add({
        category: 'sources', level: 'warning', blockIndex: i, chars,
        title: 'Точка после отсылки к источнику',
        detail: 'В отсылке к источнику точка стоит после квадратных скобок.',
        recommend: 'Ставьте точку до квадратных скобок, без пробела: «текст.[7]».',
        ref: 'СТУ 7.5–07–2021, п. 7.9.4',
      });
    }
  }

  function checkPunctuation(blocks, roles, add) {
    const CITES = /\s*(?:\[\s*\d+(?:\s*[;,\u2013-]\s*\d+)*\s*\])+\s*$/;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const r = roles[i];
      if (r.role !== 'body' && r.role !== 'where') continue;
      const raw = normText(b.text); if (!raw) continue;
      if (/\uFFFC/.test(b.text || '')) continue;
      const next = blocks[i + 1];
      const nr = next ? roles[i + 1] : null;
      const nextIsList = next && next.kind === 'p' && (nr.role === 'list' || nr.role === 'manualList');
      const t = raw.replace(CITES, '').trim();   // убираем хвостовые ссылки [N]
      const last = t.slice(-1);
      const endsColon = last === ':';
      const endsPeriod = /[.!?\u2026]/.test(last);
      if (nextIsList) {
        if (!endsColon) add({
          category: 'punct', level: 'warning', blockIndex: i,
          title: 'Нет двоеточия перед перечислением',
          detail: `Перед перечислением абзац заканчивается «${last}».`,
          recommend: 'Перед перечислением ставят двоеточие.',
          ref: 'СТУ 7.5–07–2021, п. 7.4.3',
        });
      } else if (!endsPeriod) {
        add({
          category: 'punct', level: 'warning', blockIndex: i,
          title: 'В конце абзаца нет точки',
          detail: `Абзац заканчивается «…${t.slice(-24)}».`,
          recommend: 'Заканчивайте абзац точкой (или знаком !, ?).',
          ref: 'СТУ 7.5–07–2021, п. 7.3.1',
        });
      }
    }
  }

  function checkNumberingSequence(roles, blocks, add) {
    const exists = new Set();   // полные номера встреченных заголовков
    const childLast = {};       // родитель -> последний номер ребёнка
    for (let i = 0; i < blocks.length; i++) {
      const r = roles[i]; if (r.role !== 'heading' || !r.number) continue;
      const parts = r.number.split('.').map(Number);
      const parent = parts.slice(0, -1).join('.');
      let ok = true;
      for (let d = 0; d < parts.length - 1; d++) {
        const pref = parts.slice(0, d + 1).join('.');
        if (!exists.has(pref)) ok = false;
      }
      const expected = (childLast[parent] || 0) + 1;
      if (parts[parts.length - 1] !== expected) ok = false;
      if (!ok) {
        add({
          category: 'heading', level: 'warning', blockIndex: i,
          title: 'Нарушена нумерация разделов',
          detail: `Заголовок «${r.number} ${r.title}» нарушает последовательность нумерации.`,
          recommend: 'Измените нумерацию: разделы, подразделы и пункты нумеруйте арабскими цифрами последовательно (1, 2, 3; 1.1, 1.2; 1.1.1 …).',
          ref: 'СТУ 7.5–07–2021, п. 7.4.2',
        });
      }
      exists.add(r.number);
      childLast[parent] = parts[parts.length - 1];
    }
  }

  // --------------------------------------------------------------- page numbers
  function checkPageNumbers(doc, add) {
    // берём колонтитулы секции, где больше всего текста
    const withFooter = doc.sections.filter((s) => s.footers && (s.footers.default || s.footers.first));
    let hasField = false, centerOK = true, fontOK = true, firstHidden = false, found = false;
    for (const s of doc.sections) {
      const f = s.footers || {};
      const def = f.default, first = f.first;
      if (def && def.hasPageField) { hasField = true; found = true; for (const p of def.paras) if (p.hasPageField && p.jc !== 'center') centerOK = false; for (const p of def.paras) if (p.hasPageField && p.font && !isTimesNewRoman(p.font)) fontOK = false; for (const p of def.paras) if (p.hasPageField && p.size && Math.abs(p.size - 14) > 0.01) fontOK = false; }
      if (first && first.hasPageField) { hasField = true; found = true; }
      if (s.titlePg) firstHidden = true;
      if (!def && first && first.hasPageField) { }
    }
    if (!found) {
      add({
        category: 'pagenum', level: 'error', blockIndex: null, pageLevel: true,
        title: 'Не найдена нумерация страниц',
        detail: 'В колонтитулах нет поля номера страницы.',
        recommend: 'Проставьте номера страниц: внизу по центру, Times New Roman 14.',
        ref: 'СТУ 7.5–07–2021, п. 7.2.1',
      });
      return;
    }
    if (!centerOK) add({
      category: 'pagenum', level: 'error', blockIndex: null, pageLevel: true,
      title: 'Номер страницы не по центру',
      detail: 'Номер страницы в колонтитуле расположен не по центру.',
      recommend: 'Проставляйте номер страницы в центре нижней части листа.',
      ref: 'СТУ 7.5–07–2021, п. 7.2.1',
    });
    if (!fontOK) add({
      category: 'pagenum', level: 'warning', blockIndex: null, pageLevel: true,
      title: 'Оформление номера страницы',
      detail: 'Шрифт или размер номера страницы отличается от Times New Roman 14.',
      recommend: 'Оформляйте номер страницы шрифтом Times New Roman 14.',
      ref: 'СТУ 7.5–07–2021, п. 7.2.1',
    });
    if (!firstHidden) add({
      category: 'pagenum', level: 'warning', blockIndex: 0, pageLevel: true,
      title: 'Номер может стоять на титульном листе',
      detail: 'Не включён особый колонтитул первой страницы.',
      recommend: 'На титульном листе номер страницы не проставляют (но он входит в общую нумерацию).',
      ref: 'СТУ 7.5–07–2021, п. 7.2.2',
    });
  }

  // --------------------------------------------------------------- tables
  function checkTables(doc, blocks, roles, add) {
    const tableNums = [];
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind !== 'table') continue;
      if (roles[i] && roles[i].role === 'title') continue;
      let capIndex = -1;
      for (let j = i - 1; j >= 0 && j >= i - 3; j--) {
        if (blocks[j].kind === 'p' && roles[j].role === 'tableCaption') { capIndex = j; break; }
        if (blocks[j].kind === 'p' && !normText(blocks[j].text)) continue;
        break;
      }
      if (capIndex >= 0) {
        const cap = blocks[capIndex];
        const t = normText(cap.text);
        const m = t.match(/^Таблица\s+([А-ЯA-Z0-9.\-]+)\s*(–|-|—)\s*(.*)$/i);
        const fl = cap.ind ? cap.ind.firstLine : null;
        if (fl) add({ category: 'table', level: 'error', blockIndex: capIndex, title: 'Подпись таблицы с абзацным отступом', detail: `Подпись «${t}» начинается с абзацного отступа.`, recommend: 'Слово «Таблица» и подпись размещайте слева, без абзацного отступа.', ref: 'СТУ 7.5–07–2021, п. 7.7.3' });
        if (m && m[2] === '-') add({ category: 'table', level: 'info', blockIndex: capIndex, title: 'Тире в подписи таблицы', detail: 'В подписи использован дефис вместо тире.', recommend: 'Номер и наименование таблицы разделяйте тире.', ref: 'СТУ 7.5–07–2021, п. 7.7.3' });
        if (/\.[\s]*$/.test(t)) add({ category: 'table', level: 'error', blockIndex: capIndex, title: 'Точка в конце подписи таблицы', detail: `Подпись «${t}» заканчивается точкой.`, recommend: 'Точку после наименования таблицы не ставят.', ref: 'СТУ 7.5–07–2021, п. 7.7.3' });
        const capRuns = (cap.runs || []).filter((x) => normText(x.text));
        if (capRuns.length && capRuns.some((x) => Math.abs((x.size || 14) - 14) > 0.01)) add({ category: 'table', level: 'warning', blockIndex: capIndex, title: 'Размер шрифта подписи таблицы', detail: 'Подпись таблицы набрана не 14 размером.', recommend: 'Заголовок таблицы оформляйте 14 размером шрифта.', ref: 'СТУ 7.5–07–2021, п. 7.7.3' });
        // ссылка в тексте
        if (m) {
          tableNums.push(m[1]);
          const num = m[1];
          const refRe = new RegExp('таблиц\\D{0,12}' + num.replace('.', '\\.'), 'i');
          const allText = collectText(blocks);
          if (!refRe.test(allText)) add({ category: 'table', level: 'warning', blockIndex: capIndex, title: 'Нет ссылки на таблицу в тексте', detail: `В тексте нет ссылки на таблицу ${num}.`, recommend: 'На все таблицы приводите ссылки в тексте документа, например «в таблице ' + num + '».', ref: 'СТУ 7.5–07–2021, п. 7.7.6' });
          else {
            let firstRef = null;
            for (let j = 0; j < blocks.length; j++) { if (blocks[j].kind !== 'p') continue; if (refRe.test(blocks[j].text || '')) { firstRef = j; break; } }
            if (firstRef != null && firstRef > i) add({ category: 'table', level: 'warning', blockIndex: capIndex, title: 'Ссылка на таблицу после таблицы', detail: `Ссылка на таблицу ${num} встречается только после самой таблицы.`, recommend: 'Таблицу помещают после текста, в котором дана ссылка на неё.', ref: 'СТУ 7.5–07–2021, п. 7.7.2' });
          }
        }
      } else {
        add({ category: 'table', level: 'error', blockIndex: i, title: 'У таблицы нет подписи', detail: 'Перед таблицей нет подписи «Таблица N – Наименование».', recommend: 'Поместите слева над таблицей слово «Таблица», номер и наименование.', ref: 'СТУ 7.5–07–2021, п. 7.7.3' });
      }
      if (!b.hasBorders) add({ category: 'table', level: 'error', blockIndex: i, title: 'Таблица без линий', detail: 'Таблица не ограничена линиями.', recommend: 'Ограничьте таблицу линиями справа, слева и снизу.', ref: 'СТУ 7.5–07–2021, п. 7.7.4' });
      // проверка головки
      const header = b.rows[0];
      if (header) {
        let headerBold = false, headerCenter = true, forbidden = false;
        for (const cell of header.cells) {
          for (const p of cell.paras) {
            if (isBoldPara(p)) headerBold = true;
            if (normText(p.text) && p.jc && p.jc !== 'center') headerCenter = false;
            if (/номер\s+по\s+порядку|№\s*п\/?п|по\s+порядку/i.test(p.text || '')) forbidden = true;
          }
        }
        if (!headerBold) add({ category: 'table', level: 'warning', blockIndex: i, title: 'Заголовки граф не выделены', detail: 'Заголовки граф таблицы не выделены полужирным шрифтом.', recommend: 'Выделяйте заголовки граф полужирным шрифтом и выравнивайте по центру.', ref: 'СТУ 7.5–07–2021, п. 7.7.7' });
        if (!headerCenter) add({ category: 'table', level: 'info', blockIndex: i, title: 'Заголовки граф не по центру', detail: 'Заголовки граф выровнены не по центру.', recommend: 'Заголовки граф выравнивайте по центру.', ref: 'СТУ 7.5–07–2021, п. 7.7.7' });
        if (forbidden) add({ category: 'table', level: 'warning', blockIndex: i, title: 'Графа «Номер по порядку»', detail: 'В таблицу включена графа «Номер по порядку».', recommend: 'Графу «Номер по порядку» в таблицу включать не допускается.', ref: 'СТУ 7.5–07–2021, п. 7.7.4' });
      }
      // размер шрифта внутри таблицы 10–12 (рекомендуется)
      const sizes = new Set();
      for (const row of b.rows) for (const cell of row.cells) for (const p of cell.paras) for (const r of (p.runs || [])) if ((r.text || '').trim() && r.size) sizes.add(Math.round(r.size * 10) / 10);
      const badSizes = [...sizes].filter((s) => s < 9 || s > 12.5);
      if (badSizes.length) add({ category: 'table', level: 'info', blockIndex: i, title: 'Размер шрифта в таблице', detail: `Размеры шрифта в таблице: ${badSizes.map((s) => String(s).replace('.', ',')).join(', ')} пт.`, recommend: 'В таблицах рекомендуется Times New Roman 10–12.', ref: 'СТУ 7.5–07–2021, п. 7.7.13' });
    }
    checkCaptionSequence(tableNums, 'table', 'таблиц', add);
  }

  // --------------------------------------------------------------- figures
  function figureNumberValue(t) {
    const m = normText(t).match(/^(?:Рисунок|Таблица)\s+([А-ЯA-Z]?\d+(?:\.\d+)*)/i);
    return m ? m[1] : null;
  }
  function checkFigures(blocks, roles, add) {
    const allText = collectText(blocks);
    const captionNums = [];
    // 1) рисунки-изображения должны иметь подпись
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind !== 'p' || !b.hasDrawing) continue;
      if (roles[i] && roles[i].role === 'title') continue;
      let capIdx = -1;
      for (let j = i + 1; j < blocks.length && j <= i + 3; j++) {
        if (blocks[j].kind !== 'p') break;
        if (roles[j].role === 'figureCaption') { capIdx = j; break; }
        if (!normText(blocks[j].text)) continue;
        break;
      }
      if (capIdx < 0) add({
        category: 'figure', level: 'warning', blockIndex: i,
        title: 'Рисунок без подписи',
        detail: 'После изображения нет подписи «Рисунок N – Наименование».',
        recommend: 'Под иллюстрацией по центру поместите слово «Рисунок», его номер и наименование (Times New Roman 14).',
        ref: 'СТУ 7.5–07–2021, п. 7.8.1, 7.8.3',
      });
      // Размер изображения (хорошая видимость и читаемость)
      for (const r of b.runs || []) for (const p of (r.parts || [])) {
        if (p.type === 'drawing' && p.cx) {
          const wmm = p.cx / 914400 * 25.4;
          if (wmm > 0 && wmm < 50) add({
            category: 'figure', level: 'info', blockIndex: i,
            title: 'Мелкое изображение',
            detail: `Ширина изображения около ${wmm.toFixed(0)} мм.`,
            recommend: 'Иллюстрация должна быть хорошо видимой и читаемой – увеличьте изображение.',
            ref: 'СТУ 7.5–07–2021, п. 7.8.1',
          });
        }
      }
    }
    // 2) подписи рисунков
    for (let i = 0; i < blocks.length; i++) {
      if (roles[i].role !== 'figureCaption') continue;
      const b = blocks[i];
      const t = normText(b.text);
      const m = t.match(/^Рисунок\s+([А-ЯA-Z0-9.\-]+)\s*(–|-|—)\s*(.*)$/i);
      if (b.jc !== 'center') add({ category: 'figure', level: 'error', blockIndex: i, title: 'Подпись рисунка не по центру', detail: `Подпись «${t}» выровнена не по центру.`, recommend: 'Слово «Рисунок», номер и наименование печатайте по центру.', ref: 'СТУ 7.5–07–2021, п. 7.8.3' });
      const runs = (b.runs || []).filter((x) => normText(x.text));
      if (runs.length && runs.some((x) => Math.abs((x.size || 14) - 14) > 0.01)) add({ category: 'figure', level: 'warning', blockIndex: i, title: 'Размер шрифта подписи рисунка', detail: 'Подпись рисунка набрана не 14 размером.', recommend: 'Слово «Рисунок», номер и наименование печатайте Times New Roman 14.', ref: 'СТУ 7.5–07–2021, п. 7.8.3' });
      if (m) {
        const num = m[1];
        captionNums.push(num);
        const refRe = new RegExp('рис(?:\\.|унок|унки|унка)?\\D{0,12}' + num.replace('.', '\\.'), 'i');
        if (!refRe.test(allText)) add({ category: 'figure', level: 'warning', blockIndex: i, title: 'Нет ссылки на рисунок в тексте', detail: `В тексте нет ссылки на рисунок ${num}.`, recommend: 'Приводите ссылки на иллюстрации в тексте, например «в соответствии с рисунком ' + num + '».', ref: 'СТУ 7.5–07–2021, п. 7.8.7' });
        else {
          // ссылка должна предшествовать рисунку (п. 7.8.1)
          let firstRef = null;
          for (let j = 0; j < blocks.length; j++) {
            if (blocks[j].kind !== 'p') continue;
            if (refRe.test(blocks[j].text || '')) { firstRef = j; break; }
          }
          if (firstRef != null && firstRef > i) add({ category: 'figure', level: 'warning', blockIndex: i, title: 'Ссылка на рисунок после рисунка', detail: `Ссылка на рисунок ${num} встречается только после самого рисунка.`, recommend: 'Иллюстрации размещают после ссылки на них в тексте.', ref: 'СТУ 7.5–07–2021, п. 7.8.1' });
        }
      }
    }
    // 3) последовательность нумерации рисунков
    checkCaptionSequence(captionNums, 'figure', 'рисунков', add);
  }

  // Проверка последовательности номеров подписей (1, 2, 3 … или 2.1, 2.2 …).
  function checkCaptionSequence(nums, category, what, add) {
    const counters = {};
    let prev = null;
    let broken = false;
    for (const num of nums) {
      const parts = num.split('.').map(Number);
      const parent = parts.slice(0, -1).join('.');
      const expected = (counters[parent] || 0) + 1;
      if (parts[parts.length - 1] !== expected) broken = true;
      counters[parent] = parts[parts.length - 1];
      prev = num;
    }
    if (broken) add({
      category, level: 'warning', blockIndex: null,
      title: 'Нарушена нумерация ' + what,
      detail: `Номера подписей идут не по порядку: ${nums.join(', ')}.`,
      recommend: 'Нумеруйте иллюстрации (таблицы) сквозной нумерацией арабскими цифрами по порядку.',
      ref: 'СТУ 7.5–07–2021, п. 7.8.2',
    });
  }

  // ----------------------------------------------------- references & sources
  // Нестандартные символы, которые почти не набирают вручную и часто выдают
  // копирование или генерацию текста.
  const SUSPICIOUS_CHARS = {
    '\u00A0': { name: 'неразрывный пробел', ai: false, level: 'info' },
    '\u202F': { name: 'узкий неразрывный пробел', ai: true, level: 'warning' },
    '\u2009': { name: 'тонкий пробел', ai: true, level: 'warning' },
    '\u2007': { name: 'пробел-цифра (figure space)', ai: true, level: 'warning' },
    '\u2008': { name: 'пунктуационный пробел', ai: true, level: 'warning' },
    '\u200A': { name: 'волосяной пробел', ai: true, level: 'warning' },
    '\u200B': { name: 'нулевой пробел (zero-width space)', ai: true, level: 'error' },
    '\u200C': { name: 'нулевой неприсоединяющий (ZWNJ)', ai: true, level: 'error' },
    '\u200D': { name: 'нулевой присоединяющий (ZWJ)', ai: true, level: 'error' },
    '\u2060': { name: 'соединитель слов (word joiner)', ai: true, level: 'error' },
    '\uFEFF': { name: 'невидимый символ (BOM)', ai: true, level: 'error' },
    '\u00AD': { name: 'мягкий перенос', ai: true, level: 'warning' },
    '\u2011': { name: 'неразрывный дефис', ai: true, level: 'warning' },
    '\u2212': { name: 'знак «минус»', ai: true, level: 'warning' },
    '\u00B7': { name: 'средняя точка «·»', ai: true, level: 'warning' },
    '\u2022': { name: 'маркер «•»', ai: false, level: 'info' },
    '\u2192': { name: 'стрелка «→»', ai: true, level: 'warning' },
    '\u2190': { name: 'стрелка «←»', ai: true, level: 'warning' },
    '\u21D2': { name: 'стрелка «⇒»', ai: true, level: 'warning' },
    '\u00D7': { name: 'знак умножения «×»', ai: false, level: 'info' },
    '\u201C': { name: 'английская кавычка “', ai: false, level: 'info' },
    '\u201D': { name: 'английская кавычка ”', ai: false, level: 'info' },
    '\u2018': { name: 'английская кавычка ‘', ai: false, level: 'info' },
    '\u2019': { name: 'английская кавычка ’', ai: false, level: 'info' },
    '\u201E': { name: 'кавычка „', ai: false, level: 'info' },
    '\u2032': { name: 'штрих «′»', ai: true, level: 'warning' },
    '\u2033': { name: 'двойной штрих «″»', ai: true, level: 'warning' },
    '\u2044': { name: 'дробная черта «⁄»', ai: true, level: 'warning' },
    '"': { name: 'прямая кавычка «"»', ai: true, level: 'warning' },
    "'": { name: "прямая кавычка «'»", ai: true, level: 'warning' },
    '`': { name: 'обратный апостроф', ai: true, level: 'warning' },
    '\u2026': { name: 'многоточие «…»', ai: false, level: 'info' },
  };
  const HOMOGLYPH_KEY = '__homoglyph__';

  // ------------------------------------------------------------ выделения
  // Полужирный, курсив, подчёркивание, выделение цветом, цветной текст.
  function checkEmphasis(blocks, roles, add) {
    const ALLOW = new Set(['body', 'where', 'list', 'manualList']);
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      if (!ALLOW.has(roles[i].role)) continue;
      if (b.isEmpty) continue;
      const bold = [], ital = [], und = [], hl = [], col = [];
      (b.runs || []).forEach((r, ri) => {
        if (!(r.text || '').trim()) return;
        if (r.bold) bold.push(ri);
        if (r.italic) ital.push(ri);
        if (r.underline) und.push(ri);
        if (r.highlight) hl.push(ri);
        const c = r.colorVal;
        if (c && /^[0-9a-fA-F]{6}$/.test(c) && c.toUpperCase() !== '000000') col.push(ri);
      });
      const mk = (arr, level, title, rec) => {
        if (!arr.length) return;
        add({ category: 'font', level: level, blockIndex: i, runIndices: arr, title: title, detail: `Выделенных фрагментов в абзаце: ${arr.length}.`, recommend: rec, ref: 'СТУ 7.5–07–2021, п. 7.1.1' });
      };
      mk(bold, 'info', 'Полужирный шрифт в основном тексте', 'Полужирным выделяют только заголовки; в обычном тексте уберите.');
      mk(ital, 'info', 'Курсив в основном тексте', 'Курсив допустим для терминов и латыни; иначе уберите.');
      mk(und, 'warning', 'Подчёркнутый текст', 'Уберите подчёркивание в обычном тексте.');
      mk(hl, 'warning', 'Выделение цветом', 'Уберите цветовое выделение текста.');
      mk(col, 'info', 'Цветной текст', 'Используйте чёрный цвет текста.');
    }
  }

  function checkTypography(blocks, roles, add) {
    const ALLOW = new Set(['body', 'heading', 'structural', 'list', 'manualList', 'where', 'formula', 'figureCaption', 'tableCaption']);
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      if (!ALLOW.has(roles[i].role)) continue;
      const byKey = {};
      for (let ri = 0; ri < (b.runs || []).length; ri++) {
        const run = b.runs[ri]; const text = run.text || ''; if (!text) continue;
        const singleText = run.parts && run.parts.length === 1 && run.parts[0].type === 'text';
        for (let ci = 0; ci < text.length; ci++) {
          const ch = text[ci]; const info = SUSPICIOUS_CHARS[ch];
          if (!info) continue;
          // прямые кавычки отмечаем только рядом с русским текстом (в англ. – норма)
          if ((ch === '"' || ch === "'") && !/[А-Яа-яЁё]/.test(text.slice(Math.max(0, ci - 3), ci + 4))) continue;
          // знак «минус» в формулах – норма
          if (ch === '\u2212' && (roles[i].role === 'formula' || roles[i].role === 'where')) continue;
          const e = byKey[ch] || (byKey[ch] = { info, count: 0, chars: {}, whole: new Set() });
          e.count++;
          if (singleText) (e.chars[ri] = e.chars[ri] || []).push([ci, ci + 1]);
          else e.whole.add(ri);
        }
        if (singleText) {
          const words = text.match(/[A-Za-zА-Яа-яЁё]+/g) || [];
          for (const w of words) {
            const cyr = (w.match(/[А-Яа-яЁё]/g) || []).length;
            const lat = (w.match(/[A-Za-z]/g) || []).length;
            if (cyr >= 3 && lat >= 1 && lat <= 2) {
              const start = text.indexOf(w);
              const e = byKey[HOMOGLYPH_KEY] || (byKey[HOMOGLYPH_KEY] = { info: { ai: true, level: 'info' }, count: 0, chars: {}, whole: new Set(), word: w });
              e.count++; e.word = w;
              (e.chars[ri] = e.chars[ri] || []).push([start, start + w.length]);
            }
          }
        }
      }
      for (const key in byKey) {
        const e = byKey[key];
        const chars = [];
        for (const ri in e.chars) chars.push({ runIndex: +ri, ranges: e.chars[ri] });
        const isHomo = key === HOMOGLYPH_KEY;
        add({
          category: 'chars', level: e.info.level, ai: !!e.info.ai, blockIndex: i,
          chars,
          runIndices: e.whole.size ? [...e.whole] : undefined,
          title: isHomo ? 'Возможный машинный текст: смешение алфавитов'
            : (e.info.ai ? 'Возможный машинный текст: ' + e.info.name : 'Нестандартный символ: ' + e.info.name),
          detail: isHomo
            ? `В слове «${e.word}» смешаны русские и латинские буквы – типичная ошибка копирования или генерации.`
            : `Найден символ «${e.info.name}»${e.count > 1 ? ' (повторов: ' + e.count + ')' : ''}. Такой знак не набирают вручную – он появляется при вставке или генерации текста.`,
          recommend: 'Замените символ на обычный пробел или знак препинания; проверьте этот фрагмент.',
          ref: 'СТУ 7.5–07–2021, п. 7.3.1',
        });
      }
      const text = b.text || '';
      const urls = urlSpans(text);
      // Диапазоны чисел: должно быть тире «–» без пробелов
      const badRanges = [];
      if (!/TOC \\o|PAGEREF|INCLUDETEXT/i.test(text)) {
        let rm; const rangeRe = /(\d+(?:[.,]\d+)?)\s*([-–—])\s*(\d+(?:[.,]\d+)?)/g;
        const BOUND = /[0-9A-Za-zА-Яа-я_/.-]/;
        while ((rm = rangeRe.exec(text))) {
          if (inSpans(rm.index, urls)) continue; // внутри ссылки дефис не трогаем
          const before = text[rm.index - 1] || '';
          const after = text[rm.index + rm[0].length] || '';
          if (BOUND.test(before) || BOUND.test(after)) continue; // часть URL, кода или даты
          const sep = rm[2];
          const hasSpace = /\s/.test(rm[0]);
          if (sep !== '–' || hasSpace) badRanges.push(rm[0]);
        }
      }
      if (badRanges.length) {
        let rc = []; for (const r of badRanges) rc = rc.concat(phraseRanges(b, r));
        add({
          category: 'chars', level: 'info', ai: false, blockIndex: i, chars: rc,
          title: 'Знак в диапазоне чисел',
          detail: `В диапазоне чисел использовано «${badRanges[0].replace(/\s/g, '␣')}». Между числами в цифровой форме принято тире «–» без пробелов, а не дефис.`,
          recommend: 'Ставьте тире «–» без пробелов: 1,0–6,0. Дефис и длинное тире между числами не используют.',
          ref: 'Правила русской типографики (Розенталь; Мильчин–Чельцова), gramota.ru',
        });
      }
      // длинное тире (внутри ссылок не считаем)
      const emCount = (text.match(/\u2014/g) || []).length;
      if (emCount) add({
        category: 'ai', level: 'info', ai: true, blockIndex: i, chars: phraseRanges(b, '\u2014'),
        title: 'Использовано длинное тире',
        detail: `Длинное тире встречается в абзаце ${emCount} раз(а). Обычно так не пишут вручную.`,
        recommend: 'Используйте короткое тире «–»; в диапазонах чисел без пробелов.',
        ref: 'Правила русской типографики',
      });
      // длинное тире вплотную к словам (без пробелов)
      if (/\S\u2014\S/.test(text)) add({
        category: 'chars', level: 'warning', ai: false, blockIndex: i, chars: phraseRanges(b, '\u2014'),
        title: 'Длинное тире без пробелов',
        detail: 'Длинное тире стоит вплотную к словам.',
        recommend: 'В русском тексте тире между словами отделяют пробелами.',
        ref: 'Правила русской типографики',
      });
      // Двойные пробелы (подсвечиваем сами пробелы)
      if (/[^ ] {2,}\S/.test(text)) add({
        category: 'chars', level: 'info', ai: true, blockIndex: i, chars: spaceRanges(b),
        title: 'Двойные пробелы', detail: 'В тексте подряд идут несколько пробелов.',
        recommend: 'Оставьте один пробел.', ref: 'СТУ 7.5–07–2021, п. 7.3.1',
      });
      // Пробел перед знаком препинания (кроме двоеточия – оно используется в библиографии)
      if (/\S\s+[,.;!?](?=\s|$)/.test(text)) add({
        category: 'chars', level: 'warning', ai: true, blockIndex: i, chars: spaceBeforePunctRanges(b),
        title: 'Пробел перед знаком препинания', detail: 'Перед знаком препинания стоит пробел.',
        recommend: 'Уберите пробел перед знаком препинания.', ref: 'СТУ 7.5–07–2021, п. 7.3.1',
      });
    }
    // Нестандартные символы в таблицах (картинки в таблицах не проверяем)
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'table') continue;
      if (roles[i] && roles[i].role === 'title') continue;
      const byKey = {};
      for (const row of b.rows) for (const cell of row.cells) for (const p of cell.paras) {
        const text = p.text || '';
        for (let ci = 0; ci < text.length; ci++) {
          const ch = text[ci]; const info = SUSPICIOUS_CHARS[ch]; if (!info) continue;
          const e = byKey[ch] || (byKey[ch] = { info, count: 0 }); e.count++;
        }
      }
      for (const key in byKey) {
        const e = byKey[key];
        add({
          category: 'chars', level: e.info.level, ai: !!e.info.ai, blockIndex: i,
          title: e.info.ai ? 'Возможный машинный текст: ' + e.info.name : 'Нестандартный символ: ' + e.info.name,
          detail: `В таблице найден символ «${e.info.name}»${e.count > 1 ? ' (повторов: ' + e.count + ')' : ''}.`,
          recommend: 'Замените символ на обычный пробел или знак препинания.',
          ref: 'СТУ 7.5–07–2021, п. 7.3.1',
        });
      }
    }
  }

  // ----------------------------------------------------- references & sources
  // Строгая проверка библиографических записей по ГОСТ Р 7.0.100 (по категориям).
  function bibCategory(t) {
    const s = t.toLowerCase();
    if (/url\s*:|https?:\/\/|www\./i.test(t)) return 'electronic';
    if (/^\s*(гост|сп|снип|санпин|iso|гост р)/i.test(t) || /национальный стандарт/.test(s)) return 'standard';
    if (/патент/.test(s)) return 'patent';
    if (/диссертац|автореферат/.test(s)) return 'dissertation';
    if (/ \/\/ /.test(t) || /материалы .*конференц|сборник/i.test(s)) return 'article';
    return 'book';
  }
  function checkBibliographyFormat(entries, start, add) {
    const CATNAME = { book: 'книга / учебник', article: 'статья', electronic: 'электронный ресурс', standard: 'стандарт', dissertation: 'диссертация / автореферат', patent: 'патент' };
    for (const e of entries) {
      const t = e.text.replace(/^\d{1,3}\.\s*/, '');
      const cat = bibCategory(t);
      const miss = [];
      if (cat === 'book') {
        if (!/^\s*[А-ЯЁA-Z][а-яёa-z'\-]+,\s*[А-ЯЁA-Z]\.\s?[А-ЯЁA-Z]?\./.test(t) && !/под (общ|ред)|под редакцией|\/\s*[А-ЯЁA-Z][а-яёa-z]+,/.test(t)) miss.push('автор или ответственный');
        if (!/ \/ /.test(t)) miss.push('«/» перед сведениями об ответственности');
        if (!/\s[–-]\s[А-ЯA-Z][^:]*:\s*[^,]+,\s*(19|20)\d{2}/.test(t)) miss.push('место, издательство и год');
        if (!/(19|20)\d{2}\.\s*[–-]\s*\d+\s*с\./i.test(t)) miss.push('объём в страницах («N с.»)');
      } else if (cat === 'article') {
        if (!/ \/\/ /.test(t)) miss.push('«//» перед названием источника');
        if (!/(19|20)\d{2}/.test(t)) miss.push('год');
        if (!/С\.\s*\d/.test(t)) miss.push('страницы («С. N»)');
      } else if (cat === 'electronic') {
        if (!/(url\s*:|https?:\/\/|www\.)/i.test(t)) miss.push('URL');
        if (!/дата\s+обращения/i.test(t)) miss.push('дата обращения');
      } else if (cat === 'standard') {
        if (!/(19|20)\d{2}/.test(t)) miss.push('год');
      } else if (cat === 'dissertation') {
        if (!/специальност/i.test(t)) miss.push('код специальности');
        if (!/(19|20)\d{2}/.test(t)) miss.push('год');
      } else if (cat === 'patent') {
        if (!/№/.test(t)) miss.push('номер');
        if (!/(19|20)\d{2}/.test(t)) miss.push('год');
      }
      if (miss.length) add({
        category: 'sources', level: 'warning', blockIndex: e.index,
        title: 'Неполная библиографическая запись (' + CATNAME[cat] + ')',
        detail: `Не хватает: ${miss.join('; ')}.`,
        recommend: 'Оформляйте запись по ГОСТ Р 7.0.100 (примеры в приложении У СТУ).',
        ref: 'СТУ 7.5–07–2021, п. 7.10.2',
      });
      // слэши: у книги один «/», у статьи двойной «//»
      if (cat === 'book' && / \/\/ /.test(t)) add({ category: 'sources', level: 'warning', blockIndex: e.index, title: 'Двойной слэш в описании книги', detail: 'В описании книги использован двойной слэш «//».', recommend: 'Для книги используют один слэш «/».', ref: 'СТУ 7.5–07–2021, п. 7.10.2' });
      if (cat === 'article' && / \/ /.test(t) && !/ \/\/ /.test(t)) add({ category: 'sources', level: 'warning', blockIndex: e.index, title: 'Нужен двойной слэш «//»', detail: 'В описании статьи использован одиночный слэш «/».', recommend: 'Перед названием источника ставят двойной слэш «//».', ref: 'СТУ 7.5–07–2021, п. 7.10.2' });
    }
  }

  // Список сокращений (п. 6.8): столбец «сокращение – расшифровка», по алфавиту.
  function checkAbbreviations(blocks, roles, add) {
    let start = -1, end = blocks.length;
    for (let i = 0; i < blocks.length; i++) {
      if (roles[i].role === 'structural' && roles[i].structKey === 'sokrashcheniya') { start = i; continue; }
      if (start >= 0 && roles[i].role === 'structural') { end = i; break; }
    }
    if (start < 0) return;
    const entries = [];
    for (let i = start + 1; i < end; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const t = normText(b.text); if (!t) continue;
      entries.push({ index: i, text: t });
    }
    if (!entries.length) return;
    let noSep = 0, missingIndent = 0; const abbrs = [];
    for (const e of entries) {
      const m = e.text.match(/^([^–—\-:]{1,24})\s*[–—\-:]\s*\S/);
      if (m) abbrs.push({ index: e.index, abbr: m[1].trim() }); else noSep++;
      const b = blocks[e.index]; const fl = b.ind ? b.ind.firstLine : null;
      if (!fl) missingIndent++;
    }
    if (noSep) add({ category: 'abbrev', level: 'warning', blockIndex: start, title: 'Неверное оформление списка сокращений', detail: `Записей без разделителя «сокращение – расшифровка»: ${noSep}.`, recommend: 'Каждое сокращение приводите с расшифровкой, например «АИС – автоматизированная информационная система».', ref: 'СТУ 7.5–07–2021, п. 6.8' });
    if (missingIndent) add({ category: 'abbrev', level: 'info', blockIndex: start, title: 'Записи списка сокращений без абзацного отступа', detail: `Записей без абзацного отступа: ${missingIndent}.`, recommend: 'Печатайте записи списка сокращений столбцом с абзацного отступа.', ref: 'СТУ 7.5–07–2021, п. 6.8' });
    const lower = abbrs.map((a) => a.abbr.toLowerCase());
    for (let i = 1; i < lower.length; i++) {
      if (lower[i - 1] > lower[i]) {
        add({ category: 'abbrev', level: 'warning', blockIndex: abbrs[i].index, title: 'Нарушен алфавитный порядок сокращений', detail: `«${abbrs[i].abbr}» стоит не по алфавиту.`, recommend: 'Располагайте сокращения слева столбцом в алфавитном порядке.', ref: 'СТУ 7.5–07–2021, п. 6.8' });
        break;
      }
    }
  }

  // ---------------------------------------------------------- ИИ-признаки
  // Источники: Википедия «Признаки сгенерированности текста», gramota.ru,
  // разборы на Habr. Список устойчивых оборотов, markdown-остатков и меток.
  const AI_CLICHES = [
    // обороты, которые чаще всего встречаются в сгенерированном тексте
    'играет важную роль', 'играет ключевую роль', 'играет значительную роль',
    'важную роль в', 'ключевую роль в', 'значительную роль в',
    'важно отметить', 'необходимо отметить', 'стоит отметить', 'следует отметить',
    'важно подчеркнуть', 'стоит подчеркнуть', 'следует подчеркнуть', 'необходимо подчеркнуть',
    'важно понимать', 'стоит заметить', 'следует заметить', 'примечательно, что', 'интересно, что',
    'критично отметить', 'стоит учесть', 'стоит учитывать', 'необходимо учитывать',
    'подводя итог', 'в заключение', 'в целом можно сказать', 'таким образом',
    'как итог', 'в конечном итоге', 'в конечном счёте',
    'в современном мире', 'на сегодняшний день', 'в нынешних условиях', 'в эпоху цифровизации',
    'в данной статье', 'в этой статье', 'в данной работе', 'в рамках данного', 'данный подход',
    'широкий спектр', 'ключевой аспект', 'неотъемлемой частью', 'неотъемлемая часть',
    'представляет собой важный', 'выводит на новый уровень', 'приобретает особую актуальность',
    'имеет большое значение', 'имеет важное значение', 'является одним из',
    'выступает символом', 'является напоминанием', 'символизирует', 'неизгладимый след',
    'ключевой момент', 'поворотный момент', 'динамично развивающ', 'широкое распространение',
    'комплексный подход', 'эффективное решение', 'оптимальное решение',
    'расположенный в самом сердце', 'богатым культурным наследием', 'вековое наследие',
    'сталкивается с рядом проблем', 'перспективы на будущее', 'могут варьироваться',
    'я надеюсь, это помогло', 'дайте мне знать', 'сообщите мне', 'вы абсолютно правы',
    'в этой статье мы рассмотрим', 'более подробный анализ',
  ];
  // «Вода»: слова и связки, которые часто заполняют текст без смысла.
  const AI_FILLER = ['таким образом', 'в целом', 'кроме того', 'более того', 'в частности',
    'в свою очередь', 'следует отметить', 'стоит отметить', 'важно отметить', 'необходимо',
    'в рамках', 'является', 'позволяет', 'обеспечивает', 'данный', 'данная', 'данные', 'также', 'однако'];

  // Диапазоны строки внутри прогонов (для точечной подсветки фразы).
  function phraseRanges(block, phrase) {
    const res = []; const needle = phrase.toLowerCase();
    for (let ri = 0; ri < (block.runs || []).length; ri++) {
      const run = block.runs[ri];
      if (!(run.parts && run.parts.length === 1 && run.parts[0].type === 'text')) continue;
      const t = (run.text || '').toLowerCase(); let idx = t.indexOf(needle);
      while (idx >= 0) { res.push({ runIndex: ri, ranges: [[idx, idx + needle.length]] }); idx = t.indexOf(needle, idx + needle.length); }
    }
    return res;
  }

  // Диапазон символов по смещению в склеенном тексте абзаца (для точечной подсветки).
  function rangeAt(block, start, end) {
    const res = []; let off = 0;
    for (let ri = 0; ri < (block.runs || []).length; ri++) {
      const run = block.runs[ri]; const runLen = (run.text || '').length;
      if (end > off && start < off + runLen) {
        const a = Math.max(start, off), b = Math.min(end, off + runLen);
        if (b > a) res.push({ runIndex: ri, ranges: [[a - off, b - off]] });
      }
      off += runLen;
    }
    return res;
  }

  // Адреса ссылок: внутри них дефисы и тире не трогаем, там это нормально.
  function urlSpans(text) {
    const re = /(?:https?:\/\/|ftp:\/\/|www\.)[^\s<>"']+/gi;
    const out = []; let m;
    while ((m = re.exec(text))) out.push([m.index, m.index + m[0].length]);
    return out;
  }
  function inSpans(idx, spans) { for (const s of spans) if (idx >= s[0] && idx < s[1]) return true; return false; }

  // Подсветка серий из двух и более пробелов.
  function spaceRanges(block) {
    const res = [];
    for (let ri = 0; ri < (block.runs || []).length; ri++) {
      const run = block.runs[ri];
      if (!(run.parts && run.parts.length === 1 && run.parts[0].type === 'text')) continue;
      const t = run.text || ''; const re = / {2,}/g; let m;
      while ((m = re.exec(t))) res.push({ runIndex: ri, ranges: [[m.index, m.index + m[0].length]] });
    }
    return res;
  }
  // Подсветка пробелов перед знаком препинания.
  function spaceBeforePunctRanges(block) {
    const res = [];
    for (let ri = 0; ri < (block.runs || []).length; ri++) {
      const run = block.runs[ri];
      if (!(run.parts && run.parts.length === 1 && run.parts[0].type === 'text')) continue;
      const t = run.text || ''; const re = / +(?=[,.;!?](\s|$))/g; let m;
      while ((m = re.exec(t))) if (m[0].length) res.push({ runIndex: ri, ranges: [[m.index, m.index + m[0].length]] });
    }
    return res;
  }

  function checkAiMarkers(blocks, roles, add) {
    const ALLOW = new Set(['body', 'heading', 'structural', 'list', 'manualList', 'where', 'figureCaption', 'tableCaption']);
    const notOnlyBlocks = [];
    const phraseCount = {};
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const t = b.text || ''; if (!t) continue;
      const low = normText(t).toLowerCase();
      if (!ALLOW.has(roles[i].role)) continue;
      // устойчивые обороты с подсветкой самой фразы
      for (const ph of AI_CLICHES) {
        let idx = low.indexOf(ph);
        if (idx < 0) continue;
        let c = 0; while (idx >= 0) { c++; idx = low.indexOf(ph, idx + ph.length); }
        phraseCount[ph] = (phraseCount[ph] || 0) + c;
        add({
          category: 'ai', level: 'info', ai: true, blockIndex: i, chars: phraseRanges(b, ph),
          title: 'Оборот, частый у машинного текста',
          detail: `Найден оборот «${ph}»${c > 1 ? ' (повторов: ' + c + ')' : ''}.`,
          recommend: 'Перепишите своими словами или уберите этот оборот.',
          ref: 'Правила русской типографики',
        });
      }
      // «не только ..., но и ...» — считаем все места, помечаем при частом повторе
      if (/не только/.test(low) && /но и/.test(low)) notOnlyBlocks.push(i);
      // вода: много служебных связок в абзаце
      if (normText(b.text).length > 260) {
        const found = AI_FILLER.filter((w) => low.indexOf(w) >= 0);
        if (found.length >= 4) add({ category: 'ai', level: 'info', ai: true, blockIndex: i, title: 'Много вводных слов (вода)', detail: `Служебные связки: ${found.slice(0, 6).join(', ')}.`, recommend: 'Сократите вводные слова, оставьте суть.', ref: 'Правила русской типографики' });
      }
      // markdown
      let md = null;
      if (/\*\*/.test(t)) md = '**'; else if (/^\s*#{1,6}\s/.test(b.text) || /(^|\s)##\s/.test(t)) md = '##'; else if (/```/.test(t)) md = '```';
      if (md) add({ category: 'ai', level: 'warning', ai: true, blockIndex: i, title: 'Разметка Markdown в тексте', detail: `Найден символ разметки «${md}».`, recommend: 'Уберите символы Markdown.', ref: 'Правила русской типографики' });
    }
    // документ в целом: повторяющийся оборот это признак
    const firstAt = {};
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p' || !b.text) continue;
      const low = b.text.toLowerCase();
      for (const ph of AI_CLICHES) if (firstAt[ph] == null && low.indexOf(ph) >= 0) firstAt[ph] = i;
    }
    for (const ph in phraseCount) {
      if (phraseCount[ph] >= 4) add({ category: 'ai', level: 'info', ai: true, blockIndex: firstAt[ph], title: 'Оборот часто повторяется', detail: `«${ph}» встречается в документе ${phraseCount[ph]} раз(а).`, recommend: 'Замените повторяющийся оборот синонимами или перепишите.', ref: 'Правила русской типографики' });
    }
    if (notOnlyBlocks.length >= 3) {
      notOnlyBlocks.forEach((bi, k) => {
        const b0 = blocks[bi];
        const chars = phraseRanges(b0, 'не только').concat(phraseRanges(b0, 'но и'));
        add({
          category: 'ai', level: 'info', ai: true, blockIndex: bi, chars: chars.length ? chars : undefined,
          title: 'Оборот «не только …, но и …»',
          detail: `Место ${k + 1} из ${notOnlyBlocks.length}: конструкция повторяется в документе ${notOnlyBlocks.length} раз(а).`,
          recommend: 'Разнообразьте формулировки, уберите повторяющийся шаблон.',
          ref: 'Правила русской типографики',
        });
      });
    }
  }

  // Несоответствия оформления по всему документу: указываем конкретное место.
  function checkConsistency(blocks, roles, add) {
    let hyphenRange = 0, endashRange = 0;
    let guillemet = 0, straight = 0, curly = 0;
    let firstHyphen = -1, firstHyphenText = null;
    let firstBadQuote = -1, firstBadQuoteText = null;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      if (roles[i].role === 'title') continue;
      const t = b.text || '';
      let m; const re = /(\d+(?:[.,]\d+)?)\s*([-–—])\s*(\d+(?:[.,]\d+)?)/g;
      while ((m = re.exec(t))) {
        if (m[2] === '-') { hyphenRange++; if (firstHyphen < 0) { firstHyphen = i; firstHyphenText = m[0]; } }
        else if (m[2] === '–') endashRange++;
      }
      guillemet += (t.match(/«/g) || []).length;
      const sq = (t.match(/"/g) || []).length + (t.match(/[“”]/g) || []).length;
      straight += (t.match(/"/g) || []).length; curly += (t.match(/[“”]/g) || []).length;
      if (sq && firstBadQuote < 0) { firstBadQuote = i; firstBadQuoteText = (t.match(/"[^"]{0,30}"|[“”]/) || [''])[0]; }
    }
    if (hyphenRange > 0 && endashRange > 0) add({
      category: 'chars', level: 'info', ai: false, blockIndex: firstHyphen >= 0 ? firstHyphen : null,
      chars: firstHyphen >= 0 ? phraseRanges(blocks[firstHyphen], firstHyphenText) : undefined,
      title: 'Непоследовательные тире в диапазонах',
      detail: `Часть диапазонов через дефис (${hyphenRange}), часть через тире (${endashRange}). Начало несоответствия ниже.`,
      recommend: 'Приведите все диапазоны к одному виду: тире без пробелов.',
      ref: 'Правила русской типографики',
    });
    if ((straight + curly) > 0 && guillemet > 0) add({
      category: 'chars', level: 'info', ai: false, blockIndex: firstBadQuote >= 0 ? firstBadQuote : null,
      chars: firstBadQuote >= 0 && firstBadQuoteText ? phraseRanges(blocks[firstBadQuote], firstBadQuoteText) : undefined,
      title: 'Разнобой в кавычках',
      detail: `Встречаются «ёлочки» (${guillemet}) и прямые или английские кавычки (${straight + curly}).`,
      recommend: 'Используйте единый вид кавычек, в русском тексте «ёлочки».',
      ref: 'Правила русской типографики',
    });
  }

  // Проверка сокращений: расшифровка при первом упоминании и наличие в списке.
  function checkAbbrevUsage(blocks, roles, add) {
    let start = -1, end = blocks.length;
    for (let i = 0; i < blocks.length; i++) {
      if (roles[i].role === 'structural' && roles[i].structKey === 'sokrashcheniya') { start = i; continue; }
      if (start >= 0 && roles[i].role === 'structural') { end = i; break; }
    }
    const list = [];
    if (start >= 0) for (let i = start + 1; i < end; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const m = normText(b.text).match(/^([^–—\-:]{1,24})\s*[–—\-:]\s*\S/);
      if (m) list.push({ abbr: m[1].trim(), index: i });
    }
    if (!list.length) return;
    const bodyIdx = [];
    for (let i = 0; i < blocks.length; i++) {
      if (blocks[i].kind !== 'p') continue;
      if (roles[i].role === 'title' || roles[i].role === 'toc') continue;
      if (start >= 0 && i > start && i < end) continue;
      bodyIdx.push(i);
    }
    const textOf = (i) => normText(blocks[i].text);
    for (const a of list) {
      let first = -1, pos = -1;
      for (const i of bodyIdx) { const p = textOf(i).indexOf(a.abbr); if (p >= 0) { first = i; pos = p; break; } }
      if (first < 0) { add({ category: 'abbrev', level: 'info', blockIndex: a.index, title: 'Сокращение не используется в тексте', detail: `«${a.abbr}» есть в списке сокращений, но не встречается в тексте.`, recommend: 'Уберите лишнее сокращение из списка.', ref: 'СТУ 7.5–07–2021, п. 6.8' }); continue; }
      const t = textOf(first); const before = t.slice(Math.max(0, pos - 90), pos);
      const expanded = /\([^)]{5,80}\)\s*$/.test(before);
      if (!expanded) add({ category: 'abbrev', level: 'warning', blockIndex: first, title: 'Сокращение не расшифровано при первом упоминании', detail: `«${a.abbr}» впервые встречается без расшифровки.`, recommend: 'При первом упоминании приводите расшифровку: полное название (АБВ).', ref: 'СТУ 7.5–07–2021, п. 7.3.3' });
    }
    const known = new Set(list.map((x) => x.abbr));
    const common = new Set(['ГОСТ', 'РФ', 'СФУ', 'ИИ', 'ВКР', 'МД', 'ДР', 'БР', 'КР', 'КП', 'ПЭВМ', 'РГЗ', 'РГР', 'ООО', 'АО', 'ПАО', 'БПЛА', 'API', 'URL', 'HTTP', 'HTTPS', 'IT', 'AI', 'ML', 'VR', 'AR', 'ТАСС', 'РИА', 'ИТ']);
    const re = /(^|[^А-ЯЁA-Z])([А-ЯЁA-Z]{2,6})(?![А-ЯЁA-Z])/g;
    const seen = new Map();
    for (const i of bodyIdx) {
      const t = textOf(i); let m;
      re.lastIndex = 0;
      while ((m = re.exec(t))) { const w = m[2]; if (known.has(w) || common.has(w)) continue; if (!seen.has(w)) seen.set(w, i); }
    }
    let n = 0;
    for (const [w, i] of seen) { if (n >= 5) break; add({ category: 'abbrev', level: 'info', blockIndex: i, title: 'Сокращение не внесено в список', detail: `«${w}» встречается в тексте, но отсутствует в списке сокращений.`, recommend: 'Внесите сокращение в СПИСОК СОКРАЩЕНИЙ или расшифруйте при первом упоминании.', ref: 'СТУ 7.5–07–2021, п. 6.8' }); n++; }
  }

  // Нумерация формул и ссылки на них.
  function checkFormulas(blocks, roles, add) {
    const nums = [];
    for (let i = 0; i < blocks.length; i++) {
      if (blocks[i].kind !== 'p' || roles[i].role !== 'formula') continue;
      const m = normText(blocks[i].text).match(/\((\d+(?:\.\d+)*)\)\s*$/);
      if (m) nums.push({ index: i, num: m[1] });
    }
    const counters = {}; let broken = false;
    for (const f of nums) {
      const parts = f.num.split('.').map(Number);
      const parent = parts.slice(0, -1).join('.');
      const expected = (counters[parent] || 0) + 1;
      if (parts[parts.length - 1] !== expected) broken = true;
      counters[parent] = parts[parts.length - 1];
    }
    if (broken) add({ category: 'formula', level: 'warning', blockIndex: nums[0] ? nums[0].index : null, title: 'Нарушена нумерация формул', detail: `Номера формул идут не по порядку: ${nums.map((f) => f.num).join(', ')}.`, recommend: 'Нумеруйте формулы последовательно арабскими цифрами.', ref: 'СТУ 7.5–07–2021, п. 7.6.2' });
    const allText = collectText(blocks);
    for (const f of nums) {
      const re = new RegExp('формул\\w*\\D{0,12}\\(?' + f.num.replace('.', '\\.') + '\\)?', 'i');
      if (!re.test(allText)) add({ category: 'formula', level: 'info', blockIndex: f.index, title: 'Нет ссылки на формулу', detail: `В тексте нет ссылки на формулу (${f.num}).`, recommend: 'Ссылайтесь на формулу: «по формуле (' + f.num + ')».', ref: 'СТУ 7.5–07–2021, п. 7.6.8' });
    }
  }

  // Верхний колонтитул.
  function checkHeaders(doc, add) {
    for (const s of doc.sections) {
      const h = s.headers || {};
      for (const k in h) {
        const hf = h[k]; if (!hf) continue;
        if (hf.hasPageField) add({ category: 'pagenum', level: 'warning', blockIndex: null, pageLevel: true, title: 'Номер страницы в верхнем колонтитуле', detail: 'Номер страницы расположен в верхнем колонтитуле.', recommend: 'Номер страницы ставят внизу по центру.', ref: 'СТУ 7.5–07–2021, п. 7.2.1' });
        else if ((hf.text || '').trim()) add({ category: 'page', level: 'info', blockIndex: null, pageLevel: true, title: 'Верхний колонтитул содержит текст', detail: `«${(hf.text || '').trim().slice(0, 60)}».`, recommend: 'Стандарт не предусматривает текст в верхнем колонтитуле.', ref: 'СТУ 7.5–07–2021, п. 7.2' });
      }
    }
  }

  // Пустая строка вокруг рисунков и таблиц (по стандарту объекты отделяют
  // интервалом в одну строку).
  function checkBlankLines(blocks, roles, add) {
    const isBlank = (i) => blocks[i] && blocks[i].kind === 'p' && blocks[i].isEmpty;
    const spAfter = (b) => (b && b.spacing && b.spacing.after) || 0;
    const spBefore = (b) => (b && b.spacing && b.spacing.before) || 0;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i]; const role = roles[i];
      // рисунок: перед картинкой и после подписи
      if (b.kind === 'p' && (role.role === 'image')) {
        const prev = blocks[i - 1];
        if (prev && prev.kind === 'p' && !prev.isEmpty && Math.max(spAfter(prev), spBefore(b)) < 200) add({ category: 'figure', level: 'warning', blockIndex: i, pageRel: [i - 1, i], blankInsert: { at: i, where: 'before' }, title: 'Нет пустой строки перед рисунком', detail: 'Перед рисунком нет пустой строки.', recommend: 'Отделяйте рисунок от текста интервалом в одну строку сверху.', ref: 'СТУ 7.5–07–2021, п. 7.8.1' });
        const cap = (roles[i + 1] && roles[i + 1].role === 'figureCaption') ? i + 1 : null;
        const afterIdx = cap != null ? cap + 1 : i + 1;
        const nx = blocks[afterIdx];
        if (nx && nx.kind === 'p' && !nx.isEmpty && Math.max(spAfter(blocks[afterIdx - 1]), spBefore(nx)) < 200) add({ category: 'figure', level: 'warning', blockIndex: cap != null ? cap : i, pageRel: [afterIdx - 1, afterIdx], blankInsert: { at: afterIdx, where: 'before' }, title: 'Нет пустой строки после рисунка', detail: 'После подписи рисунка сразу идёт текст.', recommend: 'Отделяйте рисунок от текста интервалом в одну строку снизу.', ref: 'СТУ 7.5–07–2021, п. 7.8.1' });
      }
      // таблица: перед подписью и после таблицы
      if (b.kind === 'table') {
        const capIdx = (i - 1 >= 0 && roles[i - 1].role === 'tableCaption') ? i - 1 : null;
        const beforeIdx = capIdx != null ? capIdx - 1 : i - 1;
        const pb = blocks[beforeIdx];
        if (pb && pb.kind === 'p' && !pb.isEmpty && Math.max(spAfter(pb), capIdx != null ? spBefore(blocks[capIdx]) : spBefore(b)) < 200) add({ category: 'table', level: 'warning', blockIndex: capIdx != null ? capIdx : i, pageRel: [beforeIdx, capIdx != null ? capIdx : i], blankInsert: { at: capIdx != null ? capIdx : i, where: 'before' }, title: 'Нет пустой строки перед таблицей', detail: 'Перед таблицей нет пустой строки.', recommend: 'Отделяйте таблицу от текста интервалом в одну строку.', ref: 'СТУ 7.5–07–2021, п. 7.7.2' });
        const nx = blocks[i + 1];
        if (nx && nx.kind === 'p' && !nx.isEmpty && spAfter(b) < 200) add({ category: 'table', level: 'warning', blockIndex: i, pageRel: [i, i + 1], blankInsert: { at: i + 1, where: 'before' }, title: 'Нет пустой строки после таблицы', detail: 'После таблицы сразу идёт текст.', recommend: 'Отделяйте таблицу от текста интервалом в одну строку.', ref: 'СТУ 7.5–07–2021, п. 7.7.2' });
      }
    }
  }

  function checkReferencesAndSources(doc, blocks, roles, presentKeys, add) {
    const URL_RE = /https?:\/\/[^\s)\]"»«]+/gi;
    const normUrl = (u) => u.toLowerCase().replace(/[.,;:]+$/, '').replace(/^https?:\/\//, '').replace(/\/$/, '');

    // Диапазоны символов строки внутри прогонов – для точечной подсветки ссылки.
    function charRangesFor(p, needle) {
      const res = [];
      if (!needle) return res;
      for (let ri = 0; ri < (p.runs || []).length; ri++) {
        const run = p.runs[ri];
        if (!(run.parts && run.parts.length === 1 && run.parts[0].type === 'text')) continue;
        const t = run.text || ''; let idx = t.indexOf(needle);
        while (idx >= 0) { res.push({ runIndex: ri, ranges: [[idx, idx + needle.length]] }); idx = t.indexOf(needle, idx + needle.length); }
      }
      return res;
    }

    // Ссылки вида [1], [1, 2], [59; 67, с. 40-46; 82], [1–3] и URL http(s).
    function parseCitations(t) {
      const out = [];
      const re = /\[([^\]]{1,48})\]/g; let m;
      while ((m = re.exec(t))) {
        for (const part of m[1].split(';')) {
          const p = part.trim();
          const rng = p.match(/^(\d{1,3})\s*[-–—]\s*(\d{1,3})$/);
          if (rng) { for (let k = +rng[1]; k <= +rng[2] && k - rng[1] < 50; k++) out.push(k); continue; }
          const num = p.match(/^(\d{1,3})\b/);
          if (num) out.push(+num[1]);
        }
      }
      return out;
    }

    // диапазон списка источников
    let start = -1, end = blocks.length;
    for (let i = 0; i < blocks.length; i++) {
      if (roles[i].role === 'structural' && roles[i].structKey === 'istochniki') { start = i; continue; }
      if (start >= 0 && roles[i].role === 'structural') { end = i; break; }
    }
    const limit = start >= 0 ? start : blocks.length;

    const citations = [];              // { num, block } – в порядке появления
    const urlBlocks = new Map();       // нормированный URL → первый блок
    for (let i = 0; i < limit; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const t = b.text || '';
      for (const num of parseCitations(t)) citations.push({ num, block: i });
      const ure = new RegExp(URL_RE.source, 'gi'); let um;
      while ((um = ure.exec(t))) { const k = normUrl(um[0]); if (!urlBlocks.has(k)) urlBlocks.set(k, i); }
    }
    const refs = new Set(citations.map((c) => c.num));

    if (refs.size && start < 0) {
      add({ category: 'refs', level: 'error', blockIndex: null, title: 'Нет списка использованных источников', detail: `В тексте есть ссылки [${[...refs].sort((a, b) => a - b).join(', ')}], но нет структурного элемента «СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ».`, recommend: 'Добавьте список использованных источников.', ref: 'СТУ 7.5–07–2021, п. 6.9, 7.10' });
    }
    if (start < 0) return;

    const entries = [];
    let n = 0, seqEntry = null, indentEntry = null;
    for (let i = start + 1; i < end; i++) {
      const b = blocks[i]; if (b.kind !== 'p') continue;
      const t = normText(b.text); if (!t) continue;
      let num = null;
      const m = t.match(/^(\d{1,3})[.)]?\s+/);
      if (b.numPr && b.numPr.numId) { n++; num = n; }
      else if (m) { num = +m[1]; if (num !== n + 1 && seqEntry == null) seqEntry = i; n = num; }
      const fl = b.ind ? b.ind.firstLine : (b.numPr ? 708 : null);
      const leftI = b.ind ? (b.ind.left || 0) : 0;
      if (!fl && !leftI && !b.numPr && indentEntry == null) indentEntry = i;
      const urls = t.match(new RegExp(URL_RE.source, 'gi')) || [];
      entries.push({ index: i, num: num, text: t, urls: urls });
    }
    if (seqEntry != null) add({ category: 'sources', level: 'warning', blockIndex: seqEntry, title: 'Нарушена нумерация записей', detail: 'Нумерация библиографических записей не сквозная или нарушена.', recommend: 'Пронумеруйте записи арабскими цифрами сквозной нумерацией.', ref: 'СТУ 7.5–07–2021, п. 7.10.2' });
    if (indentEntry != null) add({ category: 'sources', level: 'info', blockIndex: indentEntry, title: 'Запись без абзацного отступа', detail: 'Библиографическая запись напечатана без абзацного отступа.', recommend: 'Печатайте библиографические записи с абзацного отступа.', ref: 'СТУ 7.5–07–2021, п. 7.10.2' });
    // В конце библиографической записи ставят точку.
    for (const e of entries) {
      if (e.text && !/\.$/.test(e.text)) add({ category: 'sources', level: 'warning', blockIndex: e.index, title: 'Нет точки в конце записи источника', detail: `Запись «${e.text.slice(0, 60)}${e.text.length > 60 ? '…' : ''}» не заканчивается точкой.`, recommend: 'Заканчивайте библиографическую запись точкой.', ref: 'СТУ 7.5–07–2021, п. 7.10.2' });
    }
    checkBibliographyFormat(entries, start, add);

    // Ссылки внутри записей: метка чат-бота и дата обращения – с подсветкой URL.
    for (const e of entries) {
      const u = e.urls && e.urls[0];
      if (!u) continue;
      const b = blocks[e.index];
      if (/utm_source=(openai|chatgpt)/i.test(u)) add({ category: 'sources', level: 'warning', blockIndex: e.index, chars: charRangesFor(b, u), title: 'Ссылка с меткой чат-бота', detail: 'В адресе есть utm_source=openai / chatgpt.', recommend: 'Удалите UTM-метку из ссылки.', ref: 'СТУ 7.5–07–2021, п. 7.10' });
      if (!/дата\s+обращения/i.test(e.text)) add({ category: 'sources', level: 'warning', blockIndex: e.index, chars: charRangesFor(b, u), title: 'Ссылка без даты обращения', detail: `Запись «${e.text.slice(0, 60)}…» содержит адрес, но нет даты обращения.`, recommend: 'Для электронных ресурсов указывайте URL и дату обращения.', ref: 'СТУ 7.5–07–2021, п. 7.10.2' });
    }

    // --- порядок появления в тексте ---
    const firstSeen = []; const firstIdx = {};
    for (const c of citations) { if (firstIdx[c.num] == null) { firstIdx[c.num] = c.block; firstSeen.push(c.num); } }
    const systematic = firstSeen.length > 1 && firstSeen.every((v, idx) => v === idx + 1);

    // --- алфавитный порядок (с указанием нарушения) ---
    const stripNum = (s) => normText(s.replace(/^\d{1,3}\.?\s*/, '')).toLowerCase();
    const alphas = entries.map((e) => stripNum(e.text));
    let alphaViolation = -1;
    for (let i = 1; i < alphas.length; i++) { if (alphas[i - 1] > alphas[i]) { alphaViolation = i; break; } }
    const alphabetical = alphas.length > 1 && alphaViolation < 0;

    // --- хронологический порядок ---
    const years = entries.map((e) => { const y = e.text.match(/(19|20)\d{2}/); return y ? +y[0] : null; });
    let chronological = years.length > 1 && years.every((y) => y != null);
    if (chronological) for (let i = 1; i < years.length; i++) if (years[i] < years[i - 1]) { chronological = false; break; }

    let method = null;
    if (entries.length >= 2) method = systematic ? 'в порядке появления в тексте' : (alphabetical ? 'по алфавиту' : (chronological ? 'по годам' : null));

    // Конкретные нарушения порядка появления (даже если выбран другой способ – информативно).
    if (firstSeen.length > 1) {
      let expected = 1; const viol = [];
      for (const v of firstSeen) {
        if (v === expected) { expected++; continue; }
        if (v > expected) { viol.push({ appeared: v, before: expected }); expected = v + 1; }
      }
      for (let k = 0; k < viol.length && k < 6; k++) {
        const v = viol[k];
        const lvl = method ? 'info' : 'warning';
        add({
          category: 'sources', level: lvl, blockIndex: firstIdx[v.appeared],
          title: 'Порядок появления источников нарушен',
          detail: `Источник [${v.appeared}] впервые упомянут в тексте раньше источника [${v.before}].`,
          recommend: 'Расположите источники в порядке первого упоминания в тексте – тогда номера ссылок возрастают по ходу текста. Либо выдерживайте алфавитный или хронологический порядок.',
          ref: 'СТУ 7.5–07–2021, п. 7.10.3',
        });
      }
    }

    if (!method) {
      add({
        category: 'sources', level: 'warning', blockIndex: null,
        title: 'Записи не упорядочены',
        detail: 'Записи не упорядочены ни по появлению в тексте, ни по алфавиту, ни по годам.',
        recommend: 'Упорядочьте записи любым одним способом: по появлению в тексте, по алфавиту или по годам.',
        ref: 'СТУ 7.5–07–2021, п. 7.10.3',
      });
    }

    // ссылка на отсутствующий источник – у места ссылки
    const count = entries.length;
    const seenMissing = new Set();
    for (const c of citations) {
      if (count > 0 && c.num > count && !seenMissing.has(c.num)) {
        seenMissing.add(c.num);
        add({ category: 'sources', level: 'warning', blockIndex: c.block, title: 'Ссылка на отсутствующий источник', detail: `В тексте ссылка [${c.num}], а в списке ${count} записей.`, recommend: 'Проверьте номера ссылок и список источников.', ref: 'СТУ 7.5–07–2021, п. 7.9.4' });
      }
    }

    // источники без ссылок в тексте – у конкретной записи
    const entryCited = (e) => {
      if (e.num != null && refs.has(e.num)) return true;
      for (const u of (e.urls || [])) if (urlBlocks.has(normUrl(u))) return true;
      return false;
    };
    for (const e of entries) {
      if (!entryCited(e)) add({ category: 'sources', level: 'info', blockIndex: e.index, title: 'Источник без ссылки в тексте', detail: `Не найдена ссылка на «${e.text.slice(0, 60)}…».`, recommend: 'На каждый источник должна быть ссылка в тексте.', ref: 'СТУ 7.5–07–2021, п. 6.9, 7.9.4' });
    }
  }

  // --------------------------------------------------------------- appendices
  function checkAppendices(blocks, roles, forcedPageStarts, add) {
    const seen = new Set();
    for (let i = 0; i < blocks.length; i++) {
      if (roles[i].role !== 'structural' || roles[i].structKey !== 'prilozheniya') continue;
      const b = blocks[i];
      const t = normText(b.text);
      const m = t.toUpperCase().match(/^ПРИЛОЖЕНИЕ\s+([А-ЯЁ])/);
      if (m) {
        const letter = m[1];
        if (!MAX_APPENDIX_LETTERS.includes(letter)) add({ category: 'appendix', level: 'error', blockIndex: i, title: 'Недопустимая буква приложения', detail: `Приложение обозначено буквой «${letter}».`, recommend: 'Приложения обозначайте прописными буквами русского алфавита, начиная с А, кроме Ё, З, Й, О, Ч, Ь, Ы, Ъ.', ref: 'СТУ 7.5–07–2021, п. 7.11.2' });
        if (seen.has(letter)) add({ category: 'appendix', level: 'warning', blockIndex: i, title: 'Повтор обозначения приложения', detail: `Буква «${letter}» используется повторно.`, recommend: 'Каждое приложение обозначайте отдельной буквой.', ref: 'СТУ 7.5–07–2021, п. 7.11.2' });
        seen.add(letter);
      }
      if (i > 0 && !forcedPageStarts.has(i)) add({ category: 'appendix', level: 'warning', blockIndex: i, title: 'Приложение без новой страницы', detail: `«${t}» не начинается с новой страницы.`, recommend: 'Каждое приложение начинайте с новой страницы.', ref: 'СТУ 7.5–07–2021, п. 7.11.2' });
      if (!isBoldPara(b)) add({ category: 'appendix', level: 'warning', blockIndex: i, title: 'Заголовок приложения не полужирный', detail: `«${t}» не выделен полужирным.`, recommend: 'Заголовок приложения записывайте прописной буквы полужирным шрифтом, симметрично тексту.', ref: 'СТУ 7.5–07–2021, п. 7.11.3' });
    }
  }

  // --------------------------------------------------------------- title page
  // Титульный лист не проверяем: он служит только для определения вида работы.

  // ------------------------------------------------------------ содержание
  // Отступы по уровням (п. 6.4.4) и присутствие разделов в содержании (п. 6.4.2).
  function parseToc(blocks, roles, add) {
    let start = -1, end = blocks.length;
    for (let i = 0; i < blocks.length; i++) {
      if (roles[i].role === 'structural' && roles[i].structKey === 'soderzhanie') { start = i; continue; }
      if (start >= 0 && roles[i].role === 'structural') { end = i; break; }
    }
    const entries = [];
    if (start < 0) return entries;
    for (let i = start + 1; i < end; i++) {
      if (roles[i].role !== 'toc') continue;
      const b = blocks[i]; const t = normText(b.text); if (!t) continue;
      const numM = t.match(/^(\d+(?:\.\d+)*)/);
      const number = numM ? numM[1] : null;
      const depth = number ? number.split('.').length : 1;
      const pgM = t.match(/(\d{1,4})\s*$/);
      const page = pgM ? +pgM[1] : null;
      let title = t;
      if (number) title = title.slice(number.length);
      title = title.replace(/[\t]/g, ' ').replace(/[\s.\u2026]+\d{0,4}\s*$/, '').trim();
      const fl = b.ind ? b.ind.firstLine : null;
      entries.push({ index: i, number: number, depth: depth, page: page, title: title, firstLine: fl });
    }
    // В содержании заголовки записывают с первой прописной буквы, остальные строчные.
    for (const e of entries) {
      const t = e.title || '';
      if (t.length > 3 && /[А-ЯЁ]/.test(t) && !/[а-яё]/.test(t)) add({
        category: 'structure', level: 'warning', blockIndex: e.index,
        title: 'Заголовок в содержании прописными буквами',
        detail: `«${t}» набран прописными буквами.`,
        recommend: 'В содержании заголовки записывают строчными буквами, с первой прописной (аббревиатуры сохраняют написание).',
        ref: 'СТУ 7.5–07–2021, п. 6.4.2',
      });
    }
    // отступы по уровням: подраздел на два знака глубже раздела, пункт ещё на два
    for (const e of entries) {
      if (e.depth < 2) continue;
      const expected = (e.depth - 1) * 200;   // примерно два знака на уровень
      const fl = e.firstLine == null ? 0 : e.firstLine;
      if (fl < expected * 0.4) add({ category: 'structure', level: 'warning', blockIndex: e.index, title: 'Нет отступа в содержании', detail: `Строка «${e.title}» уровня ${e.depth} в содержании без отступа.`, recommend: 'Номера и заголовки подразделов приводят после абзацного отступа (два знака), пунктов ещё на два знака глубже.', ref: 'СТУ 7.5–07–2021, п. 6.4.4' });
    }
    // присутствие разделов в содержании
    const tocNums = new Set(entries.map((e) => e.number).filter(Boolean));
    const tocTitles = new Set(entries.map((e) => e.title.toUpperCase()));
    const KEY = { referat: 1, vvedenie: 1, zaklyuchenie: 1, istochniki: 1, prilozheniya: 1 };
    for (let i = 0; i < blocks.length; i++) {
      const r = roles[i];
      if (r.role === 'heading' && r.number && !tocNums.has(r.number)) add({ category: 'structure', level: 'warning', blockIndex: i, title: 'Раздел не указан в содержании', detail: `Заголовок «${r.number} ${r.title}» отсутствует в содержании.`, recommend: 'Включите раздел в содержание.', ref: 'СТУ 7.5–07–2021, п. 6.4.2' });
      if (r.role === 'structural' && KEY[r.structKey] && !tocTitles.has(normText(blocks[i].text).toUpperCase())) add({ category: 'structure', level: 'info', blockIndex: i, title: 'Раздел не указан в содержании', detail: `«${normText(blocks[i].text)}» отсутствует в содержании.`, recommend: 'Включите элемент в содержание.', ref: 'СТУ 7.5–07–2021, п. 6.4.2' });
    }
    return entries;
  }

  // Сверка номеров страниц из содержания с реальным началом раздела (по развёртке).
  function crossCheckToc(doc, roles, tocEntries, pageForBlock, add) {
    if (!tocEntries || !tocEntries.length) return;
    const blocks = doc.blocks;
    const byNum = {}, byTitle = {};
    for (let i = 0; i < blocks.length; i++) {
      const r = roles[i];
      if (r.role === 'heading' && r.number) byNum[r.number] = i;
      if (r.role === 'structural') byTitle[normText(blocks[i].text).toUpperCase()] = i;
    }
    for (const e of tocEntries) {
      if (!e.page) continue;
      let bi = null;
      if (e.number != null && byNum[e.number] != null) bi = byNum[e.number];
      else if (e.title && byTitle[e.title.toUpperCase()] != null) bi = byTitle[e.title.toUpperCase()];
      if (bi == null) continue;
      const p = pageForBlock[bi];
      if (p == null) continue;
      const real = p + 1;
      if (real !== e.page) add({ category: 'structure', level: 'warning', blockIndex: e.index, title: 'Страница в содержании не совпадает', detail: `В содержании у «${e.title}» указана страница ${e.page}, а раздел начинается на странице ${real}.`, recommend: 'Обновите номера страниц в содержании (обновите поле оглавления в Word).', ref: 'СТУ 7.5–07–2021, п. 6.4.2' });
    }
  }

  global.STURules = { runChecks, classify, DOC_REQUIREMENTS, STRUCTURAL, computeForcedPageStarts, detectDocType, crossCheckToc, parseToc };

  // Автоопределение вида документа по титульному листу и тексту.
  function detectDocType(doc) {
    const blocks = doc.blocks || [];
    let head = '';
    for (let i = 0; i < Math.min(blocks.length, 60); i++) head += ' ' + (blocks[i].text || '');
    const up = head.toUpperCase().replace(/\s+/g, ' ');
    const RULES = [
      [/ЛЕТН[А-ЯЁ]*[^.]{0,25}ПРАКТИК/, 'praktika_letnyaya'],
      [/ОТЧ[ЕЁ]Т[^.]{0,40}ПРАКТИК/, 'praktika'],
      [/(УЧЕБН|ПРОИЗВОДСТВЕНН|ПРЕДДИПЛОМН)[А-ЯЁ]*[^.]{0,25}ПРАКТИК/, 'praktika'],
      [/ПРАКТИЧЕСКАЯ РАБОТА|ЛАБОРАТОРНАЯ РАБОТА|ВЫПОЛНЕНИИ ПРОЕКТА|НАУЧНО-ИССЛЕДОВАТЕЛЬСК/, 'otchet'],
      [/КУРСОВОЙ ПРОЕКТ|КУРСОВАЯ РАБОТА|КУРСОВОЙ РАБОТЫ/, 'kursovaya'],
      [/ДИПЛОМНЫЙ ПРОЕКТ|ДИПЛОМНАЯ РАБОТА|ДИПЛОМНОЙ РАБОТЫ/, 'diplom'],
      [/МАГИСТЕРСКАЯ ДИССЕРТАЦИЯ|МАГИСТЕРСКОЙ ДИССЕРТАЦИИ/, 'magister'],
      [/БАКАЛАВРСКАЯ РАБОТА|ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ/, 'vkr'],
      [/КОНТРОЛЬНАЯ РАБОТА|КОНТРОЛЬНОЙ РАБОТЫ/, 'kontrolnaya'],
      [/РАСЧ[ЕЁ]ТНО-ГРАФИЧЕСК|РАСЧ[ЕЁ]ТНОЕ ЗАДАНИЕ/, 'rgz'],
      [/(^|[\s«"(,.:;])ЭССЕ([\s»"),.:;]|$)/, 'esse'],
      [/(^|[\s«"(,.:;])РЕФЕРАТ([\s»"),.:;]|$)/, 'referat'],
    ];
    for (const [re, type] of RULES) if (re.test(up)) return { type: type, confidence: 'high' };
    if (/ПРАКТИК/.test(up)) return { type: 'praktika', confidence: 'low' };
    return null;
  }
})(window);

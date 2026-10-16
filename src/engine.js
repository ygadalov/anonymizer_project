/*
 * Ядро обезличивания Redact: разбор листа, определение типов столбцов и замены.
 * Не зависит от DOM — работает и в браузере, и в Node (для тестов).
 */
(function (root) {
  'use strict';

  // ---------- Детерминированный ГПСЧ (ключ + значение → одинаковая замена) ----------
  function cyrb128(str) {
    let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
    for (let i = 0, k; i < str.length; i++) {
      k = str.charCodeAt(i);
      h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
      h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
      h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
      h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
    }
    h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
    h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
    h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
    h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
    h1 ^= (h2 ^ h3 ^ h4); h2 ^= h1; h3 ^= h1; h4 ^= h1;
    return [h1 >>> 0, h2 >>> 0, h3 >>> 0, h4 >>> 0];
  }
  function sfc32(a, b, c, d) {
    return function () {
      a |= 0; b |= 0; c |= 0; d |= 0;
      const t = (((a + b) | 0) + d) | 0;
      d = (d + 1) | 0;
      a = b ^ (b >>> 9);
      b = (c + (c << 3)) | 0;
      c = (c << 21) | (c >>> 11);
      c = (c + t) | 0;
      return (t >>> 0) / 4294967296;
    };
  }
  function makeRng(seed, ...parts) {
    const r = sfc32(...cyrb128(String(seed) + '␟' + parts.join('␟')));
    for (let i = 0; i < 12; i++) r();
    return r;
  }
  const pick = (rnd, arr) => arr[Math.floor(rnd() * arr.length)];
  const randInt = (rnd, lo, hi) => lo + Math.floor(rnd() * (hi - lo + 1));

  function randomSeed() {
    const abc = 'abcdefghjkmnpqrstuvwxyz23456789';
    let s = '';
    const buf = new Uint32Array(16);
    (root.crypto || globalThis.crypto).getRandomValues(buf);
    for (const x of buf) s += abc[x % abc.length];
    return s;
  }
  function factorFromSeed(seed) {
    const r = makeRng(seed, 'factor');
    // 0.55–0.85 или 1.15–1.45: заметно отличается от 1
    const k = r() < 0.5 ? 0.55 + r() * 0.3 : 1.15 + r() * 0.3;
    return Math.round(k * 10000) / 10000;
  }
  function shiftFromSeed(seed) {
    const r = makeRng(seed, 'shift');
    const d = randInt(r, 45, 400);
    return r() < 0.5 ? -d : d;
  }

  // ---------- Значения ячеек ----------
  const isEmpty = v => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const isDate = v => v instanceof Date && !isNaN(v);

  // Текстовые даты: 30.09.2026, 30/09/2026 23:59:59, 2026-09-30
  const RE_DMY = /^(\d{1,2})([./-])(\d{1,2})\2(\d{4}|\d{2})(\s+\d{1,2}:\d{2}(?::\d{2})?)?$/;
  const RE_YMD = /^(\d{4})-(\d{2})-(\d{2})([T\s]\d{1,2}:\d{2}(?::\d{2})?)?$/;
  function parseTextDate(s) {
    if (typeof s !== 'string') return null;
    const t = s.trim();
    let m = RE_DMY.exec(t);
    if (m) {
      const d = +m[1], mo = +m[3];
      let y = +m[4];
      if (m[4].length === 2) y += y < 50 ? 2000 : 1900;
      if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
      return { kind: 'dmy', d, m: mo, y, sep: m[2], dPad: m[1].length === 2, mPad: m[3].length === 2, yLen: m[4].length, rest: m[5] || '' };
    }
    m = RE_YMD.exec(t);
    if (m) {
      const y = +m[1], mo = +m[2], d = +m[3];
      if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
      return { kind: 'ymd', d, m: mo, y, rest: m[4] || '' };
    }
    return null;
  }
  const pad2 = n => String(n).padStart(2, '0');
  function formatTextDate(info, ms, keepTime = true) {
    const dt = new Date(ms);
    const d = dt.getUTCDate(), m = dt.getUTCMonth() + 1, y = dt.getUTCFullYear();
    const rest = keepTime ? info.rest : '';
    if (info.kind === 'ymd') return `${y}-${pad2(m)}-${pad2(d)}${rest}`;
    const ds = info.dPad ? pad2(d) : String(d);
    const ms_ = info.mPad ? pad2(m) : String(m);
    const ys = info.yLen === 2 ? pad2(y % 100) : String(y);
    return `${ds}${info.sep}${ms_}${info.sep}${ys}${rest}`;
  }
  const infoToMs = i => Date.UTC(i.y, i.m - 1, i.d);
  // Дата из значения ячейки → миллисекунды UTC (без времени для текстовых дат)
  function dateMs(v) {
    if (isDate(v)) return v.getTime();
    const i = parseTextDate(v);
    return i ? infoToMs(i) : null;
  }

  // Коды и идентификаторы: ИНН, «00/001455», «ЗУП002117», «76.02.1»
  const RE_ID = /^[A-Za-zА-Яа-яЁё]{0,5}[-/.]?\d[\d/.\-]*$/;
  const isIdLike = v => (typeof v === 'string' && v.length >= 3 && v.length <= 30 && RE_ID.test(v.trim()));

  // Склонение по числу: plural(5, 'ячейка', 'ячейки', 'ячеек') → 'ячеек'
  function plural(n, one, few, many) {
    const a = Math.abs(n) % 100, b = a % 10;
    if (a > 10 && a < 20) return many;
    if (b === 1) return one;
    if (b >= 2 && b <= 4) return few;
    return many;
  }

  // ---------- Чтение листа ----------
  function plainValue(v) {
    if (v === null || v === undefined) return { v: null, f: false };
    if (typeof v !== 'object' || isDate(v)) return { v, f: false };
    if (v.formula !== undefined || v.sharedFormula !== undefined) {
      const r = v.result;
      return { v: (r && typeof r === 'object' && !isDate(r)) ? null : (r === undefined ? null : r), f: true };
    }
    if (Array.isArray(v.richText)) return { v: v.richText.map(t => t.text).join(''), f: false, rich: true };
    if (v.text !== undefined) return { v: typeof v.text === 'object' ? plainValue(v.text).v : v.text, f: false, rich: true };
    if (v.error !== undefined) return { v: null, f: false };
    return { v: null, f: false };
  }

  function gridFromWorksheet(ws) {
    const maxRow = ws.rowCount, maxCol = ws.columnCount;
    const rows = new Array(maxRow + 1);
    for (let r = 1; r <= maxRow; r++) {
      const row = ws.findRow(r);
      const arr = new Array(maxCol + 1).fill(null);
      if (row) {
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if (cell.isMerged && cell.master !== cell) { arr[c] = { v: null, slave: true }; return; }
          const p = plainValue(cell.value);
          if (p.v !== null || p.f) arr[c] = p;
        });
      }
      rows[r] = arr;
    }
    return { rows, maxRow, maxCol };
  }

  // ---------- CSV ----------
  function detectDelimiter(text) {
    const lines = text.split(/\r?\n/).slice(0, 20);
    let best = ',', bestScore = -1;
    for (const d of [';', ',', '\t', '|']) {
      const counts = lines.map(l => l.split(d).length - 1).filter(n => n > 0);
      const score = counts.length ? counts.reduce((a, b) => a + b, 0) / lines.length : 0;
      if (score > bestScore) { bestScore = score; best = d; }
    }
    return best;
  }
  function parseCsv(text) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const delim = detectDelimiter(text);
    const out = [];
    let row = [], field = '', q = false, quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
        else field += ch;
      } else if (ch === '"' && field === '') { q = true; quoted = true; }
      else if (ch === delim) { row.push({ s: field, quoted }); field = ''; quoted = false; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push({ s: field, quoted }); out.push(row); row = []; field = ''; quoted = false;
      } else field += ch;
    }
    if (field !== '' || row.length) { row.push({ s: field, quoted }); out.push(row); }
    // Числа «1234,56» / «1234.56» превращаем в number, запоминаем десятичный разделитель
    let decSep = delim === ',' ? ',' : null;
    const maxCol = Math.max(1, ...out.map(r => r.length));
    const rows = [null];
    for (const r of out) {
      const arr = new Array(maxCol + 1).fill(null);
      r.forEach((cell, i) => {
        const s = cell.s;
        if (s === '') return;
        const m = /^(-?\d+)(?:([.,])(\d+))?$/.exec(s.trim());
        if (m && !(m[1].replace('-', '').length > 1 && m[1].replace('-', '')[0] === '0' && !m[2])) {
          if (m[2]) { if (!decSep) decSep = m[2]; if (m[2] !== decSep) { arr[i + 1] = { v: s }; return; } }
          arr[i + 1] = { v: Number(m[1] + (m[3] ? '.' + m[3] : '')), num: true };
        } else arr[i + 1] = { v: s };
      });
      rows.push(arr);
    }
    return { grid: { rows, maxRow: rows.length - 1, maxCol }, delim, decSep: decSep || '.' };
  }
  function gridToCsv(grid, changes, delim, decSep) {
    const lines = [];
    for (let r = 1; r <= grid.maxRow; r++) {
      const cells = [];
      let last = 0;
      for (let c = 1; c <= grid.maxCol; c++) if (grid.rows[r][c] || changes.has(r + ':' + c)) last = c;
      for (let c = 1; c <= last; c++) {
        const key = r + ':' + c;
        let v = changes.has(key) ? changes.get(key) : (grid.rows[r][c] ? grid.rows[r][c].v : null);
        if (v === null || v === undefined) { cells.push(''); continue; }
        if (typeof v === 'number') v = String(v).replace('.', decSep);
        else if (isDate(v)) v = v.toISOString().slice(0, 10);
        v = String(v);
        if (v.includes(delim) || /["\r\n]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
        cells.push(v);
      }
      lines.push(cells.join(delim));
    }
    return '﻿' + lines.join('\r\n') + '\r\n';
  }

  // ---------- Анализ структуры ----------
  const RE_TOTAL = /^\s*(итого|всего|total|grand total)(?!\p{L})/iu;

  function detectHeaderRow(grid) {
    let best = 1, bestScore = -1;
    const lim = Math.min(grid.maxRow, 30);
    for (let r = 1; r <= lim; r++) {
      const row = grid.rows[r];
      let strings = 0, filled = 0;
      for (let c = 1; c <= grid.maxCol; c++) {
        const cell = row[c];
        if (!cell || isEmpty(cell.v)) continue;
        filled++;
        if (typeof cell.v === 'string' && !parseTextDate(cell.v)) strings++;
      }
      // Заголовок — строка с наибольшим числом текстовых ячеек, причём их несколько
      const score = filled >= 2 ? strings * 10 + filled : 0;
      if (score > bestScore) { bestScore = score; best = r; }
    }
    return best;
  }

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function analyze(grid, headerRow) {
    const dataRows = [], totalsRows = new Set();
    for (let r = headerRow + 1; r <= grid.maxRow; r++) {
      const row = grid.rows[r];
      let any = false, total = false;
      for (let c = 1; c <= grid.maxCol; c++) {
        const cell = row[c];
        if (!cell || isEmpty(cell.v)) continue;
        any = true;
        if (c <= 3 && typeof cell.v === 'string' && RE_TOTAL.test(cell.v)) total = true;
      }
      if (!any) continue;
      dataRows.push(r);
      if (total) totalsRows.add(r);
    }
    const columns = [];
    for (let c = 1; c <= grid.maxCol; c++) {
      const h = grid.rows[headerRow] && grid.rows[headerRow][c];
      let slaveCount = 0, nonEmpty = 0, nNum = 0, nDate = 0, nId = 0, nInt = 0;
      const uniq = new Set(), samples = [];
      let min = Infinity, max = -Infinity, dmin = Infinity, dmax = -Infinity;
      for (const r of dataRows) {
        if (totalsRows.has(r)) continue;
        const cell = grid.rows[r][c];
        if (cell && cell.slave) { slaveCount++; continue; }
        if (!cell || isEmpty(cell.v)) continue;
        const v = cell.v;
        nonEmpty++;
        const key = isDate(v) ? 'd' + v.getTime() : typeof v + ':' + v;
        if (!uniq.has(key)) { uniq.add(key); if (samples.length < 4) samples.push(v); }
        if (typeof v === 'number') {
          nNum++; if (Number.isInteger(v)) nInt++;
          if (v < min) min = v; if (v > max) max = v;
        } else {
          const ms = dateMs(v);
          if (ms !== null) { nDate++; if (ms < dmin) dmin = ms; if (ms > dmax) dmax = ms; }
          else if (isIdLike(v)) nId++;
        }
      }
      const headerText = h && !isEmpty(h.v) ? String(isDate(h.v) ? h.v.toISOString().slice(0, 10) : h.v).trim() : '';
      if (!headerText && nonEmpty === 0) continue; // пустой столбец или вторая половина объединённых ячеек
      let type = 'text';
      const share = x => nonEmpty ? x / nonEmpty : 0;
      if (share(nNum) >= 0.9) type = 'number';
      else if (share(nDate) >= 0.9) type = 'date';
      else if (share(nId) >= 0.9) type = 'id';
      columns.push({
        c, letter: colLetter(c), name: headerText || 'Столбец ' + colLetter(c),
        detected: type, nonEmpty, uniques: uniq.size, samples,
        isInt: nNum > 0 && nInt === nNum, min, max, dmin, dmax,
      });
    }
    return { headerRow, dataRows, totalsRows, columns };
  }

  // ---------- Методы ----------
  const METHODS = {
    text: [
      ['keep', 'Оставить'],
      ['pseudo', 'Псевдоним с номером'],
      ['fake', 'Вымышленное название'],
      ['mask', 'Маска символов'],
      ['smart', 'Замена внутри текста'],
      ['hash', 'Хэш-код'],
      ['clear', 'Очистить'],
    ],
    id: [
      ['keep', 'Оставить'],
      ['digits', 'Случайные цифры'],
      ['mask', 'Маска символов'],
      ['pseudo', 'Псевдоним с номером'],
      ['hash', 'Хэш-код'],
      ['clear', 'Очистить'],
    ],
    number: [
      ['keep', 'Оставить'],
      ['scale', 'Умножить на общий K'],
      ['noise', 'Случайный шум ±%'],
      ['round', 'Округлить до шага'],
      ['random', 'Случайное в диапазоне столбца'],
      ['clear', 'Очистить'],
    ],
    date: [
      ['keep', 'Оставить'],
      ['shift', 'Сдвинуть на общий D дней'],
      ['month', 'Обобщить до месяца (1-е число)'],
      ['drandom', 'Случайная дата в диапазоне столбца'],
      ['clear', 'Очистить'],
    ],
  };
  // Подсказки к методам (всплывают при наведении на список)
  const METHOD_HINTS = {
    keep: 'Значения не меняются',
    pseudo: 'ОПТТРЕЙД ООО → Контрагент 017 ООО; одинаковые значения получают одинаковый номер',
    fake: 'ОПТТРЕЙД ООО → ТРАНСГАРАНТ ООО',
    mask: 'Каждая буква → случайная буква, цифра → цифра; длина и знаки сохраняются',
    smart: 'В свободном тексте заменяются названия и коды из других столбцов, даты сдвигаются на D, номера из 3+ цифр меняются',
    hash: 'Короткий код вида ID-3K9F2A7B',
    digits: '000001658 → 000003708: цифры случайные, длина, разделители и ведущие нули сохраняются',
    scale: 'Все числа умножаются на одно и то же K — доли и суммы сохраняются',
    noise: 'Каждое число меняется на случайные ±N %',
    round: 'Округление до шага, например до 1000',
    random: 'Случайное число между минимумом и максимумом столбца',
    shift: 'Все даты сдвигаются на одно и то же число дней D — интервалы сохраняются',
    month: '15.03.2024 → 01.03.2024',
    drandom: 'Случайная дата между минимальной и максимальной датой столбца',
    clear: 'Ячейка становится пустой',
  };
  const TYPE_NAMES = { text: 'Текст', id: 'Код / ИНН', number: 'Число', date: 'Дата' };
  const DICT_METHODS = new Set(['pseudo', 'fake', 'mask', 'digits', 'hash']);

  function defaultPrefix(name) {
    if (/контрагент|counterpart/i.test(name)) return 'Контрагент';
    if (/покупател|клиент|customer|client/i.test(name)) return 'Клиент';
    if (/поставщик|vendor|supplier/i.test(name)) return 'Поставщик';
    if (/организац|компан|company/i.test(name)) return 'Организация';
    if (/фио|сотрудник|employee|ответствен|менеджер/i.test(name)) return 'Сотрудник';
    return name.replace(/\s+/g, ' ').slice(0, 24);
  }

  // Предложение метода по названию столбца
  function suggest(col) {
    const n = col.name.toLowerCase();
    const base = { prefix: defaultPrefix(col.name), group: col.letter, pct: 10, step: 1000 };
    if (/инн|кпп|огрн|окпо|снилс|inn\b|tax id|vat|р\/с|расч[её]тн.*сч[её]т|бик|iban|swift|телефон|e-?mail|паспорт/.test(n))
      return { ...base, type: col.detected === 'text' ? 'id' : col.detected, method: col.detected === 'number' ? 'keep' : 'digits' };
    if (/^код\s+(контрагент|клиент|покупател|поставщик)/.test(n))
      return { ...base, type: 'id', method: 'digits' };
    if (/контрагент|покупател|поставщик|клиент|дебитор|кредитор|организац|компан|наименован|фио|сотрудник|counterpart|customer|client|company|vendor|supplier/.test(n) && col.detected !== 'number' && col.detected !== 'date')
      return { ...base, type: 'text', method: 'pseudo' };
    if (/^(договор|документ|основание|комментар|описание|назначение)/.test(n) && col.detected === 'text')
      return { ...base, type: 'text', method: 'smart' };
    return { ...base, type: col.detected, method: 'keep' };
  }

  // ---------- Генераторы замен ----------
  const LEGAL_FORMS = new Set(['ООО', 'ОАО', 'ЗАО', 'АО', 'ПАО', 'НАО', 'ИП', 'АНО', 'ФГУП', 'ГУП', 'МУП', 'ФГБУ', 'ГБУ', 'МБУ',
    'ТОО', 'ЧУП', 'СПК', 'КФХ', 'ГК', 'ТД', 'НКО', 'LLC', 'LTD', 'LTD.', 'GMBH', 'OU', 'OÜ', 'AG', 'SA', 'S.A.', 'INC', 'INC.',
    'LLP', 'PLC', 'BV', 'B.V.', 'AS', 'SIA', 'UAB', 'SRL', 'SPA', 'CO', 'CORP', 'КО']);
  function splitLegalForm(name) {
    const tokens = String(name).trim().split(/\s+/);
    const clean = t => t.replace(/[«»"'(),]/g, '').toUpperCase();
    const lead = [], tail = [];
    while (tokens.length > 1 && LEGAL_FORMS.has(clean(tokens[0]))) lead.push(clean(tokens.shift()));
    while (tokens.length > 1 && LEGAL_FORMS.has(clean(tokens[tokens.length - 1]))) tail.unshift(clean(tokens.pop()));
    // ОПФ в середине («СИНЕРГИЯ ООО ТЗК») переносим в конец
    if (!lead.length && !tail.length) {
      const i = tokens.findIndex((t, k) => k > 0 && LEGAL_FORMS.has(clean(t)));
      if (i > 0) tail.push(clean(tokens[i]));
    }
    return { lead: lead.join(' '), tail: tail.join(' '), core: tokens.join(' ') };
  }
  function withForm(orig, label, keepForm) {
    if (!keepForm) return label;
    const f = splitLegalForm(orig);
    return [f.lead, label, f.tail].filter(Boolean).join(' ');
  }
  const isUpper = s => /\p{L}/u.test(s) && s === s.toUpperCase();

  const FAKE_A = ['Агро', 'Транс', 'Пром', 'Строй', 'Энерго', 'Техно', 'Инвест', 'Логист', 'Металл', 'Нефте', 'Зерно', 'Хим', 'Эко', 'Гео', 'Евро', 'Сиб', 'Урал', 'Волга', 'Дон', 'Нева', 'Альфа', 'Бета', 'Вега', 'Север', 'Юг'];
  const FAKE_B = ['Альянс', 'Вектор', 'Гранит', 'Импульс', 'Ресурс', 'Стандарт', 'Терминал', 'Сервис', 'Трейд', 'Капитал', 'Гарант', 'Союз', 'Регион', 'Маршрут', 'Поток', 'Партнёр', 'Холдинг', 'Систем', 'Снаб', 'Комплекс', 'Профи', 'Лайн', 'Групп', 'Маркет', 'Форвард'];

  const LAT_U = 'ABCDEFGHJKLMNPQRSTUVWXYZ', LAT_L = 'abcdefghjkmnpqrstuvwxyz';
  const CYR_U = 'АБВГДЕЖЗИКЛМНОПРСТУФХЦШЭЮЯ', CYR_L = 'абвгдежзиклмнопрстуфхцшэюя';
  function maskString(s, rnd, lettersToo) {
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch >= '0' && ch <= '9') {
        // ведущие нули сохраняются, первая значащая цифра остаётся ненулевой — формат кода не меняется
        let j = i - 1;
        while (j >= 0 && s[j] === '0') j--;
        const leading = j < 0 || !/\d/.test(s[j]);
        if (leading && ch === '0') out += '0';
        else out += leading ? String(randInt(rnd, 1, 9)) : String(randInt(rnd, 0, 9));
      } else if (lettersToo && /[A-Z]/.test(ch)) out += pick(rnd, LAT_U);
      else if (lettersToo && /[a-z]/.test(ch)) out += pick(rnd, LAT_L);
      else if (lettersToo && /[А-ЯЁ]/.test(ch)) out += pick(rnd, CYR_U);
      else if (lettersToo && /[а-яё]/.test(ch)) out += pick(rnd, CYR_L);
      else out += ch;
    }
    return out;
  }
  function shortHash(seed, group, v) {
    const h = cyrb128(seed + '␟hash␟' + group + '␟' + v);
    return 'ID-' + (h[0].toString(36).toUpperCase() + h[1].toString(36).toUpperCase()).slice(0, 8).padStart(8, '0');
  }
  const escapeRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  // ---------- Преобразование ----------
  /*
   * config = { seed, factor, shiftDays, keepForm, keepTotals,
   *            columns: { [c]: { type, method, prefix, pct, step, group } } }
   * Возвращает { changes: Map('r:c' → новое значение), mapping: [...], stats, warnings }
   */
  function transform(grid, an, config) {
    const seed = config.seed;
    const K = Number(config.factor) || 1;
    const D = Math.round(Number(config.shiftDays) || 0);
    const DAY = 86400000;
    const active = an.columns
      .map(col => ({ col, cfg: config.columns[col.c] }))
      .filter(x => x.cfg && x.cfg.method && x.cfg.method !== 'keep');
    const changes = new Map();
    const mapping = [];
    const warnings = [];
    const leaks = [];      // исходные значения остаются в файле — показываются как ошибка
    const skipRow = r => config.keepTotals && an.totalsRows.has(r);
    const cellAt = (r, c) => { const x = grid.rows[r][c]; return (x && !x.slave && !isEmpty(x.v)) ? x : null; };
    const asText = v => isDate(v) ? v.toISOString().slice(0, 10) : String(v);

    // 1. Словари для текстовых/кодовых методов (общие для столбцов с одинаковым «словарём»)
    const groups = new Map();
    for (const { col, cfg } of active) {
      if (!DICT_METHODS.has(cfg.method)) continue;
      const gkey = cfg.method + '␟' + (String(cfg.group || '').trim() || col.letter);
      if (!groups.has(gkey)) groups.set(gkey, { method: cfg.method, name: String(cfg.group || '').trim() || col.letter, prefix: cfg.prefix, cols: [], values: new Set(), map: new Map() });
      const g = groups.get(gkey);
      g.cols.push(col);
      for (const r of an.dataRows) {
        if (skipRow(r)) continue;
        const cell = cellAt(r, col.c);
        if (cell && !isDate(cell.v)) g.values.add(asText(cell.v).trim());
      }
    }
    for (const g of groups.values()) {
      const used = new Set();
      const vals = [...g.values];
      if (g.method === 'pseudo') {
        // Номер зависит от ключа, а не от порядка строк — по номеру нельзя восстановить сортировку
        vals.sort((a, b) => {
          const ha = cyrb128(seed + '␟order␟' + g.name + '␟' + a)[0];
          const hb = cyrb128(seed + '␟order␟' + g.name + '␟' + b)[0];
          return ha - hb || (a < b ? -1 : 1);
        });
        const width = Math.max(3, String(vals.length).length);
        const prefix = (g.prefix || '').trim() || 'Объект';
        vals.forEach((v, i) => g.map.set(v, withForm(v, `${prefix} ${String(i + 1).padStart(width, '0')}`, config.keepForm)));
      } else {
        for (const v of vals) {
          let out = null;
          for (let attempt = 0; attempt < 50 && (out === null || used.has(out) || out === v); attempt++) {
            const rnd = makeRng(seed, g.method, g.name, v, attempt);
            if (g.method === 'fake') {
              let name = pick(rnd, FAKE_A) + pick(rnd, FAKE_B);
              if (attempt >= 10) name += ' ' + randInt(rnd, 2, 99);
              const core = splitLegalForm(v).core;
              if (isUpper(core)) name = name.toUpperCase();
              out = withForm(v, name, config.keepForm);
            } else if (g.method === 'mask') out = maskString(v, rnd, true);
            else if (g.method === 'digits') out = maskString(v, rnd, false);
            else if (g.method === 'hash') out = shortHash(seed, g.name, v);
          }
          if (g.method === 'digits' && out === v) out = maskString(v, makeRng(seed, 'x', v), true);
          used.add(out);
          g.map.set(v, out);
        }
      }
      for (const [o, n] of g.map) mapping.push({ dict: g.name, columns: g.cols.map(c => c.name).join(', '), method: g.method, orig: o, repl: n });
    }

    // 2. Словарь для «замены внутри текста»: исходные имена/коды из других столбцов → их замены
    const cross = new Map();
    for (const g of groups.values()) {
      for (const [o, n] of g.map) {
        const digitsOnly = /^[\d\s/.\-]+$/.test(o);
        if (o.length >= (digitsOnly ? 5 : 4)) cross.set(o, n);
        if (g.method === 'pseudo' || g.method === 'fake') {
          // название без ОПФ тоже может встречаться в тексте («... с Синергия ООО ТЗК»)
          const core = splitLegalForm(o).core;
          if (core.length >= 4 && core !== o && !cross.has(core)) cross.set(core, splitLegalForm(n).core);
        }
      }
    }
    const crossRe = cross.size ? new RegExp([...cross.keys()].sort((a, b) => b.length - a.length).map(escapeRe).join('|'), 'gi') : null;
    const crossLower = new Map([...cross].map(([k, v]) => [k.toLowerCase(), v]));
    const smartDigits = new Map();
    const RE_DATE_IN_TEXT = /(?<!\d)(\d{1,2})([./])(\d{1,2})\2(\d{4})(?!\d)/g;

    function smartText(s) {
      const parts = [];
      const hold = x => { parts.push(x); return '' + (parts.length - 1) + ''; };
      let t = s;
      if (crossRe) t = t.replace(crossRe, m => hold(crossLower.get(m.toLowerCase()) ?? m));
      t = t.replace(RE_DATE_IN_TEXT, (m, d, sep, mo, y) => {
        if (+mo < 1 || +mo > 12 || +d < 1 || +d > 31) return m;
        const info = { kind: 'dmy', sep, dPad: d.length === 2, mPad: mo.length === 2, yLen: 4, rest: '' };
        return hold(formatTextDate(info, Date.UTC(+y, +mo - 1, +d) + D * DAY));
      });
      t = t.replace(/\d{3,}/g, m => {
        if (!smartDigits.has(m)) smartDigits.set(m, maskString(m, makeRng(seed, 'smartdigits', m), false));
        return smartDigits.get(m);
      });
      return t.replace(/(\d+)/g, (_, i) => parts[+i]);
    }

    // 3. Ячейки
    let changed = 0;
    for (const { col, cfg } of active) {
      const m = cfg.method;
      const g = DICT_METHODS.has(m) ? groups.get(m + '␟' + (String(cfg.group || '').trim() || col.letter)) : null;
      let skippedType = 0;
      for (const r of an.dataRows) {
        const cell = cellAt(r, col.c);
        if (!cell) continue;
        const v = cell.v;
        const total = an.totalsRows.has(r);
        if (total && config.keepTotals && !(m === 'scale' || m === 'round')) continue;
        let nv;
        switch (m) {
          case 'clear': nv = null; break;
          case 'pseudo': case 'fake': case 'mask': case 'digits': case 'hash': {
            if (isDate(v)) { skippedType++; continue; }
            nv = g.map.get(asText(v).trim());
            if (nv === undefined) continue;
            if (typeof v === 'number' && /^\d{1,15}$/.test(nv) && nv[0] !== '0') nv = Number(nv);
            break;
          }
          case 'smart':
            if (typeof v !== 'string') { skippedType++; continue; }
            nv = smartText(v); break;
          case 'scale': case 'noise': case 'round': case 'random': {
            if (typeof v !== 'number') { skippedType++; continue; }
            const rnd = makeRng(seed, 'num', col.c, r);
            if (m === 'scale') nv = v * K;
            else if (m === 'noise') nv = v * (1 + (rnd() * 2 - 1) * (Number(cfg.pct) || 0) / 100);
            else if (m === 'round') { const st = Number(cfg.step) || 1; nv = Math.round(v / st) * st; }
            else nv = col.min + rnd() * (col.max - col.min);
            // целые столбцы (дни, штуки) остаются целыми, суммы — до копеек
            nv = col.isInt ? Math.round(nv) : Math.round(nv * 100) / 100;
            if (m === 'round') nv = Number(nv.toFixed(10));
            break;
          }
          case 'shift': case 'month': case 'drandom': {
            const isD = isDate(v);
            const info = isD ? null : parseTextDate(v);
            if (!isD && !info) { skippedType++; continue; }
            const ms = isD ? v.getTime() : infoToMs(info);
            let out;
            if (m === 'shift') out = ms + D * DAY;
            else if (m === 'month') { const dt = new Date(ms); out = Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), 1); }
            else {
              const rnd = makeRng(seed, 'date', col.c, r);
              out = col.dmin + Math.floor(rnd() * ((col.dmax - col.dmin) / DAY + 1)) * DAY;
            }
            // при сдвиге время суток сохраняется, при обобщении — отбрасывается
            nv = isD ? new Date(out) : formatTextDate(info, out, m === 'shift');
            break;
          }
          default: continue;
        }
        const same = (isDate(nv) && isDate(v)) ? nv.getTime() === v.getTime() : nv === v;
        if (!same || cell.f || cell.rich) { changes.set(r + ':' + col.c, nv); changed++; }
      }
      if (skippedType) warnings.push(`«${col.name}»: ${skippedType} ${plural(skippedType, 'значение не подходит', 'значения не подходят', 'значений не подходят')} под тип «${TYPE_NAMES[cfg.type] || cfg.type}» и ${plural(skippedType, 'оставлено', 'оставлены', 'оставлены')} без изменений.`);
    }

    // 4. Предупреждения об утечках: исходные значения в столбцах, которые оставлены как есть
    if (cross.size) {
      const keptCols = an.columns.filter(col => !active.some(a => a.col.c === col.c));
      const leakRe = new RegExp([...cross.keys()].filter(k => k.length >= 5).sort((a, b) => b.length - a.length).map(escapeRe).join('|') || '(?!)', 'i');
      for (const col of keptCols) {
        let hits = 0, example = null;
        for (const r of an.dataRows) {
          const cell = cellAt(r, col.c);
          if (cell && typeof cell.v === 'string' && leakRe.test(cell.v)) { hits++; if (!example) example = cell.v; }
        }
        if (hits) leaks.push(`Столбец «${col.name}» не обезличивается, но в ${hits} ${plural(hits, 'ячейке', 'ячейках', 'ячейках')} есть исходные значения из других столбцов (например: «${String(example).slice(0, 80)}»). Выберите для него метод «Замена внутри текста».`);
      }
    }
    return { changes, mapping, stats: { changed, columns: active.length }, warnings, leaks };
  }

  // ---------- Починка «нестандартных» xlsx (например, выгрузок из 1С) ----------
  /*
   * ExcelJS падает на некоторых файлах, которые Excel открывает без проблем. Известный случай —
   * рисунок листа (xl/drawings/drawingN.xml), записанный без префикса пространства имён:
   * <wsDr xmlns="..."> вместо <xdr:wsDr ...>. Такой рисунок отвязываем от листа и удаляем.
   * Также исправляет булевы свойства шрифтов (см. normalizeFontBooleans).
 * Возвращает { buffer, notes, changed } — файл (исправленный или исходный) и список заметных исправлений.
   */
  const FONT_BOOLEANS = 'b|i|strike|outline|shadow|condense|extend';
  function normalizeFontBooleans(xml) {
    return xml
      .replace(new RegExp(`<(?:${FONT_BOOLEANS})\\s+val="(?:false|0)"\\s*/>`, 'g'), '')
      .replace(new RegExp(`<(${FONT_BOOLEANS})\\s+val="(?:true|1)"\\s*/>`, 'g'), '<$1/>')
      .replace(/<u\s+val="none"\s*\/>/g, '');
  }

  async function repairXlsx(buffer, JSZip) {
    const zip = await JSZip.loadAsync(buffer);
    const notes = [];
    const text = async name => (await zip.file(name).async('string')).replace(/^\uFEFF/, '');
    const relsOf = name => name.replace(/([^/]+)$/, '_rels/$1.rels');
    const resolve = (base, target) => {
      if (target.startsWith('/')) return target.slice(1);
      const parts = base.split('/').slice(0, -1);
      for (const seg of target.split('/')) {
        if (seg === '..') parts.pop();
        else if (seg !== '.') parts.push(seg);
      }
      return parts.join('/');
    };
    const removed = new Set();
    for (const sheet of Object.keys(zip.files).filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
      const relsName = relsOf(sheet);
      if (!zip.file(relsName)) continue;
      let rels = await text(relsName);
      let sheetXml = null;
      const re = /<Relationship\b[^>]*\/>/g;
      for (const rel of rels.match(re) || []) {
        if (!/relationships\/drawing"/.test(rel)) continue;
        const id = (/\bId="([^"]+)"/.exec(rel) || [])[1];
        const target = (/\bTarget="([^"]+)"/.exec(rel) || [])[1];
        if (!id || !target) continue;
        const drawing = resolve(sheet, target);
        let ok = false, anchors = 0;
        if (zip.file(drawing)) {
          const xml = await text(drawing);
          const root = /<(?!\?|!)([\w:.-]+)/.exec(xml);
          ok = !!root && root[1] === 'xdr:wsDr';
          anchors = (xml.match(/Anchor\b/g) || []).length / 2;
        }
        if (ok) continue;
        if (sheetXml === null) sheetXml = await text(sheet);
        sheetXml = sheetXml.replace(new RegExp(`<drawing\\b[^>]*r:id="${id}"[^>]*/>`), '');
        rels = rels.replace(rel, '');
        removed.add(drawing);
        notes.push(anchors > 0
          ? `Рисунки на листе (${drawing}) записаны в нестандартном формате и удалены — данные не затронуты.`
          : `Пустой нестандартный рисунок (${drawing}) удалён — данные не затронуты.`);
      }
      if (sheetXml !== null) {
        zip.file(sheet, sheetXml);
        zip.file(relsName, rels);
      }
    }
    if (removed.size) {
      for (const d of removed) {
        zip.remove(d);
        zip.remove(relsOf(d));
      }
      let ct = await text('[Content_Types].xml');
      for (const d of removed) ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
      zip.file('[Content_Types].xml', ct);
    }
    // Булевы свойства шрифта вида <b val="false"/> (так пишет 1С): ExcelJS считает включённым любой такой тег,
    // и шрифт становится жирным, курсивным или зачёркнутым. Оставляем только действительно включённые.
    let fontsFixed = false;
    for (const name of ['xl/styles.xml', 'xl/sharedStrings.xml']) {
      if (!zip.file(name)) continue;
      const xml = await text(name);
      const fixed = normalizeFontBooleans(xml);
      if (fixed !== xml) { zip.file(name, fixed); fontsFixed = true; }
    }
    if (!removed.size && !fontsFixed) return { buffer, notes, changed: false };
    const out = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
    return { buffer: out, notes, changed: true };
  }

  // ---------- Запись в книгу ExcelJS ----------
  function applyToWorksheet(ws, changes) {
    for (const [key, nv] of changes) {
      const [r, c] = key.split(':').map(Number);
      const cell = ws.getCell(r, c);
      const old = cell.value;
      if (old && typeof old === 'object' && !isDate(old) && (old.formula !== undefined || old.sharedFormula !== undefined)) {
        // Формулу оставляем, кэшированный результат убираем — Excel пересчитает при открытии
        const copy = { ...old };
        delete copy.result;
        cell.value = copy;
        continue;
      }
      cell.value = nv;
    }
  }
  // Колонтитулы: текст из ExcelJS (&L, &C, &R — секции; &P, &N, &D… — поля) в читаемом виде
  const HF_KEYS = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'];
  function headerFooterText(ws) {
    const hf = ws.headerFooter || {};
    const out = [];
    for (const k of HF_KEYS) {
      if (!hf[k]) continue;
      const t = String(hf[k])
        .replace(/&"[^"]*"/g, '')          // шрифт
        .replace(/&\d+/g, '')              // размер шрифта
        .replace(/&[LCR]/g, ' ')            // секции
        .replace(/&[PN]/g, '#').replace(/&[DT]/g, '…').replace(/&[AFZG]/g, '')
        .replace(/&[BIUESXYKH]/g, '').replace(/&&/g, '&')
        .replace(/\s+/g, ' ').trim();
      if (t) out.push(t);
    }
    return out;
  }
  function clearHeaderFooter(ws) {
    // Текст колонтитулов удаляется полностью; картинки-подложки в колонтитулах ExcelJS при записи не сохраняет
    const hf = ws.headerFooter || {};
    ws.headerFooter = { differentFirst: false, differentOddEven: false };
    for (const k of HF_KEYS) if (hf[k]) ws.headerFooter[k] = '';
  }
  function clearNotes(ws) {
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => { if (cell.note) cell.note = undefined; }));
  }
  function clearWorkbookMeta(wb) {
    for (const k of ['creator', 'lastModifiedBy', 'company', 'manager', 'title', 'subject', 'keywords', 'category', 'description']) wb[k] = '';
    wb.lastPrinted = undefined;
  }

  const api = {
    VERSION: root.ANON_VERSION || 'dev',
    plural, cyrb128, makeRng, randomSeed, factorFromSeed, shiftFromSeed,
    parseTextDate, formatTextDate, gridFromWorksheet, parseCsv, gridToCsv,
    detectHeaderRow, analyze, suggest, transform, applyToWorksheet, clearWorkbookMeta, clearNotes, repairXlsx, normalizeFontBooleans, headerFooterText, clearHeaderFooter,
    splitLegalForm, METHODS, METHOD_HINTS, TYPE_NAMES, DICT_METHODS, colLetter, isDate,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Anon = api;
})(typeof window !== 'undefined' ? window : globalThis);

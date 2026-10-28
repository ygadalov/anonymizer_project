/*
 * «Свободный лист»: обезличивание листа без столбцов — для книг, где таблицы разбросаны по листу, есть заголовки,
 * подписи и выводы (источники для презентаций). Каждая ячейка с текстом проходит тот же поиск, что документы Word
 * (src/docx.js): организации, ФИО, ИНН, счета, телефоны, почта, суммы, даты, плюс «Добавить своё».
 * Числа в ячейках умножаются на коэффициент, даты сдвигаются — как в таблицах. Названия, уже обезличенные в столбцах
 * других листов, получают те же замены: одна компания — один псевдоним во всей книге.
 * Без DOM: работает в браузере и в Node (тесты).
 */
(function (root) {
  'use strict';
  const A = () => root.Anon || require('./engine.js');
  const D = () => root.RedactDocx || require('./docx.js');
  const FREE_DICT = 'Свободные листы';
  const DAY = 86400000;

  // Ячейки листа: [строка, столбец, ячейка] — без пустых и без второй половины объединённых
  function cellsOf(grid) {
    const out = [];
    for (let r = 1; r <= grid.maxRow; r++) {
      const row = grid.rows[r];
      if (!row) continue;
      for (let c = 1; c <= grid.maxCol; c++) {
        const x = row[c];
        if (x && !x.slave && x.v !== null && x.v !== undefined && x.v !== '') out.push([r, c, x]);
      }
    }
    return out;
  }

  // Номер без подписи («7720019233» в ячейке под заголовком «ИНН») шаблоны не распознают — подпись берётся
  // из ближайшей текстовой ячейки выше в том же столбце. Такие числа не умножаются, а заменяются как номера
  const ID_LABEL = /(?:^|[\s(«"])(инн|кпп|бик|огрн|снилс|телефон|тел\.?|р\/с|к\/с|расч[её]тный сч[её]т|сч[её]т|паспорт)(?![\p{L}])/iu;
  const DIGITS = /^\+?[\d\s()\-]{5,}$/;
  function idLabel(grid, r, c) {
    for (let rr = r - 1; rr >= 1 && rr >= r - 60; rr--) {
      const x = grid.rows[rr] && grid.rows[rr][c];
      if (!x || x.slave || x.v === null || x.v === '' || typeof x.v !== 'string') continue;
      if (DIGITS.test(x.v.trim())) continue;
      const m = ID_LABEL.exec(x.v);
      return m ? m[1].toUpperCase() : '';
    }
    return '';
  }
  const idText = (grid, r, c, v) => {
    const t = typeof v === 'number' ? (Number.isInteger(v) && v >= 0 ? String(v) : '') : String(v).trim();
    if (!t || !DIGITS.test(t)) return '';
    const label = idLabel(grid, r, c);
    return label ? `${label} ${t}` : '';
  };

  // Какие числа умножать: суммы — от 1000 или дробные от 1. Годы (1900–2100), доли и проценты (меньше 1)
  // и небольшие целые (количество, номера пунктов) не меняются
  const scalable = v => typeof v === 'number' && Number.isFinite(v) &&
    (Math.abs(v) >= 1000 ? !(Number.isInteger(v) && v >= 1900 && v <= 2100) : !Number.isInteger(v) && Math.abs(v) >= 1);

  /*
   * План: что найдено в тексте свободных листов и чем заменяется.
   * sheets = [{ name, grid }]; opts — как у RedactDocx.plan (seed, factor, shiftDays, types, skip, custom, customCase,
   * customForms) плюс known: Map «исходное → замена» из столбцов других листов (для общих псевдонимов).
   * Возвращает { rows, repl, keyPairs } — как RedactDocx.plan.
   */
  function planFree(sheets, opts = {}) {
    const parts = sheets.map(sh => ({
      name: sh.name,
      segs: cellsOf(sh.grid).map(([r, c, x]) => {
        // номер под подписью («ИНН» выше в столбце) ищется вместе с подписью: «ИНН 7720019233»
        const text = idText(sh.grid, r, c, x.v) || (typeof x.v === 'string' ? x.v : '');
        return text && { kind: 'text', text, nodes: [{ start: r * 16384 + c }] };
      }).filter(Boolean),
    }));
    // известные названия из столбцов ищутся и здесь (во всех падежах), чтобы получить ту же замену.
    // Название без формы организации и кавычек — тоже: «АО «Нордлиния»» → «Нордлиния» → «Клиент 004»
    const bare = x => A().splitLegalForm(String(x).trim()).core.replace(/^[«"„“]+|[»"“”]+$/g, '').trim();
    const known = new Map();
    for (const [o, n] of opts.known || []) {
      const s = String(o).trim();
      if (s.length < 4 || /^[\d\s/.\-]+$/.test(s)) continue;
      known.set(s.toLowerCase(), String(n));
      const core = bare(s);
      if (core.length >= 4 && !known.has(core.toLowerCase())) known.set(core.toLowerCase(), bare(n));
    }
    const custom = [...(opts.custom || []), ...[...(opts.known || [])].map(([o]) => String(o).trim()).filter(s => known.has(s.toLowerCase()))];
    // замены из столбцов заняты: «Контрагент 001» на листе-таблице и на свободном листе — одна и та же компания
    const reserved = new Set();
    for (const [, n] of opts.known || []) { reserved.add(String(n)); reserved.add(bare(n)); }
    const p = D().plan(parts, { ...opts, custom, reserved });
    if (!known.size) return p;
    // замены известных значений — как в столбцах; ключ — без повторов.
    // Человек из столбца в другой форме («Лаптева И. В.» при «Лаптев Илья Викторович») — та же замена
    const userCustom = new Set((opts.custom || []).map(x => String(x).trim().toLowerCase()));
    const persons = new Map();
    for (const r of p.rows) {
      const k = known.get(r.value.toLowerCase()) || known.get(bare(r.value).toLowerCase());
      if (!k) continue;
      r.repl = k;
      if (r.type === 'custom' && !userCustom.has(r.value.toLowerCase())) r.fromTable = true;
      const pk = personKey(r.value);
      if (pk && !persons.has(pk)) persons.set(pk, k);
    }
    // другая форма того же человека — псевдоним со случайными инициалами («Сотрудник 001 К. М.»), как в таблицах:
    // замена своя, поэтому по ключу возвращается именно эта форма
    const usedRepl = new Set(p.rows.map(r => r.repl));
    for (const r of p.rows) {
      if (r.type !== 'person' || known.has(r.value.toLowerCase())) continue;
      const pk = personKey(r.value);
      const k = persons.get(pk);
      if (!k) continue;
      const real = pk.split('|')[1].toUpperCase();
      const LET = 'АБВГДЕЖЗИКЛМНОПРСТУФЭЮЯ';
      let out, attempt = 0;
      do {
        const rnd = A().makeRng(opts.seed || 'free', 'initials', r.value, attempt++);
        const fake = [...real].map(() => LET[Math.floor(rnd() * LET.length)]).join('');
        out = fake === real ? '' : `${k} ${[...fake].map(x => x + '.').join(' ')}`;
      } while ((!out || usedRepl.has(out)) && attempt < 50);
      usedRepl.add(out);
      r.repl = out;
    }
    const repl = new Map(), keyPairs = [];
    const order = [...p.rows].sort((a, b) => (known.has(b.value.toLowerCase()) ? 1 : 0) - (known.has(a.value.toLowerCase()) ? 1 : 0));
    for (const r of order) {
      if (!r.on) continue;
      repl.set(r.value, r.repl);
      for (const f of r.forms || []) repl.set(f, r.repl);
      // одна пара на замену: иначе у «Сотрудник 001» два исходных значения, и по ключу не вернуть ни одно.
      // Первой идёт форма из столбцов (полное ФИО) — она и вернётся
      if (!keyPairs.some(([, x]) => x === r.repl)) keyPairs.push([r.keyValue || r.value, r.repl]);
    }
    return { rows: p.rows, repl, keyPairs };
  }

  // Человек без учёта падежа и формы записи: основа фамилии + инициалы («лаптев|ив» для «Лаптева И. В.»
  // и «Лаптев Илья Викторович»). Не ФИО — пустая строка
  function personKey(v) {
    const t = String(v).trim().split(/\s+/);
    let sur = '', ini = '';
    if (t.length === 3 && t.every(w => /^[А-ЯЁ][а-яё-]+$/.test(w))) { sur = t[0]; ini = t[1][0] + t[2][0]; }
    else {
      const words = t.filter(w => /^[А-ЯЁ][а-яё-]{2,}$/.test(w));
      const letters = t.join(' ').match(/(?:^|\s|\.)([А-ЯЁ])\./g) || [];
      if (words.length !== 1 || letters.length < 1) return '';
      sur = words[0]; ini = letters.map(x => x.replace(/[\s.]/g, '')).join('');
    }
    return sur.toLowerCase().replace(/(ого|ому|ым|ой|ою|ую|ая|а|у|е|ом)$/, '') + '|' + ini.toLowerCase();
  }

  /*
   * Применить план к листам: { bySheet: { [лист]: { changes: Map('r:c' → значение), changed } }, numberPairs, stats }.
   * opts: factor, shiftDays, numbers (умножать числа), dates (сдвигать даты).
   * numberPairs — «новое → исходное» для чисел: по ключу они возвращаются точно, без ошибки округления.
   */
  function applyFree(sheets, plan, opts = {}) {
    const finder = D().dictFinder(plan.repl);
    const K = Number(opts.factor) || 1, S = Math.round(Number(opts.shiftDays) || 0);
    const n2o = new Map(), o2n = new Map();
    const bySheet = {};
    const stats = { text: 0, numbers: 0, dates: 0 };
    for (const sh of sheets) {
      const out = bySheet[sh.name] = { changes: new Map(), changed: 0 };
      for (const [r, c, x] of cellsOf(sh.grid)) {
        const v = x.v;
        let nv = v;
        if (typeof v === 'string') {
          const f = finder(v);
          if (!f.length) continue;
          let s = '', at = 0;
          for (const m of f) { s += v.slice(at, m.start) + m.repl; at = m.end; }
          nv = s + v.slice(at);
          stats.text++;
        } else if (x.f) {
          continue;   // формула пересчитается в Excel из обезличенных ячеек
        } else if (typeof v === 'number' && plan.repl.has(String(v)) && idText(sh.grid, r, c, v)) {
          // ИНН, счёт или телефон, записанный числом: замена — тоже число
          const t = plan.repl.get(String(v));
          nv = /^\d{1,15}$/.test(t) && t[0] !== '0' ? Number(t) : t;
          stats.text++;
        } else if (typeof v === 'number' && opts.numbers !== false && scalable(v)) {
          if (o2n.has(v)) nv = o2n.get(v);
          else {
            const dec = Number.isInteger(v) && Number.isInteger(v * K) ? 0 : 2;
            nv = Math.round(v * K * 10 ** dec) / 10 ** dec;
            // замены уникальны — иначе по ключу не вернуть
            while (n2o.has(nv) || nv === v) nv = Math.round((nv + (dec ? 0.01 : 1)) * 100) / 100;
            n2o.set(nv, v); o2n.set(v, nv);
          }
          stats.numbers++;
        } else if (A().isDate(v) && opts.dates !== false && S) {
          nv = new Date(v.getTime() + S * DAY);
          stats.dates++;
        } else continue;
        out.changes.set(r + ':' + c, nv);
        out.changed++;
      }
    }
    return { bySheet, numberPairs: [...n2o].map(([n, o]) => [n, o]), stats };
  }

  // Текст ячеек, оставшийся после обезличивания, в котором поиск Redact ещё что-то находит (для предупреждений)
  function findLeft(texts) {
    const out = [];
    for (const t of texts) {
      const hits = D().detect(t).filter(h => h.type !== 'date' && h.type !== 'amount');
      if (hits.length) out.push({ text: t, hits });
    }
    return out;
  }

  const api = { FREE_DICT, cellsOf, scalable, personKey, planFree, applyFree, findLeft };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RedactFree = api;
})(typeof window !== 'undefined' ? window : globalThis);

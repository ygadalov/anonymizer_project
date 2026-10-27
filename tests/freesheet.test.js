// «Свободный лист»: обезличивание по ячейкам (src/freesheet.js), общие замены с листами-таблицами, восстановление по ключу.
// Запуск: node --test tests/*.test.js
'use strict';
const test = require('node:test');
const assert = require('node:assert');
const A = require('../src/engine.js');
const F = require('../src/freesheet.js');

// лист из списка ячеек { 'B1': значение }; формула — { f: true }
function grid(cells) {
  const pos = k => { const m = /^([A-Z]+)(\d+)$/.exec(k); return [+m[2], [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)]; };
  const list = Object.entries(cells).map(([k, v]) => [...pos(k), v]);
  const maxRow = Math.max(...list.map(x => x[0])), maxCol = Math.max(...list.map(x => x[1]));
  const rows = Array.from({ length: maxRow + 1 }, () => new Array(maxCol + 1).fill(null));
  for (const [r, c, v] of list) rows[r][c] = v && v.f ? { v: null, f: true } : { v, f: false };
  return { rows, maxRow, maxCol };
}

const slide = () => grid({
  B1: 'Продажи ООО «Вересков» за I квартал 2025 г.',
  B2: 'Подготовил: Мирошев Г. А., тел. +7 (495) 123-45-67',
  B4: 'Клиент', C4: 'Выручка, руб', D4: 'ИНН',
  B5: 'АО «Нордлиния»', C5: 4250000, D5: 7720019233,
  B6: 'ООО «Стройлес»', C6: 2980500.5, D6: '7714123459',
  B7: 'Итого', C7: { f: true },
  F4: 'Выводы', F5: 'Крупнейший клиент — АО «Нордлиния», договор от 12.03.2025.',
  F6: 'Ответственный — Лаптева И. В., сделок: 12, рост 0,35 к 2024 году.',
});

test('свободный лист: текст, номера под подписью, суммы и годы', () => {
  const sheets = [{ name: 'Слайд', grid: slide() }];
  const plan = F.planFree(sheets, { seed: 's', factor: 1.2, shiftDays: -10 });
  const { bySheet, numberPairs } = F.applyFree(sheets, plan, { factor: 1.2, shiftDays: -10 });
  const ch = bySheet['Слайд'].changes;
  const all = [...ch.values()].map(String).join('\n');
  for (const s of ['Вересков', 'Нордлиния', 'Стройлес', 'Мирошев', '123-45-67', 'Лаптева', '12.03.2025'])
    assert.ok(!all.includes(s), `осталось: ${s}`);
  // одна компания — одна замена везде
  const nord = plan.repl.get('Нордлиния');
  assert.ok(ch.get('5:2').includes(nord) && ch.get('5:6').includes(nord));
  // ИНН под подписью «ИНН» — и числом, и текстом; число остаётся числом
  assert.strictEqual(typeof ch.get('5:4'), 'number');
  assert.notStrictEqual(ch.get('5:4'), 7720019233);
  assert.ok(ch.has('6:4') && ch.get('6:4') !== '7714123459');
  // суммы умножены, формула не тронута; годы, проценты и количество в тексте не меняются
  assert.strictEqual(ch.get('5:3'), 5100000);
  assert.ok(!ch.has('7:3'));
  assert.ok(ch.get('6:6').includes('сделок: 12') && ch.get('6:6').includes('2024'));
  assert.deepStrictEqual(numberPairs.find(([, o]) => o === 4250000), [5100000, 4250000]);
  assert.ok(F.scalable(2980500.5) && !F.scalable(2025) && !F.scalable(0.35) && !F.scalable(12));
});

test('свободный лист и лист-таблица: общие замены, занятые номера не повторяются, ключ восстанавливает всё', () => {
  // лист-таблица: контрагенты псевдонимами, менеджер — псевдонимом
  const data = grid({ A1: 'Контрагент', B1: 'Менеджер', C1: 'Сумма', A2: 'АО «Нордлиния»', B2: 'Лаптев Илья Викторович', C2: 1500000, A3: 'ООО «Альфа»', B3: 'Сорокина Елена Петровна', C3: 980000 });
  const an = A.analyze(data, 1);
  const cols = { 1: { type: 'text', method: 'pseudo', prefix: 'Контрагент' }, 2: { type: 'text', method: 'pseudo', prefix: 'Сотрудник' }, 3: { type: 'number', method: 'scale' } };
  const cfg = { seed: 'k', factor: 1.2, shiftDays: -10, keepForm: true };
  const result = A.transformBook([{ name: 'Данные', grid: data, an, columns: cols }, { name: 'Слайд', grid: slide(), an: A.analyze(slide(), 4), columns: {} }], cfg);
  const known = new Map(result.groups.flatMap(g => g.pairs));
  const sheets = [{ name: 'Слайд', grid: slide() }];
  const plan = F.planFree(sheets, { ...cfg, known });
  const applied = F.applyFree(sheets, plan, cfg);
  const ch = applied.bySheet['Слайд'].changes;
  const tableNord = A.splitLegalForm(known.get('АО «Нордлиния»')).core;
  // «Нордлиния» на слайде — та же замена, что в таблице
  assert.ok(ch.get('5:6').includes(tableNord), ch.get('5:6'));
  // «Вересков» только на слайде: номер не совпадает с занятыми в таблице
  const ver = plan.repl.get('Вересков');
  assert.ok(![...known.values()].some(v => A.splitLegalForm(v).core === ver), ver);
  // «Лаптева И. В.» — тот же человек, что «Лаптев Илья Викторович» в таблице: псевдоним с другими инициалами
  const lap = known.get('Лаптев Илья Викторович');
  assert.ok(new RegExp(`${lap} [А-Я]\\. [А-Я]\\.`).test(ch.get('6:6')), ch.get('6:6'));
  // в ключе одна пара на замену
  const backs = plan.keyPairs.map(([, r]) => r);
  assert.strictEqual(new Set(backs).size, backs.length);

  // восстановление: ключ с таблицей и свободным листом
  const key = A.buildKeyBook({ sheets: [{ name: 'Данные', an, columns: cols }], config: cfg, result, sourceFile: 't.xlsx', appVersion: 'test' });
  key.freeSheets = ['Слайд'];
  key.dictionaries.push({ name: F.FREE_DICT, method: 'doc', prefix: '', columns: ['Слайд'], pairs: plan.keyPairs });
  key.numbers[F.FREE_DICT] = applied.numberPairs;
  const prep = A.prepareRestore(key);
  const red = slide();
  for (const [k, v] of ch) { const [r, c] = k.split(':').map(Number); red.rows[r][c] = { v, f: false }; }
  const ran = A.analyze(red, 0);
  const acts = Object.fromEntries(ran.columns.map(col => [col.c, 'free']));
  const back = A.restoreTransform(red, ran, prep, acts, 'Слайд').changes;
  const orig = slide();
  for (const [k] of ch) {
    const [r, c] = k.split(':').map(Number);
    assert.strictEqual(back.get(k), orig.rows[r][c].v, `ячейка ${k}`);
  }
});

test('строки над заголовками таблицы: названия из столбцов заменяются, остальное видно поиску', () => {
  const g = grid({ A1: 'Отчёт по АО «Нордлиния», подготовил Мирошев Г. А.', A3: 'Контрагент', B3: 'Сумма', A4: 'АО «Нордлиния»', B4: 100000 });
  const an = A.analyze(g, 3);
  const r = A.transformBook([{ name: '', grid: g, an, columns: { 1: { type: 'text', method: 'pseudo', prefix: 'Контрагент' } } }], { seed: 'h', factor: 1, shiftDays: 0, keepForm: true });
  const t = r.bySheet[''].changes.get('1:1');
  assert.ok(t && !t.includes('Нордлиния'), t);
  // ФИО в шапке столбцы не закрывают — его находит поиск (приложение показывает это как возможную утечку)
  assert.strictEqual(F.findLeft([t]).length, 1);
  assert.strictEqual(F.findLeft([t])[0].hits[0].type, 'person');
});

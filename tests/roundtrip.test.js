// Обезличивание → ключ → восстановление должно возвращать данные один в один.
// Запуск: node --test tests/   (зависимости не нужны; данные синтетические)
'use strict';
const test = require('node:test');
const assert = require('node:assert');
const A = require('../src/engine.js');

const FORMS = ['ООО', 'АО', 'ПАО', 'ИП'];
const WORDS = ['СИНЕРГИЯ', 'ОПТТРЕЙД', 'ПРОМЕТЕЙ', 'АГРО-ТЕХНИК', 'МЦОЗ', 'ВЕКТОР', 'ГРАНИТ', 'УРАЛСНАБ', 'ДОНЗЕРНО', 'ТРАНСЛАЙН'];
const pad = (n, w) => String(n).padStart(w, '0');

// Синтетическая выгрузка в духе 1С: шапка в 3-й строке, итог внизу, «трудные» случаи нарочно
function makeGrid(n = 240) {
  const rows = [null, [], [], []];
  const header = ['Дата отчета', 'Организация', 'ИНН', 'Код контрагента', 'Контрагент', 'Код договора', 'Договор', 'Сумма', 'Сумма, руб', 'Дата документа', 'Дни'];
  rows[1] = [null, { v: 'Дебиторская задолженность' }];
  rows[2] = [null];
  rows[3] = [null, ...header.map(v => ({ v }))];
  let total = 0;
  for (let i = 0; i < n; i++) {
    const k = i % 37;
    const word = WORDS[k % WORDS.length] + (k >= WORDS.length ? ' ' + k : '');
    // ОПФ в конце, в начале и в середине названия
    const name = k % 5 === 0 ? `${FORMS[k % 4]} ${word}` : k % 7 === 0 ? `${word} ООО ТЗК` : `${word} ${FORMS[k % 4]}`;
    const amount = [0.01, 0.02, 400316, 400315.99, 15.8, 274800][i % 6] + (i % 6 > 3 ? i : 0);
    total += amount;
    const d = 1 + (i % 28), m = 1 + (i % 12);
    rows.push([null,
      { v: '30.09.2026' },
      { v: 'ООО "Ромашка"' },
      k % 3 ? { v: 7700000000 + k * 7919 } : { v: String(2300000000 + k * 104729) }, // ИНН числом и строкой
      { v: pad(k * 13, 9) },                                                        // код с ведущими нулями
      { v: name + (i % 11 === 0 ? ' ' : '') },                                      // пробел в конце
      { v: '00/' + pad(k * 13, 6) },                                                // формат как у кодов контрагентов
      { v: `ТЛС/${word.split(' ')[0]}-2013/ЦО${pad(i, 6)} от ${pad(d, 2)}.${pad(m, 2)}.2015 аренда` },
      { v: amount },
      { v: Math.round(amount * 92.5 * 100) / 100 },
      i % 2 ? { v: `${pad(d, 2)}.${pad(m, 2)}.2017` } : { v: new Date(Date.UTC(2017, m - 1, d, 23, 59, 59)) },
      { v: i % 40 },
    ]);
  }
  rows.push([null, { v: 'Итого' }, null, null, null, null, null, null, { v: Math.round(total * 100) / 100 }]);
  const maxCol = header.length;
  for (const r of rows) if (r) while (r.length <= maxCol) r.push(null);
  return { rows, maxRow: rows.length - 1, maxCol };
}

function anonymize(grid, seed, variant) {
  const an = A.analyze(grid, A.detectHeaderRow(grid));
  const cfg = { seed, factor: A.factorFromSeed(seed), shiftDays: A.shiftFromSeed(seed), keepForm: true, keepTotals: true, columns: {} };
  for (const c of an.columns) {
    cfg.columns[c.c] = A.suggest(c);
    if (c.detected === 'number' && c.name !== 'Дни') cfg.columns[c.c].method = 'scale';
    if (c.detected === 'date') cfg.columns[c.c].method = 'shift';
    if (c.name === 'Контрагент') cfg.columns[c.c].method = ['pseudo', 'fake', 'mask', 'hash'][variant % 4];
    if (c.name === 'Код договора') cfg.columns[c.c].method = variant % 2 ? 'mask' : 'keep';
  }
  const res = A.transform(grid, an, cfg);
  const key = JSON.parse(JSON.stringify(A.buildKey({ an, config: cfg, result: res, sourceFile: 'test.xlsx', sheet: 'Лист1' })));
  const anon = { ...grid, rows: grid.rows.map((row, r) => row && row.map((cell, c) => res.changes.has(r + ':' + c) ? { v: res.changes.get(r + ':' + c) } : cell)) };
  return { an, key, anon, res };
}

function restore(anon, key) {
  const an = A.analyze(anon, A.detectHeaderRow(anon));
  const prep = A.prepareRestore(key);
  const actions = {};
  for (const c of an.columns) actions[c.c] = A.defaultRestoreAction(prep, c);
  return A.restoreTransform(anon, an, prep, actions);
}

const val = v => v instanceof Date ? 'D' + v.getTime() : JSON.stringify(v);

test('восстановление по ключу возвращает файл один в один', () => {
  const grid = makeGrid();
  for (let i = 0; i < 24; i++) {
    const { an, key, anon, res } = anonymize(grid, 'seed-' + i, i);
    assert.ok(res.stats.changed > 0);
    const back = restore(anon, key);
    for (const r of an.dataRows) for (const c of an.columns) {
      const k = r + ':' + c.c;
      const got = back.changes.has(k) ? back.changes.get(k) : (anon.rows[r][c.c] ? anon.rows[r][c.c].v : null);
      const orig = grid.rows[r][c.c] ? grid.rows[r][c.c].v : null;
      assert.strictEqual(val(got), val(orig), `seed-${i}, строка ${r}, «${c.name}»`);
    }
  }
});

test('в ответном файле названия возвращаются и внутри нового текста', () => {
  const grid = makeGrid(60);
  const { key, anon } = anonymize(grid, 'reply', 0);
  const nameCol = 5, innCol = 3;
  const reply = { rows: [null, [null, { v: 'Комментарий' }, { v: 'Контрагент' }]], maxRow: 0, maxCol: 2 };
  const expected = [];
  for (let r = 4; r < 20; r++) {
    const pseudo = String(anon.rows[r][nameCol].v).trim(), inn = anon.rows[r][innCol].v;
    reply.rows.push([null, { v: `${pseudo} (ИНН ${inn}): просрочка` }, { v: pseudo }]);
    expected.push([`${String(grid.rows[r][nameCol].v).trim()} (ИНН ${grid.rows[r][innCol].v}): просрочка`, String(grid.rows[r][nameCol].v).trim()]);
  }
  reply.rows.push([null, { v: 'Контрагент 999 ООО — нет в ключе' }, null]);
  reply.maxRow = reply.rows.length - 1;
  const an = A.analyze(reply, 1);
  const prep = A.prepareRestore(key);
  const actions = {};
  for (const c of an.columns) actions[c.c] = A.defaultRestoreAction(prep, c);
  const back = A.restoreTransform(reply, an, prep, actions);
  expected.forEach(([comment, name], i) => {
    assert.strictEqual(back.changes.get(`${i + 2}:1`), comment);
    assert.strictEqual(back.changes.get(`${i + 2}:2`), name);
  });
  assert.deepStrictEqual(back.leftovers.map(x => x.text), ['Контрагент 999']);
});

test('ключ шифруется паролем и не открывается чужим', async () => {
  const { key } = anonymize(makeGrid(30), 'crypto', 0);
  const env = await A.encryptKey(key, 'Пароль-2026');
  assert.strictEqual(env.kind, 'redact-key-encrypted');
  assert.ok(!JSON.stringify(env).includes('Ромашка'));
  assert.deepStrictEqual(await A.decryptKey(env, 'Пароль-2026'), key);
  await assert.rejects(A.decryptKey(env, 'не тот пароль'), /неверный пароль/);
});

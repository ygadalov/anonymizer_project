// Чтение .xls (Excel 97–2003, XML Excel 2003, HTML-таблица из учётных систем) через SheetJS.
// Запуск: node --test tests/*.test.js   (зависимости не нужны: книги для проверки пишет сам SheetJS из vendor/)
'use strict';
const test = require('node:test');
const assert = require('node:assert');
global.cptable = require('../vendor/cpexcel.cyr.js');
const X = require('../vendor/xlsx.core.min.js');
const XL = require('../src/xls.js');

const ab = buf => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
function sampleBook() {
  const ws = X.utils.aoa_to_sheet([['Контрагент', 'ИНН', 'Сумма'], ['ВЕРЕСКОВ ООО', '7748214918', 1000.5], ['АО НОРДЛИНИЯ', '7720019233', 2000]]);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }];
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, ws, 'ДЗ');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['Поставщик'], ['ЛУГОРАН ООО']]), 'КЗ');
  return wb;
}
const values = sh => sh.cells.map(c => c[3]);

test('Excel 97–2003 и XML Excel 2003: листы, кириллица, числа и тексты', () => {
  for (const bookType of ['biff8', 'xlml']) {
    const buf = ab(X.write(sampleBook(), { bookType, type: 'buffer' }));
    assert.ok(XL.isXls(buf, 'выгрузка.xls'), bookType);
    const data = XL.extractBook(X, new Uint8Array(buf), false);
    assert.deepStrictEqual(data.sheets.map(s => s.name), ['ДЗ', 'КЗ']);
    assert.deepStrictEqual(values(data.sheets[0]), ['Контрагент', 'ИНН', 'Сумма', 'ВЕРЕСКОВ ООО', '7748214918', 1000.5, 'АО НОРДЛИНИЯ', '7720019233', 2000]);
  }
});

test('HTML-таблица с расширением .xls: числа по-русски, коды с нулями остаются текстом', () => {
  const html = '<html><body><table><tr><td>Сумма</td><td>Код</td><td>Дата</td></tr>' +
    '<tr><td>1 000,50</td><td>000154</td><td>12.03.2025</td></tr><tr><td>-2000</td><td>00/2210</td><td>01.02.2025</td></tr></table></body></html>';
  const buf = ab(Buffer.from(html, 'utf-8'));
  assert.strictEqual(XL.sniff(buf), 'html');
  assert.ok(XL.isXls(buf, '1c.xls'));
  const sh = XL.extractBook(X, new Uint8Array(buf), true).sheets[0];
  assert.deepStrictEqual(values(sh), ['Сумма', 'Код', 'Дата', 1000.5, '000154', '12.03.2025', -2000, '00/2210', '01.02.2025']);
});

test('не книга Excel с расширением .xls — понятная ошибка; .xlsx под именем .xls открывается как обычно', () => {
  assert.throws(() => XL.isXls(ab(Buffer.from('просто текст, не таблица')), 'отчёт.xls'), /не книга Excel/);
  assert.strictEqual(XL.isXls(ab(Buffer.from('PK\u0003\u0004 остальное')), 'отчёт.xls'), false);
  assert.strictEqual(XL.isXls(ab(Buffer.from('a;b\n1;2')), 'data.csv'), false);
});

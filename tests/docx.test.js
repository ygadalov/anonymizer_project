// Обезличивание документов Word (.docx): поиск, замена с раздробленными прогонами, очистка, восстановление по ключу.
// Запуск: node --test tests/*.test.js   (зависимости не нужны: JSZip — из vendor/, документ собирается в тесте)
'use strict';
const test = require('node:test');
const assert = require('node:assert');
const JSZip = require('../vendor/jszip.min.js');
const D = require('../src/docx.js');

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const run = (t, b) => `<w:r>${b ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (...runs) => `<w:p>${runs.join('')}</w:p>`;

// документ с выдуманными данными; названия и ФИО нарочно разбиты на несколько прогонов, как это делает Word
function makeDocx() {
  const zip = new JSZip();
  const body = [
    para(run('Договор поставки от «12» марта 2025 г.')),
    para(run('ООО «Вер'), run('еско', true), run('в» (ИНН 7748214918, КПП 774801001) в лице директора '), run('Мирошева Глеба'), run(' Аркадьевича')),
    para(run('и ИП Образцова Анна Ильинична, тел. +7 (495) 123-45-67, '), '<w:hyperlink r:id="rIdMail">', run('mail@veraskov.ru'), '</w:hyperlink>'),
    para(run('Цена 1 250 000,00 руб. (Один миллион двести пятьдесят тысяч рублей 00 копеек).')),
    para('<w:commentRangeStart w:id="0"/>', run('Срок поставки — до 03.06.2025.'), '<w:commentRangeEnd w:id="0"/>',
      '<w:r><w:commentReference w:id="0"/></w:r>'),
    para(run('Было: '), '<w:del w:id="1" w:author="Сидоров Пётр"><w:r><w:delText>АО «Тайна» 7720019233</w:delText></w:r></w:del>',
      '<w:ins w:id="2" w:author="Сидоров Пётр"><w:r><w:t>новое условие</w:t></w:r></w:ins>'),
    para(run('Подписи: '), run('Мирошев Г.'), run(' А.'), '<w:tab/>', run('А.И. Образцова')),
  ].join('');
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  zip.file('word/header1.xml', `<?xml version="1.0"?><w:hdr ${W}>${para(run('ООО «Вересков» — договор поставки'))}</w:hdr>`);
  zip.file('word/comments.xml', `<?xml version="1.0"?><w:comments ${W}><w:comment w:id="0" w:author="Сидоров Пётр" w:initials="СП">${para(run('Проверить ИНН 7748214918'))}</w:comment></w:comments>`);
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0"?><Relationships><Relationship Id="rIdMail" Type="hyperlink" Target="mailto:mail@veraskov.ru" TargetMode="External"/></Relationships>');
  zip.file('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="c" xmlns:dc="d"><dc:creator>Сидоров Пётр</dc:creator><cp:lastModifiedBy>Сидоров Пётр</cp:lastModifiedBy></cp:coreProperties>');
  zip.file('docProps/app.xml', '<?xml version="1.0"?><Properties><Company>ООО Вересков</Company></Properties>');
  return zip;
}
const textOf = xml => D.segmentsOf(xml).filter(s => s.kind === 'text').map(s => s.text.split('\u0001').join('\t')).join('\n');
async function allXml(zip) {
  let s = '';
  for (const n of Object.keys(zip.files)) if (!zip.files[n].dir) s += '\n' + await zip.file(n).async('string');
  return s;
}

test('docx: находит данные в раздробленном тексте, заменяет везде и чистит исправления, примечания и свойства', async () => {
  const zip = makeDocx();
  const a = await D.analyzeDocx(zip, { seed: 'test', factor: 1.3, shiftDays: 45 });
  const byType = t => a.rows.filter(r => r.type === t).map(r => r.value);
  assert.deepStrictEqual(byType('org'), ['Вересков']);
  assert.deepStrictEqual(byType('person').sort(), ['А.И. Образцова', 'Мирошев Г. А.', 'Мирошева Глеба Аркадьевича', 'Образцова Анна Ильинична'].sort());
  for (const t of ['inn', 'kpp', 'phone', 'email', 'amount', 'words', 'date']) assert.ok(byType(t).length, t);
  assert.strictEqual(a.info.comments, 1);
  assert.ok(a.info.revisions >= 2);

  await D.redactDocx(zip, { repl: a.repl, clean: { revisions: true, comments: true, meta: true } });
  const xml = await allXml(zip);
  for (const secret of ['Вересков', 'Вер', 'Мирошев', 'Образцова', '7748214918', '774801001', '123-45-67', 'veraskov', 'Сидоров',
    'Тайна', '7720019233', '1 250 000,00', 'Один миллион', '03.06.2025', '«12» марта 2025']) {
    assert.ok(!xml.includes(secret), `осталось «${secret}»`);
  }
  const doc = textOf(await zip.file('word/document.xml').async('string'));
  assert.match(doc, /ООО «Контрагент 001»/);
  assert.match(doc, /новое условие/);
  // ФИО в родительном падеже заменено ФИО в родительном падеже, подпись — тем же человеком с инициалами
  const gen = a.repl.get('Мирошева Глеба Аркадьевича'), sig = a.repl.get('Мирошев Г. А.');
  assert.match(gen, /^[А-ЯЁ][а-яё]+ов[а]? [А-ЯЁ][а-яё]+а [А-ЯЁ][а-яё]+ича$|^[А-ЯЁ][а-яё]+(ин|ев|ёв)а [А-ЯЁ][а-яё]+а [А-ЯЁ][а-яё]+ича$/);
  assert.strictEqual(sig.split(' ')[0] + 'а', gen.split(' ')[0]);
  assert.match(textOf(await zip.file('word/header1.xml').async('string')), /Контрагент 001/);
  assert.match(await zip.file('word/_rels/document.xml.rels').async('string'), /mailto:contact\d+@example\.com/);
});

test('docx: восстановление по ключу возвращает текст документа один в один', async () => {
  const zip = makeDocx();
  const a = await D.analyzeDocx(zip, { seed: 'restore', factor: 0.77, shiftDays: -120 });
  await D.redactDocx(zip, { repl: a.repl, clean: { revisions: true, comments: true, meta: true } });
  const back = new Map([...a.repl].map(([o, r]) => [r, o]));
  await D.restoreDocx(zip, back);
  // эталон: исходный документ с принятыми исправлениями и без примечаний
  const ref = makeDocx();
  const refXml = D.acceptRevisions(await ref.file('word/document.xml').async('string'));
  assert.strictEqual(textOf(await zip.file('word/document.xml').async('string')), textOf(refXml).replace(/\n?$/, ''));
  assert.strictEqual(textOf(await zip.file('word/header1.xml').async('string')), textOf(await ref.file('word/header1.xml').async('string')));
  assert.match(await zip.file('word/_rels/document.xml.rels').async('string'), /mailto:mail@veraskov\.ru/);
});

test('docx: отключённый тип и снятая галочка не заменяются; свои значения — заменяются', async () => {
  const zip = makeDocx();
  const a = await D.analyzeDocx(zip, { seed: 'opts', types: { date: false }, skip: ['Мирошев Г. А.'], custom: ['поставки'] });
  assert.ok(![...a.repl.keys()].includes('03.06.2025'));
  assert.ok(!a.repl.has('Мирошев Г. А.'));
  assert.ok(/^Значение \d{3}$/.test(a.repl.get('поставки')));
  await D.redactDocx(zip, { repl: a.repl, clean: {} });
  const doc = textOf(await zip.file('word/document.xml').async('string'));
  assert.match(doc, /03\.06\.2025/);
  assert.match(doc, /Мирошев Г\. А\./);
  assert.ok(!doc.includes('поставки'));
});

test('docx: ИНН-замена — с неверной контрольной суммой; суммы и даты уникальны', () => {
  const t = 'ИНН 7748214918 и 507915252185, суммы 100 руб., 100,01 руб., даты 01.02.2025 и 02.02.2025.';
  const { rows } = D.plan([{ name: 'word/document.xml', segs: [{ kind: 'text', text: t }] }], { seed: 'x', factor: 1.0001, shiftDays: 1 });
  for (const r of rows.filter(r => r.type === 'inn')) assert.ok(!D.innValid(r.repl), r.repl);
  const reps = rows.map(r => r.repl);
  assert.strictEqual(new Set(reps).size, reps.length);
  assert.strictEqual(D.shiftDate('«31» декабря 2025', 1), '«01» января 2026');
});

test('«Добавить своё»: без учёта регистра и во всех падежах — одна замена, в ключе одна пара', () => {
  const xml = `<w:document ${W}><w:body>${para(run('Склад на ул. Лесная, 5. Доступ на Лесной открыт; ЛЕСНАЯ, 5 — адрес Иванова, а Ивану передали ключи.'))}</w:body></w:document>`;
  const parts = [{ name: 'word/document.xml', xml, segs: D.segmentsOf(xml) }];
  const p = D.plan(parts, { custom: ['ул. Лесная', 'Лесная', 'Иван'] });
  const row = v => p.rows.find(r => r.value === v);
  // «ул. Лесная» забирает свои вхождения первой, остальные формы — у «Лесная»
  assert.deepStrictEqual(row('ул. Лесная').forms, []);
  assert.deepStrictEqual(row('Лесная').forms.sort(), ['ЛЕСНАЯ', 'Лесной']);
  assert.strictEqual(p.repl.get('Лесной'), p.repl.get('ЛЕСНАЯ'));
  // «Иван» находит «Ивану», но не «Иванова»
  assert.deepStrictEqual(row('Иван').forms, ['Ивану']);
  assert.strictEqual(p.repl.has('Иванова'), false);
  // ключ: одна пара на значение, исходное — самая частая форма, если введённой в тексте нет
  const backs = p.keyPairs.map(([, r]) => r);
  assert.strictEqual(new Set(backs).size, backs.length);
  assert.deepStrictEqual(p.keyPairs.find(([, r]) => r === row('Иван').repl), ['Ивану', row('Иван').repl]);

  const exact = D.plan(parts, { custom: ['Лесная'], customCase: true, customForms: false });
  assert.deepStrictEqual(exact.rows.find(r => r.value === 'Лесная').count, 1);
});

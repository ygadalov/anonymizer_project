/*
 * Обезличивание документов Word (.docx) — без DOM, работает и в браузере, и в Node (тесты).
 *
 * .docx — ZIP-архив с XML. Текст лежит в элементах <w:t> внутри «прогонов» <w:r>; Word часто дробит одно
 * слово на несколько прогонов (проверка орфографии, правки), поэтому поиск идёт по склеенному тексту абзаца,
 * а замена раскладывается обратно: новый текст встаёт в первый прогон совпадения (с его оформлением),
 * остальные части совпадения из соседних прогонов убираются.
 *
 * Все замены — пары «исходное → замена»: они хранятся в ключе, и восстановление — обратная подстановка.
 */
(function (root) {
  'use strict';
  // движок Redact (src/engine.js) берётся в момент вызова: на странице он может подключаться позже этого файла
  const engine = () => root.Anon || require('./engine.js');

  // ---------- XML ----------
  const unescapeXml = s => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e.toLowerCase()];
  });
  const escapeXml = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escapeAttr = s => escapeXml(s).replace(/"/g, '&quot;');
  const escapeRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  // Части документа с текстом: основной текст, колонтитулы, сноски, примечания
  const TEXT_PART = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments(Extended)?)\.xml$/;

  // Разметка части на сегменты: сегмент — текст одного абзаца (или кода поля), склеенный из прогонов.
  // Табуляция и разрыв строки становятся символом \u0001: ни один шаблон поиска его не захватывает.
  const SEP = '\u0001';
  const TOKENS = /<w:(t|instrText)(\s[^>]*)?>([\s\S]*?)<\/w:\1>|<w:(?:t|instrText)(?:\s[^>]*)?\/>|<\/w:p>|<w:p[\s>]|<w:tab\/>|<w:br\b[^>]*\/>|<w:cr\/>/g;
  function segmentsOf(xml) {
    const segs = [];
    let cur = null;
    const close = () => { if (cur && cur.nodes.length) segs.push(cur); cur = null; };
    const open = kind => { if (!cur || cur.kind !== kind) { close(); cur = { kind, nodes: [], text: '' }; } };
    TOKENS.lastIndex = 0;
    let m;
    while ((m = TOKENS.exec(xml))) {
      const tok = m[0];
      if (m[1]) {
        const kind = m[1] === 't' ? 'text' : 'field';
        open(kind);
        const text = unescapeXml(m[3]);
        cur.nodes.push({ start: m.index, end: m.index + tok.length, tag: m[1], attrs: m[2] || '', from: cur.text.length, text });
        cur.text += text;
      } else if (/^<w:(tab|br|cr)/.test(tok)) {
        if (cur) cur.text += SEP;
      } else if (tok.startsWith('<w:p') || tok === '</w:p>') {
        close();
      }
    }
    close();
    return segs;
  }

  // Применить к XML части замены: finder(text) → [{ start, end, repl }] без пересечений, по возрастанию
  function rewrite(xml, finder) {
    const edits = [];   // { start, end, xml } — замены элементов <w:t> целиком
    let count = 0;
    for (const seg of segmentsOf(xml)) {
      const found = finder(seg.text, seg.kind);
      if (!found.length) continue;
      count += found.length;
      // новый текст каждого узла: символы вне совпадений остаются, совпадение целиком уходит в узел его начала
      const out = seg.nodes.map(() => '');
      const owner = i => { let k = 0; while (k + 1 < seg.nodes.length && seg.nodes[k + 1].from <= i) k++; return k; };
      let pos = 0;
      const copy = (a, b) => { for (let i = a; i < b; i++) { const k = owner(i); out[k] += seg.text[i]; } };
      for (const f of found) {
        copy(pos, f.start);
        out[owner(f.start)] += f.repl;
        pos = f.end;
      }
      copy(pos, seg.text.length);
      seg.nodes.forEach((n, k) => {
        const t = out[k].split(SEP).join('');
        if (t === n.text) return;
        let attrs = n.attrs;
        if (/^\s|\s$/.test(t) && !/xml:space=/.test(attrs)) attrs += ' xml:space="preserve"';
        edits.push({ start: n.start, end: n.end, xml: `<w:${n.tag}${attrs}>${escapeXml(t)}</w:${n.tag}>` });
      });
    }
    if (!edits.length) return { xml, count };
    edits.sort((a, b) => a.start - b.start);
    let res = '', at = 0;
    for (const e of edits) { res += xml.slice(at, e.start) + e.xml; at = e.end; }
    return { xml: res + xml.slice(at), count };
  }

  // Поиск по словарю «исходное → замена»: самые длинные ключи первыми, совпадение — отдельное слово/число
  function dictFinder(dict) {
    const keys = [...dict.keys()].filter(Boolean).sort((a, b) => b.length - a.length);
    if (!keys.length) return () => [];
    const re = new RegExp(`(?<![\\p{L}\\p{N}])(?:${keys.map(escapeRe).join('|')})(?![\\p{L}\\p{N}])`, 'gu');
    return text => {
      const out = [];
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(text))) {
        if (!m[0]) { re.lastIndex++; continue; }
        out.push({ start: m.index, end: m.index + m[0].length, repl: dict.get(m[0]) });
      }
      return out;
    };
  }

  // ---------- Поиск данных в тексте ----------
  const MONTHS = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const L = 'А-ЯЁа-яё';
  const innValid = s => {
    const d = [...s].map(Number);
    const chk = w => w.reduce((a, x, i) => a + x * d[i], 0) % 11 % 10;
    if (d.length === 10) return chk([2, 4, 10, 3, 5, 9, 4, 6, 8]) === d[9];
    if (d.length === 12) return chk([7, 2, 4, 10, 3, 5, 9, 4, 6, 8]) === d[10] && chk([3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]) === d[11];
    return false;
  };
  const before = (text, i, n) => text.slice(Math.max(0, i - n), i);
  const ORG_FORMS = 'ООО|ОАО|ЗАО|ПАО|НАО|АО|АНО|НКО|ГУП|МУП|ФГУП|ТОО|Общество с ограниченной ответственностью|Акционерное общество|Публичное акционерное общество|Закрытое акционерное общество|Открытое акционерное общество';
  const PATRONYMIC = '(?:ович|евич|ич|овн|евн|ичн|инич)[а-я]{0,3}';

  // Типы данных, которые находит поиск, — в порядке показа в интерфейсе
  const TYPES = {
    org: 'Организации', person: 'ФИО', inn: 'ИНН', ogrn: 'ОГРН', kpp: 'КПП', bik: 'БИК', account: 'Банковские счета',
    snils: 'СНИЛС', passport: 'Паспорта', phone: 'Телефоны', email: 'Эл. почта', amount: 'Суммы', words: 'Суммы прописью',
    date: 'Даты', custom: 'Добавлено вручную',
  };

  // Совпадения одного сегмента: [{ start, end, type, value, info }]
  function detect(text) {
    const hits = [];
    const add = (start, value, type, info) => hits.push({ start, end: start + value.length, value, type, info });
    const each = (re, fn) => { re.lastIndex = 0; let m; while ((m = re.exec(text))) { fn(m); if (!m[0]) re.lastIndex++; } };

    each(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, m => add(m.index, m[0], 'email'));
    each(/(?<![\d+])(?:\+7|8)[\s\u00a0-]*\(?\d{3}\)?[\s\u00a0-]*\d{3}[\s\u00a0-]*\d{2}[\s\u00a0-]*\d{2}(?!\d)/g, m => add(m.index, m[0], 'phone'));
    each(/(?<!\d)\d{3}-\d{3}-\d{3}[\s\u00a0-]\d{2}(?!\d)/g, m => add(m.index, m[0], 'snils'));
    each(/(?<!\d)\d{20}(?!\d)/g, m => add(m.index, m[0], 'account'));
    each(/(?<!\d)\d{2}[\s\u00a0]?\d{2}[\s\u00a0]+(?:№[\s\u00a0]*)?\d{6}(?!\d)/g, m => {
      if (/паспорт/i.test(before(text, m.index, 60))) add(m.index, m[0], 'passport');
    });
    each(/ОГРН(?:ИП)?[\s:\u00a0]*(\d{15}|\d{13})(?!\d)/g, m => add(m.index + m[0].length - m[1].length, m[1], 'ogrn'));
    each(/КПП[\s:\u00a0]*(\d{9})(?!\d)/g, m => add(m.index + m[0].length - m[1].length, m[1], 'kpp'));
    each(/БИК[\s:\u00a0]*(\d{9})(?!\d)/g, m => add(m.index + m[0].length - m[1].length, m[1], 'bik'));
    each(/(?<!\d)(\d{12}|\d{10})(?!\d)/g, m => {
      if (innValid(m[1]) || /ИНН[\s:\u00a0/№]*$/.test(before(text, m.index, 14))) add(m.index, m[1], 'inn');
    });
    // организации: форма + название в кавычках или заглавными буквами; заменяется само название
    each(new RegExp(`(?:${ORG_FORMS})[\\s\\u00a0]*[«"“„]([^«»"“”„${SEP}]{2,80})[»"”“]`, 'g'), m => {
      add(m.index + m[0].length - m[1].length - 1, m[1], 'org');
    });
    each(new RegExp(`(?<![${L}])(?:ООО|ОАО|ЗАО|ПАО|АО)[\\s\\u00a0]+([А-ЯЁ][А-ЯЁ-]{2,}(?:[\\s\\u00a0][А-ЯЁ][А-ЯЁ-]{2,})?)(?![${L}])`, 'g'), m => {
      add(m.index + m[0].length - m[1].length, m[1], 'org');
    });
    // ФИО: полное с отчеством (в любом падеже) и с инициалами
    each(new RegExp(`(?<![${L}])([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)[\\s\\u00a0]+([А-ЯЁ][а-яё]+)[\\s\\u00a0]+([А-ЯЁ][а-яё]*${PATRONYMIC})(?![${L}])`, 'g'), m => {
      add(m.index, m[0], 'person', { kind: 'full', surname: m[1], name: m[2], patronymic: m[3] });
    });
    each(new RegExp(`(?<![${L}])([А-ЯЁ][а-яё]{1,}(?:-[А-ЯЁ][а-яё]+)?)[\\s\\u00a0]+([А-ЯЁ])\\.[\\s\\u00a0]?([А-ЯЁ])\\.`, 'g'), m => {
      add(m.index, m[0], 'person', { kind: 'after', surname: m[1], i1: m[2], i2: m[3], tight: !/\.[\s\u00a0][А-ЯЁ]\.$/.test(m[0]) });
    });
    each(new RegExp(`(?<![${L}])([А-ЯЁ])\\.[\\s\\u00a0]?([А-ЯЁ])\\.[\\s\\u00a0]?([А-ЯЁ][а-яё]{1,}(?:-[А-ЯЁ][а-яё]+)?)(?![${L}])`, 'g'), m => {
      add(m.index, m[0], 'person', { kind: 'before', surname: m[3], i1: m[1], i2: m[2], tight: /^[А-ЯЁ]\.[А-ЯЁ]\./.test(m[0]) });
    });
    // суммы с валютой и сумма прописью в скобках сразу после неё
    each(/(?<![\d,.])(\d{1,3}(?:[ \u00a0]\d{3})+|\d+)(?:,(\d{1,2}))?(?=[\s\u00a0]*(?:руб|₽|RUB|USD|EUR|\$|€|долл))/g, m => {
      add(m.index, m[0], 'amount');
      const rest = text.slice(m.index + m[0].length);
      const w = /^[\s\u00a0]*(?:руб(?:лей|ля|ль)?\.?|₽|RUB)?[\s\u00a0]*\(([^()]{3,200})\)/.exec(rest);
      // в скобках — слова (копейки бывают цифрами: «… рублей 00 копеек»)
      if (w && /[а-яё]{3,}[\s\u00a0]+[а-яё]{3,}/i.test(w[1])) add(m.index + m[0].length + w[0].length - w[1].length - 1, w[1], 'words');
    });
    // даты: 12.03.2025 и «12» марта 2025
    each(/(?<![\d.])(\d{1,2})\.(\d{1,2})\.(\d{4})(?![\d])/g, m => {
      if (+m[2] >= 1 && +m[2] <= 12 && +m[1] >= 1 && +m[1] <= 31) add(m.index, m[0], 'date');
    });
    each(new RegExp(`(?<!\\d)([«"]?)(\\d{1,2})([»"]?)[\\s\\u00a0]+(${MONTHS.join('|')})[\\s\\u00a0]+(\\d{4})`, 'g'), m => add(m.index, m[0], 'date'));

    // пересечения: раньше начавшееся и более длинное побеждает
    hits.sort((a, b) => a.start - b.start || (b.end - b.start) - (a.end - a.start));
    const out = [];
    let end = -1;
    for (const h of hits) { if (h.start >= end) { out.push(h); end = h.end; } }
    return out;
  }

  // ---------- Генераторы замен ----------
  // Вымышленные ФИО: только формы с правильным склонением (фамилии на -ов/-ев/-ин, имена на согласную и на -а)
  const FAKE_SURNAMES = ['Лебедев', 'Соколов', 'Морозов', 'Волков', 'Зайцев', 'Павлов', 'Семёнов', 'Голубев', 'Виноградов', 'Богданов',
    'Воробьёв', 'Фёдоров', 'Михайлов', 'Беляев', 'Тарасов', 'Белов', 'Комаров', 'Орлов', 'Киселёв', 'Макаров', 'Андреев', 'Ковалёв',
    'Ильин', 'Гусев', 'Титов', 'Кузьмин', 'Кудрявцев', 'Баранов', 'Куликов', 'Алексеев', 'Степанов', 'Яковлев', 'Сорокин', 'Сергеев',
    'Романов', 'Захаров', 'Борисов', 'Королёв', 'Герасимов', 'Пономарёв', 'Григорьев', 'Лазарев', 'Медведев', 'Ершов', 'Никитин'];
  const FAKE_MALE = ['Олег', 'Глеб', 'Роман', 'Борис', 'Денис', 'Вадим', 'Степан', 'Тимур', 'Руслан', 'Артём', 'Максим', 'Антон', 'Матвей'];
  const FAKE_FEMALE = ['Анна', 'Ирина', 'Елена', 'Марина', 'Алина', 'Вера', 'Нина', 'Полина', 'Кристина', 'Светлана', 'Татьяна', 'Галина'];
  const FAKE_PATRO = ['Сергеев', 'Андреев', 'Олегов', 'Петров', 'Иванов', 'Михайлов', 'Романов', 'Борисов', 'Викторов', 'Алексеев', 'Павлов'];
  const CASES = ['nom', 'gen', 'dat', 'acc', 'ins', 'pre'];
  // окончания по падежам
  const END = {
    surM: { nom: '', gen: 'а', dat: 'у', acc: 'а', ins: 'ым', pre: 'е' },
    surF: { nom: 'а', gen: 'ой', dat: 'ой', acc: 'у', ins: 'ой', pre: 'ой' },
    patM: { nom: 'ич', gen: 'ича', dat: 'ичу', acc: 'ича', ins: 'ичем', pre: 'иче' },
    patF: { nom: 'на', gen: 'ны', dat: 'не', acc: 'ну', ins: 'ной', pre: 'не' },
  };
  function declineName(name, female, cs) {
    if (female) {   // все женские имена списка оканчиваются на -а после н/р/л/т
      return name.slice(0, -1) + { nom: 'а', gen: 'ы', dat: 'е', acc: 'у', ins: 'ой', pre: 'е' }[cs];
    }
    if (name.endsWith('й')) return name.slice(0, -1) + { nom: 'й', gen: 'я', dat: 'ю', acc: 'я', ins: 'ем', pre: 'е' }[cs];
    return name + { nom: '', gen: 'а', dat: 'у', acc: 'а', ins: 'ом', pre: 'е' }[cs];
  }
  // падеж и пол по окончанию отчества
  function caseFromPatronymic(p) {
    const f = /(овн|евн|ичн|инич)(а|ы|е|у|ой)$/.exec(p);
    if (f) return { female: true, cs: { а: 'nom', ы: 'gen', е: 'dat', у: 'acc', ой: 'ins' }[f[2]] };
    const m = /ич(|а|у|ем|е)$/.exec(p);
    if (m) return { female: false, cs: { '': 'nom', а: 'gen', у: 'dat', ем: 'ins', е: 'pre' }[m[1]] };
    return { female: false, cs: 'nom' };
  }
  // основа фамилии (для узнавания одного человека в разных падежах) и падеж по окончанию
  function surnameInfo(s, female) {
    const m = /^(.+?(?:ов|ев|ёв|ин|ын))(а|у|ым|е|ой|ом)?$/.exec(s);
    if (!m) return { stem: s, cs: 'nom', regular: false };
    const e = m[2] || '';
    if (female === true) return { stem: m[1], cs: { а: 'nom', ой: 'gen', у: 'acc' }[e] || 'nom', regular: true };
    if (female === false) return { stem: m[1], cs: { '': 'nom', а: 'gen', у: 'dat', ым: 'ins', е: 'pre', ом: 'ins' }[e] || 'nom', regular: true };
    // пол неизвестен (только инициалы): -а — скорее женская в им. падеже, -ой — женская в косвенном
    if (e === 'ой') return { stem: m[1], cs: 'gen', female: true, regular: true };
    if (e === 'а') return { stem: m[1], cs: 'nom', female: true, regular: true };
    return { stem: m[1], cs: { '': 'nom', у: 'dat', ым: 'ins', е: 'pre', ом: 'ins' }[e] || 'nom', female: false, regular: true };
  }

  // Свои значения: слова на кириллице ищутся по основе с падежными окончаниями («Вересков» → «Верескова», «Вересковым»;
  // «Москва» → «Москве», «Москвой»). Окончания «-ов», «-ев», «-ин» не добавляются: «Иван» не должен найти «Иванов»
  const CUT_END = ['ая', 'яя', 'ое', 'ее', 'ый', 'ий', 'ой', 'ые', 'ие', 'а', 'я', 'о', 'е', 'ь', 'й', 'ы', 'и'];
  const CASE_END = ['а', 'я', 'о', 'е', 'и', 'ы', 'у', 'ю', 'ь', 'й', 'ой', 'ей', 'ою', 'ею', 'ом', 'ем', 'ам', 'ям', 'ами', 'ями',
    'ах', 'ях', 'ым', 'им', 'ую', 'юю', 'ая', 'яя', 'ое', 'ее', 'ые', 'ие', 'ого', 'его', 'ому', 'ему', 'ых', 'их', 'ыми', 'ими'];
  function customPattern(value, { matchCase = false, forms = true } = {}) {
    const words = value.split(/[\s\u00a0]+/).filter(Boolean).map(w => {
      if (!forms || !/^[А-ЯЁа-яё]{4,}$/.test(w)) return escapeRe(w);
      const end = CUT_END.find(e => w.toLowerCase().endsWith(e) && w.length - e.length >= 3);
      const stem = end ? w.slice(0, -end.length) : w;
      return `${escapeRe(stem)}(?:${CASE_END.join('|')})?`;
    });
    return new RegExp(`(?<![\\p{L}\\p{N}])${words.join('[\\s\\u00a0]+')}(?![\\p{L}\\p{N}])`, matchCase ? 'gu' : 'giu');
  }

  // Сегменты с текстом, кроме удаляемых разделов и абзацев (cut — диапазоны [начало, конец) в word/document.xml)
  function liveSegs(parts, cut = []) {
    const out = [];
    for (const p of parts) for (const s of p.segs) {
      if (s.kind !== 'text') continue;
      if (cut.length && p.name === 'word/document.xml' && cut.some(([a, b]) => s.nodes[0].start >= a && s.nodes[0].start < b)) continue;
      out.push([p, s]);
    }
    return out;
  }

  function plan(parts, opts) {
    const seed = opts.seed || 'redact-docx';
    const factor = Number(opts.factor) || 1, shift = Math.round(Number(opts.shiftDays) || 0);
    const enabled = opts.types || Object.fromEntries(Object.keys(TYPES).map(t => [t, true]));
    const skip = new Set(opts.skip || []);
    const live = liveSegs(parts, opts.cut);
    const allText = live.map(([, s]) => s.text).join('\n');

    // 1. Что найдено: значение → { type, count, info }
    const found = new Map();
    const note = (h) => {
      const f = found.get(h.value);
      if (f) f.count++; else found.set(h.value, { type: h.type, count: 1, info: h.info, value: h.value });
    };
    for (const [, s] of live) detect(s.text).forEach(note);
    // «Добавить своё»: без учёта регистра (по умолчанию) и во всех падежах — все найденные формы получают одну замену
    const claimed = new Set(found.keys());
    for (const c of opts.custom || []) {
      const v = String(c).trim();
      if (!v || found.has(v)) continue;
      const re = customPattern(v, { matchCase: !!opts.customCase, forms: opts.customForms !== false });
      const forms = new Map();
      let m;
      while ((m = re.exec(allText))) {
        if (!m[0]) { re.lastIndex++; continue; }
        if (!claimed.has(m[0])) forms.set(m[0], (forms.get(m[0]) || 0) + 1);
      }
      forms.forEach((n, f) => claimed.add(f));
      const count = [...forms.values()].reduce((a, b) => a + b, 0);
      // в ключ — само значение, если оно есть в тексте, иначе самая частая форма: по ключу вернётся она
      const keyValue = forms.has(v) || !forms.size ? v : [...forms].sort((a, b) => b[1] - a[1])[0][0];
      found.set(v, { type: 'custom', count, value: v, forms: [...forms.keys()], keyValue });
    }

    // 2. Замены. Ни одна замена не совпадает с исходным значением, с другой заменой или с текстом документа
    const used = new Set(found.keys());
    const taken = s => used.has(s) || allText.includes(s);
    const rng = engine().makeRng(seed, 'docx');
    const rnd = (lo, hi) => lo + Math.floor(rng() * (hi - lo + 1));
    const pickOf = arr => arr[Math.floor(rng() * arr.length)];
    const counters = {};
    const numbered = (prefix, type) => {
      let s;
      do { counters[type] = (counters[type] || 0) + 1; s = `${prefix} ${String(counters[type]).padStart(3, '0')}`; } while (taken(s));
      return s;
    };
    const digits = (v, keep = 1, fixLast) => {
      let s, tries = 0;
      do {
        let k = 0;
        s = v.replace(/\d/g, d => (k++ < keep ? d : String(rnd(0, 9))));
        if (fixLast) s = fixLast(s);
        tries++;
      } while ((s === v || taken(s)) && tries < 200);
      return s;
    };
    // ИНН — с заведомо неверной контрольной цифрой: такого номера нет у реальной организации или человека
    const fakeInn = v => digits(v, 2, s => {
      const d = s.slice(0, -1);
      for (let x = 0; x < 10; x++) { const c = d + x; if (!innValid(c)) return c; }
      return s;
    });
    // люди: один человек (основа фамилии + инициалы) — один вымышленный, в нужном падеже и роде
    const persons = new Map();
    const fakeFor = (key, female) => {
      if (persons.has(key)) return persons.get(key);
      let p, tries = 0;
      do {
        p = { sur: pickOf(FAKE_SURNAMES), name: female ? pickOf(FAKE_FEMALE) : pickOf(FAKE_MALE), patro: pickOf(FAKE_PATRO), female };
        tries++;
      } while (tries < 300 && [...persons.values()].some(q => q.sur === p.sur && q.name[0] === p.name[0] && q.patro[0] === p.patro[0]));
      persons.set(key, p);
      return p;
    };
    // «Лебедев» → «Лебедеву», «Лебедева» → «Лебедевой»; отчество из основы: «Сергеев» → «Сергеевичу», «Сергеевной»
    const surForm = (p, cs) => p.sur + (p.female ? END.surF[cs] : END.surM[cs]);
    const fullPatro = (p, cs) => p.patro + (p.female ? END.patF[cs] : END.patM[cs]);
    // ключ человека по полной форме и по инициалам, пол — по отчеству или фамилии
    const personKey = f => {
      const i = f.info;
      if (i.kind === 'full') {
        const c = caseFromPatronymic(i.patronymic);
        const s = surnameInfo(i.surname, c.female);
        return { key: `${s.stem}|${i.name[0]}|${i.patronymic[0]}`, female: c.female, cs: c.cs };
      }
      const s = surnameInfo(i.surname);
      return { key: `${s.stem}|${i.i1}|${i.i2}`, female: !!s.female, cs: s.cs };
    };
    // сначала полные формы — по ним точнее определяется пол
    const order = [...found.values()].sort((a, b) => (a.type === 'person' && a.info.kind === 'full' ? 0 : 1) - (b.type === 'person' && b.info.kind === 'full' ? 0 : 1));
    const genderOf = new Map();
    for (const f of order) if (f.type === 'person' && f.info.kind === 'full') { const k = personKey(f); genderOf.set(k.key, k.female); }

    const amounts = new Map();   // новое значение суммы → исходное (замены сумм уникальны)
    const repl = new Map();      // исходное → замена (только включённые)
    const rows = [];
    for (const f of order) {
      const on = enabled[f.type] !== false && !skip.has(f.value);
      let r;
      switch (f.type) {
        case 'org': r = numbered('Контрагент', 'org'); break;
        case 'custom': r = numbered('Значение', 'custom'); break;
        case 'words': r = numbered('сумма прописью', 'words'); break;
        case 'email': { let s; do { counters.email = (counters.email || 0) + 1; s = `contact${String(counters.email).padStart(3, '0')}@example.com`; } while (taken(s)); r = s; break; }
        case 'inn': r = fakeInn(f.value); break;
        case 'phone': r = digits(f.value, 1); break;
        case 'account': r = digits(f.value, 5); break;
        case 'bik': r = digits(f.value, 2); break;
        case 'kpp': case 'ogrn': case 'snils': case 'passport': r = digits(f.value, 1); break;
        case 'person': {
          const k = personKey(f);
          const female = genderOf.has(k.key) ? genderOf.get(k.key) : k.female;
          const p = fakeFor(k.key, female);
          const i = f.info;
          if (i.kind === 'full') r = `${surForm(p, k.cs)} ${declineName(p.name, female, k.cs)} ${fullPatro(p, k.cs)}`;
          else {
            const ini = i.tight ? `${p.name[0]}.${p.patro[0]}.` : `${p.name[0]}. ${p.patro[0]}.`;
            r = i.kind === 'after' ? `${surForm(p, k.cs)} ${ini}` : `${ini} ${surForm(p, k.cs)}`;
          }
          break;
        }
        case 'amount': {
          const m = /^(\d{1,3}(?:([ \u00a0])\d{3})+|\d+)(?:,(\d{1,2}))?$/.exec(f.value);
          const sep = m[2] || '', dec = m[3];
          let v = Number(m[1].replace(/[ \u00a0]/g, '') + (dec ? '.' + dec : ''));
          let nv = Math.round(v * factor * 100) / 100;
          // формат как у исходной: тот же разделитель тысяч, копейки — если были или появились после умножения
          const fmt = x => {
            const [ip, fp] = x.toFixed(2).split('.');
            const g = ip.replace(/\B(?=(\d{3})+(?!\d))/g, sep);
            return dec === undefined && fp === '00' ? g : `${g},${fp}`;
          };
          let s = fmt(nv);
          while (s === f.value || amounts.has(s) || taken(s)) { nv = Math.round((nv + 0.01) * 100) / 100; s = fmt(nv); }
          amounts.set(s, f.value);
          r = s;
          break;
        }
        case 'date': r = shiftDate(f.value, shift); break;
        default: r = numbered('Значение', 'custom');
      }
      if (r === f.value) r = numbered('Значение', 'custom');
      used.add(r);
      const forms = f.forms ? f.forms.filter(x => x !== f.value) : [];
      rows.push({ value: f.value, type: f.type, count: f.count, repl: r, on, forms, keyValue: f.keyValue || f.value });
      if (on) {
        repl.set(f.value, r);
        forms.forEach(x => repl.set(x, r));
      }
    }
    rows.sort((a, b) => Object.keys(TYPES).indexOf(a.type) - Object.keys(TYPES).indexOf(b.type) || b.count - a.count);
    // пары для ключа: одна на значение, без падежных форм — иначе у одной замены было бы несколько исходных
    const keyPairs = rows.filter(r => r.on).map(r => [r.keyValue, r.repl]);
    return { rows, repl, keyPairs };
  }

  function shiftDate(v, days) {
    const DAY = 86400000;
    let m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(v);
    if (m) {
      const d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]) + days * DAY);
      const pad = (x, w) => String(x).padStart(w, '0');
      return `${m[1].length === 2 ? pad(d.getUTCDate(), 2) : d.getUTCDate()}.${m[2].length === 2 ? pad(d.getUTCMonth() + 1, 2) : d.getUTCMonth() + 1}.${d.getUTCFullYear()}`;
    }
    m = new RegExp(`^([«"]?)(\\d{1,2})([»"]?)([\\s\\u00a0]+)(${MONTHS.join('|')})([\\s\\u00a0]+)(\\d{4})$`).exec(v);
    if (m) {
      const d = new Date(Date.UTC(+m[7], MONTHS.indexOf(m[5]), +m[2]) + days * DAY);
      const day = m[2].length === 2 ? String(d.getUTCDate()).padStart(2, '0') : String(d.getUTCDate());
      return `${m[1]}${day}${m[3]}${m[4]}${MONTHS[d.getUTCMonth()]}${m[6]}${d.getUTCFullYear()}`;
    }
    return v;
  }

  // ---------- Документ целиком (zip — экземпляр JSZip) ----------
  async function readParts(zip) {
    const names = Object.keys(zip.files).filter(n => TEXT_PART.test(n)).sort();
    const parts = [];
    for (const name of names) {
      const xml = await zip.file(name).async('string');
      parts.push({ name, xml, segs: segmentsOf(xml) });
    }
    return parts;
  }

  // Анализ: что найдено и чем будет заменено. Плюс то, что Redact не может прочитать (картинки, вложения)
  async function analyzeDocx(zip, opts = {}) {
    const parts = await readParts(zip);
    if (!parts.some(p => p.name === 'word/document.xml')) throw new Error('это не документ Word: в архиве нет word/document.xml');
    const { rows, repl } = plan(parts, opts);
    const styles = zip.file('word/styles.xml') ? headingStyles(await zip.file('word/styles.xml').async('string')) : new Map();
    const names = Object.keys(zip.files);
    const doc = parts.find(p => p.name === 'word/document.xml').xml;
    const info = {
      images: names.filter(n => /^word\/media\//.test(n) && !zip.files[n].dir).length,
      embeddings: names.filter(n => /^word\/embeddings\//.test(n) && !zip.files[n].dir).length,
      comments: (parts.find(p => p.name === 'word/comments.xml') || { xml: '' }).xml.split('<w:comment ').length - 1,
      revisions: (doc.match(/<w:(del|ins|moveFrom|moveTo)\b/g) || []).length,
      paragraphs: parts.reduce((n, p) => n + p.segs.filter(s => s.kind === 'text').length, 0),
    };
    return { parts, rows, repl, info, blocks: blocksOf(doc, styles) };
  }

  // Записанные исправления: удалённый текст убирается из файла, вставленный становится обычным
  const acceptRevisions = xml => xml
    .replace(/<w:(del|moveFrom)\b[^>]*\/>/g, '')
    .replace(/<w:(del|moveFrom)\b[^>]*>[\s\S]*?<\/w:\1>/g, '')
    .replace(/<w:(ins|moveTo)\b[^>]*\/>/g, '')
    .replace(/<\/?w:(ins|moveTo)\b[^>]*>/g, '')
    .replace(/<w:(moveFromRange|moveToRange)(Start|End)\b[^>]*\/>/g, '')
    .replace(/<w:(rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)\b[^>]*>[\s\S]*?<\/w:\1>/g, '');
  // Примечания: метки в тексте и сами примечания
  const dropCommentMarks = xml => xml
    .replace(/<w:comment(RangeStart|RangeEnd)\b[^>]*\/>/g, '')
    .replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:commentReference\b[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, '');
  // Авторы правок и примечаний
  const dropAuthors = xml => xml.replace(/\sw:author="[^"]*"/g, ' w:author="Автор"').replace(/\sw:initials="[^"]*"/g, ' w:initials=""');

  // Обезличивание: zip изменяется на месте. opts: repl (Map), clean: { revisions, comments, meta },
  // cut — номера удаляемых блоков (blocksOf), mode — 'full' или 'mark'
  async function redactDocx(zip, { repl, clean = {}, cut = new Set(), mode = 'full' }) {
    const finder = dictFinder(repl);
    let replaced = 0, removedFiles = 0;
    // сначала удаление разделов: номера блоков — по исходному тексту, тому же, что видел пользователь
    if (cut.size) {
      const doc = await zip.file('word/document.xml').async('string');
      const styles = zip.file('word/styles.xml') ? headingStyles(await zip.file('word/styles.xml').async('string')) : new Map();
      const c = cutBlocks(doc, blocksOf(doc, styles), cut, mode);
      const after = await cleanAfterCut(zip, c.xml, c.removed);
      removedFiles = after.removedFiles;
      zip.file('word/document.xml', after.xml);
    }
    for (const name of Object.keys(zip.files).filter(n => TEXT_PART.test(n))) {
      let xml = await zip.file(name).async('string');
      if (clean.revisions) xml = acceptRevisions(xml);
      if (clean.comments) {
        if (/^word\/comments/.test(name)) xml = xml.replace(/<w:comment\b[\s\S]*?<\/w:comment>/g, '').replace(/<w15:commentEx\b[^>]*\/>/g, '');
        else xml = dropCommentMarks(xml);
      }
      xml = dropAuthors(xml);
      const r = rewrite(xml, finder);
      replaced += r.count;
      zip.file(name, r.xml);
    }
    // ссылки mailto: и другие адреса в связях документа
    for (const name of Object.keys(zip.files).filter(n => /^word\/_rels\/.*\.rels$/.test(n))) {
      let xml = await zip.file(name).async('string');
      xml = xml.replace(/Target="([^"]*)"/g, (m, t) => {
        const plain = unescapeXml(t);
        const out = plain.replace(/(mailto:)(.+)/i, (x, p, addr) => p + (repl.get(addr) || addr));
        return out === plain ? m : `Target="${escapeAttr(out)}"`;
      });
      zip.file(name, xml);
    }
    if (zip.file('word/people.xml')) zip.file('word/people.xml', (await zip.file('word/people.xml').async('string')).replace(/<w15:person\b[\s\S]*?<\/w15:person>/g, ''));
    if (clean.meta) await clearMeta(zip);
    return { replaced, removedFiles };
  }

  async function clearMeta(zip) {
    const core = zip.file('docProps/core.xml');
    if (core) {
      let xml = await core.async('string');
      for (const tag of ['dc:creator', 'cp:lastModifiedBy', 'dc:title', 'dc:subject', 'dc:description', 'cp:keywords', 'cp:category']) {
        xml = xml.replace(new RegExp(`<${tag}(\\s[^>]*)?>[\\s\\S]*?</${tag}>`, 'g'), `<${tag}$1></${tag}>`);
      }
      zip.file('docProps/core.xml', xml);
    }
    const app = zip.file('docProps/app.xml');
    if (app) {
      let xml = await app.async('string');
      for (const tag of ['Company', 'Manager']) xml = xml.replace(new RegExp(`<${tag}>[\\s\\S]*?</${tag}>`, 'g'), `<${tag}></${tag}>`);
      zip.file('docProps/app.xml', xml);
    }
    // свои свойства документа: сам файл остаётся (на него ссылается архив), его значения убираются
    const custom = zip.file('docProps/custom.xml');
    if (custom) zip.file('docProps/custom.xml', (await custom.async('string')).replace(/<property\b[\s\S]*?<\/property>/g, ''));
  }

  // Восстановление: обратная подстановка «замена → исходное» во всех частях с текстом и в ссылках
  async function restoreDocx(zip, back) {
    const finder = dictFinder(back);
    let restored = 0;
    for (const name of Object.keys(zip.files).filter(n => TEXT_PART.test(n))) {
      const r = rewrite(await zip.file(name).async('string'), finder);
      restored += r.count;
      zip.file(name, r.xml);
    }
    for (const name of Object.keys(zip.files).filter(n => /^word\/_rels\/.*\.rels$/.test(n))) {
      let xml = await zip.file(name).async('string');
      xml = xml.replace(/Target="mailto:([^"]*)"/gi, (m, a) => { const o = back.get(unescapeXml(a)); return o ? `Target="mailto:${escapeAttr(o)}"` : m; });
      zip.file(name, xml);
    }
    return { restored };
  }

  // Текст документа для предпросмотра: абзацы с найденным, разметка [исходное → замена]
  function previewParagraphs(parts, repl, limit = 60, cut = []) {
    const finder = dictFinder(repl);
    const out = [];
    for (const [p, s] of liveSegs(parts, cut)) {
      const f = finder(s.text);
      if (!f.length) continue;
      const pieces = [];
      let at = 0;
      for (const x of f) {
        if (x.start > at) pieces.push({ text: s.text.slice(at, x.start).split(SEP).join(' ') });
        pieces.push({ orig: s.text.slice(x.start, x.end), repl: x.repl });
        at = x.end;
      }
      if (at < s.text.length) pieces.push({ text: s.text.slice(at).split(SEP).join(' ') });
      out.push({ part: p.name, pieces });
      if (out.length >= limit) return out;
    }
    return out;
  }

  // Предпросмотр всего документа по порядку: абзацы основного текста (с заголовками и таблицами), затем колонтитулы,
  // сноски и примечания. Удалённое не показывается, а в режиме «С пометкой» — одна строка на месте удалённого участка.
  // onlyChanged — только абзацы с заменами. Абзац: { part, pieces, heading, table, mark }
  function previewDocument(parts, repl, { blocks = [], cut = new Set(), mode = 'full', onlyChanged = false, limit = 3000 } = {}) {
    const finder = dictFinder(repl);
    const out = [];
    const piecesOf = text => {
      const f = finder(text);
      const pieces = [];
      let at = 0;
      for (const x of f) {
        if (x.start > at) pieces.push({ text: text.slice(at, x.start).split(SEP).join(' ') });
        pieces.push({ orig: text.slice(x.start, x.end), repl: x.repl });
        at = x.end;
      }
      if (at < text.length) pieces.push({ text: text.slice(at).split(SEP).join(' ') });
      return { pieces, changed: f.length > 0 };
    };
    const push = (part, text, extra = {}) => {
      if (!text.split(SEP).join('').trim()) return;
      const { pieces, changed } = piecesOf(text);
      if (onlyChanged && !changed) return;
      out.push({ part, pieces, ...extra });
    };
    const doc = parts.find(p => p.name === 'word/document.xml');
    if (doc) {
      const segs = doc.segs.filter(s => s.kind === 'text');
      let k = 0;
      blocks.forEach((b, i) => {
        const inBlock = [];
        while (k < segs.length && segs[k].nodes[0].start < b.end) { if (segs[k].nodes[0].start >= b.start) inBlock.push(segs[k]); k++; }
        if (cut.has(i)) {
          if (mode === 'mark' && !onlyChanged && !cut.has(i - 1)) {
            let j = i; while (cut.has(j + 1)) j++;
            const label = b.level ? '[Раздел удалён]' : b.kind === 'tbl' && i === j ? '[Таблица удалена]' : i === j ? '[Абзац удалён]' : '[Фрагмент удалён]';
            out.push({ part: doc.name, pieces: [{ text: label }], mark: true });
          }
          return;
        }
        if (b.kind !== 'tbl') { for (const s of inBlock) push(doc.name, s.text, { heading: !!b.level }); return; }
        // таблица — построчно: ячейки строки через « · »
        const rows = [...doc.xml.slice(b.start, b.end).matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => [b.start + m.index, b.start + m.index + m[0].length]);
        const groups = rows.map(() => []);
        for (const s of inBlock) {
          const r = rows.findIndex(([a, e]) => s.nodes[0].start >= a && s.nodes[0].start < e);
          if (r >= 0) groups[r].push(s); else groups.push([s]);
        }
        for (const g of groups) {
          const cells = g.map(s => piecesOf(s.text)).filter(c => c.pieces.some(x => (x.text || x.orig || '').trim()));
          if (!cells.length || (onlyChanged && !cells.some(c => c.changed))) continue;
          out.push({ part: doc.name, table: true, pieces: cells.flatMap((c, n) => n ? [{ text: ' · ' }, ...c.pieces] : c.pieces) });
        }
      });
    }
    for (const p of parts) if (p !== doc) for (const s of p.segs) if (s.kind === 'text') push(p.name, s.text);
    return out.length > limit ? out.slice(0, limit) : out;
  }

  // ---------- Удаление разделов и абзацев ----------
  // Блок — абзац, таблица или элемент управления содержимым прямо в <w:body>. Раздел — заголовок и всё до следующего
  // заголовка того же или более высокого уровня. Уровень заголовка — из стиля («heading N», уровень структуры)
  // или по виду строки: «5. ОТВЕТСТВЕННОСТЬ СТОРОН», «Приложение» прописными — такие заголовки помечены как угаданные.
  function headingStyles(stylesXml) {
    const map = new Map();
    for (const m of (stylesXml || '').matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
      const name = /<w:name w:val="([^"]*)"/.exec(m[2]);
      const lvl = /<w:outlineLvl w:val="(\d)"/.exec(m[2]);
      const h = name && /^(?:heading|заголовок)\s*(\d)$/i.exec(name[1]);
      if (h) map.set(m[1], +h[1]);
      else if (lvl && +lvl[1] < 9) map.set(m[1], +lvl[1] + 1);
      else if (name && /^title$/i.test(name[1])) map.set(m[1], 1);
    }
    return map;
  }
  function guessHeading(text) {
    const t = text.split(SEP).join(' ').trim();
    if (t.length < 3 || t.length > 90 || /[,;:]$/.test(t)) return 0;
    if (/^\d{1,2}\.?\s+[А-ЯЁA-Z]/.test(t) && !/^\d+\.\d/.test(t) && !/\.$/.test(t)) return 1;
    const letters = t.replace(/[^А-ЯЁа-яёA-Za-z]/g, '');
    if (letters.length >= 4 && letters === letters.toUpperCase()) return 1;
    return 0;
  }
  function blocksOf(xml, styles = new Map()) {
    const body = /<w:body\b[^>]*>/.exec(xml);
    if (!body) return [];
    const blocks = [];
    const re = /<(\/?)w:(p|tbl|sdt|body)\b[^>]*?(\/?)>/g;
    re.lastIndex = body.index + body[0].length;
    let depth = 0, start = -1, kind = '', m;
    while ((m = re.exec(xml))) {
      const [tag, close, name, self] = m;
      if (name === 'body') break;
      if (!close && !self) { if (depth === 0) { start = m.index; kind = name; } depth++; }
      else if (close) { depth--; if (depth === 0) blocks.push({ start, end: m.index + tag.length, kind }); }
      else if (self && depth === 0) blocks.push({ start: m.index, end: m.index + tag.length, kind: name });
    }
    for (const b of blocks) {
      const x = xml.slice(b.start, b.end);
      b.text = segmentsOf(x).filter(s => s.kind === 'text').map(s => s.text.split(SEP).join(' ').trim()).filter(Boolean)
        .join(b.kind === 'tbl' ? ' · ' : ' ');
      b.sect = /<w:sectPr\b/.test(x);
      b.image = /<w:drawing\b|<w:pict\b|<w:object\b/.test(x);
      b.level = 0;
      if (b.kind === 'p') {
        const st = /<w:pStyle w:val="([^"]+)"/.exec(x);
        const ol = /<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*?<w:outlineLvl w:val="(\d)"/.exec(x);
        b.level = (st && styles.get(st[1])) || (ol && +ol[1] < 9 ? +ol[1] + 1 : 0);
        // строка оглавления повторяет текст заголовка, но заголовком не считается
        b.toc = /w:anchor="_Toc|PAGEREF\s+_Toc/.test(x);
        if (!b.level && b.text && !b.toc) { b.level = guessHeading(b.text); b.guessed = !!b.level; }
      }
    }
    return blocks;
  }
  // Конец раздела, который начинается заголовком i (не включительно)
  function sectionEnd(blocks, i) {
    const lvl = blocks[i].level;
    if (!lvl) return i + 1;
    let j = i + 1;
    while (j < blocks.length && !(blocks[j].level && blocks[j].level <= lvl)) j++;
    return j;
  }
  // Правила: блок, в тексте которого есть фраза (без учёта регистра), удаляется; если это заголовок — весь раздел
  function ruleCuts(blocks, rules) {
    const cut = new Set();
    const rs = (rules || []).map(r => r.trim().toLowerCase()).filter(r => r.length >= 2);
    if (!rs.length) return cut;
    blocks.forEach((b, i) => {
      const t = b.text.toLowerCase();
      if (!rs.some(r => t.includes(r))) return;
      for (let j = i, e = sectionEnd(blocks, i); j < e; j++) cut.add(j);
    });
    return cut;
  }
  // Диапазоны удаляемого в document.xml — чтобы «Что найдено» и предпросмотр не показывали удалённое
  const cutRanges = (blocks, cut) => [...cut].map(i => [blocks[i].start, blocks[i].end]);

  // Удаление блоков из document.xml. mode: 'full' — без следа, 'mark' — на месте каждого удалённого участка
  // абзац «[… удалён]» с оформлением первого абзаца (номер пункта при автонумерации сохраняется)
  function cutBlocks(xml, blocks, cut, mode) {
    const idx = [...cut].sort((a, b) => a - b);
    if (!idx.length) return { xml, removed: '' };
    const runs = [];
    for (const i of idx) { const r = runs[runs.length - 1]; if (r && r[1] === i - 1) r[1] = i; else runs.push([i, i]); }
    let out = '', at = 0, removed = '';
    for (const [a, b] of runs) {
      out += xml.slice(at, blocks[a].start);
      removed += xml.slice(blocks[a].start, blocks[b].end);
      const keep = [];
      // разрыв раздела (параметры страницы) хранится в абзаце — сам абзац уходит, разрыв остаётся пустым абзацем
      for (let i = a; i <= b; i++) if (blocks[i].sect && blocks[i].kind === 'p') {
        const pPr = /<w:pPr\b[\s\S]*?<\/w:pPr>/.exec(xml.slice(blocks[i].start, blocks[i].end));
        if (pPr) keep.push(`<w:p>${pPr[0]}</w:p>`);
      }
      if (mode === 'mark') {
        const first = blocks[a];
        const label = first.level ? '[Раздел удалён]' : first.kind === 'tbl' && a === b ? '[Таблица удалена]' : a === b ? '[Абзац удалён]' : '[Фрагмент удалён]';
        let pPr = first.kind === 'p' ? (/<w:pPr\b[\s\S]*?<\/w:pPr>/.exec(xml.slice(first.start, first.end)) || [''])[0] : '';
        pPr = pPr.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '').replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g, '');
        out += `<w:p>${pPr}<w:r><w:t>${label}</w:t></w:r></w:p>`;
      }
      out += keep.join('');
      at = blocks[b].end;
    }
    out += xml.slice(at);
    // в теле документа должен остаться хотя бы один абзац
    if (!/<w:body\b[^>]*>[\s\S]*<w:p[\s>/]/.test(out)) out = out.replace(/(<w:body\b[^>]*>)/, '$1<w:p/>');
    return { xml: out, removed };
  }

  // Хвосты удалённого: сноски, примечания, закладки, строки оглавления, картинки и вложенные объекты
  async function cleanAfterCut(zip, docXml, removed) {
    const ids = re => [...removed.matchAll(re)].map(m => m[1]);
    // сноски и концевые сноски
    for (const [kind, part] of [['footnote', 'word/footnotes.xml'], ['endnote', 'word/endnotes.xml']]) {
      const gone = ids(new RegExp(`<w:${kind}Reference\\b[^>]*w:id="(-?\\d+)"`, 'g'));
      if (!gone.length || !zip.file(part)) continue;
      let x = await zip.file(part).async('string');
      for (const id of gone) x = x.replace(new RegExp(`<w:${kind}\\b[^>]*w:id="${id}"[\\s\\S]*?</w:${kind}>`), '');
      zip.file(part, x);
    }
    // примечания к удалённому тексту и их метки, оставшиеся снаружи
    const comments = [...new Set(ids(/<w:comment(?:Reference|RangeStart|RangeEnd)\b[^>]*w:id="(\d+)"/g))];
    if (comments.length) {
      for (const id of comments) docXml = docXml.replace(new RegExp(`<w:comment(?:RangeStart|RangeEnd)\\b[^>]*w:id="${id}"[^>]*/>`, 'g'), '')
        .replace(new RegExp(`<w:r\\b[^>]*>(?:(?!</w:r>)[\\s\\S])*?<w:commentReference\\b[^>]*w:id="${id}"[^>]*/>(?:(?!</w:r>)[\\s\\S])*?</w:r>`, 'g'), '');
      if (zip.file('word/comments.xml')) {
        let x = await zip.file('word/comments.xml').async('string');
        for (const id of comments) x = x.replace(new RegExp(`<w:comment\\b[^>]*w:id="${id}"[\\s\\S]*?</w:comment>`), '');
        zip.file('word/comments.xml', x);
      }
    }
    // закладки удалённого текста; строки оглавления и перекрёстные ссылки на них показывают удалённые заголовки
    const marks = [...removed.matchAll(/<w:bookmarkStart\b[^>]*w:name="([^"]+)"/g)].map(m => m[1]);
    for (const name of marks) {
      const e = escapeRe(name);
      for (const re of [new RegExp(`w:anchor="${e}"`), new RegExp(`PAGEREF\\s+${e}\\b`), new RegExp(`REF\\s+${e}\\b`)]) {
        let m;
        while ((m = re.exec(docXml))) {
          const a = Math.max(docXml.lastIndexOf('<w:p>', m.index), docXml.lastIndexOf('<w:p ', m.index));
          const b = docXml.indexOf('</w:p>', m.index);
          if (a < 0 || b < 0) break;
          docXml = docXml.slice(0, a) + docXml.slice(b + 6);
        }
      }
    }
    const starts = new Set([...docXml.matchAll(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/g)].map(m => m[1]));
    const ends = new Set([...docXml.matchAll(/<w:bookmarkEnd\b[^>]*w:id="(\d+)"/g)].map(m => m[1]));
    docXml = docXml.replace(/<w:bookmark(Start|End)\b[^>]*w:id="(\d+)"[^>]*\/>/g, (t, k, id) => (k === 'Start' ? ends : starts).has(id) ? t : '');
    // связи документа, на которые больше никто не ссылается: картинки, вложенные таблицы, диаграммы
    const relsName = 'word/_rels/document.xml.rels';
    const dropped = [];
    if (zip.file(relsName)) {
      let rels = await zip.file(relsName).async('string');
      const used = new Set([...docXml.matchAll(/\br:(?:id|embed|link|pict|dm|lo|qs|cs)="([^"]+)"/g)].map(m => m[1]));
      const otherParts = Object.keys(zip.files).filter(n => TEXT_PART.test(n) && n !== 'word/document.xml');
      for (const n of otherParts) for (const m of (await zip.file(n).async('string')).matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)) used.add(m[1]);
      rels = rels.replace(/<Relationship\b[^>]*\/>/g, r => {
        const id = /Id="([^"]+)"/.exec(r), type = /Type="([^"]+)"/.exec(r), target = /Target="([^"]+)"/.exec(r);
        if (!id || used.has(id[1]) || /TargetMode="External"/.test(r) && !/hyperlink/.test(type[1])) return r;
        if (!/\/(image|oleObject|package|chart|diagram\w*|video|audio|media|hyperlink)$/.test(type[1])) return r;
        if (target && !/TargetMode="External"/.test(r)) dropped.push('word/' + target[1].replace(/^\.\//, '').replace(/^\/word\//, ''));
        return '';
      });
      zip.file(relsName, rels);
    }
    // сами файлы — если на них не ссылается больше ни одна связь; у диаграмм — и их вложенные данные
    const allRels = async () => (await Promise.all(Object.keys(zip.files).filter(n => /\.rels$/.test(n)).map(n => zip.file(n).async('string')))).join('\n');
    let removedFiles = 0;
    for (let k = 0; k < dropped.length; k++) {
      const f = dropped[k].replace(/\/[^/]+\/\.\.\//g, '/');
      if (!zip.file(f)) continue;
      const base = f.split('/').pop();
      if ((await allRels()).includes(base)) continue;
      const own = f.replace(/([^/]+)$/, '_rels/$1.rels');
      if (zip.file(own)) {
        const dir = f.replace(/[^/]+$/, '');
        for (const m of (await zip.file(own).async('string')).matchAll(/Target="([^"]+)"/g)) if (!/^(https?|mailto):/.test(m[1])) dropped.push((dir + m[1]).replace(/[^/]+\/\.\.\//g, ''));
        zip.remove(own);
      }
      zip.remove(f);
      removedFiles++;
      if (zip.file('[Content_Types].xml')) {
        const ct = await zip.file('[Content_Types].xml').async('string');
        zip.file('[Content_Types].xml', ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${escapeRe(f)}"[^>]*/>`), ''));
      }
    }
    return { xml: docXml, removedFiles };
  }

  const api = { TYPES, detect, customPattern, headingStyles, blocksOf, sectionEnd, ruleCuts, cutRanges, cutBlocks, segmentsOf, rewrite, dictFinder, plan, shiftDate, innValid, analyzeDocx, redactDocx, restoreDocx, previewParagraphs, previewDocument, acceptRevisions };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RedactDocx = api;
})(typeof window !== 'undefined' ? window : globalThis);

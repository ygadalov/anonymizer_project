
/*
 * Файлы Excel 97–2003 (.xls) и Excel 5.0/95: чтение через SheetJS (vendor/xlsx.full.min.js, Apache-2.0)
 * и преобразование в .xlsx, с которым дальше работает весь Redact (ExcelJS).
 *
 * SheetJS 0.20.3 (полная сборка с кодировками) разбирает файл в отдельном Web Worker.
 * Ограничения размера, времени и отмена снижают риск зависания. Worker не является
 * полной границей безопасности. Из потока возвращаются только простые данные — значения,
 * форматы чисел, формулы, объединения, ширина столбцов и высота строк.
 *
 * Без DOM: extractBook и buildWorkbook можно вызывать в Node (тесты).
 */
(function (root) {
  'use strict';

  // Что внутри файла: 'ole' — Excel 97–2003 и 5.0/95 (составной документ D0 CF 11 E0), 'xml' — таблица XML Excel 2003,
  // 'html' — HTML-таблица с расширением .xls (так сохраняют некоторые учётные системы), 'zip' — на самом деле .xlsx
  function sniff(buffer) {
    const b = new Uint8Array(buffer, 0, Math.min(1024, buffer.byteLength));
    if (b.length >= 4 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0) return 'ole';
    if (b.length >= 2 && b[0] === 0x50 && b[1] === 0x4B) return 'zip';
    let head = '';
    try { head = new TextDecoder('utf-8').decode(b).replace(/^\uFEFF/, '').trimStart().slice(0, 600).toLowerCase(); } catch (e) { /* не текст */ }
    if (/^<\?xml/.test(head) && head.includes('urn:schemas-microsoft-com:office:spreadsheet')) return 'xml';
    if (/^<(!doctype html|html|table|meta)/.test(head)) return 'html';
    return null;
  }
  // Нужно ли преобразование .xls → .xlsx. Файл .xls, который не похож ни на один формат Excel, — ошибка
  function isXls(buffer, fileName) {
    const kind = sniff(buffer);
    if (kind === 'ole' || kind === 'xml' || kind === 'html') return /\.(xls|xml|html?)$/i.test(fileName || '') || kind === 'ole';
    if (/\.xls$/i.test(fileName || '') && kind !== 'zip') throw new Error('это не книга Excel. Откройте его в Excel и сохраните как .xlsx');
    return false;
  }

  // Выполняется внутри фонового потока (исходный текст функции передаётся туда как есть) и в тестах
  // html — файл с HTML-таблицей: числа в нём читаются по-русски («1 000,5»), SheetJS принял бы запятую за разделитель тысяч
  function extractBook(XLSX, bytes, html) {
    const wb = XLSX.read(bytes, { type: 'array', cellNF: true, cellStyles: true, cellFormula: true, cellDates: false, raw: !!html });
    const ruNumber = s => {
      const t = String(s).trim();
      if (!/^-?\d{1,3}(?:[ \u00a0]\d{3})*(?:,\d+)?$|^-?\d+(?:,\d+)?$/.test(t)) return null;
      // ведущие нули — это код или номер, а не число
      if (/^-?0\d/.test(t)) return null;
      return Number(t.replace(/[ \u00a0]/g, '').replace(',', '.'));
    };
    if (wb.SheetNames.length > 100) throw new Error('В книге более 100 листов. Разделите книгу.');
    let bookCells = 0;
    const props = (wb.Workbook && wb.Workbook.WBProps) || {};
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    return {
      date1904: !!props.date1904,
      sheets: wb.SheetNames.map((name, i) => {
        const ws = wb.Sheets[name] || {};
        const cells = [];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        let maxR = range.e.r + 1, maxC = range.e.c + 1;
        for (const m of ws['!merges'] || []) { maxR = Math.max(maxR, m.e.r + 1); maxC = Math.max(maxC, m.e.c + 1); }
        if (maxR > 100000 || maxC > 2048 || maxR * maxC > 500000 || (ws['!cols'] || []).length > 2048 || (ws['!rows'] || []).length > 100000)
          throw new Error('Лист XLS слишком большой. Разделите файл.');
        bookCells += maxR * maxC;
        if (bookCells > 1000000) throw new Error('Книга XLS содержит слишком много ячеек.');
        for (const k of Object.keys(ws)) {
          if (k[0] === '!') continue;
          const c = ws[k], a = XLSX.utils.decode_cell(k);
          if (a.r < 0 || a.c < 0 || a.r >= 100000 || a.c >= 2048 || (a.r + 1) * (a.c + 1) > 500000 || cells.length >= 500000) throw new Error('Слишком много ячеек XLS.');
          if (c.t === 'z' && !c.f) continue;
          // ошибки (#Н/Д, #ДЕЛ/0!) — как текст, который показывал Excel
          let v = c.t === 'e' ? (c.w || '#N/A') : (c.v === undefined ? null : c.v), t = c.t === 'e' ? 's' : c.t;
          if (html && t === 's' && v !== null) { const n = ruNumber(v); if (n !== null) { v = n; t = 'n'; } }
          cells.push([a.r, a.c, t, v, c.z && c.z !== 'General' ? String(c.z) : null, c.f ? String(c.f) : null]);
        }
        return {
          name,
          hidden: meta[i] ? Number(meta[i].Hidden) || 0 : 0,
          cells,
          merges: (ws['!merges'] || []).map(m => [m.s.r, m.s.c, m.e.r, m.e.c]),
          // width — ширина в единицах Excel, как в исходном файле (wch у SheetJS меньше на поля ячейки)
          cols: (ws['!cols'] || []).map(x => x ? [x.width || (x.wch > 0 ? x.wch : null) || (x.wpx ? x.wpx / 7 : null), !!x.hidden] : null),
          rows: (ws['!rows'] || []).map(x => x ? [x.hpt || (x.hpx ? x.hpx * 0.75 : null), !!x.hidden] : null),
        };
      }),
    };
  }

  // Книга ExcelJS из разобранных данных
  function buildWorkbook(data, ExcelJS) {
    const wb = new ExcelJS.Workbook();
    if (data.date1904) wb.properties.date1904 = true;
    const used = new Set();
    let formulas = 0;
    for (const sh of data.sheets) {
      // имя листа: допустимое в .xlsx и уникальное
      let name = String(sh.name || 'Лист').replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Лист';
      for (let n = 2; used.has(name.toLowerCase()); n++) name = `${name.slice(0, 27)} (${n})`;
      used.add(name.toLowerCase());
      const ws = wb.addWorksheet(name, { state: sh.hidden === 2 ? 'veryHidden' : sh.hidden ? 'hidden' : 'visible' });
      for (const [r, c, t, v, z, f] of sh.cells) {
        const cell = ws.getCell(r + 1, c + 1);
        let value = v;
        if (t === 'b') value = !!v;
        else if (t === 's' && v !== null) value = String(v);
        else if (t === 'n' && typeof v !== 'number') value = Number(v);
        if (f) { cell.value = { formula: f, result: value === null ? undefined : value }; formulas++; }
        else cell.value = value;
        if (z) cell.numFmt = z;
      }
      sh.cols.forEach((x, i) => {
        if (!x) return;
        const col = ws.getColumn(i + 1);
        if (x[0]) col.width = Math.round(x[0] * 100) / 100;
        if (x[1]) col.hidden = true;
      });
      sh.rows.forEach((x, i) => {
        if (!x) return;
        const row = ws.getRow(i + 1);
        if (x[0]) row.height = Math.round(x[0] * 100) / 100;
        if (x[1]) row.hidden = true;
      });
      for (const [r1, c1, r2, c2] of sh.merges) {
        try { ws.mergeCells(r1 + 1, c1 + 1, r2 + 1, c2 + 1); } catch (e) { /* пересекающиеся объединения пропускаются */ }
      }
    }
    return { wb, formulas };
  }

  // ---------- Браузер: разбор в фоновом потоке ----------
  const TIMEOUT_MS = 60000;
  let workerBody = null;
  function makeWorkerBody() {
    if (workerBody) return workerBody;
    // текст библиотек встроен в страницу при сборке (build.py) как неисполняемые <script type="text/x-worker">
    const lib = id => { const el = document.getElementById(id); return el ? el.textContent : ''; };
    const sheetjs = lib('libSheetjs');
    if (!sheetjs) throw new Error('чтение .xls доступно только в собранном файле Redact_vX.Y.html');
    const body = `${sheetjs}\n;\n` +
      `const extractBook = ${extractBook.toString()};\n` +
      `self.onmessage = e => { try { self.postMessage({ ok: true, data: extractBook(XLSX, new Uint8Array(e.data.buf), e.data.html) }); }` +
      ` catch (err) { self.postMessage({ ok: false, error: String(err && err.message || err) }); } };`;
    workerBody = body;
    return workerBody;
  }
  function parseInWorker(buffer, html) {
    root.RedactLimits.checkFile(buffer.byteLength);
    return root.RedactLimits.runWorker(makeWorkerBody(), { buf: buffer.slice(0), html }, TIMEOUT_MS);
  }

  // .xls → { buffer: ArrayBuffer .xlsx, notes: [что сказать пользователю] }
  async function toXlsx(buffer) {
    let data;
    try { data = await parseInWorker(buffer, sniff(buffer) === 'html'); }
    catch (e) {
      console.error('xls:', e);
      throw new Error(/password|encrypt/i.test(e.message) ? 'файл .xls защищён паролем — снимите защиту в Excel и попробуйте снова'
        : /долго|лимит|время|отмен|слишком|более|много/.test(e.message) ? e.message
        : 'файл .xls повреждён или сохранён в формате, который Redact не знает. Откройте его в Excel и сохраните как .xlsx');
    }
    if (!data.sheets.length) throw new Error('В книге нет листов.');
    const { wb } = buildWorkbook(data, root.ExcelJS);
    const out = await wb.xlsx.writeBuffer();
    return {
      buffer: out instanceof ArrayBuffer ? out : out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength),
      notes: ['Файл .xls (Excel 97–2003) преобразован в .xlsx: значения, формулы, форматы чисел, ширина столбцов, высота строк ' +
        'и объединённые ячейки перенесены, а шрифты, заливка и рамки — нет. Результат сохранится в .xlsx.'],
    };
  }

  const api = { sniff, isXls, extractBook, buildWorkbook, toXlsx };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RedactXls = api;
})(typeof window !== 'undefined' ? window : globalThis);


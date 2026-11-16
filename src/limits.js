
/* Локальные ограничения ресурсов. Проверка ZIP выполняется в прерываемом Worker до ExcelJS/Word. */
(function (root) {
  'use strict';
  const LIMITS = Object.freeze({ fileBytes: 25 * 1024 * 1024, zipEntries: 4096, entryBytes: 32 * 1024 * 1024,
    totalBytes: 128 * 1024 * 1024, ratio: 200, sheets: 100, rows: 100000, cols: 2048,
    cells: 500000, bookCells: 1000000, xmlTags: 500000, xmlDepth: 256, timeoutMs: 15000 });
  function checkFile(size) {
    if (!Number.isSafeInteger(size) || size < 0 || size > LIMITS.fileBytes) throw new Error('Файл слишком большой: максимум 25 МиБ. Разделите его на несколько файлов.');
  }
  function checkGrid(rows, cols) {
    if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 0 || cols < 0 || rows > LIMITS.rows || cols > LIMITS.cols || rows * cols > LIMITS.cells)
      throw new Error('Лист слишком большой: максимум 100 000 строк, 2048 столбцов и 500 000 ячеек в прямоугольнике данных. Разделите лист.');
  }
  function checkZip(buffer) {
    checkFile(buffer.byteLength);
    const b = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let end = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (v.getUint32(i, true) === 0x06054b50 && i + 22 + v.getUint16(i + 20, true) === b.length) { end = i; break; }
    const bad = () => { throw new Error('Повреждённый ZIP или неподдерживаемый ZIP64/многотомный архив. Пересохраните документ.'); };
    if (end < 0) bad();
    const count = v.getUint16(end + 10, true), size = v.getUint32(end + 12, true), offset = v.getUint32(end + 16, true);
    if (v.getUint16(end + 4, true) || v.getUint16(end + 6, true) || count !== v.getUint16(end + 8, true) || count === 65535 || size === 0xffffffff || offset === 0xffffffff) bad();
    if (count > LIMITS.zipEntries) throw new Error('В архиве слишком много частей: максимум 4096. Упростите документ.');
    if (offset + size > end) bad();
    let pos = offset, total = 0;
    const names = new Set();
    for (let i = 0; i < count; i++) {
      if (pos + 46 > offset + size || v.getUint32(pos, true) !== 0x02014b50) bad();
      const compressed = v.getUint32(pos + 20, true), unpacked = v.getUint32(pos + 24, true);
      const len = v.getUint16(pos + 28, true), extra = v.getUint16(pos + 30, true), comment = v.getUint16(pos + 32, true), local = v.getUint32(pos + 42, true);
      if (pos + 46 + len + extra + comment > offset + size || local + 30 > offset || v.getUint32(local, true) !== 0x04034b50 || v.getUint16(pos + 34, true) || v.getUint16(pos + 8, true) & 1) bad();
      const name = new TextDecoder().decode(b.subarray(pos + 46, pos + 46 + len));
      if (name.includes('\\') || name.startsWith('/') || name.split('/').includes('..') || name.includes('\0') || names.has(name)) bad();
      names.add(name); total += unpacked;
      if (unpacked > LIMITS.entryBytes || total > LIMITS.totalBytes || unpacked > Math.max(1, compressed) * LIMITS.ratio)
        throw new Error('Архив превышает ограничения распаковки: 32 МиБ на часть, 128 МиБ всего, сжатие не более 200:1. Разделите документ.');
      pos += 46 + len + extra + comment;
    }
    if (pos !== offset + size) bad();
    return { entries: count, bytes: total };
  }
  function inspectSheetXml(xml) {
    let rows = 0, cols = 0;
    for (const m of xml.matchAll(/\b(?:r|ref)\s*=\s*["'](?:[A-Z]+\d+:)?([A-Z]+)(\d+)["']/g)) {
      const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      rows = Math.max(rows, Number(m[2])); cols = Math.max(cols, c);
    }
    for (const m of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*\br\s*=\s*["'](\d+)["']/g)) rows = Math.max(rows, Number(m[1]));
    checkGrid(rows, cols);
    return rows * cols;
  }
  function inspectXml(text) {
    let tags = 0, depth = 0;
    for (const m of text.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?[^>"']*(?:"[^"]*"[^>"']*|'[^']*'[^>"']*)*>/g)) {
      const t = m[0];
      if (/^<[!?]/.test(t)) continue;
      if (++tags > LIMITS.xmlTags) throw new Error('XML содержит слишком много элементов. Разделите документ.');
      if (t.startsWith('</')) depth--;
      else if (!/\/\s*>$/.test(t)) depth++;
      if (depth > LIMITS.xmlDepth || depth < 0) throw new Error('XML слишком глубоко вложен или повреждён. Пересохраните документ.');
    }
    if (depth) throw new Error('XML содержит незакрытые элементы. Пересохраните документ.');
  }
  async function validateZip(buffer, JSZip) {
    const summary = checkZip(buffer), zip = await JSZip.loadAsync(buffer);
    let total = 0, bookCells = 0, sheets = 0;
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      const sheet = /^xl\/worksheets\/[^/]+\.xml$/.test(name), xml = /\.(xml|rels)$/.test(name);
      let bytes = 0; const chunks = [];
      await new Promise((resolve, reject) => {
        const stream = entry.internalStream('uint8array');
        stream.on('data', chunk => {
          bytes += chunk.length; total += chunk.length;
          if (bytes > LIMITS.entryBytes || total > LIMITS.totalBytes) { stream.pause(); reject(new Error('Реальный объём распаковки превышает лимит.')); return; }
          if (xml) chunks.push(chunk);
        }).on('error', reject).on('end', resolve).resume();
      });
      if (xml) {
        const data = new Uint8Array(bytes); let at = 0;
        for (const chunk of chunks) { data.set(chunk, at); at += chunk.length; }
        const text = new TextDecoder().decode(data);
        inspectXml(text);
        if (sheet) {
          if (++sheets > LIMITS.sheets) throw new Error('В книге более 100 листов. Разделите книгу.');
          bookCells += inspectSheetXml(text);
          if (bookCells > LIMITS.bookCells) throw new Error('В книге более 1 000 000 ячеек в прямоугольниках данных. Разделите книгу.');
        }
      }
    }
    return summary;
  }
  const running = new Set();
  function runWorker(body, data, timeoutMs = LIMITS.timeoutMs) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(new Blob([body], { type: 'text/javascript' }));
      let worker;
      try { worker = new Worker(url); } catch (e) { URL.revokeObjectURL(url); reject(e); return; }
      const button = typeof document !== 'undefined' && document.getElementById('cancelRead');
      const finish = (error, result) => {
        clearTimeout(timer); worker.terminate(); URL.revokeObjectURL(url); running.delete(cancel);
        if (button) button.classList.toggle('hidden', !running.size);
        if (error) reject(error); else resolve(result);
      };
      const cancel = () => finish(new Error('Проверка/чтение файла отменены.'));
      const timer = setTimeout(() => finish(new Error('Превышено время проверки/чтения файла. Разделите документ или пересохраните его.')), timeoutMs);
      running.add(cancel);
      if (button) { button.classList.remove('hidden'); button.onclick = () => [...running].forEach(fn => fn()); }
      worker.onmessage = e => e.data && e.data.ok ? finish(null, e.data.data) : finish(new Error(e.data && e.data.error || 'Ошибка проверки файла.'));
      worker.onerror = e => { e.preventDefault(); finish(new Error(e.message || 'Ошибка фонового чтения.')); };
      try { worker.postMessage(data); } catch (e) { finish(e); }
    });
  }
  async function preflight(buffer, JSZip) {
    checkZip(buffer);
    if (typeof document === 'undefined') return validateZip(buffer, JSZip);
    const lib = document.getElementById('libZipGuard');
    if (!lib || !lib.textContent) throw new Error('Проверка документов доступна в собранном HTML-файле Redact.');
    const body = lib.textContent + `\nconst LIMITS=${JSON.stringify(LIMITS)};\n` +
      [checkFile, checkGrid, checkZip, inspectSheetXml, inspectXml, validateZip].map(fn => fn.toString()).join('\n') +
      '\nself.onmessage=async e=>{try{self.postMessage({ok:true,data:await validateZip(e.data,JSZip)})}catch(err){self.postMessage({ok:false,error:err.message})}};';
    return runWorker(body, buffer);
  }
  const api = { LIMITS, checkFile, checkGrid, checkZip, inspectSheetXml, inspectXml, validateZip, preflight, runWorker };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RedactLimits = api;
})(typeof window !== 'undefined' ? window : globalThis);


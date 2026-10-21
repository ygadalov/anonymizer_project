/* Режим «Восстановить по ключу»: обезличенный (или присланный в ответ) файл + ключ → исходные данные. */
(function () {
  'use strict';
  const A = window.Anon;
  const { $, icon, esc, fmt, toast, download, KEY_SHEET } = window.RedactUI;

  const R = {
    fileName: '', kind: null, buffer: null, wb: null, csv: null,
    sheetIndex: 0, grid: null, an: null, notes: [],
    key: null, keyName: '', keyNotes: [], envelope: null, prep: null,
    actions: {}, result: null,
    sheetActions: new Map(),   // действия, изменённые вручную, по листам: имя → { headerRow, actions }
  };
  const curName = () => R.kind === 'xlsx' && R.wb ? R.wb.worksheets[R.sheetIndex].name : null;
  // лист есть в ключе — он был обезличен вместе с книгой и восстанавливается полностью
  const inKey = name => !!(R.prep && name != null && R.prep.keySheets.has(name));

  // ---------- Переключение режимов ----------
  // вкладки по порядку: Таблица Excel · Документ Word · Восстановить по ключу
  const MODES = [['redact', 'tabRedact', 'modeRedact'], ['doc', 'tabDoc', 'modeDoc'], ['restore', 'tabRestore', 'modeRestore']];
  function setMode(mode) {
    for (const [m, tab, panel] of MODES) {
      $(panel).classList.toggle('hidden', m !== mode);
      $(tab).setAttribute('aria-selected', String(m === mode));
    }
  }
  for (const [m, tab] of MODES) $(tab).addEventListener('click', () => setMode(m));
  $('tabRedact').parentElement.addEventListener('keydown', ev => {
    if (ev.key !== 'ArrowLeft' && ev.key !== 'ArrowRight') return;
    const i = MODES.findIndex(([, tab]) => $(tab).getAttribute('aria-selected') === 'true');
    const next = MODES[(i + (ev.key === 'ArrowRight' ? 1 : MODES.length - 1)) % MODES.length];
    setMode(next[0]);
    $(next[1]).focus();
  });

  // ---------- Загрузка файлов ----------
  function wireDrop(dropId, inputId, handler) {
    const drop = $(dropId);
    ['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', ev => { const f = ev.dataTransfer.files[0]; if (f) handler(f); });
    $(inputId).addEventListener('change', ev => { const f = ev.target.files[0]; ev.target.value = ''; if (f) handler(f); });
  }
  wireDrop('rDropFile', 'rFile', loadData);
  wireDrop('rDropKey', 'rKey', loadKey);

  const ext = name => (name.split('.').pop() || '').toLowerCase();

  async function loadData(file) {
    $('rFileInfo').textContent = 'Читаем файл…';
    try {
      const buf = await file.arrayBuffer();
      const e = ext(file.name);
      const docx = e === 'docx';
      const xls = !docx && window.RedactXls.isXls(buf, file.name);
      R.fileName = file.name;
      R.buffer = buf;
      R.notes = [];
      if (docx) {
        // документ Word: столбцов нет, восстанавливается весь текст по словарю ключа
        R.kind = 'docx';
        R.wb = null; R.grid = null; R.an = null;
        R.docParts = (await window.RedactDocx.analyzeDocx(await JSZip.loadAsync(buf), { types: {} })).parts;
      } else if (!xls && (e === 'csv' || e === 'txt')) {
        R.kind = 'csv';
        let text = new TextDecoder('utf-8').decode(buf);
        if (text.includes('�')) text = new TextDecoder('windows-1251').decode(buf);
        const parsed = A.parseCsv(text);
        R.wb = null;
        R.grid = parsed.grid;
        R.csv = { delim: parsed.delim, decSep: parsed.decSep };
        $('rSheet').innerHTML = '<option>CSV</option>';
        $('rSheet').disabled = true;
      } else {
        R.kind = 'xlsx';
        // .xls (Excel 97–2003) сначала преобразуется в .xlsx
        const conv = xls ? await window.RedactXls.toXlsx(buf) : null;
        if (conv) R.buffer = conv.buffer;
        const fixed = await A.repairXlsx(R.buffer, JSZip);
        if (fixed.changed) R.buffer = fixed.buffer;
        R.notes = [...(conv ? conv.notes : []), ...fixed.notes];
        R.wb = new ExcelJS.Workbook();
        await R.wb.xlsx.load(R.buffer);
        if (!R.wb.worksheets.length) throw new Error('в книге нет листов');
        R.sheetActions = new Map();
        R.sheetIndex = pickSheet();
        renderSheetOptions();
        R.grid = A.gridFromWorksheet(R.wb.worksheets[R.sheetIndex]);
      }
      $('rClearRow').classList.remove('hidden');
      $('rFileInfo').innerHTML = `${icon(docx ? 'file-text' : 'file-spreadsheet', 'i i20')}<b>${esc(file.name)}</b><span class="muted">${(file.size / 1024).toFixed(0)} КБ</span>`;
      setup(true);
    } catch (e) {
      console.error(e);
      $('rFileInfo').textContent = '';
      toast('Не удалось прочитать файл: ' + e.message, 7000, 'error');
    }
  }
  // первый лист из ключа, если он есть в файле; иначе самый большой
  function pickSheet() {
    const wss = R.wb.worksheets;
    const names = R.key ? (R.key.sheets || []).map(x => x.name).concat(R.key.sheet ? [R.key.sheet] : []) : [];
    for (const n of names) {
      const i = wss.findIndex(ws => ws.name === n);
      if (i >= 0) return i;
    }
    let best = 0;
    wss.forEach((ws, i) => { if (ws.rowCount * ws.columnCount > wss[best].rowCount * wss[best].columnCount) best = i; });
    return best;
  }

  async function loadKey(file) {
    $('rKeyInfo').textContent = 'Читаем ключ…';
    $('rPassRow').classList.add('hidden');
    R.envelope = null;
    try {
      const e = ext(file.name);
      R.keyName = file.name;
      R.keyNotes = [];
      if (e === 'xlsx') {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await file.arrayBuffer());
        const ks = wb.worksheets.find(ws => ws.name === KEY_SHEET);
        if (ks) {
          const parts = [];
          for (let r = 2; r <= ks.rowCount; r++) { const v = ks.getCell(r, 1).value; if (v) parts.push(String(v)); }
          setKey(JSON.parse(parts.join('')));
        } else {
          // ключ версии до 2.1: только таблица соответствий и параметры
          const map = wb.getWorksheet('Соответствия'), par = wb.getWorksheet('Параметры');
          if (!map || !par) throw new Error('в файле нет листов «Соответствия» и «Параметры» — это не ключ Redact');
          const rows = (ws, n) => { const out = []; ws.eachRow((row, r) => { if (r > 1) out.push(Array.from({ length: n }, (_, i) => row.getCell(i + 1).value)); }); return out; };
          R.keyNotes.push('Ключ сохранён версией до 2.1: названия и коды вернутся, но номера внутри текста не восстанавливаются, а суммы делятся на коэффициент с точностью до копейки.');
          setKey(A.legacyKeyFromRows(rows(map, 5), rows(par, 2)));
        }
      } else {
        const obj = JSON.parse(await file.text());
        if (obj && obj.kind === 'redact-key-encrypted') {
          R.envelope = obj;
          R.key = null; R.prep = null;
          $('rKeyInfo').innerHTML = `${icon('lock', 'i i20')}<b>${esc(file.name)}</b><span class="muted">зашифрован паролем</span>`;
          $('rPassRow').classList.remove('hidden');
          $('rClearRow').classList.remove('hidden');
          $('rPass').value = '';
          $('rPass').focus();
          return;
        }
        if (!obj || obj.kind !== 'redact-key') throw new Error('это не ключ Redact');
        setKey(obj);
      }
    } catch (e) {
      console.error(e);
      $('rKeyInfo').textContent = '';
      toast('Не удалось прочитать ключ: ' + e.message, 7000, 'error');
    }
  }

  // ---------- Очистить ----------
  function clearRestore() {
    Object.assign(R, {
      fileName: '', kind: null, buffer: null, wb: null, csv: null, sheetIndex: 0, grid: null, an: null, notes: [],
      key: null, keyName: '', keyNotes: [], envelope: null, prep: null, actions: {}, result: null,
      sheetActions: new Map(), docParts: null,
    });
    ['r2', 'r3', 'rPassRow', 'rClearRow'].forEach(id => $(id).classList.add('hidden'));
    for (const id of ['rFileInfo', 'rKeyInfo', 'rRowsInfo', 'rStats', 'rPreview', 'rWarnings', 'rSheet']) $(id).innerHTML = '';
    $('rColsTable').tBodies[0].innerHTML = '';
    $('rPass').value = '';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    toast('Готово к новому файлу и ключу');
  }
  $('btnRClear').addEventListener('click', async () => {
    const ok = await window.RedactUI.confirm({
      title: 'Очистить и начать заново?',
      text: 'Будут убраны выбранный файл, ключ соответствий и предпросмотр. Сами файлы на компьютере не затрагиваются.',
      confirm: 'Да, очистить',
    });
    if (ok) clearRestore();
  });

  async function unlock() {
    if (!R.envelope) return;
    const btn = $('btnUnlock');
    btn.disabled = true;
    try {
      const key = await A.decryptKey(R.envelope, $('rPass').value);
      $('rPassRow').classList.add('hidden');
      $('rPass').value = '';
      setKey(key);
    } catch (e) {
      toast('Не удалось открыть ключ: ' + e.message, 6000, 'error');
    } finally { btn.disabled = false; }
  }
  $('btnUnlock').addEventListener('click', unlock);
  $('rPass').addEventListener('keydown', ev => { if (ev.key === 'Enter') unlock(); });

  function setKey(key) {
    $('rClearRow').classList.remove('hidden');
    R.key = key;
    R.prep = A.prepareRestore(key);
    R.sheetActions = new Map(); // действия по умолчанию зависят от ключа
    const pairs = (key.dictionaries || []).reduce((n, d) => n + d.pairs.length, 0);
    const locked = !!R.envelope;
    $('rKeyInfo').innerHTML = `${icon(locked ? 'lock-open' : 'key-round', 'i i20')}<b>${esc(R.keyName)}</b>` +
      `<span class="muted">${key.sourceFile ? `для «${esc(key.sourceFile)}» · ` : ''}замен в ключе: ${pairs.toLocaleString('ru-RU')} · коэффициент ${esc(key.factor)} · сдвиг дат ${esc(key.shiftDays)} дн.</span>`;
    R.envelope = null;
    // если файл уже загружен, а лист из ключа другой — переключаемся на лист из ключа
    if (R.kind === 'xlsx' && R.wb) {
      const i = pickSheet();
      if (i !== R.sheetIndex) { R.sheetIndex = i; $('rSheet').value = String(i); R.grid = A.gridFromWorksheet(R.wb.worksheets[i]); }
    }
    setup(true);
  }

  // ---------- Структура ----------
  function renderSheetOptions() {
    if (R.kind !== 'xlsx' || !R.wb) { $('rSheet').innerHTML = '<option>CSV</option>'; $('rSheet').disabled = true; return; }
    $('rSheet').innerHTML = R.wb.worksheets.map((ws, i) =>
      `<option value="${i}">${esc(ws.name)}${R.prep ? (inKey(ws.name) ? ' · из ключа' : ' · нет в ключе') : ''}</option>`).join('');
    $('rSheet').value = String(R.sheetIndex);
    $('rSheet').disabled = R.wb.worksheets.length < 2;
  }
  $('rSheet').addEventListener('change', () => {
    R.sheetIndex = +$('rSheet').value;
    R.grid = A.gridFromWorksheet(R.wb.worksheets[R.sheetIndex]);
    setup(true);
  });
  $('rHeaderRow').addEventListener('change', () => {
    const v = Math.max(1, Math.min(R.grid.maxRow, parseInt($('rHeaderRow').value, 10) || 1));
    $('rHeaderRow').value = v;
    setup(false, v);
  });
  $('rOtherSheets').addEventListener('change', preview);

  // ---------- Документ Word ----------
  // словарь «замена → исходное» из ключа (документа или таблицы): псевдонимы, коды, ФИО с инициалами
  const backMap = () => new Map([...R.prep.global, ...(R.prep.names || [])]);
  function setupDoc() {
    if (!R.docParts || !R.prep) { ['r2', 'r3'].forEach(id => $(id).classList.add('hidden')); return; }
    $('r2').classList.add('hidden');
    $('r3').classList.remove('hidden');
    const back = backMap();
    const finder = window.RedactDocx.dictFinder(back);
    let n = 0;
    for (const p of R.docParts) for (const sg of p.segs) n += finder(sg.text).length;
    $('rStats').textContent = n
      ? `Будет восстановлено вхождений: ${n.toLocaleString('ru-RU')} — в тексте, таблицах, колонтитулах и сносках документа`
      : 'В документе не найдено значений из ключа. Проверьте, что ключ от этого документа или от таблицы, по которой он написан.';
    const paras = window.RedactDocx.previewParagraphs(R.docParts, back, 30);
    $('rPreview').innerHTML = paras.length ? '<tbody>' + paras.map(p => `<tr><td class="doc-cell">${p.pieces.map(x => x.text !== undefined ? esc(x.text)
      : `<span class="pair-inline" title="${esc(x.orig)} → ${esc(x.repl)}"><span class="rv rv-token">${esc(x.orig)}</span><span class="rv-arrow" aria-hidden="true">→</span><span class="rv rv-original">${esc(x.repl)}</span></span>`).join('')}</td></tr>`).join('') + '</tbody>' : '';
    const info = [...R.keyNotes];
    if (R.key && !R.key.document) info.push('Ключ от таблицы: в документе вернутся названия, коды и ФИО из неё. Суммы и даты в тексте документа по такому ключу не меняются.');
    $('rWarnings').innerHTML = info.length ? `<div class="notice"><div class="notice-title">${icon('circle-alert', 'i i20')}Проверьте перед выгрузкой</div><ul>${info.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
  }

  function setup(autoHeader, headerRow) {
    if (R.kind === 'docx') { setupDoc(); return; }
    if (!R.grid || !R.prep) {
      ['r2', 'r3'].forEach(id => $(id).classList.add('hidden'));
      return;
    }
    const name = curName();
    const kept = R.sheetActions.get(name);
    const hr = autoHeader ? (kept ? kept.headerRow : A.detectHeaderRow(R.grid)) : headerRow;
    $('rHeaderRow').value = hr;
    R.an = A.analyze(R.grid, hr);
    if (kept && kept.headerRow === hr) R.actions = kept.actions;
    else {
      R.actions = {};
      for (const col of R.an.columns) R.actions[col.c] = A.defaultRestoreAction(R.prep, col, name);
      R.sheetActions.set(name, { headerRow: hr, actions: R.actions });
    }
    renderSheetOptions();
    const known = R.an.columns.filter(col => A.matchKeyColumn(R.prep, col, name)).length;
    $('rRowsInfo').textContent = `Строк: ${R.an.dataRows.length} · столбцов: ${R.an.columns.length}, из них из исходного файла: ${known}`;
    renderCols();
    ['r2', 'r3'].forEach(id => $(id).classList.remove('hidden'));
    preview();
  }

  const methodLabel = m => {
    for (const list of Object.values(A.METHODS)) for (const [k, n] of list) if (k === m) return n;
    return m;
  };

  function renderCols() {
    $('rColsTable').tBodies[0].innerHTML = R.an.columns.map(col => {
      const kc = A.matchKeyColumn(R.prep, col, curName());
      const inKey = !kc ? '<span class="muted">нет — новый столбец</span>'
        : kc.method === 'keep' ? '<span class="muted">не обезличивался</span>'
        : `${esc(methodLabel(kc.method))}${A.REVERSE_OF[kc.method] ? '' : ' <span class="text-danger">· необратимо</span>'}`;
      const action = R.actions[col.c];
      const opts = A.RESTORE_ACTIONS.map(([k, n]) => `<option value="${k}" title="${esc(A.RESTORE_HINTS[k])}"${k === action ? ' selected' : ''}>${n}</option>`).join('');
      const samples = col.samples.map(v => `<span title="${esc(fmt(v))}">${esc(fmt(v))}</span>`).join('');
      return `<tr data-c="${col.c}" class="${action !== 'keep' ? 'active' : ''}">
        <td class="colname"><b>${col.letter}</b>${esc(col.name)}</td>
        <td class="samples">${samples}</td>
        <td class="small">${inKey}</td>
        <td><select data-k="action" title="${esc(A.RESTORE_HINTS[action])}">${opts}</select></td>
      </tr>`;
    }).join('');
  }
  $('rColsTable').addEventListener('change', ev => {
    const tr = ev.target.closest('tr[data-c]');
    if (!tr || ev.target.dataset.k !== 'action') return;
    R.actions[+tr.dataset.c] = ev.target.value;
    tr.classList.toggle('active', ev.target.value !== 'keep');
    ev.target.title = A.RESTORE_HINTS[ev.target.value];
    preview();
  });

  // ---------- Предпросмотр ----------
  function preview() {
    if (!R.an || !R.prep) return;
    R.result = A.restoreTransform(R.grid, R.an, R.prep, R.actions, curName());
    const { changes, stats, leftovers } = R.result;
    $('rStats').textContent = stats.restored
      ? `Будет восстановлено значений: ${stats.restored.toLocaleString('ru-RU')} в ${stats.columns} ${A.plural(stats.columns, 'столбце', 'столбцах', 'столбцах')}` +
        (R.kind === 'xlsx' && R.wb.worksheets.length > 1 ? ` на листе «${curName()}»` : '')
      : 'В файле не найдено значений из ключа. Проверьте, что ключ от этого файла и строка заголовков выбрана верно.';
    const active = R.an.columns.filter(col => R.result.perColumn.some(p => p.col.c === col.c && p.restored));
    const rows = R.an.dataRows.filter(r => active.some(col => changes.has(r + ':' + col.c))).slice(0, 15);
    $('rPreview').innerHTML = active.length ? `<thead><tr><th class="r">Строка</th>${active.map(col => `<th>${esc(col.letter)} · ${esc(col.name)}</th>`).join('')}</tr></thead><tbody>` +
      rows.map(r => `<tr><td class="r rownum">${r}</td>${active.map(col => {
        const cell = R.grid.rows[r][col.c];
        const old = cell && !cell.slave ? cell.v : null;
        const k = r + ':' + col.c;
        if (!changes.has(k)) return `<td>${old === null ? '' : `<span class="rv rv-original">${esc(fmt(old))}</span>`}</td>`;
        const nv = changes.get(k);
        return `<td title="${esc(fmt(old))} → ${esc(fmt(nv))}"><div class="pair"><span class="rv rv-token">${esc(fmt(old))}</span><span class="line"><span class="rv-arrow" aria-hidden="true">→</span><span class="rv rv-original">${esc(fmt(nv))}</span></span></div></td>`;
      }).join('')}</tr>`).join('') + '</tbody>' : '';

    const danger = [], info = [...R.notes, ...R.keyNotes];
    if (leftovers.length) {
      const ex = [...new Set(leftovers.map(x => x.text))].slice(0, 5).map(t => `«${t}»`).join(', ');
      danger.push(`Не найдено в ключе: ${leftovers.length} ${A.plural(leftovers.length, 'значение', 'значения', 'значений')} (${ex}). Возможно, файл обезличен другим ключом или значения добавлены вручную.`);
    }
    const irreversible = R.an.columns.map(col => A.matchKeyColumn(R.prep, col, curName())).filter(kc => kc && kc.method !== 'keep' && !A.REVERSE_OF[kc.method]);
    if (irreversible.length) info.push(`Необратимые методы: ${irreversible.map(kc => `«${kc.name}» — ${methodLabel(kc.method).toLowerCase()}`).join('; ')}. Эти значения не восстановить.`);
    if (R.prep.ambiguous) info.push(`Замен, которые совпали в разных группах: ${R.prep.ambiguous}. Внутри текста они не заменяются; в своих столбцах восстанавливаются по группе столбца.`);
    if (R.kind === 'xlsx' && R.wb.worksheets.length > 1) {
      const others = R.wb.worksheets.filter((_, i) => i !== R.sheetIndex);
      const fromKey = others.filter(ws => inKey(ws.name)).map(ws => `«${ws.name}»`);
      const rest = others.filter(ws => !inKey(ws.name)).map(ws => `«${ws.name}»`);
      if (fromKey.length) info.push(fromKey.length > 1
        ? `Листы ${fromKey.join(', ')} из ключа тоже будут восстановлены полностью — названия, коды, суммы и даты. Чтобы проверить или изменить действия, выберите лист в списке «Лист».`
        : `Лист ${fromKey[0]} из ключа тоже будет восстановлен полностью — названия, коды, суммы и даты. Чтобы проверить или изменить действия, выберите его в списке «Лист».`);
      if (rest.length) info.push($('rOtherSheets').checked
        ? `На ${rest.length > 1 ? 'листах' : 'листе'} ${rest.join(', ')} (нет в ключе) будут возвращены названия и коды. Числа и даты там не меняются.`
        : `${rest.length > 1 ? 'Листы' : 'Лист'} ${rest.join(', ')} (нет в ключе) останутся без изменений.`);
    }
    const notice = (cls, ic, title, items) => items.length
      ? `<div class="notice ${cls}"><div class="notice-title">${icon(ic, 'i i20')}${title}</div><ul>${items.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
    $('rWarnings').innerHTML = notice('danger', 'triangle-alert', 'Восстановлено не всё', danger) + notice('', 'circle-alert', 'Проверьте перед выгрузкой', info);
  }

  // ---------- Выгрузка ----------
  $('btnRestore').addEventListener('click', async () => {
    if (R.kind === 'docx' && R.prep && R.docParts) {
      const btn = $('btnRestore');
      btn.disabled = true;
      try {
        const zip = await JSZip.loadAsync(R.buffer);
        const { restored } = await window.RedactDocx.restoreDocx(zip, backMap());
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        download(blob, R.fileName.replace(/\.[^.]+$/, '').replace(/_redact$/, '') + '_restored.docx');
        toast(`Восстановлено вхождений: ${restored.toLocaleString('ru-RU')}. Документ сохранён в «Загрузки»`);
      } catch (e) {
        console.error(e);
        toast('Не удалось сохранить документ: ' + e.message, 7000, 'error');
      } finally { btn.disabled = false; }
      return;
    }
    if (!R.an || !R.prep) return;
    const btn = $('btnRestore');
    btn.disabled = true;
    try {
      preview();
      const base = R.fileName.replace(/\.[^.]+$/, '').replace(/_redact$/, '') + '_restored';
      let total = R.result.stats.restored;
      if (R.kind === 'csv') {
        download(new Blob([A.gridToCsv(R.grid, R.result.changes, R.csv.delim, R.csv.decSep)], { type: 'text/csv;charset=utf-8' }), base + '.csv');
      } else {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(R.buffer);
        wb.worksheets.forEach((ws, i) => {
          if (i === R.sheetIndex) { A.applyToWorksheet(ws, R.result.changes); return; }
          const g = A.gridFromWorksheet(ws);
          let an, acts = {};
          if (inKey(ws.name)) {
            // лист из ключа — полностью, с действиями, выбранными для него (или по умолчанию)
            const kept = R.sheetActions.get(ws.name);
            an = A.analyze(g, kept ? kept.headerRow : A.detectHeaderRow(g));
            if (kept) acts = kept.actions;
            else for (const col of an.columns) acts[col.c] = A.defaultRestoreAction(R.prep, col, ws.name);
          } else {
            if (!$('rOtherSheets').checked) return;
            // лист не из ключа — только названия и коды (числа и даты там могут быть уже своими)
            an = A.analyze(g, 0);
            for (const col of an.columns) acts[col.c] = 'text';
          }
          const res = A.restoreTransform(g, an, R.prep, acts, ws.name);
          A.applyToWorksheet(ws, res.changes);
          total += res.stats.restored;
        });
        wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
        if (!wb._themes || !Object.keys(wb._themes).length) wb._themes = undefined;
        const out = await wb.xlsx.writeBuffer();
        download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), base + '.xlsx');
      }
      toast(`Восстановлено значений: ${total.toLocaleString('ru-RU')}. Файл сохранён в «Загрузки»`);
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить файл: ' + e.message, 7000, 'error');
    } finally { btn.disabled = false; }
  });
})();

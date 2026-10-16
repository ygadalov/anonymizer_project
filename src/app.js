/* Интерфейс Redact (обезличивание таблиц). Логика замен — в engine.js (window.Anon). */
(function () {
  'use strict';
  const A = window.Anon;
  const $ = id => document.getElementById(id);
  const icon = (name, cls = 'i') => `<svg class="${cls}" aria-hidden="true"><use href="#i-${name}"></use></svg>`;
  const esc = s => String(s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

  const S = {
    pending: null,    // настройки, загруженные до файла с данными
    loadNotes: [],    // что пришлось исправить в файле при чтении
    kind: null,       // 'xlsx' | 'csv'
    fileName: '',
    buffer: null,     // исходный файл — для записи результата в «чистую» копию книги
    wb: null,
    sheetIndex: 0,
    grid: null,
    csv: null,        // { delim, decSep }
    an: null,
    cols: {},         // настройки по номеру столбца
    result: null,
  };

  $('appVersion').textContent = 'v' + A.VERSION;

  // ---------- Тема: системная по умолчанию, выбор пользователя запоминается ----------
  $('btnTheme').addEventListener('click', () => {
    const cur = document.documentElement.dataset.theme ||
      (window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem('redact-theme', next); } catch (e) { /* без сохранения */ }
  });

  // ---------- Уведомления ----------
  let toastTimer;
  function toast(msg, ms = 3500, kind) {
    const t = $('toast');
    kind = kind || (/^(ошибк|не удалось)/i.test(msg) ? 'error' : 'ok');
    t.className = 'toast ' + kind;
    t.innerHTML = icon(kind === 'error' ? 'circle-alert' : 'circle-check') + `<span>${esc(msg)}</span>`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add('hidden'), ms);
  }

  // ---------- localStorage: запоминаем настройки для одинаковой структуры файла ----------
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* недоступно — не страшно */ } },
  };
  const layoutKey = () => S.an ? 'anon-layout:' + A.cyrb128(S.an.columns.map(c => c.name).join('␟'))[0].toString(36) : null;

  // ---------- Загрузка файла ----------
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', ev => { const f = ev.dataTransfer.files[0]; if (f) loadFile(f); });
  $('file').addEventListener('change', ev => { const f = ev.target.files[0]; if (f) loadFile(f); ev.target.value = ''; });

  async function loadFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    $('fileInfo').textContent = 'Читаем файл…';
    try {
      const buf = await file.arrayBuffer();
      S.fileName = file.name;
      S.buffer = buf;
      S.loadNotes = [];
      if (ext === 'xls') throw new Error('Формат .xls (Excel 97–2003) не поддерживается. Откройте файл в Excel и сохраните как .xlsx.');
      if (ext === 'csv' || ext === 'txt') {
        S.kind = 'csv';
        let text = new TextDecoder('utf-8').decode(buf);
        if (text.includes('�')) text = new TextDecoder('windows-1251').decode(buf);
        const parsed = A.parseCsv(text);
        S.wb = null;
        S.grid = parsed.grid;
        S.csv = { delim: parsed.delim, decSep: parsed.decSep };
        $('sheet').innerHTML = '<option>CSV</option>';
        $('sheet').disabled = true;
      } else {
        S.kind = 'xlsx';
        // Файл заранее приводится к виду, который ExcelJS читает без потерь (выгрузки из 1С и т.п.)
        const fixed = await A.repairXlsx(buf, JSZip);
        if (fixed.changed) S.buffer = fixed.buffer;
        S.loadNotes = fixed.notes;
        S.wb = new ExcelJS.Workbook();
        await S.wb.xlsx.load(S.buffer);
        if (!S.wb.worksheets.length) throw new Error('В книге нет листов.');
        $('sheet').innerHTML = S.wb.worksheets.map((ws, i) => `<option value="${i}">${esc(ws.name)}${ws.state && ws.state !== 'visible' ? ' (скрытый)' : ''}</option>`).join('');
        $('sheet').disabled = S.wb.worksheets.length < 2;
        // по умолчанию — самый большой лист
        let best = 0;
        S.wb.worksheets.forEach((ws, i) => { if (ws.rowCount * ws.columnCount > S.wb.worksheets[best].rowCount * S.wb.worksheets[best].columnCount) best = i; });
        $('sheet').value = String(best);
        S.sheetIndex = best;
        S.grid = A.gridFromWorksheet(S.wb.worksheets[best]);
      }
      $('fileInfo').innerHTML = `${icon('file-spreadsheet', 'i i20')}<b>${esc(file.name)}</b><span class="muted">${(file.size / 1024).toFixed(0)} КБ` +
        (S.loadNotes.length ? ' · файл исправлен при чтении, подробности в шаге 4' : '') + '</span>';
      if (!$('seed').value) newSeed();
      $('settingsInfo').textContent = '';
      setupSheet(true);
      ['step2', 'step3', 'step4'].forEach(id => $(id).classList.remove('hidden'));
      if (S.pending) { applySettings(S.pending.p, S.pending.name); }
    } catch (e) {
      console.error(e);
      $('fileInfo').textContent = '';
      toast('Не удалось прочитать файл: ' + e.message, 7000);
    }
  }

  $('sheet').addEventListener('change', () => {
    S.sheetIndex = +$('sheet').value;
    S.grid = A.gridFromWorksheet(S.wb.worksheets[S.sheetIndex]);
    setupSheet(true);
  });
  $('headerRow').addEventListener('change', () => {
    const v = Math.max(1, Math.min(S.grid.maxRow, parseInt($('headerRow').value, 10) || 1));
    $('headerRow').value = v;
    setupSheet(false, v);
  });
  $('keepTotals').addEventListener('change', schedulePreview);

  function setupSheet(autoHeader, headerRow) {
    const hr = autoHeader ? A.detectHeaderRow(S.grid) : headerRow;
    $('headerRow').value = hr;
    S.an = A.analyze(S.grid, hr);
    S.cols = {};
    const saved = store.get(layoutKey());
    for (const col of S.an.columns) S.cols[col.c] = A.suggest(col);
    if (saved && saved.columns) applySettingsColumns(saved.columns);
    const tot = S.an.totalsRows.size;
    $('rowsInfo').textContent = `Строк данных: ${S.an.dataRows.length - tot}${tot ? `, строк итогов: ${tot}` : ''} · столбцов: ${S.an.columns.length}` + (saved ? ' · применены настройки, сохранённые для этой структуры' : '');
    renderColumns();
    schedulePreview();
  }

  // ---------- Таблица столбцов ----------
  function fmt(v) {
    if (v === null || v === undefined) return '';
    if (A.isDate(v)) {
      const d = v;
      const p = n => String(n).padStart(2, '0');
      const time = (d.getUTCHours() || d.getUTCMinutes()) ? ` ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}` : '';
      return `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}${time}`;
    }
    if (typeof v === 'number') return v.toLocaleString('ru-RU', { maximumFractionDigits: 2 });
    return String(v);
  }

  function paramsHtml(col, cfg) {
    const m = cfg.method;
    const parts = [];
    if (m === 'pseudo') parts.push(`<label>префикс <input type="text" data-p="prefix" value="${esc(cfg.prefix || '')}"></label>`);
    if (m === 'noise') parts.push(`<label>шум, ± % <input type="number" data-p="pct" min="0" step="1" value="${esc(cfg.pct)}"></label>`);
    if (m === 'round') parts.push(`<label>шаг <input type="number" data-p="step" min="0" step="any" value="${esc(cfg.step)}"></label>`);
    if (A.DICT_METHODS.has(m)) parts.push(`<label title="Столбцы с одинаковым именем словаря получают одинаковые замены для одинаковых значений">словарь <input type="text" data-p="group" value="${esc(cfg.group || col.letter)}"></label>`);
    if (m === 'scale') parts.push('<span class="muted">K из шага 3</span>');
    if (m === 'shift' || m === 'smart') parts.push('<span class="muted">D из шага 3</span>');
    return `<div class="params">${parts.join('')}</div>`;
  }

  function rowHtml(col) {
    const cfg = S.cols[col.c];
    const types = Object.entries(A.TYPE_NAMES).map(([k, n]) => `<option value="${k}"${k === cfg.type ? ' selected' : ''}>${n}${k === col.detected ? ' •' : ''}</option>`).join('');
    const methods = A.METHODS[cfg.type].map(([k, n]) => `<option value="${k}" title="${esc(A.METHOD_HINTS[k] || '')}"${k === cfg.method ? ' selected' : ''}>${n}</option>`).join('');
    const samples = col.samples.map(v => `<span title="${esc(fmt(v))}">${esc(fmt(v))}</span>`).join('');
    return `<tr data-c="${col.c}" class="${cfg.method !== 'keep' ? 'active' : ''}">
      <td class="colname"><b>${col.letter}</b>${esc(col.name)}</td>
      <td><select data-k="type" title="• — тип, определённый автоматически">${types}</select></td>
      <td class="samples">${samples}</td>
      <td class="r counts">${col.nonEmpty} / ${col.uniques}</td>
      <td><select data-k="method" title="${esc(A.METHOD_HINTS[cfg.method] || '')}">${methods}</select></td>
      <td>${paramsHtml(col, cfg)}</td>
    </tr>`;
  }

  function renderColumns() {
    $('colsTable').tBodies[0].innerHTML = S.an.columns.map(rowHtml).join('');
  }
  function rerenderRow(c) {
    const tr = $('colsTable').querySelector(`tr[data-c="${c}"]`);
    const col = S.an.columns.find(x => x.c === c);
    const tmp = document.createElement('tbody');
    tmp.innerHTML = rowHtml(col);
    tr.replaceWith(tmp.firstElementChild);
  }

  $('colsTable').addEventListener('change', ev => {
    const tr = ev.target.closest('tr[data-c]');
    if (!tr) return;
    const c = +tr.dataset.c;
    const cfg = S.cols[c];
    if (ev.target.dataset.k === 'type') {
      cfg.type = ev.target.value;
      if (!A.METHODS[cfg.type].some(([k]) => k === cfg.method)) cfg.method = 'keep';
      rerenderRow(c);
    } else if (ev.target.dataset.k === 'method') {
      cfg.method = ev.target.value;
      rerenderRow(c);
    } else if (ev.target.dataset.p) {
      cfg[ev.target.dataset.p] = ev.target.value;
    }
    rememberLayout();
    schedulePreview();
  });

  $('btnAuto').addEventListener('click', () => {
    for (const col of S.an.columns) S.cols[col.c] = A.suggest(col);
    renderColumns(); rememberLayout(); schedulePreview();
  });
  $('btnReset').addEventListener('click', () => {
    for (const col of S.an.columns) S.cols[col.c].method = 'keep';
    renderColumns(); rememberLayout(); schedulePreview();
  });

  // ---------- Общие параметры ----------
  function newSeed() {
    const seed = A.randomSeed();
    $('seed').value = seed;
    $('factor').value = A.factorFromSeed(seed);
    $('shiftDays').value = A.shiftFromSeed(seed);
  }
  $('btnSeed').addEventListener('click', () => { newSeed(); schedulePreview(); });
  $('btnFactor').addEventListener('click', () => { $('factor').value = A.factorFromSeed(A.randomSeed()); schedulePreview(); });
  $('btnShift').addEventListener('click', () => { $('shiftDays').value = A.shiftFromSeed(A.randomSeed()); schedulePreview(); });
  ['seed', 'factor', 'shiftDays', 'keepForm', 'clearHF'].forEach(id => $(id).addEventListener('change', schedulePreview));

  function config() {
    return {
      seed: $('seed').value || 'default',
      factor: parseFloat($('factor').value) || 1,
      shiftDays: parseInt($('shiftDays').value, 10) || 0,
      keepForm: $('keepForm').checked,
      keepTotals: $('keepTotals').checked,
      columns: S.cols,
    };
  }

  // ---------- Файл настроек (JSON) ----------
  const COL_KEYS = ['type', 'method', 'prefix', 'pct', 'step', 'group'];
  function settingsColumns() {
    const out = {};
    for (const col of S.an.columns) {
      const cfg = S.cols[col.c];
      const item = { letter: col.letter };
      for (const k of COL_KEYS) if (cfg[k] !== undefined && cfg[k] !== '') item[k] = cfg[k];
      out[col.name] = item;
    }
    return out;
  }
  // Столбцы сопоставляются по названию, а если названия нет в файле — по букве столбца
  function applySettingsColumns(columns) {
    const names = new Set(S.an.columns.map(c => c.name));
    const byLetter = {};
    for (const [name, item] of Object.entries(columns)) if (!names.has(name) && item.letter) byLetter[item.letter] = name;
    const used = new Set();
    let hit = 0;
    for (const col of S.an.columns) {
      const name = columns[col.name] ? col.name : byLetter[col.letter];
      if (!name) continue;
      const item = columns[name];
      const cfg = { ...S.cols[col.c] };
      for (const k of COL_KEYS) if (item[k] !== undefined) cfg[k] = item[k];
      if (!A.METHODS[cfg.type] || !A.METHODS[cfg.type].some(([k]) => k === cfg.method)) continue;
      S.cols[col.c] = cfg;
      used.add(name);
      hit++;
    }
    return { hit, missing: Object.keys(columns).filter(n => !used.has(n)) };
  }
  function rememberLayout() {
    const k = layoutKey();
    if (k) store.set(k, { columns: settingsColumns() });
  }
  function collectSettings() {
    const p = {
      app: 'Redact',
      appVersion: A.VERSION,
      format: 1,
      savedAt: new Date().toISOString(),
      sourceFile: S.fileName,
      sheet: S.kind === 'xlsx' ? S.wb.worksheets[S.sheetIndex].name : null,
      headerRow: S.an.headerRow,
      options: {
        keepTotals: $('keepTotals').checked,
        keepForm: $('keepForm').checked,
        clearMeta: $('clearMeta').checked,
        clearNotes: $('clearNotes').checked,
        clearHF: $('clearHF').checked,
      },
      columns: settingsColumns(),
    };
    if ($('profileWithKey').checked) p.secret = { seed: $('seed').value, factor: +$('factor').value, shiftDays: +$('shiftDays').value };
    return p;
  }
  function applyOptions(p) {
    const o = p.options || p; // формат ранней версии хранил флаги в корне
    for (const id of ['keepTotals', 'keepForm', 'clearMeta', 'clearNotes', 'clearHF']) if (typeof o[id] === 'boolean') $(id).checked = o[id];
    const sec = p.secret || (p.seed ? p : null);
    if (sec && sec.seed) {
      $('seed').value = sec.seed;
      if (sec.factor !== undefined) $('factor').value = sec.factor;
      if (sec.shiftDays !== undefined) $('shiftDays').value = sec.shiftDays;
    }
  }
  // Применить настройки к загруженному файлу
  function applySettings(p, name) {
    applyOptions(p);
    if (S.kind === 'xlsx' && p.sheet) {
      const i = S.wb.worksheets.findIndex(ws => ws.name === p.sheet);
      if (i >= 0 && i !== S.sheetIndex) {
        S.sheetIndex = i;
        $('sheet').value = String(i);
        S.grid = A.gridFromWorksheet(S.wb.worksheets[i]);
        setupSheet(true);
      }
    }
    if (p.headerRow && p.headerRow !== S.an.headerRow && p.headerRow <= S.grid.maxRow) setupSheet(false, p.headerRow);
    const { hit, missing } = applySettingsColumns(p.columns || {});
    renderColumns(); rememberLayout(); schedulePreview();
    $('settingsInfo').innerHTML = `Настройки <b>${esc(name)}</b> применены: столбцов ${hit} из ${S.an.columns.length}` +
      (missing.length ? `. <span class="text-danger">Нет в файле: ${missing.map(esc).join(', ')}</span>` : '');
    toast(`Настройки применены: ${hit} из ${S.an.columns.length} ${A.plural(S.an.columns.length, 'столбца', 'столбцов', 'столбцов')}`);
  }
  async function readSettingsFile(f) {
    const p = JSON.parse(await f.text());
    if (!p || typeof p !== 'object' || !p.columns) throw new Error('это не файл настроек Redact (нет раздела columns)');
    return p;
  }

  $('btnSaveProfile').addEventListener('click', () => {
    if (!S.an) return;
    const p = collectSettings();
    download(new Blob([JSON.stringify(p, null, 2)], { type: 'application/json' }), `Redact_settings_${baseName()}.json`);
    toast(p.secret ? 'Настройки сохранены вместе с ключом — не передавайте этот файл вместе с данными' : 'Настройки сохранены');
  });
  // Загрузить настройки можно и до файла с данными — тогда они применятся при его загрузке
  for (const id of ['profileFile', 'profileFile1']) {
    $(id).addEventListener('change', async ev => {
      const f = ev.target.files[0];
      ev.target.value = '';
      if (!f) return;
      try {
        const p = await readSettingsFile(f);
        if (S.an) applySettings(p, f.name);
        else {
          S.pending = { p, name: f.name };
          $('pendingInfo').innerHTML = `Загружены настройки <b>${esc(f.name)}</b> — применятся к файлу с данными.`;
        }
      } catch (e) { toast('Не удалось прочитать настройки: ' + e.message, 6000); }
    });
  }

  // ---------- Предпросмотр ----------
  let previewTimer;
  function schedulePreview() { clearTimeout(previewTimer); previewTimer = setTimeout(runPreview, 250); }
  $('btnPreview').addEventListener('click', runPreview);

  function runPreview() {
    if (!S.an) return;
    const t0 = performance.now();
    S.result = A.transform(S.grid, S.an, config());
    const ms = Math.round(performance.now() - t0);
    const active = S.an.columns.filter(col => S.cols[col.c].method !== 'keep');
    const { changes, mapping, warnings, stats } = S.result;
    $('previewStats').textContent = active.length
      ? `Будет обезличено ячеек: ${stats.changed.toLocaleString('ru-RU')} в ${stats.columns} ${A.plural(stats.columns, 'столбце', 'столбцах', 'столбцах')} · уникальных замен: ${mapping.length.toLocaleString('ru-RU')} · ${ms} мс`
      : 'Столбцы для обезличивания не выбраны. Выберите метод хотя бы для одного столбца на шаге 2.';
    const table = $('previewTable');
    $('legend').classList.toggle('hidden', !active.length);
    // «очищено» в легенде — только если в предпросмотре есть очищенные ячейки
    $('legendCleared').classList.toggle('hidden', !active.some(col => S.cols[col.c].method === 'clear'));
    if (!active.length) { table.innerHTML = ''; }
    else {
      const rows = S.an.dataRows.filter(r => active.some(col => changes.has(r + ':' + col.c))).slice(0, 15);
      const head = `<thead><tr><th class="r">Строка</th>${active.map(col => `<th>${esc(col.letter)} · ${esc(col.name)}</th>`).join('')}</tr></thead>`;
      // RedactedValue: найдено (маркер) → заменено (токен) или скрыто (плашка)
      const body = rows.map(r => `<tr><td class="r rownum">${r}</td>${active.map(col => {
        const cell = S.grid.rows[r][col.c];
        const old = cell && !cell.slave ? cell.v : null;
        const key = r + ':' + col.c;
        if (!changes.has(key)) return `<td>${old === null ? '' : `<span class="rv rv-original">${esc(fmt(old))}</span>`}</td>`;
        const nv = changes.get(key);
        const after = (nv === null || nv === '')
          ? `<span class="rv rv-redacted" style="width:${Math.min(40, Math.max(3, fmt(old).length))}ch" role="img" aria-label="Скрыто: ${esc(col.name)}"></span>`
          : `<span class="rv rv-token">${esc(fmt(nv))}</span>`;
        return `<td title="${esc(fmt(old))} → ${esc(fmt(nv) || 'скрыто')}"><div class="pair"><span class="rv rv-detected">${esc(fmt(old))}</span><span class="line"><span class="rv-arrow" aria-hidden="true">→</span>${after}</span></div></td>`;
      }).join('')}</tr>`).join('');
      table.innerHTML = head + `<tbody>${body}</tbody>`;
    }
    const w = [...S.loadNotes, ...warnings];
    const leaks = [...S.result.leaks];
    if (S.kind === 'xlsx') {
      const hf = A.headerFooterText(S.wb.worksheets[S.sheetIndex]);
      if (hf.length) {
        if ($('clearHF').checked) w.push(`В колонтитулах есть текст («${hf.join('», «').slice(0, 150)}»). Он будет удалён.`);
        else leaks.push(`В колонтитулах есть текст («${hf.join('», «').slice(0, 150)}»), и он останется в файле. Включите «Очистить колонтитулы» на шаге 3.`);
      }
    }
    if (S.kind === 'xlsx' && S.wb.worksheets.length > 1) {
      const others = S.wb.worksheets.filter((_, i) => i !== S.sheetIndex).map(ws => `«${ws.name}»`).join(', ');
      w.push(`В книге есть другие листы (${others}). Они не обезличиваются и в результат не попадут.`);
    }
    const notice = (cls, ic, title, items) => items.length
      ? `<div class="notice ${cls}"><div class="notice-title">${icon(ic, 'i i20')}${title}</div><ul>${items.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
    $('warnings').innerHTML =
      notice('danger', 'triangle-alert', `Возможна утечка данных: ${leaks.length}`, leaks) +
      notice('', 'circle-alert', 'Проверьте перед выгрузкой', w);
  }

  // ---------- Выгрузка ----------
  const baseName = () => S.fileName.replace(/\.[^.]+$/, '');
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  $('btnDownload').addEventListener('click', async () => {
    if (!S.an) return;
    const btn = $('btnDownload');
    btn.disabled = true;
    try {
      runPreview();
      const { changes } = S.result;
      if (S.kind === 'csv') {
        const text = A.gridToCsv(S.grid, changes, S.csv.delim, S.csv.decSep);
        download(new Blob([text], { type: 'text/csv;charset=utf-8' }), baseName() + '_redact.csv');
      } else {
        // Пишем в свежую копию исходной книги — так сохраняется всё оформление
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(S.buffer);
        const ws = wb.worksheets[S.sheetIndex];
        const others = wb.worksheets.filter(x => x !== ws);
        if (others.length) {
          others.forEach(x => wb.removeWorksheet(x.id));
          wb.definedNames.model = []; // имена могли ссылаться на удалённые листы
        }
        A.applyToWorksheet(ws, changes);
        if ($('clearMeta').checked) A.clearWorkbookMeta(wb);
        if ($('clearNotes').checked) A.clearNotes(ws);
        if ($('clearHF').checked) A.clearHeaderFooter(ws);
        wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
        // ExcelJS ссылается на тему оформления, даже если в исходной книге её не было, но сам файл темы не пишет —
        // Excel тогда предлагает «восстановить» книгу. Без своей темы записывается стандартная.
        if (!wb._themes || !Object.keys(wb._themes).length) wb._themes = undefined;
        const out = await wb.xlsx.writeBuffer();
        download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '_redact.xlsx');
      }
      toast(`Обезличено ячеек: ${S.result.stats.changed.toLocaleString('ru-RU')}. Файл сохранён в «Загрузки»`);
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить файл: ' + e.message, 7000);
    } finally { btn.disabled = false; }
  });

  $('btnKey').addEventListener('click', async () => {
    if (!S.an) return;
    runPreview();
    const cfg = config();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Соответствия');
    ws.columns = [
      { header: 'Словарь', key: 'dict', width: 10 },
      { header: 'Столбцы', key: 'columns', width: 28 },
      { header: 'Метод', key: 'method', width: 10 },
      { header: 'Было', key: 'orig', width: 45 },
      { header: 'Стало', key: 'repl', width: 45 },
    ];
    ws.getRow(1).font = { bold: true };
    ws.addRows(S.result.mapping);
    ws.autoFilter = 'A1:E1';
    const ps = wb.addWorksheet('Параметры');
    ps.columns = [{ header: 'Параметр', key: 'k', width: 34 }, { header: 'Значение', key: 'v', width: 40 }];
    ps.getRow(1).font = { bold: true };
    ps.addRows([
      { k: 'Исходный файл', v: S.fileName },
      { k: 'Дата обработки', v: new Date().toLocaleString('ru-RU') },
      { k: 'Секретный ключ', v: cfg.seed },
      { k: 'Коэффициент K (число_анонимное / K = исходное)', v: cfg.factor },
      { k: 'Сдвиг дат D, дней (исходная = анонимная − D)', v: cfg.shiftDays },
      ...S.an.columns.filter(col => S.cols[col.c].method !== 'keep').map(col => ({ k: `${col.letter} · ${col.name}`, v: A.METHODS[S.cols[col.c].type].find(([k]) => k === S.cols[col.c].method)[1] })),
    ]);
    const out = await wb.xlsx.writeBuffer();
    download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '_KEY.xlsx');
    toast('Ключ соответствий сохранён. Храните его отдельно от обезличенного файла', 5000);
  });
})();

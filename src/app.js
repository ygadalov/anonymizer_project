/* Интерфейс Redact (обезличивание таблиц). Логика замен — в engine.js (window.Anon). */
(function () {
  'use strict';
  const A = window.Anon;
  const $ = id => document.getElementById(id);
  const icon = (name, cls = 'i') => `<svg class="${cls}" aria-hidden="true"><use href="#i-${name}"></use></svg>`;
  const KEY_SHEET = 'redact-key';
  const esc = s => String(s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

  const S = {
    pending: null,    // настройки, загруженные до файла с данными
    loadNotes: [],    // что пришлось исправить в файле при чтении
    kind: null,       // 'xlsx' | 'csv'
    fileName: '',
    buffer: null,     // исходный файл — для записи результата в «чистую» копию книги
    wb: null,
    sheets: [],       // состояние каждого листа: { name, grid, an, cols, saved }
    sheetIndex: 0,    // текущий лист; grid, an и cols ниже — его же, для краткости
    grid: null,
    csv: null,        // { delim, decSep }
    an: null,
    cols: {},         // настройки текущего листа по номеру столбца
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
    t.innerHTML = icon(kind === 'error' ? 'circle-alert' : kind === 'warn' ? 'triangle-alert' : 'circle-check') + `<span>${esc(msg)}</span>`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add('hidden'), ms);
  }

  // ---------- localStorage: запоминаем настройки для одинаковой структуры файла ----------
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* недоступно — не страшно */ } },
  };
  const layoutKey = (an = S.an) => an ? 'anon-layout:' + A.cyrb128(an.columns.map(c => c.name).join('␟'))[0].toString(36) : null;

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
      if (ext === 'docx' || ext === 'doc') throw new Error('это документ Word — откройте вкладку «Документ Word»');
      const xls = window.RedactXls.isXls(buf, file.name);
      if (!xls && (ext === 'csv' || ext === 'txt')) {
        S.kind = 'csv';
        let text = new TextDecoder('utf-8').decode(buf);
        if (text.includes('�')) text = new TextDecoder('windows-1251').decode(buf);
        const parsed = A.parseCsv(text);
        S.wb = null;
        S.csv = { delim: parsed.delim, decSep: parsed.decSep };
        S.sheets = [makeState('CSV', parsed.grid)];
        S.sheetIndex = 0;
      } else {
        S.kind = 'xlsx';
        // .xls (Excel 97–2003) сначала преобразуется в .xlsx — дальше всё как с обычной книгой
        const conv = xls ? await window.RedactXls.toXlsx(buf) : null;
        if (conv) S.buffer = conv.buffer;
        // Файл заранее приводится к виду, который ExcelJS читает без потерь (выгрузки из 1С и т.п.)
        const fixed = await A.repairXlsx(S.buffer, JSZip);
        if (fixed.changed) S.buffer = fixed.buffer;
        S.loadNotes = [...(conv ? conv.notes : []), ...fixed.notes];
        S.wb = new ExcelJS.Workbook();
        await S.wb.xlsx.load(S.buffer);
        if (!S.wb.worksheets.length) throw new Error('В книге нет листов.');
        // все листы разбираются сразу: методы подбираются для каждого, а в списке листов видно, какие обезличиваются
        S.sheets = S.wb.worksheets.map(ws => makeState(ws.name, A.gridFromWorksheet(ws)));
        linkGroupsAcrossSheets();
        // открывается самый большой лист
        let best = 0;
        S.wb.worksheets.forEach((ws, i) => { if (ws.rowCount * ws.columnCount > S.wb.worksheets[best].rowCount * S.wb.worksheets[best].columnCount) best = i; });
        S.sheetIndex = best;
      }
      $('fileInfo').innerHTML = `${icon('file-spreadsheet', 'i i20')}<b>${esc(file.name)}</b><span class="muted">${(file.size / 1024).toFixed(0)} КБ` +
        (xls ? ' · преобразован из .xls в .xlsx, подробности в шаге 4' : S.loadNotes.length ? ' · файл исправлен при чтении, подробности в шаге 4' : '') + '</span>';
      if (!$('seed').value) newSeed();
      $('settingsInfo').textContent = '';
      $('btnClear').classList.remove('hidden');
      ['step2', 'step3', 'step4'].forEach(id => $(id).classList.remove('hidden'));
      showSheet(S.sheetIndex);
      if (S.pending) { applySettings(S.pending.p, S.pending.name); }
    } catch (e) {
      console.error(e);
      $('fileInfo').textContent = '';
      toast('Не удалось прочитать файл: ' + e.message, 7000);
    }
  }

  // ---------- Очистить: начать работу с другим файлом без перезагрузки страницы ----------
  // Убираются файл, настройки столбцов, предпросмотр и пароль ключа; общие параметры шага 3 остаются,
  // чтобы несколько выгрузок можно было обезличить одним ключом
  function clearRedact() {
    Object.assign(S, {
      pending: null, loadNotes: [], kind: null, fileName: '', buffer: null, wb: null,
      sheets: [], sheetIndex: 0, grid: null, csv: null, an: null, cols: {}, result: null,
    });
    clearTimeout(previewTimer);
    ['step2', 'step3', 'step4'].forEach(id => $(id).classList.add('hidden'));
    $('btnClear').classList.add('hidden');
    $('fileInfo').textContent = '';
    $('pendingInfo').textContent = 'Сохранённые настройки (JSON) можно загрузить до или после файла с данными.';
    $('sheet').innerHTML = '';
    $('colsTable').tBodies[0].innerHTML = '';
    for (const id of ['rowsInfo', 'settingsInfo', 'previewStats', 'warnings', 'previewTable']) $(id).innerHTML = '';
    $('legend').classList.add('hidden');
    $('keyPass').value = '';
    $('keyPass2').value = '';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    toast('Готово к новому файлу. Общие параметры сохранены');
  }
  $('btnClear').addEventListener('click', async () => {
    const ok = await window.RedactUI.confirm({
      title: 'Очистить и начать с другим файлом?',
      text: 'Будут убраны загруженный файл, настройки столбцов и предпросмотр. Секретная фраза, коэффициент, сдвиг дат и флажки шага 3 сохранятся.',
      confirm: 'Да, очистить',
    });
    if (ok) clearRedact();
  });

  // ---------- Листы ----------
  // Состояние листа: разбор, подобранные методы (или сохранённые для такой же структуры столбцов)
  function makeState(name, grid, headerRow) {
    const an = A.analyze(grid, headerRow || A.detectHeaderRow(grid));
    const cols = {};
    for (const col of an.columns) cols[col.c] = A.suggest(col);
    const saved = store.get(layoutKey(an));
    const st = { name, grid, an, cols, saved: !!(saved && saved.columns) };
    if (st.saved) applySettingsColumns(saved.columns, st);
    return st;
  }
  const isActive = cfg => cfg && cfg.method && cfg.method !== 'keep';
  const activeCount = st => st.an.columns.filter(col => isActive(st.cols[col.c])).length;
  const usedSheets = () => S.sheets.filter(st => activeCount(st) > 0);
  const resultSheets = () => S.sheets.filter(st => activeCount(st) > 0 || !$('dropUnused').checked);

  // Однотипные столбцы на разных листах (контрагенты, ИНН, организация) получают общую группу замен —
  // одна и та же компания на «ДЗ» и «КЗ» станет одним псевдонимом. Группы, заданные вручную, не трогаются
  const GROUP_KINDS = [
    [/инн/i, 'ИНН'],
    [/^код\s+(контрагент|клиент|покупател|поставщик)/i, 'Код контрагента'],
    [/организац/i, 'Организация'],
    [/контрагент|покупател|поставщик|клиент|дебитор|кредитор|counterpart|customer|client|vendor|supplier/i, 'Контрагент'],
  ];
  function linkGroupsAcrossSheets() {
    if (usedSheets().length < 2) return;
    const byKind = new Map();
    S.sheets.forEach(st => st.an.columns.forEach(col => {
      const cfg = st.cols[col.c];
      if (!A.DICT_METHODS.has(cfg.method) || (cfg.group && cfg.group !== col.letter)) return;
      const kind = (GROUP_KINDS.find(([re]) => re.test(col.name)) || [])[1];
      if (!kind) return;
      if (!byKind.has(kind)) byKind.set(kind, []);
      byKind.get(kind).push({ st, cfg });
    }));
    for (const [kind, list] of byKind) {
      if (new Set(list.map(x => x.st)).size < 2) continue; // только если такие столбцы есть на разных листах
      for (const { cfg } of list) cfg.group = kind;
    }
  }

  // Подпись листа в списке: обезличивается ли он и сколько столбцов
  function renderSheetOptions() {
    const sel = $('sheet');
    sel.innerHTML = S.sheets.map((st, i) => {
      const n = activeCount(st);
      const hidden = S.wb && S.wb.worksheets[i].state && S.wb.worksheets[i].state !== 'visible' ? ' (скрытый)' : '';
      const status = n ? `${n} ${A.plural(n, 'столбец', 'столбца', 'столбцов')}` : 'не обезличивается';
      return `<option value="${i}">${esc(st.name)}${hidden} · ${status}</option>`;
    }).join('');
    sel.value = String(S.sheetIndex);
    sel.disabled = S.sheets.length < 2;
  }

  function showSheet(i) {
    S.sheetIndex = i;
    const st = S.sheets[i];
    S.grid = st.grid; S.an = st.an; S.cols = st.cols;
    $('headerRow').value = st.an.headerRow;
    const tot = st.an.totalsRows.size;
    $('rowsInfo').textContent = `Строк данных: ${st.an.dataRows.length - tot}${tot ? `, строк итогов: ${tot}` : ''} · столбцов: ${st.an.columns.length}` + (st.saved ? ' · применены настройки, сохранённые для этой структуры' : '');
    renderSheetOptions();
    renderColumns();
    schedulePreview();
  }

  $('sheet').addEventListener('change', () => showSheet(+$('sheet').value));
  $('headerRow').addEventListener('change', () => {
    const v = Math.max(1, Math.min(S.grid.maxRow, parseInt($('headerRow').value, 10) || 1));
    $('headerRow').value = v;
    setupSheet(false, v);
  });
  $('keepTotals').addEventListener('change', schedulePreview);
  $('dropUnused').addEventListener('change', schedulePreview);

  // Строка заголовков изменена — текущий лист разбирается заново
  function setupSheet(autoHeader, headerRow) {
    const st = S.sheets[S.sheetIndex];
    S.sheets[S.sheetIndex] = makeState(st.name, st.grid, autoHeader ? null : headerRow);
    showSheet(S.sheetIndex);
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
    const tipAttr = t => `title="${esc(t)}"`;
    if (m === 'pseudo') parts.push(`<label ${tipAttr('Начало псевдонима: «Контрагент» → «Контрагент 017 ООО». Номер подставляется автоматически')}>префикс <input type="text" data-p="prefix" value="${esc(cfg.prefix || '')}"></label>`);
    if (m === 'noise') parts.push(`<label ${tipAttr('На сколько процентов в обе стороны случайно меняется каждое число')}>отклонение, ± % <input type="number" data-p="pct" min="0" step="1" value="${esc(cfg.pct)}"></label>`);
    if (m === 'round') parts.push(`<label ${tipAttr('До какого шага округлять: 1000 → 274 800 станет 275 000')}>шаг <input type="number" data-p="step" min="0" step="any" value="${esc(cfg.step)}"></label>`);
    if (A.DICT_METHODS.has(m)) parts.push(`<label ${tipAttr(`Одинаковые значения в столбцах одной группы получают одинаковую замену. Например, укажите одну группу у «Покупателя» и «Грузополучателя» — одна и та же компания в обоих станет одним псевдонимом. По умолчанию у каждого столбца своя группа — его буква (${col.letter})`)}>группа замен <input type="text" data-p="group" value="${esc(cfg.group || col.letter)}"></label>`);
    if (m === 'scale') parts.push('<span class="muted">коэффициент — в шаге 3</span>');
    if (m === 'shift' || m === 'smart') parts.push('<span class="muted">сдвиг дат — в шаге 3</span>');
    // необратимые методы: пометка остаётся, пока метод выбран
    if (A.IRREVERSIBLE.has(m)) parts.push(`<span class="irrev">${icon('triangle-alert', 'i')}${m === 'clear'
      ? 'Значения удаляются — по ключу их не вернуть'
      : 'Необратимо — по ключу исходные значения не вернуть'}</span>`);
    return `<div class="params">${parts.join('')}</div>`;
  }

  function rowHtml(col) {
    const cfg = S.cols[col.c];
    const types = Object.entries(A.TYPE_NAMES).map(([k, n]) => `<option value="${k}"${k === cfg.type ? ' selected' : ''}${k === col.detected ? ' title="Определено автоматически по значениям столбца"' : ''}>${n}${k === col.detected ? ' · авто' : ''}</option>`).join('');
    const methods = A.METHODS[cfg.type].map(([k, n]) => `<option value="${k}" title="${esc(A.METHOD_HINTS[k] || '')}"${k === cfg.method ? ' selected' : ''}>${n}</option>`).join('');
    const samples = col.samples.map(v => `<span title="${esc(fmt(v))}">${esc(fmt(v))}</span>`).join('');
    return `<tr data-c="${col.c}" class="${cfg.method !== 'keep' ? 'active' : ''}">
      <td class="colname"><b>${col.letter}</b>${esc(col.name)}</td>
      <td><select data-k="type" title="Тип данных в столбце: от него зависят доступные методы. «авто» — тип, определённый по значениям">${types}</select></td>
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
      if (A.IRREVERSIBLE.has(cfg.method)) {
        const col = S.an.columns.find(x => x.c === c);
        toast(cfg.method === 'clear'
          ? `Столбец «${col.name}» будет очищен: значения удалятся из файла, и по ключу их не вернуть`
          : `Метод для «${col.name}» необратим: по ключу исходные значения не вернуть`, 6000, 'warn');
      }
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
  // все непустые столбцы листа — восстанавливаемыми методами (по названию или по типу данных)
  $('btnAll').addEventListener('click', () => {
    let n = 0;
    for (const col of S.an.columns) { S.cols[col.c] = A.suggestAll(col); if (S.cols[col.c].method !== 'keep') n++; }
    linkGroupsAcrossSheets();
    renderColumns(); rememberLayout(); schedulePreview();
    toast(`Метод задан для ${n} ${A.plural(n, 'столбца', 'столбцов', 'столбцов')} из ${S.an.columns.length}: все восстанавливаются по ключу. Проверьте и при необходимости измените вручную`, 6000);
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
    };
  }
  const bookForEngine = (list = S.sheets) => list.map(st => ({ name: st.name, grid: st.grid, an: st.an, columns: st.cols }));

  // ---------- Файл настроек (JSON) ----------
  const COL_KEYS = ['type', 'method', 'prefix', 'pct', 'step', 'group'];
  function settingsColumns(st = S.sheets[S.sheetIndex]) {
    const out = {};
    for (const col of st.an.columns) {
      const cfg = st.cols[col.c];
      const item = { letter: col.letter };
      for (const k of COL_KEYS) if (cfg[k] !== undefined && cfg[k] !== '') item[k] = cfg[k];
      out[col.name] = item;
    }
    return out;
  }
  // Столбцы сопоставляются по названию, а если названия нет в файле — по букве столбца
  function applySettingsColumns(columns, st = S.sheets[S.sheetIndex]) {
    const names = new Set(st.an.columns.map(c => c.name));
    const byLetter = {};
    for (const [name, item] of Object.entries(columns)) if (!names.has(name) && item.letter) byLetter[item.letter] = name;
    const used = new Set();
    let hit = 0;
    for (const col of st.an.columns) {
      const name = columns[col.name] ? col.name : byLetter[col.letter];
      if (!name) continue;
      const item = columns[name];
      const cfg = { ...st.cols[col.c] };
      for (const k of COL_KEYS) if (item[k] !== undefined) cfg[k] = item[k];
      if (!A.METHODS[cfg.type] || !A.METHODS[cfg.type].some(([k]) => k === cfg.method)) continue;
      st.cols[col.c] = cfg;
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
      sheet: S.kind === 'xlsx' ? S.sheets[S.sheetIndex].name : null,
      headerRow: S.an.headerRow,
      options: {
        keepTotals: $('keepTotals').checked,
        keepForm: $('keepForm').checked,
        clearMeta: $('clearMeta').checked,
        clearNotes: $('clearNotes').checked,
        clearHF: $('clearHF').checked,
        dropUnused: $('dropUnused').checked,
      },
      columns: settingsColumns(),   // текущий лист — для совместимости с прежними версиями
      // все листы книги: методы и строка заголовков каждого
      sheets: Object.fromEntries(S.sheets.map(st => [st.name, { headerRow: st.an.headerRow, columns: settingsColumns(st) }])),
    };
    if ($('profileWithKey').checked) p.secret = { seed: $('seed').value, factor: +$('factor').value, shiftDays: +$('shiftDays').value };
    return p;
  }
  function applyOptions(p) {
    const o = p.options || p; // формат ранней версии хранил флаги в корне
    for (const id of ['keepTotals', 'keepForm', 'clearMeta', 'clearNotes', 'clearHF', 'dropUnused']) if (typeof o[id] === 'boolean') $(id).checked = o[id];
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
    // настройки по листам (с версии 3.6) или одного листа (прежние версии)
    const entries = p.sheets ? Object.entries(p.sheets) : [[p.sheet, { headerRow: p.headerRow, columns: p.columns || {} }]];
    let hit = 0, total = 0;
    const missing = [], noSheet = [];
    for (const [sheetName, cfg] of entries) {
      let i = S.sheets.findIndex(st => st.name === sheetName);
      if (i < 0) { if (p.sheets && S.sheets.length > 1) { noSheet.push(sheetName); continue; } i = S.sheetIndex; }
      let st = S.sheets[i];
      if (cfg.headerRow && cfg.headerRow !== st.an.headerRow && cfg.headerRow <= st.grid.maxRow) st = S.sheets[i] = makeState(st.name, st.grid, cfg.headerRow);
      const r = applySettingsColumns(cfg.columns || {}, st);
      hit += r.hit; total += st.an.columns.length;
      missing.push(...r.missing.map(n => entries.length > 1 ? `${st.name} · ${n}` : n));
      if (!p.sheets && p.sheet && i !== S.sheetIndex) S.sheetIndex = i;
    }
    showSheet(S.sheetIndex);
    rememberLayout();
    $('settingsInfo').innerHTML = `Настройки <b>${esc(name)}</b> применены: столбцов ${hit} из ${total}` +
      (missing.length ? `. <span class="text-danger">Нет в файле: ${missing.map(esc).join(', ')}</span>` : '') +
      (noSheet.length ? `. <span class="text-danger">Нет листов: ${noSheet.map(esc).join(', ')}</span>` : '');
    toast(`Настройки применены: ${hit} из ${total} ${A.plural(total, 'столбца', 'столбцов', 'столбцов')}`);
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
    toast(p.secret ? 'Настройки сохранены вместе с секретной фразой — не передавайте этот файл вместе с данными' : 'Настройки сохранены');
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
    S.result = A.transformBook(bookForEngine(), config());
    const ms = Math.round(performance.now() - t0);
    renderSheetOptions();
    const active = S.an.columns.filter(col => S.cols[col.c].method !== 'keep');
    const { mapping, warnings, stats } = S.result;
    const changes = S.result.bySheet[S.sheets[S.sheetIndex].name].changes;
    const where = stats.sheets > 1
      ? `в ${stats.columns} ${A.plural(stats.columns, 'столбце', 'столбцах', 'столбцах')} на ${stats.sheets} ${A.plural(stats.sheets, 'листе', 'листах', 'листах')}`
      : `в ${stats.columns} ${A.plural(stats.columns, 'столбце', 'столбцах', 'столбцах')}`;
    $('previewStats').textContent = stats.columns
      ? `Будет обезличено ячеек: ${stats.changed.toLocaleString('ru-RU')} ${where} · уникальных замен: ${mapping.length.toLocaleString('ru-RU')} · ${ms} мс` +
        (stats.sheets > 1 ? ` · ниже — лист «${S.sheets[S.sheetIndex].name}»` : '')
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
          ? `<span class="rv-empty" style="width:${Math.min(40, Math.max(6, fmt(old).length))}ch" role="img" aria-label="Пустая ячейка: значение удалено (${esc(col.name)})"></span>`
          : `<span class="rv rv-token">${esc(fmt(nv))}</span>`;
        return `<td title="${esc(fmt(old))} → ${esc(fmt(nv) || 'пустая ячейка')}"><div class="pair"><span class="rv rv-detected">${esc(fmt(old))}</span><span class="line"><span class="rv-arrow" aria-hidden="true">→</span>${after}</span></div></td>`;
      }).join('')}</tr>`).join('');
      table.innerHTML = head + `<tbody>${body}</tbody>`;
    }
    const w = [...S.loadNotes, ...warnings];
    // что не вернётся по ключу
    const lost = usedSheets().flatMap(st => st.an.columns.filter(col => A.IRREVERSIBLE.has(st.cols[col.c].method)).map(col => {
      const name = A.METHODS[st.cols[col.c].type].find(([k]) => k === st.cols[col.c].method)[1];
      return `«${S.sheets.length > 1 ? st.name + ' · ' : ''}${col.name}» — ${name.toLowerCase()}`;
    }));
    if (lost.length) w.push(`По ключу не вернутся (${lost.length > 1 ? 'эти методы необратимы' : 'метод необратим'}): ${lost.join('; ')}.`);
    const leaks = [...S.result.leaks];
    if (S.kind === 'xlsx') {
      for (const st of resultSheets()) {
        const hf = A.headerFooterText(S.wb.worksheets[S.sheets.indexOf(st)]);
        if (!hf.length) continue;
        const where = S.sheets.length > 1 ? ` листа «${st.name}»` : '';
        if ($('clearHF').checked) w.push(`В колонтитулах${where} есть текст («${hf.join('», «').slice(0, 150)}»). Он будет удалён.`);
        else leaks.push(`В колонтитулах${where} есть текст («${hf.join('», «').slice(0, 150)}»), и он останется в файле. Включите «Очистить колонтитулы» на шаге 3.`);
      }
      const unused = S.sheets.filter(st => !activeCount(st)).map(st => `«${st.name}»`);
      if (unused.length && S.sheets.length > 1) {
        const list = unused.join(', '), many = unused.length > 1;
        if ($('dropUnused').checked) w.push(many
          ? `Листы ${list} не обезличиваются и в результат не попадут. Чтобы обезличить, выберите их в списке «Лист» и задайте методы.`
          : `Лист ${list} не обезличивается и в результат не попадёт. Чтобы обезличить, выберите его в списке «Лист» и задайте методы.`);
        else leaks.push(many
          ? `Листы ${list} не обезличиваются и останутся в файле как есть. Задайте методы или включите «Удалять из результата листы, которые не обезличиваются» на шаге 3.`
          : `Лист ${list} не обезличивается и останется в файле как есть. Задайте методы или включите «Удалять из результата листы, которые не обезличиваются» на шаге 3.`);
      }
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

  // Запись результата: result — из transformBook; opt.keep — листы, которые попадут в файл;
  // opt.clean — что очистить (свойства, примечания, колонтитулы); suffix — окончание имени файла
  async function writeResult(result, { keep, clean, suffix }) {
    if (S.kind === 'csv') {
      const text = A.gridToCsv(S.grid, result.bySheet[S.sheets[0].name].changes, S.csv.delim, S.csv.decSep);
      download(new Blob([text], { type: 'text/csv;charset=utf-8' }), baseName() + suffix + '.csv');
      return;
    }
    // Пишем в свежую копию исходной книги — так сохраняется всё оформление
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(S.buffer);
    // листы с методами обезличиваются; остальные удаляются или остаются как есть (флажок на шаге 3)
    const names = new Set(keep.map(st => st.name));
    const drop = wb.worksheets.filter(ws => !names.has(ws.name));
    if (drop.length) {
      drop.forEach(x => wb.removeWorksheet(x.id));
      wb.definedNames.model = []; // имена могли ссылаться на удалённые листы
    }
    for (const ws of wb.worksheets) {
      A.applyToWorksheet(ws, result.bySheet[ws.name] ? result.bySheet[ws.name].changes : new Map());
      if (clean.notes) A.clearNotes(ws);
      if (clean.hf) A.clearHeaderFooter(ws);
    }
    if (clean.meta) A.clearWorkbookMeta(wb);
    wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
    // ExcelJS ссылается на тему оформления, даже если в исходной книге её не было, но сам файл темы не пишет —
    // Excel тогда предлагает «восстановить» книгу. Без своей темы записывается стандартная.
    if (!wb._themes || !Object.keys(wb._themes).length) wb._themes = undefined;
    const out = await wb.xlsx.writeBuffer();
    download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + suffix + '.xlsx');
  }

  $('btnDownload').addEventListener('click', async () => {
    if (!S.an) return;
    const btn = $('btnDownload');
    btn.disabled = true;
    try {
      runPreview();
      if (!S.result.stats.columns) { toast('Не удалось обезличить: не выбран метод ни для одного столбца', 6000, 'error'); return; }
      await writeResult(S.result, {
        keep: resultSheets(), suffix: '_redact',
        clean: { meta: $('clearMeta').checked, notes: $('clearNotes').checked, hf: $('clearHF').checked },
      });
      toast(`Обезличено ячеек: ${S.result.stats.changed.toLocaleString('ru-RU')}. Файл сохранён в «Загрузки»`);
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить файл: ' + e.message, 7000);
    } finally { btn.disabled = false; }
  });

  // ---------- Для внешней передачи ----------
  // Тот же файл, но без возможности вернуть данные: одноразовые секретная фраза, коэффициент и сдвиг дат
  // нигде не сохраняются и ключ не создаётся. Денежные столбцы и даты, оставленные как есть, тоже умножаются и сдвигаются;
  // свойства файла, примечания, колонтитулы и листы без методов удаляются независимо от флажков шага 3.
  $('btnExternal').addEventListener('click', async () => {
    if (!S.an) return;
    runPreview();
    if (!S.result.stats.columns) { toast('Не удалось обезличить: не выбран метод ни для одного столбца', 6000, 'error'); return; }
    const sheets = usedSheets();
    const plan = sheets.map(st => ({ st, ...A.externalColumns(st.an, st.cols) }));
    const named = list => plan.flatMap(p => p[list].map(n => sheets.length > 1 ? `${p.st.name} · ${n}` : n));
    const listOf = names => `${names.slice(0, 6).map(n => `«${n}»`).join(', ')}${names.length > 6 ? ` и ещё ${names.length - 6}` : ''}`;
    const scaled = named('scaled'), shifted = named('shifted');
    const leaks = S.result.leaks.length;
    const items = [
      'Секретная фраза, коэффициент для сумм и сдвиг дат — одноразовые: они нигде не показываются и не сохраняются. Псевдонимы будут другими, чем в файлах с вашей фразой.',
      scaled.length ? `Суммы, оставленные на шаге 2 как есть, тоже будут умножены: ${listOf(scaled)}.` : null,
      shifted.length ? `Даты, оставленные на шаге 2 как есть, тоже будут сдвинуты: ${listOf(shifted)}.` : null,
      'Свойства файла, примечания и колонтитулы очищаются, листы без методов удаляются.',
      leaks ? `В предпросмотре есть предупреждения об утечке (${leaks}) — проверьте их до выгрузки.` : null,
    ].filter(Boolean);
    const ok = await window.RedactUI.confirm({
      title: 'Скачать без ключа для внешней передачи?',
      text: 'Ключа соответствий не будет. Восстановить исходные данные по этому файлу будет нельзя, в том числе вам.',
      items, confirm: 'Да, скачать без ключа', iconName: 'shield-check',
    });
    if (!ok) return;
    const btn = $('btnExternal');
    btn.disabled = true;
    try {
      const once = A.randomSeed();
      const cfg = { ...config(), seed: once, factor: A.factorFromSeed(A.randomSeed()), shiftDays: A.shiftFromSeed(A.randomSeed()) };
      const result = A.transformBook(plan.map(p => ({ name: p.st.name, grid: p.st.grid, an: p.st.an, columns: p.columns })), cfg);
      await writeResult(result, { keep: sheets, suffix: '_external', clean: { meta: true, notes: true, hf: true } });
      toast(`Файл для внешней передачи сохранён: обезличено ячеек ${result.stats.changed.toLocaleString('ru-RU')}. Ключа нет — данные по нему не восстановить`, 6000);
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить файл: ' + e.message, 7000);
    } finally { btn.disabled = false; }
  });

  // ---------- Ключ соответствий ----------
  // Без пароля — таблица Excel (листы для чтения + скрытый лист с ключом для восстановления).
  // С паролем — зашифрованный файл .redactkey (AES-256-GCM, ключ из пароля через PBKDF2).
  $('btnKey').addEventListener('click', async () => {
    if (!S.an) return;
    const pass = $('keyPass').value, pass2 = $('keyPass2').value;
    if (pass && pass.length < 8) { toast('Не удалось сохранить ключ: пароль должен быть не короче 8 символов', 6000, 'error'); return; }
    if (pass !== pass2) { toast('Не удалось сохранить ключ: пароль и повтор не совпадают', 6000, 'error'); return; }
    const btn = $('btnKey');
    btn.disabled = true;
    try {
      runPreview();
      const cfg = config();
      const key = A.buildKeyBook({
        sheets: bookForEngine(S.kind === 'xlsx' ? resultSheets() : S.sheets), config: cfg, result: S.result,
        sourceFile: S.fileName, appVersion: A.VERSION,
      });
      if (S.kind !== 'xlsx') key.sheet = null;
      key.seed = cfg.seed;
      if (pass) {
        const env = await A.encryptKey(key, pass);
        download(new Blob([JSON.stringify(env)], { type: 'application/json' }), baseName() + '_KEY.redactkey');
        toast('Ключ сохранён и зашифрован паролем. Без пароля его не открыть — запомните или сохраните пароль отдельно', 6000);
      } else {
        const out = await keyWorkbook(key, cfg).xlsx.writeBuffer();
        download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '_KEY.xlsx');
        toast('Ключ соответствий сохранён без пароля. Храните его отдельно от обезличенного файла', 5000);
      }
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить ключ: ' + e.message, 7000, 'error');
    } finally { btn.disabled = false; }
  });

  function keyWorkbook(key, cfg) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Соответствия');
    ws.columns = [
      { header: 'Группа замен', key: 'dict', width: 14 },
      { header: 'Столбцы', key: 'columns', width: 28 },
      { header: 'Метод', key: 'method', width: 10 },
      { header: 'Было', key: 'orig', width: 45 },
      { header: 'Стало', key: 'repl', width: 45 },
    ];
    ws.getRow(1).font = { bold: true };
    ws.addRows(S.result.mapping);
    // ФИО с инициалами, найденные внутри текста («Смирнов А. В.»)
    ws.addRows((key.smartNames || []).map(([orig, repl]) => ({ dict: 'ФИО в тексте', columns: '', method: 'smart', orig, repl })));
    ws.autoFilter = 'A1:E1';
    const ps = wb.addWorksheet('Параметры');
    ps.columns = [{ header: 'Параметр', key: 'k', width: 34 }, { header: 'Значение', key: 'v', width: 40 }];
    ps.getRow(1).font = { bold: true };
    ps.addRows([
      { k: 'Исходный файл', v: S.fileName },
      { k: 'Дата обработки', v: new Date().toLocaleString('ru-RU') },
      { k: 'Секретная фраза', v: cfg.seed },
      { k: 'Коэффициент для сумм (исходное = обезличенное ÷ коэффициент)', v: cfg.factor },
      { k: 'Сдвиг дат, дней (исходная дата = обезличенная − сдвиг)', v: cfg.shiftDays },
      ...S.sheets.flatMap(st => st.an.columns.filter(col => isActive(st.cols[col.c])).map(col => ({
        k: `${S.sheets.length > 1 ? st.name + ' · ' : ''}${col.letter} · ${col.name}`,
        v: A.METHODS[st.cols[col.c].type].find(([k]) => k === st.cols[col.c].method)[1],
      }))),
    ]);
    // ключ для восстановления: JSON частями (в ячейке Excel не больше 32 767 символов), лист скрыт
    const ks = wb.addWorksheet(KEY_SHEET, { state: 'veryHidden' });
    const json = JSON.stringify(key);
    ks.getCell(1, 1).value = 'redact-key';
    for (let i = 0, row = 2; i < json.length; i += 30000, row++) ks.getCell(row, 1).value = json.slice(i, i + 30000);
    return wb;
  }

  // общие помощники для режима восстановления (restore.js)
  window.RedactUI = { $, icon, esc, fmt, toast, download, store, KEY_SHEET };
})();

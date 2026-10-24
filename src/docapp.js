/* Вкладка «Документ Word»: .docx → что найдено → обезличенный .docx и ключ соответствий. Логика — в src/docx.js. */
(function () {
  'use strict';
  const A = window.Anon;
  const D = window.RedactDocx;
  const { $, icon, esc, toast, download, KEY_SHEET } = window.RedactUI;
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const DS = {
    fileName: '', buffer: null, analysis: null, rows: [], repl: new Map(),
    seed: '', types: {}, skip: new Set(),
  };

  // ---------- Загрузка ----------
  const drop = $('dDrop');
  ['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', ev => { const f = ev.dataTransfer.files[0]; if (f) loadDoc(f); });
  $('dFile').addEventListener('change', ev => { const f = ev.target.files[0]; ev.target.value = ''; if (f) loadDoc(f); });

  async function loadDoc(file) {
    $('dFileInfo').textContent = 'Читаем документ…';
    try {
      if (/\.doc$/i.test(file.name)) throw new Error('формат .doc (Word 97–2003) не поддерживается. Откройте документ в Word и сохраните как .docx');
      const buf = await file.arrayBuffer();
      const head = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
      if (head[0] !== 0x50 || head[1] !== 0x4B) throw new Error('это не документ Word .docx');
      const zip = await JSZip.loadAsync(buf);
      DS.fileName = file.name;
      DS.buffer = buf;
      DS.seed = A.randomSeed();
      DS.types = {};
      DS.skip = new Set();
      $('dFactor').value = A.factorFromSeed(A.randomSeed());
      $('dShift').value = A.shiftFromSeed(A.randomSeed());
      DS.analysis = await D.analyzeDocx(zip, opts());
      $('dFileInfo').innerHTML = `${icon('file-text', 'i i20')}<b>${esc(file.name)}</b><span class="muted">${(file.size / 1024).toFixed(0)} КБ · абзацев с текстом: ${DS.analysis.info.paragraphs.toLocaleString('ru-RU')}</span>`;
      $('dClearRow').classList.remove('hidden');
      ['d2', 'd3'].forEach(id => $(id).classList.remove('hidden'));
      replan();
    } catch (e) {
      console.error(e);
      $('dFileInfo').textContent = '';
      toast('Не удалось прочитать документ: ' + e.message, 7000, 'error');
    }
  }

  const custom = () => $('dCustom').value.split('\n').map(s => s.trim()).filter(s => s.length >= 2);
  const opts = () => ({
    seed: DS.seed, factor: parseFloat($('dFactor').value) || 1, shiftDays: parseInt($('dShift').value, 10) || 0,
    types: DS.types, skip: [...DS.skip], custom: custom(),
  });

  // ---------- Найденное ----------
  function replan() {
    if (!DS.analysis) return;
    const { rows, repl } = D.plan(DS.analysis.parts, opts());
    DS.rows = rows;
    DS.repl = repl;
    renderTypes();
    renderFound();
    renderPreview();
  }

  function renderTypes() {
    const counts = {};
    for (const r of DS.rows) counts[r.type] = (counts[r.type] || 0) + 1;
    $('dTypes').innerHTML = Object.entries(D.TYPES).filter(([t]) => counts[t]).map(([t, name]) =>
      `<label class="chk"><input type="checkbox" data-type="${t}"${DS.types[t] === false ? '' : ' checked'}> ${esc(name)} — ${counts[t]}</label>`).join('');
    const total = DS.rows.length;
    $('dSummary').textContent = total
      ? `Найдено значений: ${total.toLocaleString('ru-RU')}. Снимите флажок у типа или у отдельного значения, чтобы оставить его как есть.`
      : 'Автоматически ничего не найдено. Добавьте то, что нужно скрыть, в «Добавить своё» ниже.';
  }

  function renderFound() {
    $('dFound').tBodies[0].innerHTML = DS.rows.map((r, i) => {
      const typeOn = DS.types[r.type] !== false;
      return `<tr data-i="${i}" class="${r.on ? 'active' : ''}">
        <td class="c-chk"><input type="checkbox" data-v="${i}"${r.on ? ' checked' : ''}${typeOn ? '' : ' disabled'} aria-label="Заменять «${esc(r.value)}»"></td>
        <td class="small muted">${esc(D.TYPES[r.type])}</td>
        <td><span class="rv ${r.on ? 'rv-detected' : 'rv-original'}">${esc(r.value)}</span></td>
        <td class="r small">${r.count}</td>
        <td>${r.on ? `<span class="rv rv-token">${esc(r.repl)}</span>` : '<span class="small muted">останется как есть</span>'}</td>
      </tr>`;
    }).join('');
  }

  $('dTypes').addEventListener('change', ev => {
    const t = ev.target.dataset.type;
    if (!t) return;
    DS.types[t] = ev.target.checked;
    replan();
  });
  $('dFound').addEventListener('change', ev => {
    const i = ev.target.dataset.v;
    if (i === undefined) return;
    const v = DS.rows[+i].value;
    if (ev.target.checked) DS.skip.delete(v); else DS.skip.add(v);
    replan();
  });
  let customTimer;
  $('dCustom').addEventListener('input', () => { clearTimeout(customTimer); customTimer = setTimeout(replan, 400); });
  ['dFactor', 'dShift'].forEach(id => $(id).addEventListener('change', replan));
  $('btnDFactor').addEventListener('click', () => { $('dFactor').value = A.factorFromSeed(A.randomSeed()); replan(); });
  $('btnDShift').addEventListener('click', () => { $('dShift').value = A.shiftFromSeed(A.randomSeed()); replan(); });

  // ---------- Предпросмотр ----------
  function renderPreview() {
    const paras = D.previewParagraphs(DS.analysis.parts, DS.repl, 40);
    const n = DS.repl.size;
    const occ = DS.rows.filter(r => r.on).reduce((s, r) => s + r.count, 0);
    $('dStats').textContent = n
      ? `Будет заменено: ${n.toLocaleString('ru-RU')} ${A.plural(n, 'значение', 'значения', 'значений')}, ${occ.toLocaleString('ru-RU')} ${A.plural(occ, 'вхождение', 'вхождения', 'вхождений')}` +
        (paras.length >= 40 ? ' · ниже — первые 40 абзацев' : '')
      : 'Замен нет: отметьте найденное или добавьте своё на шаге 2.';
    const partName = p => /header/.test(p) ? 'Колонтитул' : /footer/.test(p) ? 'Колонтитул' : /footnotes/.test(p) ? 'Сноска'
      : /endnotes/.test(p) ? 'Концевая сноска' : /comments/.test(p) ? 'Примечание' : '';
    $('dPreview').innerHTML = paras.map(p => `<p>${partName(p.part) ? `<span class="caption muted doc-part">${partName(p.part)}</span>` : ''}${p.pieces.map(x => x.text !== undefined
      ? esc(x.text)
      : `<span class="pair-inline" title="${esc(x.orig)} → ${esc(x.repl)}"><span class="rv rv-detected">${esc(x.orig)}</span><span class="rv-arrow" aria-hidden="true">→</span><span class="rv rv-token">${esc(x.repl)}</span></span>`).join('')}</p>`).join('');

    const info = [], danger = [];
    const i = DS.analysis.info;
    if (i.revisions) info.push($('dRevisions').checked
      ? `В документе есть записанные исправления (${i.revisions}): удалённый текст будет убран из файла, вставленный станет обычным.`
      : `В документе есть записанные исправления (${i.revisions}), и удалённый текст останется внутри файла. Включите «Принять исправления» на шаге 2.`);
    if (i.comments) (($('dComments').checked) ? info : danger).push($('dComments').checked
      ? `Примечания (${i.comments}) будут удалены вместе с именами авторов.`
      : `Примечания (${i.comments}) останутся; найденное в них заменится, а имена авторов — уберутся.`);
    if (i.images) danger.push(`В документе ${i.images} ${A.plural(i.images, 'картинка', 'картинки', 'картинок')} — например, сканы подписей и печатей. Redact их не читает: картинки останутся в файле как есть.`);
    if (i.embeddings) danger.push(`В документе ${i.embeddings} ${A.plural(i.embeddings, 'вложенный объект', 'вложенных объекта', 'вложенных объектов')} (например, таблица Excel). Redact их не читает: они останутся в файле как есть — проверьте их отдельно.`);
    const kept = DS.rows.filter(r => !r.on);
    if (kept.length) danger.push(`Останутся без замены: ${kept.length} ${A.plural(kept.length, 'значение', 'значения', 'значений')} (${kept.slice(0, 5).map(r => `«${r.value}»`).join(', ')}${kept.length > 5 ? '…' : ''}).`);
    info.push('Проверьте предпросмотр: Redact находит типовые данные по шаблонам, но не понимает смысл текста. Адреса, названия без формы организации и имена без отчества добавьте в «Добавить своё».');
    const notice = (cls, ic, title, items) => items.length
      ? `<div class="notice ${cls}"><div class="notice-title">${icon(ic, 'i i20')}${title}</div><ul>${items.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
    $('dWarnings').innerHTML = notice('danger', 'triangle-alert', `Возможна утечка данных: ${danger.length}`, danger) + notice('', 'circle-alert', 'Проверьте перед выгрузкой', info);
  }
  ['dRevisions', 'dComments'].forEach(id => $(id).addEventListener('change', () => DS.analysis && renderPreview()));

  // ---------- Выгрузка ----------
  const baseName = () => DS.fileName.replace(/\.[^.]+$/, '');
  $('btnDDownload').addEventListener('click', async () => {
    if (!DS.analysis) return;
    const btn = $('btnDDownload');
    btn.disabled = true;
    try {
      const zip = await JSZip.loadAsync(DS.buffer);
      const { replaced } = await D.redactDocx(zip, {
        repl: DS.repl, clean: { revisions: $('dRevisions').checked, comments: $('dComments').checked, meta: $('dMeta').checked },
      });
      const blob = await zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
      download(blob, baseName() + '_redact.docx');
      toast(`Заменено вхождений: ${replaced.toLocaleString('ru-RU')}. Документ сохранён в «Загрузки»`);
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить документ: ' + e.message, 7000, 'error');
    } finally { btn.disabled = false; }
  });

  // Ключ документа — тот же формат, что у таблиц: пары «исходное → замена»; открывается на вкладке восстановления
  function buildKey() {
    return {
      app: 'Redact', kind: 'redact-key', format: 1, appVersion: A.VERSION, createdAt: new Date().toISOString(),
      sourceFile: DS.fileName, document: true, sheet: null, sheets: [], headerRow: 1,
      factor: parseFloat($('dFactor').value) || 1, shiftDays: parseInt($('dShift').value, 10) || 0,
      columns: [],
      dictionaries: [{ name: 'Документ', method: 'doc', prefix: '', columns: [], pairs: [...DS.repl] }],
      smartDigits: [], smartNames: [], numbers: {},
    };
  }
  $('btnDKey').addEventListener('click', async () => {
    if (!DS.analysis) return;
    const pass = $('dKeyPass').value, pass2 = $('dKeyPass2').value;
    if (pass && pass.length < 8) { toast('Не удалось сохранить ключ: пароль должен быть не короче 8 символов', 6000, 'error'); return; }
    if (pass !== pass2) { toast('Не удалось сохранить ключ: пароль и повтор не совпадают', 6000, 'error'); return; }
    const btn = $('btnDKey');
    btn.disabled = true;
    try {
      const key = buildKey();
      if (pass) {
        const env = await A.encryptKey(key, pass);
        download(new Blob([JSON.stringify(env)], { type: 'application/json' }), baseName() + '_KEY.redactkey');
        toast('Ключ сохранён и зашифрован паролем. Без пароля его не открыть — запомните или сохраните пароль отдельно', 6000);
      } else {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Соответствия');
        ws.columns = [{ header: 'Тип', key: 't', width: 18 }, { header: 'Было', key: 'o', width: 50 }, { header: 'Стало', key: 'r', width: 50 }];
        ws.getRow(1).font = { bold: true };
        const typeOf = new Map(DS.rows.map(r => [r.value, D.TYPES[r.type]]));
        ws.addRows([...DS.repl].map(([o, r]) => ({ t: typeOf.get(o) || '', o, r })));
        const ks = wb.addWorksheet(KEY_SHEET, { state: 'veryHidden' });
        const json = JSON.stringify(key);
        ks.getCell(1, 1).value = 'redact-key';
        for (let i = 0, row = 2; i < json.length; i += 30000, row++) ks.getCell(row, 1).value = json.slice(i, i + 30000);
        const out = await wb.xlsx.writeBuffer();
        download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '_KEY.xlsx');
        toast('Ключ соответствий сохранён без пароля. Храните его отдельно от документа', 5000);
      }
    } catch (e) {
      console.error(e);
      toast('Не удалось сохранить ключ: ' + e.message, 7000, 'error');
    } finally { btn.disabled = false; }
  });

  // ---------- Очистить ----------
  $('btnDClear').addEventListener('click', async () => {
    const ok = await window.RedactUI.confirm({
      title: 'Очистить и начать с другим документом?',
      text: 'Будут убраны загруженный документ, найденное и предпросмотр. Сам документ на компьютере не затрагивается.',
      confirm: 'Да, очистить',
    });
    if (!ok) return;
    Object.assign(DS, { fileName: '', buffer: null, analysis: null, rows: [], repl: new Map(), seed: '', types: {}, skip: new Set() });
    ['d2', 'd3', 'dClearRow'].forEach(id => $(id).classList.add('hidden'));
    for (const id of ['dFileInfo', 'dSummary', 'dTypes', 'dPreview', 'dWarnings', 'dStats']) $(id).innerHTML = '';
    $('dFound').tBodies[0].innerHTML = '';
    $('dCustom').value = '';
    $('dKeyPass').value = $('dKeyPass2').value = '';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    toast('Готово к новому документу');
  });
})();

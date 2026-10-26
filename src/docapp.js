/* Вкладка «Документ Word»: .docx → что найдено → обезличенный .docx и ключ соответствий. Логика — в src/docx.js. */
(function () {
  'use strict';
  const A = window.Anon;
  const D = window.RedactDocx;
  const { $, icon, esc, toast, download, store, KEY_SHEET } = window.RedactUI;
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const DS = {
    fileName: '', buffer: null, analysis: null, rows: [], repl: new Map(), keyPairs: [], cutManual: new Map(), cut: new Set(), ruleCut: new Set(),
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
      DS.cutManual = new Map();
      DS.cut = new Set();
      DS.analysis = null;
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
    customCase: $('dCustomCase').checked, customForms: $('dCustomForms').checked,
    cut: DS.analysis ? D.cutRanges(DS.analysis.blocks, DS.cut) : [],
  });

  // ---------- Найденное ----------
  // ---------- Удаление разделов и абзацев ----------
  // Итог: по правилам плюс отмеченное вручную, минус возвращённое вручную (cutManual: номер блока → удалить или оставить)
  const cutRules = () => $('dCutRules').value.split('\n').map(s => s.trim()).filter(s => s.length >= 2);
  function computeCut() {
    const blocks = DS.analysis.blocks;
    DS.ruleCut = D.ruleCuts(blocks, cutRules());
    DS.cut = new Set([...DS.ruleCut].filter(i => DS.cutManual.get(i) !== false));
    DS.cutManual.forEach((v, i) => { if (v) DS.cut.add(i); });
  }
  function renderOutline() {
    const B = DS.analysis.blocks;
    let ctx = 0;
    $('dOutline').innerHTML = B.map((b, i) => {
      if (b.level) ctx = b.level;
      if (!b.text && !b.image && b.kind !== 'tbl') return '';
      const ind = b.level ? (b.level - 1) * 16 : ctx * 16 + 8;
      const label = b.kind === 'tbl' ? `Таблица: ${b.text || 'пустая'}` : b.text || 'Картинка';
      const cut = DS.cut.has(i);
      const tag = DS.ruleCut.has(i) && cut ? 'по правилу' : b.guessed ? '≈ заголовок' : '';
      const what = b.level ? 'раздел' : b.kind === 'tbl' ? 'таблицу' : 'абзац';
      return `<div class="ol-row${b.level ? ' h' : ''}${b.kind === 'tbl' ? ' tbl' : ''}${!b.text && b.image ? ' img' : ''}${cut ? ' cut' : ''}" role="listitem" style="--ind:${ind}px">
        <button type="button" class="btn ghost icon-only" data-cut="${i}" aria-pressed="${cut}" title="${cut ? 'Вернуть' : 'Удалить'} ${what}" aria-label="${cut ? 'Вернуть' : 'Удалить'} ${what}: ${esc(label.slice(0, 60))}">${icon(cut ? 'undo-2' : 'trash-2', 'i')}</button>
        <span class="ol-text">${esc(label)}</span>${tag ? `<span class="ol-tag">${tag}</span>` : ''}</div>`;
    }).join('');
    const n = DS.cut.size;
    const secs = [...DS.cut].filter(i => B[i].level).length;
    $('dCutStats').textContent = n
      ? `Будет удалено: ${secs ? `${secs} ${A.plural(secs, 'заголовок', 'заголовка', 'заголовков')} с разделами, ` : ''}всего ${n} ${A.plural(n, 'блок', 'блока', 'блоков')} (абзацы, таблицы, картинки).`
      : 'Ничего не удаляется.';
  }
  $('dOutline').addEventListener('click', ev => {
    const btn = ev.target.closest('[data-cut]');
    if (!btn) return;
    const B = DS.analysis.blocks, i = +btn.dataset.cut;
    const end = D.sectionEnd(B, i);
    const range = [...Array(end - i)].map((_, k) => i + k);
    const allCut = range.every(j => DS.cut.has(j));
    range.forEach(j => DS.cutManual.set(j, !allCut));
    replan();
    const again = $('dOutline').querySelector(`[data-cut="${i}"]`);
    if (again) again.focus();
  });
  let rulesTimer;
  $('dCutRules').addEventListener('input', () => {
    clearTimeout(rulesTimer);
    rulesTimer = setTimeout(() => { store.set('redact-doc-cut-rules', $('dCutRules').value); replan(); }, 400);
  });
  $('dCutRules').value = store.get('redact-doc-cut-rules') || '';
  document.querySelectorAll('input[name=dCutMode]').forEach(r => r.addEventListener('change', () => DS.analysis && renderPreview()));
  const cutMode = () => (document.querySelector('input[name=dCutMode]:checked') || {}).value || 'full';

  function replan() {
    if (!DS.analysis) return;
    computeCut();
    const { rows, repl, keyPairs } = D.plan(DS.analysis.parts, opts());
    DS.rows = rows;
    DS.repl = repl;
    DS.keyPairs = keyPairs;
    renderTypes();
    renderFound();
    renderOutline();
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
        <td><span class="rv ${r.on ? 'rv-detected' : 'rv-original'}">${esc(r.value)}</span>${r.forms.length
          ? `<span class="small muted found-forms">также: ${r.forms.map(esc).join(', ')}</span>` : ''}</td>
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
  ['dCustomCase', 'dCustomForms'].forEach(id => $(id).addEventListener('change', replan));
  ['dFactor', 'dShift'].forEach(id => $(id).addEventListener('change', replan));
  $('btnDFactor').addEventListener('click', () => { $('dFactor').value = A.factorFromSeed(A.randomSeed()); replan(); });
  $('btnDShift').addEventListener('click', () => { $('dShift').value = A.shiftFromSeed(A.randomSeed()); replan(); });

  // ---------- Предпросмотр ----------
  function renderPreview() {
    const only = $('dOnlyChanged').checked;
    const paras = D.previewDocument(DS.analysis.parts, DS.repl, { blocks: DS.analysis.blocks, cut: DS.cut, mode: cutMode(), onlyChanged: only });
    const n = DS.repl.size;
    const occ = DS.rows.filter(r => r.on).reduce((s, r) => s + r.count, 0);
    $('dStats').textContent = n
      ? `Будет заменено: ${n.toLocaleString('ru-RU')} ${A.plural(n, 'значение', 'значения', 'значений')}, ${occ.toLocaleString('ru-RU')} ${A.plural(occ, 'вхождение', 'вхождения', 'вхождений')}`
      : 'Замен нет: отметьте найденное или добавьте своё на шаге 2.';
    const partName = p => /header/.test(p) ? 'Колонтитул' : /footer/.test(p) ? 'Колонтитул' : /footnotes/.test(p) ? 'Сноска'
      : /endnotes/.test(p) ? 'Концевая сноска' : /comments/.test(p) ? 'Примечание' : '';
    $('dPreview').classList.toggle('full', !only);
    // удалённый участок: в режиме «С пометкой» — пометка, как в файле; название удалённого — только здесь, зачёркнутым
    const cutLine = p => `<p class="cut-line">${p.mark ? `<span class="mark">${esc(p.mark)}</span>` : '<span class="cut-note">удалено</span>'}` +
      `<span class="cut-what">${esc(p.cutWhat.length > 90 ? p.cutWhat.slice(0, 90) + '…' : p.cutWhat)}</span>` +
      `${p.cutCount > 1 ? `<span class="cut-note">и ещё ${p.cutCount - 1} ${A.plural(p.cutCount - 1, 'блок', 'блока', 'блоков')}</span>` : ''}</p>`;
    $('dPreview').innerHTML = paras.length ? paras.map(p => p.cutWhat !== undefined ? cutLine(p) : `<p class="${p.heading ? 'h' : ''}${p.table ? ' tbl' : ''}">${partName(p.part) ? `<span class="caption muted doc-part">${partName(p.part)}</span>` : ''}${p.pieces.map(x => x.text !== undefined
      ? esc(x.text)
      : `<span class="pair-inline" title="${esc(x.orig)} → ${esc(x.repl)}"><span class="rv rv-detected">${esc(x.orig)}</span><span class="rv-arrow" aria-hidden="true">→</span><span class="rv rv-token">${esc(x.repl)}</span></span>`).join('')}</p>`).join('')
      : `<p class="small muted">${only ? 'Абзацев с заменами нет.' : 'В документе не осталось текста.'}</p>`;

    const info = [], danger = [];
    const i = DS.analysis.info;
    if (i.revisions) info.push($('dRevisions').checked
      ? `В документе есть записанные исправления (${i.revisions}): удалённый текст будет убран из файла, вставленный станет обычным.`
      : `В документе есть записанные исправления (${i.revisions}), и удалённый текст останется внутри файла. Включите «Принять исправления» на шаге 2.`);
    if (i.comments) (($('dComments').checked) ? info : danger).push($('dComments').checked
      ? `Примечания (${i.comments}) будут удалены вместе с именами авторов.`
      : `Примечания (${i.comments}) останутся; найденное в них заменится, а имена авторов — уберутся.`);
    // картинки удалённых разделов уходят из файла; предупреждаем только об оставшихся (в тексте и колонтитулах)
    const B = DS.analysis.blocks;
    const cutImg = B.filter((b, k) => b.image && DS.cut.has(k)).length;
    const leftImg = B.some((b, k) => b.image && !DS.cut.has(k)) || DS.analysis.parts.some(p => p.name !== 'word/document.xml' && /<w:drawing\b|<w:pict\b/.test(p.xml));
    if (i.images && leftImg) danger.push(`В документе есть картинки — например, сканы подписей и печатей. Redact их не читает: картинки${cutImg ? ' вне удаляемых разделов' : ''} останутся в файле как есть.`);
    if (cutImg) info.push(`Картинки из удаляемых разделов (${cutImg}) будут удалены из файла вместе с текстом.`);
    if (i.embeddings) danger.push(`В документе ${i.embeddings} ${A.plural(i.embeddings, 'вложенный объект', 'вложенных объекта', 'вложенных объектов')} (например, таблица Excel). Redact их не читает: они останутся в файле как есть — проверьте их отдельно.`);
    if (DS.cut.size) danger.push(`Будут удалены ${DS.cut.size} ${A.plural(DS.cut.size, 'блок', 'блока', 'блоков')} документа${cutMode() === 'mark' ? ' (на их месте — пометка «[… удалён]»)' : ''}: по ключу их не вернуть. Проверьте список в «Удалить разделы и абзацы».`);
    const kept = DS.rows.filter(r => !r.on);
    if (kept.length) danger.push(`Останутся без замены: ${kept.length} ${A.plural(kept.length, 'значение', 'значения', 'значений')} (${kept.slice(0, 5).map(r => `«${r.value}»`).join(', ')}${kept.length > 5 ? '…' : ''}).`);
    info.push('Проверьте предпросмотр: Redact находит типовые данные по шаблонам, но не понимает смысл текста. Адреса, названия без формы организации и имена без отчества добавьте в «Добавить своё».');
    const notice = (cls, ic, title, items) => items.length
      ? `<div class="notice ${cls}"><div class="notice-title">${icon(ic, 'i i20')}${title}</div><ul>${items.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
    $('dWarnings').innerHTML = notice('danger', 'triangle-alert', `Возможна утечка данных: ${danger.length}`, danger) + notice('', 'circle-alert', 'Проверьте перед выгрузкой', info);
  }
  ['dRevisions', 'dComments', 'dOnlyChanged'].forEach(id => $(id).addEventListener('change', () => DS.analysis && renderPreview()));

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
        cut: DS.cut, mode: cutMode(),
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
      dictionaries: [{ name: 'Документ', method: 'doc', prefix: '', columns: [], pairs: DS.keyPairs }],
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
        ws.addRows(DS.rows.filter(r => r.on).map(r => ({ t: D.TYPES[r.type], o: r.keyValue, r: r.repl })));
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
    Object.assign(DS, { fileName: '', buffer: null, analysis: null, rows: [], repl: new Map(), keyPairs: [], cutManual: new Map(), cut: new Set(), ruleCut: new Set(), seed: '', types: {}, skip: new Set() });
    ['d2', 'd3', 'dClearRow'].forEach(id => $(id).classList.add('hidden'));
    for (const id of ['dFileInfo', 'dSummary', 'dTypes', 'dPreview', 'dWarnings', 'dStats']) $(id).innerHTML = '';
    $('dFound').tBodies[0].innerHTML = '';
    $('dCustom').value = '';
    $('dKeyPass').value = $('dKeyPass2').value = '';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    toast('Готово к новому документу');
  });
})();

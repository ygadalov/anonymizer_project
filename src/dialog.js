/*
 * Диалог подтверждения необратимого действия (brand/BRAND.md, «Кнопки»: danger-solid — только здесь).
 * RedactUI.confirm({ title, text, confirm }) → Promise<boolean>.
 * Фокус по умолчанию — на «Отмене», чтобы случайный Enter ничего не стёр; Esc и щелчок мимо окна — отмена.
 */
(function () {
  'use strict';
  const { icon, esc } = window.RedactUI;

  function confirmDialog({ title, text, confirm = 'Да', cancel = 'Отмена', iconName = 'eraser' }) {
    return new Promise(resolve => {
      const back = document.activeElement;
      const dlg = document.createElement('dialog');
      dlg.className = 'confirm';
      dlg.setAttribute('aria-labelledby', 'confirmTitle');
      dlg.setAttribute('aria-describedby', 'confirmText');
      dlg.innerHTML = `
        <div class="confirm-head">${icon(iconName, 'i i20 confirm-icon')}<h2 class="h2" id="confirmTitle">${esc(title)}</h2></div>
        <p class="body muted" id="confirmText">${esc(text)}</p>
        <div class="confirm-actions">
          <button type="button" class="btn secondary" data-r="0">${esc(cancel)}</button>
          <button type="button" class="btn danger-solid" data-r="1">${esc(confirm)}</button>
        </div>`;
      document.body.appendChild(dlg);
      let done = false;
      const finish = ok => {
        if (done) return;
        done = true;
        dlg.close();
        dlg.remove();
        if (back && back.isConnected) back.focus();
        resolve(ok);
      };
      dlg.addEventListener('click', ev => {
        const b = ev.target.closest('button[data-r]');
        if (b) finish(b.dataset.r === '1');
        else if (ev.target === dlg) finish(false); // щелчок по подложке
      });
      dlg.addEventListener('cancel', ev => { ev.preventDefault(); finish(false); }); // Esc
      dlg.showModal();
      dlg.querySelector('[data-r="0"]').focus();
    });
  }

  window.RedactUI.confirm = confirmDialog;
})();

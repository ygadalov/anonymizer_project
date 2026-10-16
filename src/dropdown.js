/*
 * Выпадающие списки в стиле Redact вместо системных.
 * Системный список рисует ОС — его нельзя скруглить и оформить, поэтому он выглядит чужеродно.
 * Здесь <select> остаётся на месте (значение, события change, доступность), а вместо системного
 * списка открывается свой — для всех <select> на странице, включая создаваемые позже.
 */
(function () {
  'use strict';
  const GAP = 2; // зазор между полем и списком, px
  let pop = null, sel = null, active = -1, typed = '', typedTimer = null;

  const options = () => Array.from(sel.options).filter(o => !o.hidden);

  function open(select) {
    if (select.disabled || select.options.length < 2) return;
    close();
    sel = select;
    pop = document.createElement('div');
    pop.className = 'dd-pop';
    pop.setAttribute('role', 'listbox');
    const opts = options();
    pop.innerHTML = opts.map((o, i) =>
      `<div class="dd-opt${o.selected ? ' selected' : ''}" role="option" aria-selected="${o.selected}" data-i="${i}"${o.title ? ` title="${o.title.replace(/"/g, '&quot;')}"` : ''}>` +
      `<span class="dd-text"></span><svg class="i dd-check" aria-hidden="true"><use href="#i-check"></use></svg></div>`).join('');
    pop.querySelectorAll('.dd-text').forEach((el, i) => { el.textContent = opts[i].textContent; });
    document.body.appendChild(pop);
    place();
    active = Math.max(0, opts.findIndex(o => o.selected));
    highlight(true);
    sel.setAttribute('aria-expanded', 'true');
    sel.classList.add('dd-open');
  }

  function place() {
    const r = sel.getBoundingClientRect();
    // ровно по ширине поля: список не выходит за ячейку таблицы
    pop.style.width = r.width + 'px';
    pop.style.left = Math.max(8, Math.min(r.left, window.innerWidth - pop.offsetWidth - 8)) + 'px';
    const below = window.innerHeight - r.bottom - 8, above = r.top - 8;
    const h = Math.min(pop.scrollHeight, 320);
    // вниз, а если места мало — вверх
    if (below >= h || below >= above) { pop.style.top = (r.bottom + GAP) + 'px'; pop.style.maxHeight = Math.min(320, below - GAP) + 'px'; }
    else { pop.style.top = (r.top - GAP - Math.min(h, above)) + 'px'; pop.style.maxHeight = Math.min(320, above - GAP) + 'px'; }
  }

  function highlight(scroll) {
    pop.querySelectorAll('.dd-opt').forEach((el, i) => el.classList.toggle('active', i === active));
    const el = pop.querySelector('.dd-opt.active');
    if (el && scroll) el.scrollIntoView({ block: 'nearest' });
  }

  function choose(i) {
    const opts = options();
    if (opts[i] && !opts[i].selected) {
      sel.value = opts[i].value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const s = sel;
    close();
    if (document.body.contains(s)) s.focus();
  }

  function close() {
    if (!pop) return;
    pop.remove();
    pop = null;
    if (sel) { sel.removeAttribute('aria-expanded'); sel.classList.remove('dd-open'); }
    sel = null;
    active = -1;
  }

  // Открытие: щелчок мышью по списку — системный список подавляется
  document.addEventListener('mousedown', ev => {
    const s = ev.target.closest && ev.target.closest('select');
    if (s && !s.multiple && !s.size) {
      ev.preventDefault();
      if (sel === s) { close(); return; }
      s.focus();
      open(s);
      return;
    }
    if (pop && !pop.contains(ev.target)) close();
  }, true);

  document.addEventListener('click', ev => {
    const o = ev.target.closest && ev.target.closest('.dd-opt');
    if (o && pop && pop.contains(o)) choose(+o.dataset.i);
  });
  document.addEventListener('mousemove', ev => {
    const o = ev.target.closest && ev.target.closest('.dd-opt');
    if (o && pop && pop.contains(o) && +o.dataset.i !== active) { active = +o.dataset.i; highlight(false); }
  });

  // Клавиатура: Enter, пробел, ↓/↑ открывают; стрелки, Home/End, буквы — выбор; Enter — подтвердить; Esc — закрыть
  document.addEventListener('keydown', ev => {
    const t = ev.target;
    if (!pop) {
      if (t.tagName === 'SELECT' && (ev.key === 'Enter' || ev.key === ' ' || ev.key === 'ArrowDown' || ev.key === 'ArrowUp')) {
        ev.preventDefault();
        open(t);
      }
      return;
    }
    const n = options().length;
    if (ev.key === 'ArrowDown') active = Math.min(n - 1, active + 1);
    else if (ev.key === 'ArrowUp') active = Math.max(0, active - 1);
    else if (ev.key === 'Home') active = 0;
    else if (ev.key === 'End') active = n - 1;
    else if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); choose(active); return; }
    else if (ev.key === 'Escape') { ev.preventDefault(); const s = sel; close(); s.focus(); return; }
    else if (ev.key === 'Tab') { close(); return; }
    else if (ev.key.length === 1) {
      // поиск по первым буквам
      clearTimeout(typedTimer);
      typed += ev.key.toLowerCase();
      typedTimer = setTimeout(() => { typed = ''; }, 700);
      const i = options().findIndex(o => o.textContent.trim().toLowerCase().startsWith(typed));
      if (i >= 0) active = i;
    } else return;
    ev.preventDefault();
    highlight(true);
  }, true);

  // при прокрутке список едет вместе с полем, а когда поле уходит с экрана — закрывается
  window.addEventListener('scroll', ev => {
    if (!pop || pop.contains(ev.target)) return;
    const r = sel.getBoundingClientRect();
    if (r.bottom < 0 || r.top > window.innerHeight) close(); else place();
  }, true);
  window.addEventListener('resize', close);
})();

/*
 * Подсказки в стиле Redact вместо системных (атрибут title).
 * Системную подсказку рисует ОС: её нельзя оформить, и длинный текст уходит далеко вправо.
 * Здесь title на время наведения снимается (чтобы не всплывала системная), а текст показывается
 * в карточке рядом с элементом — с переносом строк и в пределах экрана.
 */
(function () {
  'use strict';
  const DELAY = 350, GAP = 6, MARGIN = 8;
  let tip = null, target = null, timer = null;

  function textOf(el) {
    // title мог появиться заново (его меняет код страницы) — берём свежий и снимаем
    const t = el.getAttribute('title');
    if (t) {
      el.dataset.tip = t;
      el.removeAttribute('title');
      if (!el.hasAttribute('aria-label') && !el.textContent.trim()) el.setAttribute('aria-label', t);
      else el.setAttribute('aria-description', t);
    }
    return el.dataset.tip || '';
  }

  function show(el) {
    const text = textOf(el);
    const pop = document.querySelector('.dd-pop');
    if (!text || (pop && !pop.contains(el))) return; // пока открыт список, подсказки — только у его пунктов
    if (!tip) {
      tip = document.createElement('div');
      tip.className = 'tip';
      tip.setAttribute('role', 'tooltip');
      document.body.appendChild(tip);
    }
    tip.textContent = text;
    tip.style.left = '0px';
    tip.style.top = '0px';
    tip.classList.add('on');
    const r = el.getBoundingClientRect(), w = tip.offsetWidth, h = tip.offsetHeight;
    let top, left;
    if (pop && pop.contains(el)) {
      // у пунктов списка — сбоку от списка, чтобы не закрывать соседние пункты
      const pr = pop.getBoundingClientRect();
      left = pr.right + GAP + w <= window.innerWidth - MARGIN ? pr.right + GAP : pr.left - GAP - w;
      top = Math.min(r.top, window.innerHeight - h - MARGIN);
    } else {
      // под элементом, а если внизу не помещается — над ним
      top = r.bottom + GAP;
      if (top + h > window.innerHeight - MARGIN && r.top - GAP - h > MARGIN) top = r.top - GAP - h;
      left = r.left;
    }
    left = Math.max(MARGIN, Math.min(left, window.innerWidth - w - MARGIN));
    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  }

  function hide() {
    clearTimeout(timer);
    target = null;
    if (tip) tip.classList.remove('on');
  }

  const tipped = el => el && el.closest && el.closest('[title], [data-tip]');

  document.addEventListener('mouseover', ev => {
    const el = tipped(ev.target);
    if (el === target) return;
    hide();
    if (!el) return;
    target = el;
    textOf(el); // снять title сразу, чтобы системная подсказка не успела появиться
    timer = setTimeout(() => { if (target === el) show(el); }, DELAY);
  });
  document.addEventListener('mouseout', ev => {
    if (target && !target.contains(ev.relatedTarget)) hide();
  });
  // с клавиатуры подсказка показывается при фокусе
  document.addEventListener('focusin', ev => {
    const el = tipped(ev.target);
    if (el && el.matches(':focus-visible')) { target = el; show(el); }
  });
  document.addEventListener('focusout', hide);
  document.addEventListener('mousedown', hide, true);
  document.addEventListener('keydown', ev => { if (ev.key === 'Escape') hide(); }, true);
  window.addEventListener('scroll', hide, true);
  window.addEventListener('resize', hide);
})();

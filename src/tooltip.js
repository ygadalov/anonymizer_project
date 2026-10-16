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

  // Место элемента на экране. У элементов с display: contents (подписи параметров в сетке) своего
  // прямоугольника нет — берём общий прямоугольник их содержимого. null — элемента на экране нет
  function boxOf(el) {
    const r = el.getBoundingClientRect();
    if (r.width || r.height) return r;
    const rects = Array.from(el.getClientRects()).concat(...Array.from(el.childNodes).map(n => {
      if (n.nodeType === 1) return Array.from(n.getClientRects());
      if (n.nodeType === 3 && n.textContent.trim()) { const rg = document.createRange(); rg.selectNodeContents(n); return Array.from(rg.getClientRects()); }
      return [];
    })).filter(x => x.width || x.height);
    if (!rects.length) return null;
    const left = Math.min(...rects.map(x => x.left)), top = Math.min(...rects.map(x => x.top));
    const right = Math.max(...rects.map(x => x.right)), bottom = Math.max(...rects.map(x => x.bottom));
    return { left, top, right, bottom, width: right - left, height: bottom - top };
  }

  function show(el) {
    // элемент мог исчезнуть, пока ждали задержку (список закрылся, строку таблицы перерисовали):
    // у такого элемента нет координат, и подсказка оказалась бы в левом верхнем углу
    if (!el.isConnected) return;
    const rect = boxOf(el);
    if (!rect) return;
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
    const r = rect, w = tip.offsetWidth, h = tip.offsetHeight;
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

  // после щелчка или выбора под неподвижным курсором оказывается другой элемент (список закрылся,
  // строка перерисовалась) — его подсказка не нужна, пока мышь снова не сдвинется
  let quiet = false, qx = 0, qy = 0, mx = 0, my = 0;
  const hush = () => { quiet = true; qx = mx; qy = my; hide(); };
  document.addEventListener('mousedown', hush, true);
  document.addEventListener('change', hush, true);
  document.addEventListener('mousemove', ev => {
    mx = ev.clientX; my = ev.clientY;
    if (!quiet || Math.abs(mx - qx) + Math.abs(my - qy) < 3) return;
    quiet = false;
    const el = tipped(ev.target);
    if (el && el !== target) { target = el; textOf(el); timer = setTimeout(() => { if (target === el) show(el); }, DELAY); }
  }, true);

  document.addEventListener('mouseover', ev => {
    const el = tipped(ev.target);
    if (el === target) return;
    hide();
    if (!el || quiet) return;
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
  // элемент под подсказкой удалили из страницы — подсказку убираем (mouseout в этом случае не приходит)
  new MutationObserver(() => { if (target && !target.isConnected) hide(); })
    .observe(document.body, { childList: true, subtree: true });
  document.addEventListener('keydown', ev => { if (ev.key === 'Escape') hide(); }, true);
  window.addEventListener('scroll', hide, true);
  window.addEventListener('resize', hide);
})();

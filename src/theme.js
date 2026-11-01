
  // тема до отрисовки, чтобы не мигало: сохранённый выбор или системная
  try { const t = localStorage.getItem('redact-theme'); if (t) document.documentElement.dataset.theme = t; } catch (e) {}

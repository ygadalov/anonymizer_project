#!/usr/bin/env python3
"""Собирает src/ в один автономный файл Anonymizer_vX.Y.html (работает офлайн, без установки)."""
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / 'src'


ICON_MARKERS = re.compile(r'<!-- icons:start -->.*?<!-- icons:end -->', re.S)


def icon_sprite():
    """SVG-спрайт из src/icons/*.svg (иконки Lucide). В разметке: <svg class="i"><use href="#i-имя"></use></svg>"""
    symbols = []
    for f in sorted((SRC / 'icons').glob('*.svg')):
        svg = f.read_text(encoding='utf-8')
        body = re.search(r'<svg[^>]*>(.*)</svg>', svg, re.S).group(1)
        body = re.sub(r'\s*\n\s*', '', body)
        symbols.append(f'<symbol id="i-{f.stem}" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
                       f'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{body}</symbol>')
    return ('<!-- icons:start -->\n<!-- Иконки: Lucide (lucide.dev), лицензия ISC — vendor/lucide.LICENSE. '
            'Генерируется build.py из src/icons/ -->\n'
            '<svg xmlns="http://www.w3.org/2000/svg" style="display:none">\n'
            + '\n'.join(symbols) + '\n</svg>\n<!-- icons:end -->')


def main():
    index = SRC / 'index.html'
    html = index.read_text(encoding='utf-8')
    # спрайт обновляется и в исходнике, чтобы src/index.html тоже работал без сборки
    html = ICON_MARKERS.sub(lambda m: icon_sprite(), html)
    index.write_text(html, encoding='utf-8')

    def css(m):
        return '<style>\n' + (SRC / m.group(1)).read_text(encoding='utf-8') + '</style>'

    def js(m):
        code = (SRC / m.group(1)).resolve().read_text(encoding='utf-8')
        if '</script' in code.lower():
            raise SystemExit(f'{m.group(1)}: содержит </script — встроить нельзя')
        return '<script>\n' + code + '\n</script>'

    html = re.sub(r'<link rel="stylesheet" href="([^"]+)">', css, html)
    html = re.sub(r'<script src="([^"]+)"></script>', js, html)
    m = re.search(r"ANON_VERSION\s*=\s*'([\d.]+)'", (SRC / 'version.js').read_text(encoding='utf-8'))
    if not m:
        raise SystemExit('src/version.js: не найдена версия')
    out = ROOT / f'Anonymizer_v{m.group(1)}.html'
    out.write_text(html, encoding='utf-8')
    print(f'{out.name}: {out.stat().st_size / 1024:.0f} КБ')


if __name__ == '__main__':
    main()

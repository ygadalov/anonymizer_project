#!/usr/bin/env python3
"""Собирает src/ в один автономный файл Redact_vX.Y.html (работает офлайн, без установки)
и страницу проекта index.html (из src/landing.html) для GitHub Pages."""
import base64
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


def inline(name):
    """Читает src/<name>, обновляет в нём спрайт иконок и встраивает стили, шрифты, SVG и скрипты."""
    src = SRC / name
    html = src.read_text(encoding='utf-8')
    # спрайт обновляется и в исходнике, чтобы страница из src/ тоже работала без сборки
    html = ICON_MARKERS.sub(lambda m: icon_sprite(), html)
    src.write_text(html, encoding='utf-8')

    def data_uri(path):
        mime = {'.svg': 'image/svg+xml', '.woff2': 'font/woff2'}[path.suffix]
        return f'data:{mime};base64,' + base64.b64encode(path.read_bytes()).decode('ascii')

    def css(m):
        path = (SRC / m.group(1)).resolve()
        code = path.read_text(encoding='utf-8')
        # шрифты (url(fonts/...woff2)) встраиваются, чтобы файл работал без интернета
        code = re.sub(r'url\(([^)]+\.woff2)\)', lambda u: f'url({data_uri(path.parent / u.group(1))})', code)
        return '<style>\n' + code + '</style>'

    def svg_asset(m):
        return m.group(1) + data_uri((SRC / m.group(2)).resolve()) + '"'

    def js(m):
        code = (SRC / m.group(1)).resolve().read_text(encoding='utf-8')
        if '</script' in code.lower():
            raise SystemExit(f'{m.group(1)}: содержит </script — встроить нельзя')
        return '<script>\n' + code + '\n</script>'

    html = re.sub(r'<link rel="stylesheet" href="([^"]+)">', css, html)
    # логотип и фавикон бренда — из brand/logo, встраиваются как есть
    html = re.sub(r'((?:src|href)=")([^"]+\.svg)"', svg_asset, html)
    html = re.sub(r'<script src="([^"]+)"></script>', js, html)

    # библиотеки для фонового потока: текст встраивается, но на странице не выполняется (type="text/x-worker")
    def worker_lib(m):
        code = (SRC / m.group(2)).resolve().read_text(encoding='utf-8')
        if '</script' in code.lower():
            raise SystemExit(f'{m.group(2)}: содержит </script — встроить нельзя')
        return f'<script type="text/x-worker" id="{m.group(1)}">\n' + code + '\n</script>'
    html = re.sub(r'<script type="text/x-worker" id="([^"]+)" src="([^"]+)"></script>', worker_lib, html)
    return html


def main():
    html = inline('index.html')
    m = re.search(r"ANON_VERSION\s*=\s*'([\d.]+)'", (SRC / 'version.js').read_text(encoding='utf-8'))
    if not m:
        raise SystemExit('src/version.js: не найдена версия')
    out = ROOT / f'Redact_v{m.group(1)}.html'
    out.write_text(html, encoding='utf-8')
    print(f'{out.name}: {out.stat().st_size / 1024:.0f} КБ')

    # страница проекта для GitHub Pages: index.html в корне ссылается на свежую сборку
    size = out.stat().st_size / 1024 / 1024
    page = inline('landing.html')
    page = (page.replace('{{VERSION}}', m.group(1)).replace('{{APP}}', out.name)
                .replace('{{SIZE}}', f'{size:.1f} МБ'.replace('.', ',')))
    landing = ROOT / 'index.html'
    landing.write_text(page, encoding='utf-8')
    print(f'{landing.name}: {landing.stat().st_size / 1024:.0f} КБ')


if __name__ == '__main__':
    main()

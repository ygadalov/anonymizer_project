#!/usr/bin/env python3
"""Собирает src/ в один автономный файл Anonymizer_vX.Y.html (работает офлайн, без установки)."""
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / 'src'


def main():
    html = (SRC / 'index.html').read_text(encoding='utf-8')

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

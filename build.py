#!/usr/bin/env python3
"""Offline build for Redact; never writes original/."""
from pathlib import Path
import re,base64,hashlib,json
ROOT=Path(__file__).resolve().parent

def build():
    def data(m):
        path,mime=m.groups()
        return 'data:'+mime+';base64,'+base64.b64encode((ROOT/path).read_bytes()).decode()
    def expand(text):
        return re.sub(r'\{\{DATA:([^:}]+):([^}]+)\}\}',data,text)
    template=(ROOT/'src/index.html').read_text()
    html=re.sub(r'\{\{FILE:([^}]+)\}\}',lambda m:expand((ROOT/m[1]).read_text()),template)
    html=expand(html)
    version=re.search(r"ANON_VERSION\s*=\s*'([\d.]+)'",(ROOT/'src/version.js').read_text())[1]
    out=ROOT/'dist'/f'Redact_v{version}.html';out.parent.mkdir(exist_ok=True);out.write_text(html)
    digest=hashlib.sha256(out.read_bytes()).hexdigest()
    expected=json.loads((ROOT/'recovery/manifest.json').read_text())['source_sha256']
    print(json.dumps({'output':str(out.relative_to(ROOT)),'version':version,'sha256':digest,'matches_original_v6_2':digest==expected}))
    return out
if __name__=='__main__': build()

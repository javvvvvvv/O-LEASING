import os
import glob
import sys
sys.path.insert(0, '.')

from core.cfdi import parse_cfdi, REGLAS_CONCEPTO_DEFAULT

xmls = glob.glob('*.xml') + glob.glob('**/*.xml', recursive=True)
print(f"Total XML files found: {len(xmls)}")
ok = 0
errs = []
for path in xmls[:50]:
    try:
        with open(path, 'rb') as f:
            raw = f.read()
        res = parse_cfdi(raw, REGLAS_CONCEPTO_DEFAULT)
        ok += 1
        print(f"[OK] {os.path.basename(path)} -> UUID: {res['uuid']} | Folio: {res['folio']} | Total: {res['total']} | Contrato: {res['id_contrato']}")
    except Exception as e:
        errs.append((path, str(e)))

print(f"\nParsed successfully: {ok}/{min(50, len(xmls))}")
if errs:
    print("Errors:")
    for p, e in errs:
        print(f"  {p}: {e}")

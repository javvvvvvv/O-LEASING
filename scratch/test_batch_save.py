import os
import glob
import sys
import sqlite3
sys.path.insert(0, '.')

from core.cfdi import parse_cfdi, REGLAS_CONCEPTO_DEFAULT

xmls = glob.glob('*.xml') + glob.glob('**/*.xml', recursive=True)
print(f"Total XML files found: {len(xmls)}")

conn = sqlite3.connect("leasing.db")
cur = conn.cursor()

results = []
for p in xmls[:10]:
    with open(p, 'rb') as f:
        raw = f.read()
    fact = parse_cfdi(raw, REGLAS_CONCEPTO_DEFAULT)
    results.append(fact)

print(f"Parsed {len(results)} XML invoices successfully.")

rows_fact = []
for r in results:
    rows_fact.append((
        r['uuid'],
        r['id_contrato'],
        r['fecha'],
        r.get('folio', ''),
        r['subtotal'],
        r['total'],
        r['tipo'],
        r['periodo'],
        r['mes_contrato'],
        'CONCILIADO',
        'Test batch save',
        r.get('rfc_emisor'),
        r.get('rfc_receptor'),
        0.0,
        r['total'],
        0.0,
        r['id_contrato'],
        0
    ))

cur.executemany("""
    INSERT OR REPLACE INTO facturas
    (uuid, id_contrato, fecha_emision, folio, subtotal, total,
     tipo, periodo, mes_contrato, estatus, observaciones,
     rfc_emisor, rfc_receptor, esperado, facturado, diferencia,
     id_contrato_detectado, alias_aplicado)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
""", rows_fact)
conn.commit()
print(f"Successfully inserted/updated {len(rows_fact)} facturas into leasing.db!")
conn.close()

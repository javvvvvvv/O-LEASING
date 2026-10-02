import os
import glob
import sys
import sqlite3
from datetime import datetime
sys.path.insert(0, '.')

from core.cfdi import parse_cfdi, REGLAS_CONCEPTO_DEFAULT

folder = r"C:\Users\cynti\Downloads\tre\XML_Exportados_24-09-2026 12-46-34"
paths = glob.glob(os.path.join(folder, "*.xml")) + glob.glob(os.path.join(folder, "**", "*.xml"), recursive=True)
print(f"Total XML files in target folder: {len(paths)}")

conn = sqlite3.connect("leasing.db")
cur = conn.cursor()

results = []
errores = []

for i, fp in enumerate(paths):
    nombre = os.path.basename(fp)
    try:
        with open(fp, "rb") as fh:
            raw = fh.read()
        if not raw:
            errores.append(f"{nombre}: vacío")
            continue
        fact = parse_cfdi(raw, REGLAS_CONCEPTO_DEFAULT)
        results.append(fact)
    except Exception as e:
        errores.append(f"{nombre}: {e}")

print(f"Parsed XML files count: {len(results)} / {len(paths)}")
print(f"Errors count: {len(errores)}")

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
        'CONCILIADO' if r['id_contrato'] else 'SIN_CONTRATO',
        'Carga masiva anual',
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
print(f"Successfully inserted/updated all {len(rows_fact)} facturas from target folder into leasing.db!")

# Query period breakdown from DB
cur.execute("SELECT periodo, COUNT(*), SUM(total) FROM facturas GROUP BY periodo ORDER BY periodo")
breakdown = cur.fetchall()
print("\nDesglose de facturas cargadas en la BD por periodo/mes:")
for p, c, t in breakdown:
    print(f"  Mes: {p} | Facturas: {c} | Total Facturado: ${t:,.2f}")

conn.close()

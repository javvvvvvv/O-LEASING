import os
import glob
import sys
import sqlite3
sys.path.insert(0, '.')

from core.cfdi import parse_cfdi, REGLAS_CONCEPTO_DEFAULT

folder = r"C:\Users\cynti\Downloads\tre\XML_Exportados_24-09-2026 12-46-34"
xmls = glob.glob(os.path.join(folder, "*.xml")) + glob.glob(os.path.join(folder, "**", "*.xml"), recursive=True)

print(f"Total XML files in folder: {len(xmls)}")

emitters = set()
receptors = set()
for p in xmls[:50]:
    with open(p, "rb") as f:
        raw = f.read()
    fact = parse_cfdi(raw, REGLAS_CONCEPTO_DEFAULT)
    if fact.get("rfc_emisor"):
        emitters.add(fact["rfc_emisor"])
    if fact.get("rfc_receptor"):
        receptors.add(fact["rfc_receptor"])

print(f"RFC Emisores en los XMLs: {emitters}")
print(f"RFC Receptores en los XMLs: {receptors}")

# Leer rfc_arrendadora configurado en leasing.db
conn = sqlite3.connect("leasing.db")
cur = conn.cursor()

try:
    r = cur.execute("SELECT valor FROM configuracion WHERE clave='rfc_arrendadora'").fetchone()
    rfc_cfg = r[0] if r else "No configurado"
except Exception as e:
    rfc_cfg = f"Error: {e}"

print(f"RFC Arrendadora configurado en BD: '{rfc_cfg}'")
conn.close()

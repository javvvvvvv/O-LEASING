import sqlite3
import glob

for db_file in glob.glob('*.db') + glob.glob('data/*.db'):
    try:
        c = sqlite3.connect(db_file)
        c.execute("DELETE FROM contratos WHERE ID_Contrato='0472-0003'")
        c.commit()
        c.close()
    except Exception:
        pass

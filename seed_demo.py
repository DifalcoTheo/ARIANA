# -*- coding: utf-8 -*-
"""
Builds a demo database from the fictional sample emails, so the app can be
run without Outlook or any real data.

    python seed_demo.py
    python app.py          # then open http://127.0.0.1:5000  (admin / changeme)
"""
import os
import shutil
import sqlite3
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
DB = os.path.join(BASE, "cerebro.db")


def main():
    os.makedirs(DATA, exist_ok=True)
    shutil.copy(os.path.join(BASE, "sample_data", "oc_emails.json"),
                os.path.join(DATA, "oc_emails.json"))

    # Run the real extraction pipeline over the sample inbox (no COM / no Outlook).
    subprocess.run([sys.executable, os.path.join(BASE, "scripts", "extractor.py")], check=True)

    # Validate ~2/3 of them so the commercial / client views have data to show,
    # leaving the rest in the human-validation queue.
    con = sqlite3.connect(DB)
    for col, ddl in [("oc_codigo", "TEXT"), ("monto_final", "REAL")]:
        try:
            con.execute(f"ALTER TABLE oc_extraccion ADD COLUMN {col} {ddl}")
        except sqlite3.OperationalError:
            pass  # already migrated by the app
    rows = con.execute("SELECT id, monto FROM oc_extraccion WHERE estado='pendiente' ORDER BY id").fetchall()
    n = 1000
    for i, (oid, monto) in enumerate(rows):
        if i % 3 == 0:
            continue  # keep some pending to show the validation queue
        con.execute("UPDATE oc_extraccion SET estado='aprobado', oc_codigo=?, monto_final=? WHERE id=?",
                    (f"OC-{n}", monto, oid))
        n += 1
    con.commit()
    con.close()

    print("\nDemo database ready (cerebro.db).")
    print("Next:  python app.py   ->  http://127.0.0.1:5000   (login: admin / changeme)")


if __name__ == "__main__":
    main()

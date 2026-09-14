# -*- coding: utf-8 -*-
"""Muestra los correos del dump que aun NO estan en cerebro.db (por entry_id)."""
import json, os, sqlite3

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
emails = json.load(open(os.path.join(BASE, "data", "oc_emails.json"), encoding="utf-8"))
con = sqlite3.connect(os.path.join(BASE, "cerebro.db"))
en_db = {r[0] for r in con.execute("SELECT entry_id FROM oc_extraccion")}
con.close()

nuevos = [e for e in emails if e["entry_id"] not in en_db]
print(f"Total en casilla: {len(emails)} | Ya en base: {len(en_db)} | NUEVOS: {len(nuevos)}")
print("=" * 70)
for e in nuevos:
    print(f"idx {e['idx']} | {e['received'][:16]} | {e['sender_email']}")
    print(f"ASUNTO: {e['subject']}")
    print(f"ADJUNTOS ({e['num_adjuntos']}): {e['adjuntos']}")
    print("CUERPO:")
    print(e["body"])
    print("-" * 70)

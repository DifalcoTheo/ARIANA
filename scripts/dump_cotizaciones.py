# -*- coding: utf-8 -*-
"""
Baja los mails de la casilla quotes@example.com (Bandeja de entrada) a un
JSON local. SOLO LECTURA (no envia nada). Mismo patron que dump_oc.py.
Formato esperado de asunto: FECHA-Costeo-[CODIGO SP]-CLIENTE
Cuerpo: "solicitada por (VENDEDOR) para (CLIENTE)" + lista de filtros + observaciones.
"""
import json
import os
import sys
import win32com.client

MAILBOX = "quotes@example.com"
OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
OUT_FILE = os.path.join(OUT_DIR, "cotizaciones_emails.json")

olMail = 43


def find_inbox(ns, mailbox):
    for i in range(1, ns.Folders.Count + 1):
        store_folder = ns.Folders.Item(i)
        if mailbox.lower() in store_folder.Name.lower():
            for j in range(1, store_folder.Folders.Count + 1):
                sub = store_folder.Folders.Item(j)
                if sub.Name.lower() in ("inbox", "bandeja de entrada"):
                    return sub
            return store_folder
    return None


def safe(getter, default=""):
    try:
        return getter()
    except Exception:
        return default


def main():
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = find_inbox(ns, MAILBOX)
    if inbox is None:
        print(f"ERROR: no se encontro la casilla {MAILBOX}")
        sys.exit(1)

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    total = items.Count
    print(f"Casilla: {MAILBOX} | Bandeja de entrada: {total} items")

    registros = []
    for k in range(1, total + 1):
        try:
            it = items.Item(k)
        except Exception:
            continue
        if safe(lambda: it.Class, None) != olMail:
            continue
        adjuntos = []
        try:
            for a in range(1, it.Attachments.Count + 1):
                adjuntos.append(it.Attachments.Item(a).FileName)
        except Exception:
            pass
        registros.append({
            "idx": k,
            "entry_id": safe(lambda: it.EntryID),
            "received": str(safe(lambda: it.ReceivedTime)),
            "sender_name": safe(lambda: it.SenderName),
            "sender_email": safe(lambda: it.SenderEmailAddress),
            "subject": safe(lambda: it.Subject),
            "num_adjuntos": len(adjuntos),
            "adjuntos": adjuntos,
            "body": safe(lambda: it.Body),
        })

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, indent=2)

    print(f"\nGuardados {len(registros)} correos en:\n  {OUT_FILE}")
    print("\nRESUMEN (fecha | remitente | #adj | asunto):")
    print("-" * 90)
    for r in registros[:30]:
        print(f"  {r['received'][:16]} | {(r['sender_name'] or r['sender_email'])[:22].ljust(22)} "
              f"| {r['num_adjuntos']} adj | {(r['subject'] or '')[:45]}")


if __name__ == "__main__":
    main()

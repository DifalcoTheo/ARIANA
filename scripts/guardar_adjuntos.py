# -*- coding: utf-8 -*-
"""
Guarda los adjuntos de un correo especifico (por EntryID) de la casilla oc@.
SOLO LECTURA sobre Outlook; escribe los adjuntos en data/attachments/<idx>/.
Uso: python guardar_adjuntos.py <ENTRY_ID> <idx>
"""
import os
import sys
import win32com.client

MAILBOX = "orders@example.com"
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
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


def main():
    entry_id = sys.argv[1]
    idx = sys.argv[2] if len(sys.argv) > 2 else "x"
    out_dir = os.path.join(BASE, "data", "attachments", f"idx{idx}")
    os.makedirs(out_dir, exist_ok=True)

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = find_inbox(ns, MAILBOX)
    items = inbox.Items

    target = None
    for k in range(1, items.Count + 1):
        it = items.Item(k)
        try:
            if it.Class == olMail and it.EntryID == entry_id:
                target = it
                break
        except Exception:
            continue

    if target is None:
        print("No se encontro el correo con ese EntryID")
        sys.exit(1)

    print(f"Correo: {target.Subject}")
    print(f"Adjuntos: {target.Attachments.Count}")
    saved = []
    for a in range(1, target.Attachments.Count + 1):
        att = target.Attachments.Item(a)
        name = att.FileName
        if name.lower().endswith(".pdf"):
            dest = os.path.join(out_dir, name)
            att.SaveAsFile(dest)
            saved.append(dest)
            print(f"  guardado: {name}")
        else:
            print(f"  (omitido, no PDF): {name}")

    print("\nPDFs guardados en:", out_dir)
    for s in saved:
        print("  ", s)


if __name__ == "__main__":
    main()

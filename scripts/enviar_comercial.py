# -*- coding: utf-8 -*-
"""
Regenera los informes comerciales por vendedor (Alex, Sam, Diego) y
los envía por Outlook — un mail HTML por vendedor. Por ahora todos van a
DESTINATARIO (Theo); más adelante se rutea a la casilla de cada vendedor.
Reintenta una vez; si falla, deja el mail en Borradores.
Programado mensual (tarea "Ariana - Informe comercial mensual"). Requiere Outlook abierto.
"""
import datetime
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import informe_comercial as inf

# A dónde va el informe de cada vendedor. Hoy todos a Theo; luego se cambia por la
# casilla de cada uno (alex.f@..., etc.).
DESTINO = {
    "Alex Rivera": "manager@example.com",
    "Sam Carter": "manager@example.com",
    "Diego Torres":      "manager@example.com",
}

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG = os.path.join(BASE, "data", "envio_comercial.log")


def log(msg):
    linea = f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} | {msg}"
    print(linea)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
    except OSError:
        pass


def main():
    # 1) regenerar los 3 informes
    subprocess.run([sys.executable, os.path.join(BASE, "scripts", "informe_comercial.py")],
                   cwd=BASE, timeout=120)

    import pythoncom
    import win32com.client
    hoy = datetime.date.today().strftime("%m/%Y")
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        for v in inf.VENDEDORES_INFORME:
            ruta = inf.ruta_de(v)
            if not os.path.exists(ruta):
                log(f"{v}: no se generó el informe, se omite."); continue
            html = open(ruta, encoding="utf-8").read()
            dest = DESTINO.get(v, "").strip()

            def armar():
                mail = outlook.CreateItem(0)
                mail.Subject = f"Informe comercial · {v} — Ariana ({hoy})"
                mail.HTMLBody = html
                return mail

            if not dest:
                armar().Save(); log(f"{v}: sin destinatario, guardado en Borradores."); continue
            enviado = False
            for intento in (1, 2):
                try:
                    m = armar(); m.To = dest; m.Send()
                    log(f"{v}: ENVIADO a {dest} (intento {intento})."); enviado = True; break
                except Exception as e:
                    log(f"{v}: error al enviar (intento {intento}): {e}"); time.sleep(15)
            if not enviado:
                m = armar(); m.To = dest
                m.Subject = "[NO ENVIADO - revisar] " + m.Subject
                m.Save(); log(f"{v}: falló 2 veces, dejado en BORRADORES.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()

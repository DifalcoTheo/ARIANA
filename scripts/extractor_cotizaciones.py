# -*- coding: utf-8 -*-
"""
Extrae de los mails de cotizaciones@ (cotizaciones_emails.json) los campos
estructurados: fecha, cliente, vendedor, códigos SP y observaciones.
Asunto:  FECHA-Costeo-[CODIGO]-CLIENTE   (FECHA = YYYYMMDD)
Cuerpo:  "solicitada por VENDEDOR para CLIENTE de los siguientes artículos:" + lista SP + Observaciones.
Salida:  data/cotizaciones_extraidas.json (idempotente por entry_id).
Esto agrega el VENDEDOR a las cotizaciones (la base del cotizador no lo tiene) y
da el mapeo cliente->vendedor.
"""
import datetime
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comercial_util import nombre_vendedor, cliente_canonico

DATA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
IN_FILE = os.path.join(DATA, "cotizaciones_emails.json")
OUT_FILE = os.path.join(DATA, "cotizaciones_extraidas.json")

RE_SOLIC = re.compile(r"solicitad[ao]\s+por\s+(.+?)\s+para\s+(.+?)\s+(?:de\s+los\s+siguientes|del\s+siguiente|de\s+el\s+siguiente)\s+art", re.I | re.S)
RE_SOLIC2 = re.compile(r"solicitad[ao]\s+por\s+(.+?)\s+para\s+(.+?)[\.\n:]", re.I | re.S)
RE_FECHA = re.compile(r"(\d{8})")
RE_OBS = re.compile(r"Observaciones:\s*(.+?)(?:\n\s*De:|Muchas gracias|Saludos)", re.I | re.S)
RE_COD = re.compile(r"^\s*[\*\-•]\s*([A-Za-z0-9][A-Za-z0-9\-\.\/]{2,})", re.M)


def fecha_de_asunto(asunto):
    m = RE_FECHA.search(asunto or "")
    if not m:
        return None
    try:
        return datetime.datetime.strptime(m.group(1), "%Y%m%d").date().isoformat()
    except ValueError:
        return None


def limpiar(s):
    return re.sub(r"\s+", " ", (s or "").strip()).strip(" .-*,;:")


def parse(m):
    body = m.get("body") or ""
    asunto = m.get("subject") or ""
    sol = RE_SOLIC.search(body) or RE_SOLIC2.search(body)
    if not sol:
        return None                       # no es un costeo con formato estándar
    vendedor_txt = limpiar(sol.group(1))
    cliente_txt = limpiar(sol.group(2))
    # descartar capturas basura (mails RE:/reenvíos donde el regex agarró de más)
    if len(vendedor_txt) > 35 or not cliente_txt or len(cliente_txt) > 45:
        return None
    basura = ("observacion", "enviado el", "asunto:", "de:", "solicit", "http", "@")
    if any(b in (vendedor_txt + " " + cliente_txt).lower() for b in basura):
        return None
    # códigos SP: de la lista entre "artículos:" y "Observaciones:"
    seg = body
    a = body.lower().find("siguientes")
    b = body.lower().find("observaciones")
    if a >= 0 and b > a:
        seg = body[a:b]
    codigos = []
    for c in RE_COD.findall(seg):
        c = c.strip(" .-")
        if c and c.lower() not in ("de", "para") and c not in codigos:
            codigos.append(c)
    obs = RE_OBS.search(body)
    return {
        "entry_id": m.get("entry_id"),
        "fecha": fecha_de_asunto(asunto),
        "cliente": cliente_canonico(cliente_txt),
        "cliente_texto": cliente_txt,
        "vendedor": nombre_vendedor(vendedor_txt),
        "codigos": codigos,
        "observaciones": limpiar(obs.group(1)) if obs else "",
        "subject": asunto,
    }


def main():
    emails = json.load(open(IN_FILE, encoding="utf-8"))
    prev = {}
    if os.path.exists(OUT_FILE):
        for r in json.load(open(OUT_FILE, encoding="utf-8")):
            prev[r["entry_id"]] = r

    nuevos, sin_formato = 0, 0
    for m in emails:
        eid = m.get("entry_id")
        if eid in prev:
            continue
        reg = parse(m)
        if reg is None:
            sin_formato += 1
            continue
        prev[eid] = reg
        nuevos += 1

    salida = sorted(prev.values(), key=lambda r: (r.get("fecha") or ""), reverse=True)
    json.dump(salida, open(OUT_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    print(f"Cotizaciones extraídas nuevas: {nuevos} | sin formato estándar (RE:/licitaciones/etc.): {sin_formato}")
    print(f"Total en base de cotizaciones: {len(salida)}")
    # resumen por vendedor
    from collections import Counter
    porv = Counter(r["vendedor"] for r in salida)
    print("\nCotizaciones por vendedor:")
    for v, n in porv.most_common():
        print(f"  {n:>4}  {v}")


if __name__ == "__main__":
    main()

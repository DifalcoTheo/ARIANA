# -*- coding: utf-8 -*-
"""
Parsea el registro de ingeniería (INGENIERIA (NO TOCAR).html) a
data/registro_ingenieria.json. Cada fila = una fabricación:
  fecha (día-mes; año reconstruido: lista va de 2025-09-04 a hoy), lote,
  codigos SP (se antepone 'SP' a los que empiezan con dígito), vendedor, cliente.
Ignora ARMEXAS / Andrea / vacío como vendedor. Erik se registra (ex-vendedor).
Detecta duplicados de cliente por prefijo para aliasear.
"""
import json
import os
import re
import sys
from html.parser import HTMLParser
from collections import defaultdict, Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comercial_util import cliente_canonico

RUTA = r"./sample_data/engineering_log.html"
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "data", "registro_ingenieria.json")

VEND = {"alex": "Alex Rivera", "sam": "Sam Carter",
        "sam c": "Sam Carter", "sam carter": "Sam Carter",
        "nils": "Nils Becker", "nils": "Nils Becker", "diego": "Diego Torres",
        "marco": "Marco Rossi", "luca": "Luca Moretti", "erik": "Erik Solberg"}
IGNORAR_VEND = {"armexas", "andrea", "", "\u200b"}


class T(HTMLParser):
    def __init__(s): super().__init__(); s.filas=[]; s.fila=None; s.cel=None; s.b=[]
    def handle_starttag(s, t, a):
        if t == "tr": s.fila = []
        elif t in ("td", "th"): s.cel = []; s.b = []
    def handle_data(s, d):
        if s.cel is not None: s.b.append(d)
    def handle_endtag(s, t):
        if t in ("td", "th") and s.cel is not None:
            s.fila.append(" ".join(" ".join(s.b).split())); s.cel = None
        elif t == "tr" and s.fila is not None:
            s.filas.append(s.fila); s.fila = None


def norm_cod(c):
    c = c.strip().upper()
    return ("SP" + c) if c and c[0].isdigit() else c


def main():
    p = T(); p.feed(open(RUTA, encoding="utf-8", errors="replace").read())
    filas = [f for f in p.filas if len(f) >= 7]

    empezar = False
    crudas = []
    for f in filas:
        if f[2].strip().lower() == "recibido":
            empezar = True; continue
        if not empezar:
            continue
        cliente = f[6].strip()
        if not cliente:
            continue
        crudas.append({"recibido": f[2].strip(), "lote": f[3].strip(),
                       "codigo": f[4].strip(), "vend": f[5].strip(), "cliente": cliente})

    # Reconstrucción de año: rango real 2025-09-04 -> 2026-09 (lista por lote desc).
    # Regla determinística: meses 10,11,12 = 2025 ; meses 1..8 = 2026 ;
    # mes 9 = 2026 arriba (antes de entrar al bloque oct-dic 2025), 2025 abajo.
    wrapped = False
    registros = []
    for d in crudas:
        r = d["recibido"].replace("/", "-")
        parts = [x for x in r.split("-") if x.strip().isdigit()]
        fecha = None
        if len(parts) >= 2:
            dia, mes = int(parts[0]), int(parts[1])
            if mes >= 10:
                anio = 2025; wrapped = True
            elif mes <= 8:
                anio = 2026
            else:                       # mes == 9
                anio = 2025 if wrapped else 2026
            try:
                fecha = f"{anio:04d}-{mes:02d}-{dia:02d}"
            except ValueError:
                fecha = None
        codigos = [norm_cod(x) for x in re.split(r"[\s,]+", d["codigo"]) if x.strip()]
        vraw = d["vend"].strip().lower()
        vendedor = None if vraw in IGNORAR_VEND else VEND.get(vraw, d["vend"].strip())
        registros.append({
            "fecha": fecha, "lote": d["lote"],
            "codigos": codigos, "vendedor": vendedor,
            "cliente": cliente_canonico(d["cliente"]), "cliente_texto": d["cliente"],
        })

    json.dump(registros, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    fok = [r["fecha"] for r in registros if r["fecha"]]
    print(f"Registros: {len(registros)} | fechas {min(fok)} -> {max(fok)}")

    # duplicados por prefijo (una es prefijo de palabras de la otra)
    clientes = sorted(set(r["cliente"] for r in registros))
    dup = []
    for a in clientes:
        for b in clientes:
            if a != b and b.startswith(a + " ") and len(a.split()) >= 1:
                dup.append((b, a))   # b (más larga) -> a (más corta)
    print(f"\nCandidatos a duplicado por prefijo ({len(dup)}):")
    for b, a in dup:
        print(f"   {b}   ->   {a}")


if __name__ == "__main__":
    main()

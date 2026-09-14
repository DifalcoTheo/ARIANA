# -*- coding: utf-8 -*-
"""
Genera un vault de Obsidian (carpeta de .md enlazados) con la info del negocio:
Clientes, Vendedores, Órdenes de compra y Elementos filtrantes (de GASPI/cotizador).
Los notes se enlazan con [[wikilinks]] para armar la red navegable en Obsidian.
Fuentes: cerebro.db (OC), cotizaciones.db (cotizaciones), cartuchos.db (elementos).
"""
import datetime
import os
import re
import sqlite3
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comercial_util import nombre_vendedor, cliente_canonico

DESKTOP = r"."
VAULT = os.path.join(DESKTOP, "Ariana-Vault")
DB_OC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "cerebro.db")
BASES = os.path.join(DESKTOP, "BASES DE DATOS")   # external database location (2026-09-11)
DB_COT = os.path.join(BASES, "cotizaciones.db")
DB_CART = os.path.join(BASES, "cartuchos.db")

DIRS = {"cli": "Clientes", "vend": "Vendedores", "oc": "Ordenes de compra",
        "elem": "Elementos filtrantes"}


def san(s):
    """Nombre de archivo/nota válido para Obsidian/Windows."""
    s = re.sub(r'[\\/:*?"<>|#\^\[\]]+', "-", str(s or "").strip())
    s = re.sub(r"\s+", " ", s).strip(" .-")
    return s[:100] or "sin-nombre"


def w(sub, nombre, contenido):
    d = os.path.join(VAULT, DIRS[sub])
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, san(nombre) + ".md"), "w", encoding="utf-8") as f:
        f.write(contenido)


def fnum(x):
    try:
        return f"{float(x):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except (TypeError, ValueError):
        return str(x)


def main():
    os.makedirs(VAULT, exist_ok=True)
    hoy = datetime.date.today().strftime("%d/%m/%Y")

    # ── OC validadas ──
    oc = sqlite3.connect(DB_OC); oc.row_factory = sqlite3.Row
    ocs = oc.execute(
        "SELECT oc_codigo, empresa, COALESCE(vendedor_final,sender) v, oc_numero, "
        "COALESCE(monto_final,monto) monto, moneda, COALESCE(fecha_oc,received) fecha_oc, "
        "fecha_entrega, COALESCE(fabricacion,0) fab, received "
        "FROM oc_extraccion WHERE estado='aprobado'").fetchall()
    oc.close()

    clientes = defaultdict(lambda: {"vend": defaultdict(int), "ocs": [], "usd": 0.0,
                                    "fab": 0, "n": 0, "elems": []})
    for r in ocs:
        c = cliente_canonico(r["empresa"])
        d = clientes[c]
        d["vend"][nombre_vendedor(r["v"])] += 1
        d["ocs"].append(r)
        d["usd"] += (r["monto"] or 0)
        d["fab"] += r["fab"]; d["n"] += 1

    # ── Cotizaciones por cliente ──
    cot_idx = defaultdict(lambda: {"n": 0, "ultima": ""})
    if os.path.exists(DB_COT):
        cot = sqlite3.connect(DB_COT); cot.row_factory = sqlite3.Row
        for r in cot.execute("SELECT cliente, fecha_costeo FROM filtros WHERE cliente IS NOT NULL AND cliente<>''"):
            c = cliente_canonico(r["cliente"]); a = cot_idx[c]
            a["n"] += 1
            f = str(r["fecha_costeo"] or "")[:7]
            if f > a["ultima"]:
                a["ultima"] = f
        cot.close()

    # ── Elementos filtrantes (cartuchos.db) ──
    elementos = {}   # sp -> dict
    if os.path.exists(DB_CART):
        ca = sqlite3.connect(DB_CART); ca.row_factory = sqlite3.Row
        def cargar(tabla, familia, campos):
            try:
                for r in ca.execute(f"SELECT * FROM {tabla}"):
                    sp = (r["sp"] or "").strip()
                    if not sp:
                        continue
                    if sp in elementos:
                        continue
                    cli = cliente_canonico(r["cliente"]) if r["cliente"] else ""
                    specs = {k: r[k] for k in campos if k in r.keys() and r[k] not in (None, "")}
                    elementos[sp] = {"familia": familia, "cliente": cli, "spf": (r["spf"] or ""), "specs": specs}
                    if cli in clientes:
                        clientes[cli]["elems"].append(sp)
            except Exception:
                pass
        cargar("cartuchos", "cartucho", ["altura", "od", "id_interior", "plisado",
               "cant_pliegues", "altura_pliegue", "nucleo", "jaula"])
        cargar("conos", "cono", ["altura_total", "diametro_exterior", "diametro_base",
               "medio_filtrante", "material"])
        cargar("canastos", "canasto", ["altura_maxima", "brida", "material", "medio_filtrante"])
        ca.close()

    # asegurar que clientes con elementos pero sin OC también tengan ficha
    for sp, e in elementos.items():
        if e["cliente"] and e["cliente"] not in clientes:
            clientes[e["cliente"]]  # crea entrada vacía
            clientes[e["cliente"]]["elems"].append(sp)

    # ── Escribir notas de CLIENTES ──
    vend_clientes = defaultdict(list)
    for c, d in sorted(clientes.items()):
        vend_princ = max(d["vend"].items(), key=lambda x: x[1])[0] if d["vend"] else ""
        if vend_princ:
            vend_clientes[vend_princ].append(c)
        cot = cot_idx.get(c, {"n": 0, "ultima": ""})
        if d["n"] == 0:
            compra = "sin-oc"
        elif d["fab"] == d["n"]:
            compra = "fabricacion"
        elif d["fab"] == 0:
            compra = "stock"
        else:
            compra = "mixto"
        fm = ["---", "tipo: cliente",
              f"vendedor: {vend_princ}" if vend_princ else "vendedor: ",
              f"oc_total: {d['n']}", f"usd_total: {round(d['usd'],2)}",
              f"cotizaciones: {cot['n']}", f"ultima_cotizacion: {cot['ultima']}",
              f"compra: {compra}", "rubro: ", "---", ""]
        body = [f"# {c}", ""]
        body.append(f"**Vendedor:** [[{san(vend_princ)}]]" if vend_princ else "**Vendedor:** _(asignar)_")
        if len(d["vend"]) > 1:
            otros = ", ".join(f"[[{san(v)}]] ({n})" for v, n in d["vend"].items() if v != vend_princ)
            body.append(f"**También le vendió:** {otros}")
        body.append(f"**Compra:** {compra} · **Cotizaciones:** {cot['n']} (última {cot['ultima'] or 'nunca'})")
        body.append("")
        if d["ocs"]:
            body.append("## Órdenes de compra")
            for r in sorted(d["ocs"], key=lambda x: str(x["fecha_oc"]), reverse=True):
                body.append(f"- [[{san(r['oc_codigo'] or 'OC')}]] — {r['moneda'] or 'USD'} {fnum(r['monto'])} "
                            f"· entrega: {r['fecha_entrega'] or '—'}")
            body.append("")
        if d["elems"]:
            body.append("## Elementos filtrantes")
            for sp in sorted(set(d["elems"])):
                body.append(f"- [[{san(sp)}]]")
            body.append("")
        body.append("## Notas\n_(agregar requisitos, gustos, contactos, condiciones…)_\n")
        w("cli", c, "\n".join(fm) + "\n".join(body))

    # ── Notas de VENDEDORES ──
    for v, cls in sorted(vend_clientes.items()):
        fm = ["---", "tipo: vendedor", f"clientes: {len(cls)}", "---", ""]
        body = [f"# {v}", "", "## Clientes"]
        for c in sorted(cls):
            body.append(f"- [[{san(c)}]]")
        w("vend", v, "\n".join(fm) + "\n".join(body))

    # ── Notas de OC ──
    for r in ocs:
        c = cliente_canonico(r["empresa"]); vend = nombre_vendedor(r["v"])
        cod = r["oc_codigo"] or f"OC-{r['oc_numero']}"
        fm = ["---", "tipo: orden-compra", f"codigo: {cod}", f"cliente: {c}",
              f"vendedor: {vend}", f"oc_numero: {r['oc_numero'] or ''}",
              f"monto: {round(r['monto'] or 0,2)}", f"moneda: {r['moneda'] or 'USD'}",
              f"fecha_oc: {str(r['fecha_oc'])[:10]}", f"fecha_entrega: {r['fecha_entrega'] or ''}",
              f"fabricacion: {'true' if r['fab'] else 'false'}", "---", ""]
        body = [f"# {cod} — {c}", "",
                f"Cliente: [[{san(c)}]] · Vendedor: [[{san(vend)}]]",
                f"N° OC cliente: {r['oc_numero'] or '—'} · Monto: {r['moneda'] or 'USD'} {fnum(r['monto'])}",
                f"Fecha OC: {str(r['fecha_oc'])[:10]} · Entrega: {r['fecha_entrega'] or '—'} · "
                f"{'Fabricación' if r['fab'] else 'Stock'}", ""]
        w("oc", cod, "\n".join(fm) + "\n".join(body))

    # ── Notas de ELEMENTOS FILTRANTES ──
    for sp, e in sorted(elementos.items()):
        fm = ["---", "tipo: elemento-filtrante", f"sp: {sp}",
              f"spf: {e['spf']}", f"cliente: {e['cliente']}", f"familia: {e['familia']}"]
        for k, val in e["specs"].items():
            fm.append(f"{k}: {val}")
        fm += ["---", ""]
        body = [f"# {sp} — {e['familia']}", ""]
        if e["cliente"]:
            body.append(f"Cliente: [[{san(e['cliente'])}]]")
        if e["spf"]:
            body.append(f"Lote: {e['spf']}")
        if e["specs"]:
            body.append("\n## Especificaciones")
            for k, val in e["specs"].items():
                body.append(f"- **{k.replace('_',' ')}:** {val}")
        w("elem", sp, "\n".join(fm) + "\n".join(body) + "\n")

    # ── Home / índice ──
    home = [
        "# 🏠 Ariana — Vault", "",
        f"> Generado el {hoy} desde las bases (OC, cotizaciones, elementos filtrantes de GASPI).",
        "> La red se navega con la vista **Graph** de Obsidian. Los `[[enlaces]]` conectan"
        " clientes ↔ vendedores ↔ órdenes ↔ elementos.", "",
        "## Contenido",
        f"- **Clientes:** {len(clientes)} → carpeta [[Clientes]]",
        f"- **Vendedores:** {len(vend_clientes)} → carpeta [[Vendedores]]",
        f"- **Órdenes de compra:** {len(ocs)} → carpeta [[Ordenes de compra]]",
        f"- **Elementos filtrantes:** {len(elementos)} → carpeta [[Elementos filtrantes]]", "",
        "> **Esta página (Home) es solo un índice.** No hay nada que completar acá.", "",
        "## Qué completás vos (en las notas de Clientes)",
        "Abrí la carpeta **Clientes** y en cada ficha:",
        "- Completá el campo **`rubro:`** (en las propiedades) — ej. laboratorio, alimenticia, energía.",
        "- Escribí en la sección **Notas** lo cualitativo: requisitos, gustos, contactos, condiciones de pago.",
        "- A los clientes marcados **_(asignar)_** en Vendedor, ponéles el vendedor que corresponde.", "",
        "## Qué hace el asistente (Ariana / \"Cowork\")",
        "Leo estas notas al generar los informes y sumo lo que escribiste. Son archivos `.md`"
        " que se leen directo; no hay que instalar nada extra en Obsidian.", "",
    ]
    with open(os.path.join(VAULT, "Home.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(home))

    print(f"Vault creado en: {VAULT}")
    print(f"  Clientes: {len(clientes)} | Vendedores: {len(vend_clientes)} | "
          f"OC: {len(ocs)} | Elementos: {len(elementos)}")


if __name__ == "__main__":
    main()

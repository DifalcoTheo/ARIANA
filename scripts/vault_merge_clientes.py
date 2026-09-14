# -*- coding: utf-8 -*-
"""
Regenera SOLO las notas de Clientes del vault, PRESERVANDO lo cargado a mano:
  - rubro y cualquier campo propio del frontmatter
  - vendedor asignado a mano (si el dato no lo deriva)
  - la sección '## Notas' (texto libre)
Fusiona duplicados según data/clientes_alias.json (COSMACA-TH4-PM -> COSMACA) y
borra las notas variantes ya fusionadas. No toca Vendedores / OC / Elementos / Home.
"""
import os
import re
import sqlite3
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comercial_util import nombre_vendedor, cliente_canonico, VENDEDORES_ACTIVOS

DESKTOP = r"."
COT_EXTR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "data", "cotizaciones_extraidas.json")
REG_ING = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "data", "registro_ingenieria.json")
VAULT = os.path.join(DESKTOP, "Ariana-Vault")
CLI_DIR = os.path.join(VAULT, "Clientes")
DB_OC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "cerebro.db")
BASES = os.path.join(DESKTOP, "BASES DE DATOS")   # external database location (2026-09-11)
DB_COT = os.path.join(BASES, "cotizaciones.db")
DB_CART = os.path.join(BASES, "cartuchos.db")

GEN_KEYS = {"tipo", "vendedor", "oc_total", "usd_total", "cotizaciones", "ultima_cotizacion", "compra"}
PLACEHOLDER = "_(agregar requisitos, gustos, contactos, condiciones…)_"


def san(s):
    s = re.sub(r'[\\/:*?"<>|#\^\[\]]+', "-", str(s or "").strip())
    return re.sub(r"\s+", " ", s).strip(" .-")[:100] or "sin-nombre"


def fnum(x):
    try:
        return f"{float(x):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except (TypeError, ValueError):
        return str(x)


def parse_fm(texto):
    """(dict_frontmatter, cuerpo). Conserva el orden de las claves."""
    if not texto.startswith("---"):
        return {}, texto
    p = texto.split("---", 2)
    if len(p) < 3:
        return {}, texto
    fm = {}
    for ln in p[1].strip().splitlines():
        if ":" in ln:
            k, _, v = ln.partition(":")
            fm[k.strip()] = v.strip()
    return fm, p[2].lstrip("\n")


def extraer_notas(cuerpo):
    i = cuerpo.find("## Notas")
    if i < 0:
        return ""
    txt = cuerpo[i + len("## Notas"):].strip()
    if not txt or "agregar requisitos" in txt:
        return ""
    return txt


def cargar_datos():
    """canonical -> stats + elementos. Igual criterio que crear_vault."""
    clientes = defaultdict(lambda: {"vend": defaultdict(int), "ocs": [], "usd": 0.0,
                                    "fab": 0, "n": 0, "elems": set()})
    oc = sqlite3.connect(DB_OC); oc.row_factory = sqlite3.Row
    for r in oc.execute("SELECT oc_codigo, empresa, COALESCE(vendedor_final,sender) v, "
                        "COALESCE(monto_final,monto) monto, moneda, fecha_entrega, "
                        "COALESCE(fecha_oc,received) fecha_oc, COALESCE(fabricacion,0) fab "
                        "FROM oc_extraccion WHERE estado='aprobado'"):
        c = cliente_canonico(r["empresa"]); d = clientes[c]
        d["vend"][nombre_vendedor(r["v"])] += 1
        d["ocs"].append(r); d["usd"] += (r["monto"] or 0); d["fab"] += r["fab"]; d["n"] += 1
    oc.close()

    cot_idx = defaultdict(lambda: {"n": 0, "ultima": ""})
    if os.path.exists(DB_COT):
        c = sqlite3.connect(DB_COT); c.row_factory = sqlite3.Row
        for r in c.execute("SELECT cliente, fecha_costeo FROM filtros WHERE cliente IS NOT NULL AND cliente<>''"):
            k = cliente_canonico(r["cliente"]); a = cot_idx[k]
            a["n"] += 1; f = str(r["fecha_costeo"] or "")[:7]
            if f > a["ultima"]:
                a["ultima"] = f
        c.close()

    if os.path.exists(DB_CART):
        ca = sqlite3.connect(DB_CART); ca.row_factory = sqlite3.Row
        for tabla in ("cartuchos", "conos", "canastos"):
            try:
                for r in ca.execute(f"SELECT sp, cliente FROM {tabla}"):
                    sp = (r["sp"] or "").strip()
                    if sp and r["cliente"]:
                        clientes[cliente_canonico(r["cliente"])]["elems"].add(sp)
            except Exception:
                pass
        ca.close()

    # registro de ingeniería (fabricaciones históricas): agrega clientes, códigos y vendedor real
    from collections import Counter
    import json as _json
    log_vend = {}
    if os.path.exists(REG_ING):
        lv = defaultdict(Counter)
        for r in _json.load(open(REG_ING, encoding="utf-8")):
            c = r.get("cliente")
            if not c:
                continue
            clientes[c]["elems"].update(x for x in r.get("codigos", []) if x)
            v = r.get("vendedor")
            if v in VENDEDORES_ACTIVOS:
                lv[c][v] += 1
        log_vend = {c: cnt.most_common(1)[0][0] for c, cnt in lv.items() if cnt}
    return clientes, cot_idx, log_vend


def cargar_cotiz_vendedor():
    """canonical cliente -> vendedor ACTIVO más frecuente en cotizaciones (o '')."""
    import json
    from collections import Counter
    if not os.path.exists(COT_EXTR):
        return {}
    cv = defaultdict(Counter)
    for r in json.load(open(COT_EXTR, encoding="utf-8")):
        v, c = r.get("vendedor"), r.get("cliente")
        if v in VENDEDORES_ACTIVOS and c:
            cv[c][v] += 1
    return {c: cnt.most_common(1)[0][0] for c, cnt in cv.items() if cnt}


def main():
    if not os.path.isdir(CLI_DIR):
        print("No existe la carpeta Clientes del vault."); return
    clientes, cot_idx, log_vend = cargar_datos()
    cotiz_vend = cargar_cotiz_vendedor()

    # ── Preservar contenido manual de las notas existentes (agrupado por canónico) ──
    preserv = defaultdict(lambda: {"rubro": "", "extra": {}, "vend_manual": "", "notas": []})
    variantes = []   # (archivo, canonico) para borrar los que son alias
    for fn in os.listdir(CLI_DIR):
        if not fn.endswith(".md"):
            continue
        nombre = fn[:-3]
        canon = cliente_canonico(nombre)
        fm, cuerpo = parse_fm(open(os.path.join(CLI_DIR, fn), encoding="utf-8").read())
        p = preserv[canon]
        if fm.get("rubro") and not p["rubro"]:
            p["rubro"] = fm["rubro"]
        if fm.get("vendedor") and not p["vend_manual"]:
            p["vend_manual"] = fm["vendedor"]
        for k, v in fm.items():
            if k not in GEN_KEYS and k != "rubro" and v:
                p["extra"].setdefault(k, v)
        notas = extraer_notas(cuerpo)
        if notas:
            p["notas"].append(notas)
        if san(canon) != nombre:                 # es una nota variante (alias) -> se borrará
            variantes.append((fn, canon))

    # ── Reescribir notas canónicas ──
    escritos = 0
    for c, d in clientes.items():
        vend_princ = max(d["vend"].items(), key=lambda x: x[1])[0] if d["vend"] else ""
        pres = preserv.get(c, {"rubro": "", "extra": {}, "vend_manual": "", "notas": []})
        if not vend_princ and pres["vend_manual"]:
            vend_princ = pres["vend_manual"]      # respetar asignación manual
        if not vend_princ and log_vend.get(c):
            vend_princ = log_vend[c]              # fabricación real histórica (solo activos)
        if not vend_princ and cotiz_vend.get(c):
            vend_princ = cotiz_vend[c]            # inferido de cotizaciones (solo activos)
        cot = cot_idx.get(c, {"n": 0, "ultima": ""})
        compra = ("sin-oc" if d["n"] == 0 else "fabricacion" if d["fab"] == d["n"]
                  else "stock" if d["fab"] == 0 else "mixto")
        fm = ["---", "tipo: cliente", f"vendedor: {vend_princ}",
              f"oc_total: {d['n']}", f"usd_total: {round(d['usd'],2)}",
              f"cotizaciones: {cot['n']}", f"ultima_cotizacion: {cot['ultima']}",
              f"compra: {compra}", f"rubro: {pres['rubro']}"]
        for k, v in pres["extra"].items():
            fm.append(f"{k}: {v}")
        fm += ["---", ""]
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
            for sp in sorted(d["elems"]):
                body.append(f"- [[{san(sp)}]]")
            body.append("")
        notas = "\n\n".join(pres["notas"]) if pres["notas"] else PLACEHOLDER
        body.append("## Notas\n" + notas + "\n")
        with open(os.path.join(CLI_DIR, san(c) + ".md"), "w", encoding="utf-8") as f:
            f.write("\n".join(fm) + "\n".join(body))
        escritos += 1

    # ── Borrar notas variantes (alias) ya fusionadas ──
    borrados = 0
    for fn, canon in variantes:
        ruta = os.path.join(CLI_DIR, fn)
        if san(canon) != fn[:-3] and os.path.exists(ruta):
            os.remove(ruta); borrados += 1

    total = len([f for f in os.listdir(CLI_DIR) if f.endswith(".md")])
    con_notas = sum(1 for p in preserv.values() if p["notas"] or p["rubro"] or p["vend_manual"])
    print(f"Notas canónicas escritas: {escritos} | variantes borradas: {borrados} | "
          f"quedan {total} notas de clientes.")
    print(f"Notas con contenido manual preservado (rubro/vendedor/notas): {con_notas}")


if __name__ == "__main__":
    main()

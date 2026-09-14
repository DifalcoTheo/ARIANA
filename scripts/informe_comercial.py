# -*- coding: utf-8 -*-
"""
Informe comercial POR VENDEDOR en HTML — un archivo por vendedor (Alex,
Sam, Diego). Fuente: OC validadas (estricto desde MES_INICIO). Clientes
agrupados con comercial_util. Por vendedor: KPIs, ventas por mes, ranking de
clientes por OC + ticket + perfil, y una sección de OPORTUNIDADES cruzando las
cotizaciones (clientes que cotizan mucho / hace tiempo que no cotizan).
"""
import datetime
import os
import sqlite3
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comercial_util import nombre_vendedor, cliente_canonico

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = os.path.join(BASE, "cerebro.db")
OUT_DIR = os.path.join(BASE, "memoria")
MES_INICIO = "2026-09"   # el informe arranca en septiembre 2026
COTIZ_DB = r"./external_db\cotizaciones.db"   # base del cotizador (histórica, 2025+)
COTIZ_EXTR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "data", "cotizaciones_extraidas.json")   # cotizaciones con vendedor (mail, 2026+)
MESES_ES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MES_ABR = ["", "ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

# Solo estos tres, cada uno en su propio HTML
VENDEDORES_INFORME = ["Alex Rivera", "Sam Carter", "Diego Torres"]
SLUG = {"Alex Rivera": "alex_rivera",
        "Sam Carter": "sam_carter",
        "Diego Torres": "diego_torres"}
MESES_SIN_COTIZAR = 2   # umbral: si la última cotización es más vieja que esto -> "hace tiempo"


def ruta_de(vendedor):
    return os.path.join(OUT_DIR, f"informe_{SLUG[vendedor]}.html")


def fmt(x):
    return f"{x:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")


def mes_label(ym):
    a, m = ym.split("-")
    return f"{MESES_ES[int(m)].capitalize()} {a}"


def mes_corto(ym):
    if not ym or "-" not in ym:
        return "nunca"
    a, m = ym.split("-")
    return f"{MES_ABR[int(m)]} {a[2:]}"


def meses_desde(inicio, hasta):
    a, m = map(int, inicio.split("-"))
    a2, m2 = map(int, hasta.split("-"))
    out = []
    while (a, m) <= (a2, m2):
        out.append(f"{a:04d}-{m:02d}")
        m += 1
        if m > 12:
            m = 1; a += 1
    return out


def restar_meses(ym, k):
    a, m = map(int, ym.split("-"))
    m -= k
    while m < 1:
        m += 12; a -= 1
    return f"{a:04d}-{m:02d}"


def cargar_cotizaciones():
    """canonical cliente -> {n, ultima 'YYYY-MM'} desde cotizaciones.db. Vacío si no está."""
    idx = {}
    if not os.path.exists(COTIZ_DB):
        return idx
    try:
        c = sqlite3.connect(COTIZ_DB); c.row_factory = sqlite3.Row
        for r in c.execute("SELECT cliente, fecha_costeo FROM filtros "
                           "WHERE cliente IS NOT NULL AND cliente<>''"):
            k = cliente_canonico(r["cliente"])
            a = idx.setdefault(k, {"n": 0, "ultima": ""})
            a["n"] += 1
            f = str(r["fecha_costeo"] or "")[:7]
            if f > a["ultima"]:
                a["ultima"] = f
        c.close()
    except Exception:
        pass
    return idx


def cargar_cotiz_por_vendedor():
    """vendedor -> {cliente: {n, ultima}} desde cotizaciones_extraidas.json (mail, con vendedor)."""
    import json
    out = defaultdict(lambda: defaultdict(lambda: {"n": 0, "ultima": ""}))
    if not os.path.exists(COTIZ_EXTR):
        return out
    for r in json.load(open(COTIZ_EXTR, encoding="utf-8")):
        v, c, f = r.get("vendedor"), r.get("cliente"), (r.get("fecha") or "")[:7]
        if not v or not c:
            continue
        a = out[v][c]; a["n"] += 1
        if f > a["ultima"]:
            a["ultima"] = f
    return out


def _perfil(oc, ticket_cli, ticket_vend):
    if oc >= 3:
        return "Recurrente", "rec"
    if ticket_cli >= 2 * ticket_vend and ticket_vend > 0:
        return "Ticket alto", "alto"
    return "Puntual", "pun"


def _kpi(k, v, destacado=False):
    cls = "kpi big" if destacado else "kpi"
    return f'<div class="{cls}"><div class="k">{k}</div><div class="v">{v}</div></div>'


def construir_html(v, d, meses, COTIZ, generado, mes_actual, cotiz_vend, oc_cli_vend):
    H = [_ENCABEZADO.format(vendedor=v, generado=generado,
                            desde=mes_label(MES_INICIO), hasta=mes_label(mes_actual))]
    if not d or d["oc"] == 0:
        H.append('<section class="vend"><p class="vacio">Sin OC validadas en el período.</p></section>')
        H.append(_PIE)
        return "\n".join(H)

    ticket = d["total"] / d["oc"] if d["oc"] else 0
    H.append('<section class="vend">')
    H.append('<div class="kpis">')
    H.append(_kpi("Vendido (USD)", fmt(d["total"]), True))
    H.append(_kpi("Órdenes de compra", str(d["oc"])))
    H.append(_kpi("Ticket promedio", "USD " + fmt(ticket)))
    H.append(_kpi("Clientes", str(len(d["cli"]))))
    H.append('</div>')

    # Ventas por mes
    H.append('<h3>Ventas por mes</h3><table class="t"><thead><tr>'
             '<th>Mes</th><th class="n">USD</th><th class="n">OC</th></tr></thead><tbody>')
    for ym in meses:
        mm = d["mes"].get(ym, {"oc": 0, "total": 0.0})
        vacio = ' class="cero"' if mm["oc"] == 0 else ''
        H.append(f'<tr{vacio}><td>{mes_label(ym)}</td>'
                 f'<td class="n">{fmt(mm["total"])}</td><td class="n">{mm["oc"]}</td></tr>')
    H.append('</tbody></table>')

    # Clientes (sin columna de cotizaciones)
    H.append('<h3>Clientes — quién le compra más</h3>'
             '<table class="t"><thead><tr><th>#</th><th>Cliente</th>'
             '<th class="n">OC</th><th class="n">USD total</th><th class="n">Ticket</th>'
             '<th class="n">% vend.</th><th>Perfil</th></tr></thead><tbody>')
    clientes = sorted(d["cli"].items(), key=lambda x: (-x[1]["oc"], -x[1]["total"]))
    for i, (cli, c) in enumerate(clientes, 1):
        tk = c["total"] / c["oc"] if c["oc"] else 0
        pct = 100 * c["total"] / d["total"] if d["total"] else 0
        perfil, clase = _perfil(c["oc"], tk, ticket)
        H.append(f'<tr><td class="pos">{i}</td><td>{cli}</td>'
                 f'<td class="n">{c["oc"]}</td><td class="n">{fmt(c["total"])}</td>'
                 f'<td class="n">{fmt(tk)}</td><td class="n">{pct:.1f}%</td>'
                 f'<td><span class="badge {clase}">{perfil}</span></td></tr>')
    H.append('</tbody></table>')

    # ── Oportunidades (cruce con cotizaciones) ──
    cutoff = restar_meses(mes_actual, MESES_SIN_COTIZAR)
    cotiza_mucho, sin_cotizar = [], []
    for cli in d["cli"]:
        ct = COTIZ.get(cli)
        n = ct["n"] if ct else 0
        ult = ct["ultima"] if ct else ""
        if not ult or ult < cutoff:
            sin_cotizar.append((cli, n, ult))     # sin cotizar reciente -> reactivar
        elif n >= 4:
            cotiza_mucho.append((cli, n, ult))    # cotiza seguido y en volumen -> oportunidad
    cotiza_mucho.sort(key=lambda x: -x[1])
    sin_cotizar.sort(key=lambda x: (x[2] or ""))

    H.append('<h3>Oportunidades y seguimiento</h3><div class="ops">')
    H.append('<p class="ops-intro">Cruce con el historial de cotizaciones. '
             'Se irá afinando a medida que se acumulen datos.</p>')
    if cotiza_mucho:
        items = "".join(
            f'<li><b>{cli}</b> — {n} cotizaciones (última {mes_corto(ult)}). '
            f'Piden mucho presupuesto: oportunidad de cerrar más ventas.</li>'
            for cli, n, ult in cotiza_mucho[:6])
        H.append(f'<div class="ops-b op-si"><div class="ops-t">💡 Cotizan mucho</div>'
                 f'<ul>{items}</ul></div>')
    if sin_cotizar:
        items = "".join(
            f'<li><b>{cli}</b> — {"sin cotizaciones registradas" if not ult else "última cotización " + mes_corto(ult)}. '
            f'{"Compra sin pedir presupuesto (stock)." if not ult else "Hace tiempo que no cotiza: conviene reactivar."}</li>'
            for cli, n, ult in sin_cotizar[:6])
        H.append(f'<div class="ops-b op-no"><div class="ops-t">🕒 Hace tiempo que no cotizan</div>'
                 f'<ul>{items}</ul></div>')
    # Clientes que ESTE vendedor cotizó (mail) y todavía no le compraron (ninguna OC histórica)
    ya_compran = oc_cli_vend.get(v, set())
    cotizo = cotiz_vend.get(v, {})
    oportun = [(cli, info) for cli, info in cotizo.items() if cli not in ya_compran]
    oportun.sort(key=lambda x: -x[1]["n"])
    if oportun:
        items = "".join(
            f'<li><b>{cli}</b> — {info["n"]} cotización(es) (última {mes_corto(info["ultima"])}), '
            f'sin compra todavía. Oportunidad para seguir.</li>'
            for cli, info in oportun[:10])
        H.append(f'<div class="ops-b op-op"><div class="ops-t">💰 Cotizaste y todavía no te compran</div>'
                 f'<ul>{items}</ul></div>')

    if not cotiza_mucho and not sin_cotizar and not oportun:
        H.append('<p class="vacio">Sin señales para destacar este mes.</p>')
    H.append('</div>')

    H.append('</section>')
    H.append(_PIE)
    return "\n".join(H)


def main():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    rows = con.execute(
        "SELECT empresa, COALESCE(vendedor_final, sender) AS v, "
        "COALESCE(fecha_oc, received) AS f, COALESCE(monto_final, monto) AS m, "
        "COALESCE(fabricacion, 0) AS fab FROM oc_extraccion WHERE estado='aprobado'"
    ).fetchall()
    con.close()

    COTIZ = cargar_cotizaciones()
    COTIZ_VEND = cargar_cotiz_por_vendedor()
    hoy = datetime.date.today()
    mes_actual = hoy.strftime("%Y-%m")
    # Clientes a los que cada vendedor YA le vendió alguna vez (todo el historial, no solo septiembre)
    oc_cli_vend = defaultdict(set)
    for r in rows:
        oc_cli_vend[nombre_vendedor(r["v"])].add(cliente_canonico(r["empresa"]))
    rows = [r for r in rows if str(r["f"] or "")[:7] >= MES_INICIO]   # estricto desde septiembre
    meses = meses_desde(MES_INICIO, max(mes_actual, MES_INICIO))

    V = defaultdict(lambda: {"cli": defaultdict(lambda: {"oc": 0, "total": 0.0, "fab": 0}),
                             "mes": defaultdict(lambda: {"oc": 0, "total": 0.0}),
                             "total": 0.0, "oc": 0})
    for r in rows:
        v = nombre_vendedor(r["v"]); cli = cliente_canonico(r["empresa"])
        ym = str(r["f"] or "")[:7]; monto = r["m"] or 0
        d = V[v]
        c = d["cli"][cli]; c["oc"] += 1; c["total"] += monto; c["fab"] += r["fab"]
        mm = d["mes"][ym]; mm["oc"] += 1; mm["total"] += monto
        d["total"] += monto; d["oc"] += 1

    generado = hoy.strftime("%d/%m/%Y")
    os.makedirs(OUT_DIR, exist_ok=True)
    for v in VENDEDORES_INFORME:
        html = construir_html(v, V.get(v), meses, COTIZ, generado, mes_actual, COTIZ_VEND, oc_cli_vend)
        with open(ruta_de(v), "w", encoding="utf-8") as f:
            f.write(html)
        d = V.get(v)
        print(f"  {v:20} -> {os.path.basename(ruta_de(v))}  "
              f"({d['oc'] if d else 0} OC, USD {d['total'] if d else 0:,.2f})")


_ENCABEZADO = """<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Informe {vendedor} — Ariana</title>
<style>
  body{{margin:0;background:#eef2f6;color:#1a2430;font-family:Segoe UI,Arial,sans-serif;font-size:14px}}
  .wrap{{max-width:820px;margin:0 auto;padding:24px}}
  .head{{background:#0e2233;color:#fff;border-radius:14px;padding:22px 26px;margin-bottom:22px}}
  .head h1{{margin:0 0 4px;font-size:22px}}
  .head .sub{{color:#8fb6cf;font-size:12.5px;font-family:Consolas,monospace}}
  section.vend{{background:#fff;border:1px solid #dbe3ec;border-radius:14px;padding:20px 22px;margin-bottom:20px}}
  h3{{font-size:13px;text-transform:uppercase;letter-spacing:.05em;color:#5c6b7a;margin:22px 0 8px}}
  .kpis{{display:flex;gap:10px;flex-wrap:wrap}}
  .kpi{{flex:1;min-width:130px;background:#f5f8fb;border:1px solid #e2e9f0;border-radius:10px;padding:12px 14px}}
  .kpi.big{{background:#e8f7f4;border-color:#b9e7e0}}
  .kpi .k{{font-size:11px;color:#6b7a89;text-transform:uppercase;letter-spacing:.04em}}
  .kpi .v{{font-size:20px;font-weight:700;color:#0e2233;margin-top:3px}}
  .kpi.big .v{{color:#0f8f80}}
  table.t{{width:100%;border-collapse:collapse;font-size:13px}}
  table.t th,table.t td{{padding:8px 10px;border-bottom:1px solid #eef2f6;text-align:left}}
  table.t th{{font-size:11px;text-transform:uppercase;letter-spacing:.03em;color:#6b7a89;background:#f7fafc}}
  table.t td.n,table.t th.n{{text-align:right;font-variant-numeric:tabular-nums}}
  table.t td.pos{{color:#17b3a3;font-weight:700;width:26px}}
  table.t tr.cero td{{color:#aab6c2}}
  .badge{{font-size:11px;padding:2px 8px;border-radius:20px;font-weight:600}}
  .badge.rec{{background:#e3f4ec;color:#1f8a54}}
  .badge.alto{{background:#fdeede;color:#b5701a}}
  .badge.pun{{background:#eef1f5;color:#67788a}}
  .ops{{margin-top:4px}}
  .ops-intro{{color:#6b7a89;font-size:12.5px;margin:0 0 12px}}
  .ops-b{{border-radius:10px;padding:12px 16px;margin-bottom:12px}}
  .ops-b .ops-t{{font-weight:700;font-size:13.5px;margin-bottom:6px}}
  .ops-b ul{{margin:0;padding-left:18px}}
  .ops-b li{{margin:4px 0;font-size:13px;line-height:1.45}}
  .op-si{{background:#eef8f4;border:1px solid #cbe9e0}}
  .op-si .ops-t{{color:#0f8f80}}
  .op-no{{background:#fbf3ea;border:1px solid #f0ddc4}}
  .op-no .ops-t{{color:#b5701a}}
  .op-op{{background:#eef2fb;border:1px solid #cdd8f0}}
  .op-op .ops-t{{color:#3856b5}}
  .vacio{{color:#8794a3;font-size:13px}}
  .pie{{color:#7c8a99;font-size:11.5px;text-align:center;padding:10px 0 30px;font-family:Consolas,monospace}}
</style></head><body><div class="wrap">
<div class="head">
  <h1>Informe comercial · {vendedor}</h1>
  <div class="sub">Ariana · generado el {generado} · período {desde} → {hasta} · OC validadas</div>
</div>"""

_PIE = ('<div class="pie">Ariana — Acme Filtration Co. · OC validadas (desde septiembre 2026) · '
        'oportunidades cruzadas con el historial de cotizaciones · '
        'clientes agrupados automáticamente (normalización + alias)</div></div></body></html>')


if __name__ == "__main__":
    main()

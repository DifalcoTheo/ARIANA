# -*- coding: utf-8 -*-
"""
Extractor go-forward de la casilla oc@.  Corre sobre data/oc_emails.json
(generado por dump_oc.py) y carga en cerebro.db lo que aun no este.

Campos que recapitula: Vendedor, Cliente, N de OC del cliente, Monto (sin IVA),
Fecha de entrega.  Idempotente por entry_id (Regla 06).

Dos caminos:
  - FORMATO NUEVO (macro):  asunto  OC-<CLIENTE>-<dd/mm/aaaa>  + etiquetas en el
    cuerpo  ->  se parsea SOLO (deterministico, sin IA).
  - MAIL LIBRE (informal):  la extraccion la deja Cowork en
    data/extraccion_manual.json  (dict por entry_id).  Si no esta, se reporta
    como pendiente de extraccion.
"""
import json, os, re, sqlite3, datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "data")
DB = os.path.join(BASE, "cerebro.db")
CASILLA = "orders@example.com"

RE_ASUNTO = re.compile(r"^\s*OC-(.+)-(\d{2}/\d{2}/\d{4})\s*$", re.IGNORECASE)


def parse_monto(s):
    """Normaliza montos en formato AR (1.166,90) o US (3100.80) a float."""
    if not s:
        return None
    m = re.search(r"\d[\d.,]*", s)
    if not m:
        return None
    num = m.group(0)
    if "." in num and "," in num:
        if num.rfind(",") > num.rfind("."):      # la coma es el decimal (AR)
            num = num.replace(".", "").replace(",", ".")
        else:                                      # el punto es el decimal (US)
            num = num.replace(",", "")
    elif "," in num:
        num = num.replace(".", "").replace(",", ".") if len(num.split(",")[-1]) == 2 else num.replace(",", "")
    else:
        if len(num.split(".")) > 2:                # varios puntos = miles
            num = num.replace(".", "")
    try:
        return round(float(num), 2)
    except ValueError:
        return None


def valor_etiqueta(body, etiqueta_regex):
    for linea in body.splitlines():
        m = re.search(etiqueta_regex + r"\s*:?\s*(.*)$", linea, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def parse_formato_nuevo(e):
    """Devuelve un registro si el mail vino en el formato del macro, si no None."""
    asunto = e.get("subject", "") or ""
    m = RE_ASUNTO.match(asunto)
    if not m:
        return None
    body = e.get("body", "") or ""
    if "OC del cliente" not in body or "MONTO" not in body.upper():
        return None  # asunto parece formato pero el cuerpo no -> lo tratamos como libre

    cliente = m.group(1).strip()
    oc_cli = valor_etiqueta(body, r"N.*?de OC del cliente")
    monto_txt = valor_etiqueta(body, r"MONTO SIN IVA")
    entrega = valor_etiqueta(body, r"Fecha de entrega")
    fab_txt = valor_etiqueta(body, r"Fabricaci[oó]n")
    fabricacion = 1 if re.search(r"\bs[ií]\b", fab_txt, re.IGNORECASE) else 0  # SI=fab, NO/vacio=stock

    flags = []
    if not oc_cli:
        flags.append("falta_n_oc_cliente")
    if not parse_monto(monto_txt):
        flags.append("falta_monto")
    if not entrega:
        flags.append("falta_fecha_entrega")

    return {
        "empresa": cliente,
        "oc_numero": oc_cli or None,
        "moneda": "USD",
        "monto": parse_monto(monto_txt),
        "texto_original": monto_txt or None,
        "fuente": "formato",
        "confianza": "alta" if not flags else "media",
        "flags": flags,
        "fecha_entrega": entrega or None,
        "fabricacion": fabricacion,
    }


SCHEMA = """
CREATE TABLE IF NOT EXISTS oc_extraccion (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  entry_id TEXT UNIQUE NOT NULL, casilla TEXT, received TEXT, sender TEXT,
  subject TEXT, empresa TEXT, oc_numero TEXT, moneda TEXT, monto REAL,
  texto_original TEXT, fuente TEXT, confianza TEXT, flags TEXT,
  fecha_entrega TEXT, fabricacion INTEGER DEFAULT 0, estado TEXT DEFAULT 'pendiente',
  monto_final REAL, revisado_por TEXT, revisado_en TEXT, creado_en TEXT
);
"""


def asegurar_columna(con):
    cols = [r[1] for r in con.execute("PRAGMA table_info(oc_extraccion)")]
    if "fecha_entrega" not in cols:
        con.execute("ALTER TABLE oc_extraccion ADD COLUMN fecha_entrega TEXT")
    if "fabricacion" not in cols:
        con.execute("ALTER TABLE oc_extraccion ADD COLUMN fabricacion INTEGER DEFAULT 0")


def main():
    emails = json.load(open(os.path.join(DATA, "oc_emails.json"), encoding="utf-8"))
    manual_path = os.path.join(DATA, "extraccion_manual.json")
    manual = json.load(open(manual_path, encoding="utf-8")) if os.path.exists(manual_path) else {}
    # entry_ids que NO son OC nuevas (duplicados, respuestas RE:, confirmaciones) -> no molestar
    ignorar_path = os.path.join(DATA, "ignorar_extraccion.json")
    ignorar = json.load(open(ignorar_path, encoding="utf-8")) if os.path.exists(ignorar_path) else {}

    con = sqlite3.connect(DB)
    con.execute(SCHEMA)
    asegurar_columna(con)
    en_db = {r[0] for r in con.execute("SELECT entry_id FROM oc_extraccion")}
    ahora = datetime.datetime.now().isoformat(timespec="seconds")

    ins_formato, ins_manual, pendientes = 0, 0, []
    for e in emails:
        eid = e["entry_id"]
        if eid in en_db:
            continue

        reg = parse_formato_nuevo(e)
        origen = "formato-nuevo"
        if reg is None:
            reg = manual.get(eid)
            origen = "manual"
        if reg is None:
            if eid not in ignorar:                 # duplicado/no-OC conocido: se omite
                pendientes.append((e.get("subject"), e.get("sender_email")))
            continue

        con.execute("""
            INSERT INTO oc_extraccion
            (entry_id, casilla, received, sender, subject, empresa, oc_numero,
             moneda, monto, texto_original, fuente, confianza, flags,
             fecha_entrega, fabricacion, estado, creado_en)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'pendiente', ?)
        """, (
            eid, CASILLA, e.get("received"), e.get("sender_email") or e.get("sender_name"),
            e.get("subject"), reg.get("empresa"), reg.get("oc_numero"), reg.get("moneda", "USD"),
            reg.get("monto"), reg.get("texto_original"), reg.get("fuente"),
            reg.get("confianza"), json.dumps(reg.get("flags", []), ensure_ascii=False),
            reg.get("fecha_entrega"), int(reg.get("fabricacion", 0)), ahora,
        ))
        if origen == "manual":
            ins_manual += 1
        else:
            ins_formato += 1

    con.commit()
    pend = con.execute("SELECT COUNT(*) FROM oc_extraccion WHERE estado='pendiente'").fetchone()[0]
    total = con.execute("SELECT COUNT(*) FROM oc_extraccion").fetchone()[0]
    con.close()

    print(f"Insertados formato-nuevo (auto): {ins_formato}")
    print(f"Insertados mail-libre (Cowork) : {ins_manual}")
    print(f"Sin extraer todavia            : {len(pendientes)}")
    for s, rem in pendientes:
        print(f"   - PENDIENTE: {rem} | {s}")
    print(f"\nTotal en base: {total} | Pendientes de validacion: {pend}")


if __name__ == "__main__":
    main()

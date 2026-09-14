# -*- coding: utf-8 -*-
"""
El Cerebro — servidor local de revision/validacion de OC.

Flujo:
  INGESTA (usa Outlook/COM)  -> descarga .msg + adjuntos a  ORDENES DE COMPRA/EN REVISION/<id - empresa>/
  VALIDAR  (sin COM)         -> asigna OC-XXXX y MUEVE la carpeta a  VALIDADAS/OC-XXXX/
  RECHAZAR (sin COM)         -> mueve la carpeta a  RECHAZADAS/
  DESCARGAR                  -> el servidor entrega el .msg (funciona desde otra PC por LAN/VPN)

Solo la INGESTA necesita Outlook abierto. Revisar/validar/descargar son puro disco,
asi que funcionan aunque Outlook este cerrado y desde cualquier PC de la red.
"""
import os, re, json, sqlite3, datetime, shutil, subprocess, sys
from functools import wraps
from flask import Flask, jsonify, request, render_template, send_file, session, redirect, Response
from werkzeug.security import generate_password_hash, check_password_hash

import requests
try:                                 # Outlook COM is Windows-only.
    import pythoncom                 # The web UI, validation queue, reports and demo
    import win32com.client           # all run without it; only the live-mailbox
except ImportError:                  # features (fetch / open in Outlook) need it.
    pythoncom = None
    win32com = None

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "cerebro.db")
OC_DIR = os.path.join(BASE, "ORDENES DE COMPRA")
EN_REVISION = os.path.join(OC_DIR, "EN REVISION")
VALIDADAS = os.path.join(OC_DIR, "VALIDADAS")
RECHAZADAS = os.path.join(OC_DIR, "RECHAZADAS")
MAILBOX = "orders@example.com"

# Carpeta compartida que alimenta a GASPI con las OF en proceso (SPF-XXXX-CLIENTE/)
GASPI_DIR = os.path.join(os.path.dirname(BASE), "GASPI")
EN_PROCESO = os.path.join(GASPI_DIR, "datos", "en_proceso")
OC_INICIAL = 1000

# Mapa email -> nombre del vendedor (editable)
VENDEDORES = {
    "a.rivera@example.com": "Alex Rivera",
    "s.carter@example.com": "Sam Carter",
    "d.torres@example.com": "Diego Torres",
    "l.moretti@example.com": "Luca Moretti",
    "n.becker@example.com": "Nils Becker",
    "m.rossi@example.com": "Marco Rossi",
}
# Nombres visibles que aparecian en los correos viejos -> mismo canonico
ALIAS_VENDEDOR = {
    "sam carter": "Sam Carter", "sam carter": "Sam Carter",
    "nils becker": "Nils Becker", "nils becker": "Nils Becker",
    "luca moretti": "Luca Moretti", "luca moretti": "Luca Moretti",
    "alex rivera": "Alex Rivera", "diego torres": "Diego Torres",
    "marco rossi": "Marco Rossi",
}
# Correos que NO son vendedores y se asignan a un vendedor real.
# Caso Erik Solberg: el macro de Diego Torres deja predeterminado el mail de
# Erik; en realidad la venta es de Diego. Formato: mail -> (mail_real, nombre_para_aviso)
REDIR_VENDEDOR = {
    "e.solberg@example.com": ("d.torres@example.com", "Erik Solberg"),
}


def nombre_vendedor(sender):
    if not sender:
        return "(sin remitente)"
    s = str(sender).strip().lower()
    if s in REDIR_VENDEDOR:          # no es vendedor -> vendedor real
        s = REDIR_VENDEDOR[s][0]
    return VENDEDORES.get(s) or ALIAS_VENDEDOR.get(s) or sender


# ── Resolución de la fecha de entrega para el calendario ──
RE_DIAS = re.compile(r"(\d{1,3})\s*d[ií]a", re.IGNORECASE)          # "45 dias", "10 día"
RE_FECHA = re.compile(r"(\d{1,2})[/\-.](\d{1,2})(?:[/\-.](\d{2,4}))?")  # dd/mm[/aaaa]
DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


def resolver_entrega(received, txt):
    """Devuelve (date|None, [observaciones]) a partir del texto de fecha de entrega."""
    obs = []
    if not txt or not str(txt).strip():
        return None, []   # sin fecha cargada: no es discrepancia, no se anota
    t = str(txt).strip()
    low = t.lower()
    try:
        base = datetime.date.fromisoformat(str(received)[:10])
    except (ValueError, TypeError):
        base = None

    # 1) inmediato / stock
    if "inmediat" in low or "stock" in low:
        # Entrega inmediata: normal, no es discrepancia. Se ubica en la fecha del correo.
        return (base, []) if base else (None, [])

    # 2) fecha absoluta dd/mm[/aaaa]
    m = RE_FECHA.search(t)
    if m:
        d, mo = int(m.group(1)), int(m.group(2))
        y = m.group(3)
        if y:
            y = int(y); y = y + 2000 if y < 100 else y
        else:
            y = base.year if base else datetime.date.today().year
        try:
            return datetime.date(y, mo, d), obs
        except ValueError:
            return None, [f"Fecha de entrega inválida: '{t}'."]

    # 3) plazo relativo en días
    m = RE_DIAS.search(t)
    if m:
        if not base:
            return None, [f"Plazo '{t}' pero sin fecha base del correo."]
        return base + datetime.timedelta(days=int(m.group(1))), obs

    return None, [f"No se pudo interpretar la fecha de entrega: '{t}'."]


def observaciones_fecha(fecha, hoy):
    """Chequeos deterministicos sobre una fecha ya resuelta."""
    out = []
    if fecha.weekday() >= 5:  # 5=sáb, 6=dom
        out.append(f"La entrega cae {DIAS_ES[fecha.weekday()]} ({fecha.strftime('%d/%m')}).")
    if fecha < hoy:
        out.append(f"La fecha de entrega ya pasó ({fecha.strftime('%d/%m/%Y')}).")
    return out


app = Flask(__name__)

# ─────────────────────────── Autenticación ───────────────────────────
# Clave de sesión persistente (se genera una vez).
SECRET_FILE = os.path.join(BASE, "secret.key")
if not os.path.exists(SECRET_FILE):
    with open(SECRET_FILE, "wb") as f:
        f.write(os.urandom(32))
with open(SECRET_FILE, "rb") as f:
    app.secret_key = f.read()

USUARIOS_FILE = os.path.join(BASE, "usuarios.json")
# users iniciales (CAMBIAR las contraseñas). Solo se usan si no existe usuarios.json.
_DEFAULT_USERS = {
    "admin":    {"rol": "administracion", "nombre": "Administración", "pass": "changeme"},
    "usuario":  {"rol": "usuarios", "nombre": "user", "pass": "changeme"},
}

# Qué pestañas ve cada rol (el Panel principal lo ven todos)
VISTAS_ROL = {
    "administracion": ["panel", "calendario", "oc", "fabricacion", "comercial", "clientes"],
    "usuarios":       ["panel", "calendario"],
}
DEFAULT_VISTA = {"administracion": "panel", "usuarios": "panel"}


def cargar_usuarios():
    if not os.path.exists(USUARIOS_FILE):
        data = {u: {"rol": v["rol"], "nombre": v["nombre"], "hash": generate_password_hash(v["pass"])}
                for u, v in _DEFAULT_USERS.items()}
        with open(USUARIOS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    with open(USUARIOS_FILE, encoding="utf-8") as f:
        return json.load(f)


def usuario_actual():
    return session.get("usuario")


def login_required(f):
    @wraps(f)
    def wrap(*a, **k):
        if not session.get("usuario"):
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "error": "no autenticado"}), 401
            return redirect("/login")
        return f(*a, **k)
    return wrap


def requiere_vista(*vistas):
    """Protege un endpoint: el rol debe tener acceso a ALGUNA de las vistas dadas."""
    def deco(f):
        @wraps(f)
        def wrap(*a, **k):
            rol = session.get("rol")
            if not rol:
                return jsonify({"ok": False, "error": "no autenticado"}), 401
            permitidas = VISTAS_ROL.get(rol, [])
            if not any(v in permitidas for v in vistas):
                return jsonify({"ok": False, "error": "sin permiso"}), 403
            return f(*a, **k)
        return wrap
    return deco
app.config["TEMPLATES_AUTO_RELOAD"] = True  # tomar cambios de index.html sin reiniciar


# ───────────────────────── Base de datos ─────────────────────────
def get_db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = get_db()
    cols = [r["name"] for r in con.execute("PRAGMA table_info(oc_extraccion)")]
    for c, ddl in [("oc_codigo", "TEXT"), ("carpeta", "TEXT"), ("descargado", "INTEGER DEFAULT 0"),
                   ("entrega_terminada", "INTEGER DEFAULT 0"), ("entrega_terminada_en", "TEXT"),
                   ("revisado_por", "TEXT"), ("fabricacion", "INTEGER DEFAULT 0"),
                   ("fab_estado", "TEXT"), ("fab_por", "TEXT"), ("fab_en", "TEXT"),
                   ("spf", "TEXT"), ("vendedor_final", "TEXT"), ("fecha_oc", "TEXT")]:
        if c not in cols:
            con.execute(f"ALTER TABLE oc_extraccion ADD COLUMN {c} {ddl}")
    con.execute("""
        CREATE TABLE IF NOT EXISTS calendario_notas (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          fecha TEXT NOT NULL,        -- YYYY-MM-DD
          texto TEXT NOT NULL,
          creado_en TEXT
        )
    """)
    con.commit(); con.close()
    for d in (OC_DIR, EN_REVISION, VALIDADAS, RECHAZADAS, EN_PROCESO):
        os.makedirs(d, exist_ok=True)


def proximo_num(con):
    row = con.execute(
        "SELECT MAX(CAST(SUBSTR(oc_codigo,4) AS INTEGER)) AS m "
        "FROM oc_extraccion WHERE oc_codigo IS NOT NULL"
    ).fetchone()
    return (row["m"] + 1) if row and row["m"] is not None else OC_INICIAL


def sanitizar(nombre):
    nombre = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", str(nombre)).strip()
    return nombre[:80] or "sin_nombre"


# ───────────────────────── Outlook (COM) ─────────────────────────
def _get_item(ns, entry_id):
    sid = None
    for i in range(1, ns.Folders.Count + 1):
        f = ns.Folders.Item(i)
        if MAILBOX.lower() in f.Name.lower():
            sid = f.Store.StoreID
            break
    return ns.GetItemFromID(entry_id, sid) if sid else ns.GetItemFromID(entry_id)


def _descargar_item(item, carpeta):
    os.makedirs(carpeta, exist_ok=True)
    item.SaveAs(os.path.join(carpeta, "mail.msg"), 3)  # olMSG = 3
    adj = []
    for a in range(1, item.Attachments.Count + 1):
        att = item.Attachments.Item(a)
        nombre = sanitizar(att.FileName)
        dest = os.path.join(carpeta, nombre)
        if os.path.exists(dest):
            b, e = os.path.splitext(nombre); dest = os.path.join(carpeta, f"{b}_{a}{e}")
        att.SaveAsFile(dest); adj.append(os.path.basename(dest))
    with open(os.path.join(carpeta, "resumen.txt"), "w", encoding="utf-8") as f:
        f.write(f"Asunto : {item.Subject}\nDe     : {item.SenderName}\n")
        f.write(f"Fecha  : {item.ReceivedTime}\nEntryID: {item.EntryID}\n")
        f.write(f"Adjuntos ({len(adj)}): " + ", ".join(adj) + "\n")
    return adj


def ingestar_pendientes():
    """Descarga a EN REVISION todas las OC pendientes que aun no esten en disco."""
    con = get_db()
    filas = con.execute(
        "SELECT * FROM oc_extraccion WHERE estado='pendiente' AND COALESCE(descargado,0)=0"
    ).fetchall()
    if not filas:
        con.close(); return {"descargadas": 0, "detalle": []}

    pythoncom.CoInitialize()
    detalle = []
    try:
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for r in filas:
            carpeta = os.path.join(EN_REVISION, f"{r['id']:04d} - {sanitizar(r['empresa'])}")
            try:
                item = _get_item(ns, r["entry_id"])
                adj = _descargar_item(item, carpeta)
                con.execute("UPDATE oc_extraccion SET carpeta=?, descargado=1 WHERE id=?", (carpeta, r["id"]))
                detalle.append({"id": r["id"], "empresa": r["empresa"], "adjuntos": len(adj)})
            except Exception as e:
                detalle.append({"id": r["id"], "empresa": r["empresa"], "error": str(e)})
        con.commit()
    finally:
        pythoncom.CoUninitialize()
        con.close()
    return {"descargadas": len([d for d in detalle if "error" not in d]), "detalle": detalle}


def abrir_en_outlook(entry_id):
    pythoncom.CoInitialize()
    try:
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        item = _get_item(ns, entry_id); item.Display(); return item.Subject
    finally:
        pythoncom.CoUninitialize()


# ───────────────────────────── API ─────────────────────────────
def fila_dict(r):
    d = {k: r[k] for k in r.keys()}
    d["flags"] = json.loads(r["flags"] or "[]")
    d["descargado"] = bool(r["descargado"])
    d["vendedor"] = nombre_vendedor(d.get("vendedor_final") or r["sender"])
    # Fecha de la OC: la editable; por defecto el día en que se recibió el correo.
    d["fecha_oc"] = r["fecha_oc"] or (str(r["received"])[:10] if r["received"] else None)
    # Aviso: correo que no es vendedor y quedó reasignado (ej. Erik -> Diego)
    raw = str(r["sender"] or "").strip().lower()
    d["aviso_remitente"] = REDIR_VENDEDOR[raw][1] if raw in REDIR_VENDEDOR else None
    return d


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        data = request.get_json(silent=True) or request.form
        u = (data.get("usuario") or "").strip().lower()
        p = data.get("clave") or ""
        usuarios = cargar_usuarios()
        info = usuarios.get(u)
        if info and check_password_hash(info["hash"], p):
            session["usuario"] = u
            session["rol"] = info["rol"]
            session["nombre"] = info["nombre"]
            return jsonify({"ok": True})
        return jsonify({"ok": False, "error": "user o contraseña incorrectos"}), 401
    if session.get("usuario"):
        return redirect("/")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


@app.route("/api/me")
@login_required
def me():
    rol = session.get("rol")
    return jsonify({
        "usuario": session.get("usuario"), "nombre": session.get("nombre"), "rol": rol,
        "vistas": VISTAS_ROL.get(rol, []), "inicio": DEFAULT_VISTA.get(rol, "calendario"),
    })


@app.route("/")
@login_required
def index():
    return render_template("index.html")


# ── Puerta única: Ariana hace de proxy hacia GASPI (que corre en 8000) ──
GASPI_URL = "http://127.0.0.1:8000"
_HDRS_EXCLUIR = {"content-encoding", "transfer-encoding", "content-length", "connection", "host"}


def _gaspi_auth():
    """Firma con el secret de GASPI (mismo disco) para que acepte pedidos de Ariana."""
    try:
        import hmac as _h, hashlib as _hl
        with open(os.path.join(GASPI_DIR, "datos", "secret.key"), "rb") as f:
            sec = f.read()
        return {"X-Gaspi-Auth": _h.new(sec, b"ariana", _hl.sha256).hexdigest()}
    except OSError:
        return {}


@app.route("/gaspi", defaults={"ruta": ""}, methods=["GET", "POST", "PUT", "DELETE"])
@app.route("/gaspi/", defaults={"ruta": ""}, methods=["GET", "POST", "PUT", "DELETE"])
@app.route("/gaspi/<path:ruta>", methods=["GET", "POST", "PUT", "DELETE"])
def gaspi_proxy(ruta):
    # SIN @login_required a proposito: GASPI tiene su PROPIO login (ingenieria/
    # produccion/compras). Un usuario de GASPI no tiene cuenta de Ariana, asi que
    # el proxy es publico del lado de Ariana y GASPI gestiona su autenticacion
    # (portada publica; sus /api/* siguen protegidos por la cookie gaspi_sesion).
    # El acceso ya esta acotado a la subred del VPN por firewall.
    destino = f"{GASPI_URL}/{ruta}"
    try:
        up = requests.request(
            method=request.method, url=destino, params=request.args,
            data=request.get_data(),
            headers={k: v for k, v in request.headers if k.lower() != "host"},
            cookies=request.cookies, allow_redirects=False, timeout=120,
        )
    except requests.RequestException:
        html = (
            "<div style=\"font-family:system-ui,Segoe UI,sans-serif;max-width:520px;"
            "margin:90px auto;padding:30px;text-align:center;color:#2b3a44;"
            "background:#f4f7f8;border:1px solid #dbe3e6;border-radius:14px;\">"
            "<div style=\"font-size:34px;margin-bottom:6px;\">🛠️</div>"
            "<h2 style=\"margin:0 0 8px;\">GASPI no está disponible</h2>"
            "<p style=\"margin:0 0 6px;\">El sistema de fabricación no está respondiendo en este momento.</p>"
            "<p style=\"color:#667a86;font-size:14px;margin:14px 0 0;\">Si sigue, reiniciá con "
            "<b>INICIAR ARIANA.bat</b> (en el Escritorio de la PC servidor).</p>"
            "<p style=\"margin-top:18px;\"><a href=\"/gaspi/\" style=\"color:#0e7c8b;\">Reintentar</a> · "
            "<a href=\"/\" style=\"color:#0e7c8b;\">Volver a Ariana</a></p></div>"
        )
        return Response(html, status=503, content_type="text/html; charset=utf-8")

    ctype = up.headers.get("Content-Type", "")
    if "text/html" in ctype:
        # Reescribe los llamados de GASPI para que salgan por /gaspi/ (sin tocar su código)
        cuerpo = up.content.replace(b"/api/", b"/gaspi/api/").replace(b'"/web/', b'"/gaspi/web/')
        resp = Response(cuerpo, status=up.status_code, content_type=ctype)
        for k, v in up.headers.items():
            if k.lower() == "set-cookie":
                resp.headers.add("Set-Cookie", v)
        return resp
    cabeceras = [(k, v) for k, v in up.headers.items() if k.lower() not in _HDRS_EXCLUIR]
    return Response(up.content, status=up.status_code, headers=cabeceras)


@app.route("/api/resumen")
@requiere_vista("panel", "oc")
def resumen():
    con = get_db()
    def c(q): return con.execute(q).fetchone()[0]
    pend = c("SELECT COUNT(*) FROM oc_extraccion WHERE estado='pendiente'")
    sinDesc = c("SELECT COUNT(*) FROM oc_extraccion WHERE estado='pendiente' AND COALESCE(descargado,0)=0")
    apro = c("SELECT COUNT(*) FROM oc_extraccion WHERE estado='aprobado'")
    rech = c("SELECT COUNT(*) FROM oc_extraccion WHERE estado='rechazado'")
    total = c("SELECT COALESCE(SUM(COALESCE(monto_final,monto)),0) FROM oc_extraccion WHERE estado='aprobado'")
    prox = proximo_num(con)

    # Próximas entregas: OC validadas, con fecha futura y NO terminadas
    hoy = datetime.date.today()
    ocs = con.execute(
        "SELECT oc_codigo, empresa, COALESCE(vendedor_final, sender) AS sender, received, fecha_entrega "
        "FROM oc_extraccion WHERE estado='aprobado' AND COALESCE(entrega_terminada,0)=0"
    ).fetchall()
    con.close()
    proximas = []
    for o in ocs:
        fecha, _ = resolver_entrega(o["received"], o["fecha_entrega"])
        if fecha and fecha >= hoy:
            proximas.append({
                "oc": o["oc_codigo"], "cliente": o["empresa"],
                "vendedor": nombre_vendedor(o["sender"]),
                "fecha": fecha.isoformat(), "dias": (fecha - hoy).days,
            })
    proximas.sort(key=lambda x: x["fecha"])

    return jsonify({"pendientes": pend, "sin_descargar": sinDesc, "aprobadas": apro,
                    "rechazadas": rech, "total_usd_aprobado": round(total, 2),
                    "proximo_codigo": f"OC-{prox}", "proximas_entregas": proximas[:6]})


@app.route("/api/vendedores")
@requiere_vista("oc")
def vendedores_lista():
    return jsonify(VENDEDORES)


@app.route("/api/oc")
@requiere_vista("oc")
def listar():
    estado = request.args.get("estado", "pendiente")
    con = get_db()
    rows = con.execute("SELECT * FROM oc_extraccion WHERE estado=? ORDER BY received DESC", (estado,)).fetchall()
    con.close()
    return jsonify([fila_dict(r) for r in rows])


@app.route("/api/fabricacion")
@requiere_vista("fabricacion")
def fabricacion_listar():
    """OC validadas con Fabricacion=SI, para que Ingenieria decida aplica/no aplica."""
    estado = request.args.get("estado", "pendiente")  # pendiente | aplica | no_aplica
    con = get_db()
    base = ("FROM oc_extraccion WHERE estado='aprobado' AND COALESCE(fabricacion,0)=1 "
            "AND COALESCE(fab_estado,'pendiente')=?")
    rows = con.execute(f"SELECT * {base} ORDER BY received DESC", (estado,)).fetchall()
    conteos = {e: con.execute(f"SELECT COUNT(*) {base}", (e,)).fetchone()[0]
               for e in ("pendiente", "aplica", "no_aplica")}
    con.close()
    return jsonify({"conteos": conteos, "items": [fila_dict(r) for r in rows]})


def _proximo_spf():
    """Proximo N de SPF: base = GASPI (max lote + 1), reconciliado con la carpeta en proceso."""
    base_num = 0
    try:
        r = requests.get(f"{GASPI_URL}/api/of/proximo", headers=_gaspi_auth(), timeout=10).json()
        m = re.match(r"SPF-(\d+)", str(r.get("lote", "")), re.I)
        if m:
            base_num = int(m.group(1))
    except (requests.RequestException, ValueError):
        base_num = 0
    folder_max = 0
    if os.path.isdir(EN_PROCESO):
        for nombre in os.listdir(EN_PROCESO):
            m = re.match(r"SPF-(\d+)", nombre, re.I)
            if m:
                folder_max = max(folder_max, int(m.group(1)))
    num = max(base_num, folder_max + 1)
    return f"SPF-{num}" if num > 0 else None


def _quitar_carpeta_spf(spf, cliente):
    if not spf:
        return
    carpeta = os.path.join(EN_PROCESO, f"{spf}-{sanitizar(cliente)}")
    if os.path.isdir(carpeta):
        shutil.rmtree(carpeta, ignore_errors=True)


@app.route("/api/oc/<int:oid>/fabricacion", methods=["POST"])
@requiere_vista("fabricacion")
def fabricacion_decidir(oid):
    data = request.get_json(silent=True) or {}
    decision = data.get("decision")
    if decision not in ("aplica", "no_aplica", "pendiente"):
        return jsonify({"ok": False, "error": "decision invalida"}), 400
    ahora = datetime.datetime.now().isoformat(timespec="seconds")
    con = get_db()
    r = con.execute("SELECT empresa, oc_numero, COALESCE(vendedor_final, sender) AS sender, "
                    "fecha_entrega, oc_codigo, spf "
                    "FROM oc_extraccion WHERE id=?", (oid,)).fetchone()
    if not r:
        con.close(); return jsonify({"ok": False, "error": "no existe"}), 404
    cliente = r["empresa"] or "SIN-CLIENTE"
    spf = None

    if decision == "aplica":
        spf = _proximo_spf()
        if not spf:
            con.close()
            return jsonify({"ok": False, "error": "No pude determinar el N° de SPF (¿GASPI corriendo?)."}), 500
        carpeta = os.path.join(EN_PROCESO, f"{spf}-{sanitizar(cliente)}")
        os.makedirs(carpeta, exist_ok=True)
        datos = {
            "spf": spf, "cliente": cliente, "oc_numero": r["oc_numero"],
            "vendedor": nombre_vendedor(r["sender"]), "fecha_entrega": r["fecha_entrega"],
            "origen": "Ariana", "oc_ariana": r["oc_codigo"], "ariana_oc_id": oid, "creado": ahora,
        }
        with open(os.path.join(carpeta, "datos.json"), "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    else:
        # revertir (pendiente / no aplica): sacar la carpeta si existia
        _quitar_carpeta_spf(r["spf"], cliente)

    con.execute("UPDATE oc_extraccion SET fab_estado=?, fab_por=?, fab_en=?, spf=? WHERE id=?",
                (decision, usuario_actual(), ahora, spf, oid))
    con.commit(); con.close()
    return jsonify({"ok": True, "decision": decision, "spf": spf})


@app.route("/api/comercial")
@requiere_vista("comercial")
def comercial():
    """Analitica comercial: se alimenta SOLO de las OC validadas (estado='aprobado')."""
    periodo = request.args.get("periodo", "mes")
    # Selección puntual opcional (histórico): año, mes (1-12) y/o trimestre (1-4).
    # Si no vienen, el período es relativo a hoy (comportamiento de los botones rápidos).
    p_anio = request.args.get("anio", type=int)
    p_mes = request.args.get("mes", type=int)
    p_q = request.args.get("q", type=int)
    hoy = datetime.date.today()
    con = get_db()
    rows = con.execute(
        "SELECT COALESCE(vendedor_final, sender) AS sender, "
        "COALESCE(fecha_oc, received) AS fecha_ef, COALESCE(monto_final, monto) AS m "
        "FROM oc_extraccion WHERE estado='aprobado'"
    ).fetchall()
    con.close()

    def en_periodo(recibido):
        try:
            d = datetime.date.fromisoformat(str(recibido)[:10])
        except (ValueError, TypeError):
            return False
        if periodo == "hoy":
            return d == hoy
        if periodo == "mes":
            yy, mm = (p_anio or hoy.year), (p_mes or hoy.month)
            return d.year == yy and d.month == mm
        if periodo == "trimestre":
            yy = p_anio or hoy.year
            qq = p_q if p_q else (hoy.month - 1) // 3 + 1
            return d.year == yy and (d.month - 1) // 3 == qq - 1
        if periodo == "anio":
            return d.year == (p_anio or hoy.year)
        return True  # todo

    # Meses con OC validadas (para armar los selectores de mes/trimestre), más nuevo primero
    meses_disp = sorted({str(r["fecha_ef"])[:7] for r in rows if r["fecha_ef"]}, reverse=True)

    datos = [r for r in rows if en_periodo(r["fecha_ef"])]
    total = sum((r["m"] or 0) for r in datos)
    cant = len(datos)
    ticket = (total / cant) if cant else 0

    agg = {}
    for r in datos:
        v = nombre_vendedor(r["sender"])
        a = agg.setdefault(v, {"vendedor": v, "total": 0.0, "cantidad": 0})
        a["total"] += (r["m"] or 0)
        a["cantidad"] += 1
    ranking = sorted(agg.values(), key=lambda x: x["total"], reverse=True)
    for a in ranking:
        a["ticket"] = round(a["total"] / a["cantidad"], 2) if a["cantidad"] else 0
        a["total"] = round(a["total"], 2)

    return jsonify({
        "periodo": periodo, "cantidad": cant, "total_usd": round(total, 2),
        "ticket_promedio": round(ticket, 2), "ranking": ranking,
        "meses_disponibles": meses_disp,
    })


@app.route("/api/clientes_resumen")
@requiere_vista("clientes")
def clientes_resumen():
    """Ficha comercial por cliente: se alimenta de las OC validadas + notas de memoria."""
    con = get_db()
    rows = con.execute(
        "SELECT empresa, COALESCE(monto_final,monto) m, received, fecha_entrega, "
        "COALESCE(fabricacion,0) fab FROM oc_extraccion WHERE estado='aprobado'"
    ).fetchall()
    con.close()
    agg = {}
    for r in rows:
        k = (r["empresa"] or "(sin cliente)").strip().upper()
        a = agg.setdefault(k, {"cliente": k, "oc": 0, "total": 0.0, "fab": 0,
                               "ultima": "", "entregas": {}})
        a["oc"] += 1
        a["total"] += r["m"] or 0
        a["fab"] += r["fab"]
        f = str(r["received"] or "")[:10]
        if f > a["ultima"]:
            a["ultima"] = f
        e = (r["fecha_entrega"] or "").strip().upper()
        if e:
            a["entregas"][e] = a["entregas"].get(e, 0) + 1
    out = []
    for a in agg.values():
        ent = sorted(a["entregas"].items(), key=lambda x: -x[1])[:2]
        out.append({"cliente": a["cliente"], "oc": a["oc"], "total": round(a["total"], 2),
                    "ticket": round(a["total"] / a["oc"], 2) if a["oc"] else 0,
                    "fab": a["fab"], "ultima": a["ultima"],
                    "entregas": " · ".join(f"{e} ({n})" for e, n in ent) or "—"})
    out.sort(key=lambda x: -x["total"])
    notas = ""
    np_ = os.path.join(BASE, "memoria", "clientes_notas.md")
    if os.path.exists(np_):
        with open(np_, encoding="utf-8") as f:
            notas = f.read()
    return jsonify({"clientes": out, "notas": notas})


@app.route("/api/calendario")
@requiere_vista("calendario")
def calendario():
    hoy = datetime.date.today()
    anio = int(request.args.get("anio") or hoy.year)
    mes = int(request.args.get("mes") or hoy.month)
    con = get_db()
    ocs = con.execute(
        "SELECT id, oc_codigo, empresa, COALESCE(vendedor_final, sender) AS sender, oc_numero, "
        "COALESCE(monto_final,monto) AS m, "
        "received, fecha_entrega, COALESCE(entrega_terminada,0) AS terminada "
        "FROM oc_extraccion WHERE estado='aprobado'"
    ).fetchall()
    notas = con.execute(
        "SELECT id, fecha, texto FROM calendario_notas WHERE substr(fecha,1,7)=? ORDER BY fecha",
        (f"{anio:04d}-{mes:02d}",)
    ).fetchall()
    con.close()

    eventos, observaciones = [], []

    def anotar(o, fecha_iso, texto):
        observaciones.append({
            "oc": o["oc_codigo"], "cliente": o["empresa"],
            "vendedor": nombre_vendedor(o["sender"]), "fecha": fecha_iso, "texto": texto,
        })

    for o in ocs:
        fecha, obs = resolver_entrega(o["received"], o["fecha_entrega"])
        if fecha is None:
            try:
                rec = datetime.date.fromisoformat(str(o["received"])[:10])
            except (ValueError, TypeError):
                rec = None
            if rec and rec.year == anio and rec.month == mes:
                for t in obs:
                    anotar(o, None, t)
            continue
        if fecha.year == anio and fecha.month == mes:
            terminada = bool(o["terminada"])
            eventos.append({
                "id": o["id"], "fecha": fecha.isoformat(), "dia": fecha.day, "oc": o["oc_codigo"],
                "cliente": o["empresa"], "vendedor": nombre_vendedor(o["sender"]),
                "oc_numero": o["oc_numero"], "monto": o["m"],
                "fecha_entrega_txt": o["fecha_entrega"], "terminada": terminada,
            })
            if not terminada:   # una entrega terminada ya no es discrepancia
                txt = (o["fecha_entrega"] or "").lower()
                inmediata = ("inmediat" in txt or "stock" in txt)
                checks = obs if inmediata else obs + observaciones_fecha(fecha, hoy)
                for t in checks:
                    anotar(o, fecha.isoformat(), t)

    eventos.sort(key=lambda e: e["fecha"])
    return jsonify({
        "anio": anio, "mes": mes, "hoy": hoy.isoformat(), "eventos": eventos,
        "notas": [{"id": n["id"], "fecha": n["fecha"], "texto": n["texto"]} for n in notas],
        "observaciones": observaciones,
    })


@app.route("/api/calendario/nota", methods=["POST"])
@requiere_vista("calendario")
def calendario_nota():
    data = request.get_json(silent=True) or {}
    fecha = (data.get("fecha") or "").strip()
    texto = (data.get("texto") or "").strip()
    if not fecha or not texto:
        return jsonify({"ok": False, "error": "faltan datos"}), 400
    ahora = datetime.datetime.now().isoformat(timespec="seconds")
    con = get_db()
    cur = con.execute("INSERT INTO calendario_notas (fecha, texto, creado_en) VALUES (?,?,?)",
                      (fecha, texto, ahora))
    con.commit(); nid = cur.lastrowid; con.close()
    return jsonify({"ok": True, "id": nid})


@app.route("/api/calendario/nota/<int:nid>/eliminar", methods=["POST"])
@requiere_vista("calendario")
def calendario_nota_eliminar(nid):
    con = get_db()
    con.execute("DELETE FROM calendario_notas WHERE id=?", (nid,))
    con.commit(); con.close()
    return jsonify({"ok": True})


@app.route("/api/oc/<int:oid>/entrega", methods=["POST"])
@requiere_vista("calendario")
def marcar_entrega(oid):
    data = request.get_json(silent=True) or {}
    terminada = 1 if data.get("terminada", True) else 0
    cuando = datetime.datetime.now().isoformat(timespec="seconds") if terminada else None
    con = get_db()
    con.execute("UPDATE oc_extraccion SET entrega_terminada=?, entrega_terminada_en=? WHERE id=?",
                (terminada, cuando, oid))
    con.commit(); con.close()
    return jsonify({"ok": True, "terminada": bool(terminada)})


@app.route("/api/actualizar", methods=["POST"])
@requiere_vista("oc")
def actualizar_casilla():
    """Paso 1: lee la casilla (dump por COM) y extrae/carga las OC nuevas.
    Corre en el SERVIDOR (esta PC, que tiene Outlook), no en el navegador del que lo toca."""
    scripts = os.path.join(BASE, "scripts")
    py = sys.executable
    try:
        d = subprocess.run([py, os.path.join(scripts, "dump_oc.py")], cwd=BASE,
                           capture_output=True, text=True, timeout=180)
        e = subprocess.run([py, os.path.join(scripts, "extractor.py")], cwd=BASE,
                           capture_output=True, text=True, timeout=180)
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "Tardó demasiado leyendo la casilla."}), 500

    salida = (e.stdout or "") + (e.stderr or "")

    def _n(patron):
        m = re.search(patron, salida)
        return int(m.group(1)) if m else 0

    if d.returncode != 0 or e.returncode != 0:
        err = (d.stderr or e.stderr or "")[-300:]
        return jsonify({"ok": False, "error": f"Error leyendo la casilla (¿Outlook abierto?). {err}"}), 500

    nuevas = _n(r"formato-nuevo \(auto\):\s*(\d+)") + _n(r"mail-libre \(Cowork\)\s*:\s*(\d+)")
    sin_extraer = _n(r"Sin extraer todavia\s*:\s*(\d+)")
    return jsonify({"ok": True, "nuevas": nuevas, "sin_extraer": sin_extraer})


@app.route("/api/ingestar", methods=["POST"])
@requiere_vista("oc")
def ingestar():
    try:
        return jsonify({"ok": True, **ingestar_pendientes()})
    except Exception as e:
        return jsonify({"ok": False, "error": f"Outlook/ingesta: {e}"}), 500


@app.route("/api/oc/<int:oid>/validar", methods=["POST"])
@requiere_vista("oc")
def validar(oid):
    data = request.get_json(silent=True) or {}
    con = get_db()
    r = con.execute("SELECT * FROM oc_extraccion WHERE id=?", (oid,)).fetchone()
    if not r:
        con.close(); return jsonify({"ok": False, "error": "no existe"}), 404
    if r["estado"] == "aprobado":
        con.close(); return jsonify({"ok": False, "error": "ya aprobada", "oc_codigo": r["oc_codigo"]}), 409
    try:
        monto_final = float(data.get("monto_final", r["monto"]))
    except (TypeError, ValueError):
        con.close(); return jsonify({"ok": False, "error": "monto invalido"}), 400

    # Asegurar que la carpeta de revision exista (si no se ingesto antes, descargar ahora)
    origen = r["carpeta"]
    if not (origen and os.path.isdir(origen)):
        try:
            origen = os.path.join(EN_REVISION, f"{r['id']:04d} - {sanitizar(r['empresa'])}")
            pythoncom.CoInitialize()
            ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            _descargar_item(_get_item(ns, r["entry_id"]), origen)
            pythoncom.CoUninitialize()
        except Exception as e:
            con.close(); return jsonify({"ok": False, "error": f"no se pudo descargar: {e}"}), 500

    codigo = f"OC-{proximo_num(con)}"
    destino = os.path.join(VALIDADAS, codigo)
    try:
        if os.path.isdir(destino): shutil.rmtree(destino)
        shutil.move(origen, destino)
    except Exception as e:
        con.close(); return jsonify({"ok": False, "error": f"mover carpeta: {e}"}), 500

    ahora = datetime.datetime.now().isoformat(timespec="seconds")
    con.execute("UPDATE oc_extraccion SET estado='aprobado', oc_codigo=?, monto_final=?, "
                "carpeta=?, descargado=1, revisado_en=?, revisado_por=? WHERE id=?",
                (codigo, monto_final, destino, ahora, usuario_actual(), oid))
    con.commit(); con.close()
    return jsonify({"ok": True, "oc_codigo": codigo, "carpeta": destino, "monto_final": monto_final})


@app.route("/api/oc/<int:oid>/rechazar", methods=["POST"])
@requiere_vista("oc")
def rechazar(oid):
    con = get_db()
    r = con.execute("SELECT * FROM oc_extraccion WHERE id=?", (oid,)).fetchone()
    if not r:
        con.close(); return jsonify({"ok": False, "error": "no existe"}), 404
    if r["carpeta"] and os.path.isdir(r["carpeta"]):
        destino = os.path.join(RECHAZADAS, os.path.basename(r["carpeta"]))
        try:
            if os.path.isdir(destino): shutil.rmtree(destino)
            shutil.move(r["carpeta"], destino)
        except Exception:
            destino = r["carpeta"]
    else:
        destino = None
    ahora = datetime.datetime.now().isoformat(timespec="seconds")
    con.execute("UPDATE oc_extraccion SET estado='rechazado', carpeta=?, revisado_en=?, revisado_por=? WHERE id=?",
                (destino, ahora, usuario_actual(), oid))
    con.commit(); con.close()
    return jsonify({"ok": True})


# Campos que un humano puede corregir en una OC ya validada
EDITABLES = ("empresa", "oc_numero", "fecha_entrega", "moneda", "vendedor_final", "fecha_oc")


@app.route("/api/oc/<int:oid>/editar", methods=["POST"])
@requiere_vista("oc")
def editar(oid):
    data = request.get_json(silent=True) or {}
    campos = {k: data[k] for k in EDITABLES if k in data}
    if "fecha_oc" in campos:
        fv = str(campos["fecha_oc"] or "").strip()
        if fv:
            try:
                datetime.date.fromisoformat(fv)   # exige YYYY-MM-DD válido
            except ValueError:
                return jsonify({"ok": False, "error": "fecha de OC inválida (use AAAA-MM-DD)"}), 400
            campos["fecha_oc"] = fv
        else:
            campos["fecha_oc"] = None   # vaciar = vuelve a la fecha de recepción
    if "monto_final" in data:
        try:
            campos["monto_final"] = float(str(data["monto_final"]).replace(",", "."))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "monto invalido"}), 400
    if not campos:
        return jsonify({"ok": False, "error": "nada para actualizar"}), 400
    campos["revisado_en"] = datetime.datetime.now().isoformat(timespec="seconds")
    campos["revisado_por"] = usuario_actual()
    sets = ", ".join(f"{k}=?" for k in campos)  # claves de lista blanca EDITABLES
    con = get_db()
    con.execute(f"UPDATE oc_extraccion SET {sets} WHERE id=?", list(campos.values()) + [oid])
    con.commit(); con.close()
    return jsonify({"ok": True})


@app.route("/api/oc/<int:oid>/eliminar", methods=["POST"])
@requiere_vista("oc")
def eliminar(oid):
    # Borrado logico: la fila queda con estado 'eliminado' (trazabilidad, Regla 04).
    # Sale de todas las vistas y del ranking Comercial; no se pierde el historial.
    ahora = datetime.datetime.now().isoformat(timespec="seconds")
    con = get_db()
    con.execute("UPDATE oc_extraccion SET estado='eliminado', revisado_en=?, revisado_por=? WHERE id=?",
                (ahora, usuario_actual(), oid))
    con.commit(); con.close()
    return jsonify({"ok": True})


@app.route("/api/oc/<int:oid>/descargar")
@requiere_vista("oc")
def descargar(oid):
    con = get_db()
    r = con.execute("SELECT * FROM oc_extraccion WHERE id=?", (oid,)).fetchone()
    con.close()
    if not r:
        return jsonify({"ok": False, "error": "no existe"}), 404
    msg = os.path.join(r["carpeta"] or "", "mail.msg")
    if not os.path.isfile(msg):
        return jsonify({"ok": False, "error": "aun no descargado (ingesta pendiente)"}), 404
    nombre = f"{r['oc_codigo'] or r['empresa']} - mail.msg"
    return send_file(msg, as_attachment=True, download_name=sanitizar(nombre) + ".msg")


@app.route("/api/oc/<int:oid>/abrir", methods=["POST"])
@requiere_vista("oc")
def abrir(oid):
    con = get_db()
    r = con.execute("SELECT entry_id FROM oc_extraccion WHERE id=?", (oid,)).fetchone()
    con.close()
    if not r:
        return jsonify({"ok": False, "error": "no existe"}), 404
    try:
        return jsonify({"ok": True, "asunto": abrir_en_outlook(r["entry_id"])})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    init_db()
    cargar_usuarios()   # crea usuarios.json con los usuarios iniciales si no existe
    print("El Cerebro corriendo en:")
    print("  local  ->  http://127.0.0.1:5000")
    print("  en LAN ->  http://0.0.0.0:5000")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)

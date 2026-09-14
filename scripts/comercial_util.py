# -*- coding: utf-8 -*-
"""
Utilidades comerciales compartidas: resolución del vendedor y agrupación
inteligente de clientes (normalización + mapa de alias mantenido por Cowork).
"""
import json
import os
import re
import unicodedata

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ALIAS_FILE = os.path.join(BASE, "data", "clientes_alias.json")

# ───────────────────────── Vendedores ─────────────────────────
VENDEDORES = {
    "a.rivera@example.com": "Alex Rivera",
    "s.carter@example.com": "Sam Carter",
    "d.torres@example.com": "Diego Torres",
    "l.moretti@example.com": "Luca Moretti",
    "n.becker@example.com": "Nils Becker",
    "m.rossi@example.com": "Marco Rossi",
}
ALIAS_VENDEDOR = {
    "sam carter": "Sam Carter", "sam carter": "Sam Carter",
    "nils becker": "Nils Becker", "nils becker": "Nils Becker",
    "luca moretti": "Luca Moretti", "luca moretti": "Luca Moretti",
    "alex rivera": "Alex Rivera", "alex rivera": "Alex Rivera",
    "diego torres": "Diego Torres", "marco rossi": "Marco Rossi",
    "erik solberg": "Erik Solberg",   # ex-vendedor: sus cotizaciones quedan a su nombre
}
# Reasignación SOLO por email: el macro de Diego manda OC desde el mail de Erik.
# OJO: en cotizaciones "Erik Solberg" es un vendedor real (ex) -> NO se reasigna por nombre.
REDIR_VENDEDOR = {
    "e.solberg@example.com": "d.torres@example.com",
}
# Vendedores actualmente activos (Erik Solberg es ex-vendedor: sus cotizaciones quedan a su nombre)
VENDEDORES_ACTIVOS = ["Alex Rivera", "Sam Carter", "Diego Torres",
                      "Marco Rossi", "Luca Moretti", "Nils Becker"]


def _na(s):   # sin acentos, minúsculas
    import unicodedata as _u
    return "".join(c for c in _u.normalize("NFD", str(s).strip().lower())
                   if _u.category(c) != "Mn")


_ALIAS_NA = {_na(k): v for k, v in ALIAS_VENDEDOR.items()}


def nombre_vendedor(sender):
    if not sender:
        return "(sin remitente)"
    s = _na(sender)
    if s in REDIR_VENDEDOR:
        s = REDIR_VENDEDOR[s]
    return VENDEDORES.get(s) or _ALIAS_NA.get(s) or sender


# ───────────────────────── Clientes ─────────────────────────
# Sufijos societarios que se quitan para que "DALBAN PHARMA" == "DALBAN PHARMA SA".
_SUFIJOS = re.compile(r"\b(SAU|SACIFIA|SACIF|SAICF|SAIC|SCA|SRL|SA|LTDA|SH|SAS)\b")


def _sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def normalizar_cliente(nombre):
    """Nombre canónico determinista: mayúsculas, sin acentos, sufijos societarios ni puntuación."""
    s = _sin_acentos((nombre or "").upper().strip())
    s = s.replace(".", "")                       # S.A. -> SA ; S.R.L -> SRL
    s = _SUFIJOS.sub(" ", s)
    s = re.sub(r"[^\w\s&/-]", " ", s)            # fuera símbolos raros
    s = re.sub(r"\s+", " ", s).strip()
    return s or "(SIN CLIENTE)"


_alias_idx = None


def _cargar_alias():
    """Índice variante_normalizada -> nombre canónico, desde data/clientes_alias.json."""
    global _alias_idx
    if _alias_idx is None:
        _alias_idx = {}
        if os.path.exists(ALIAS_FILE):
            data = json.load(open(ALIAS_FILE, encoding="utf-8"))
            for canonico, variantes in data.items():
                if canonico.startswith("_") or not isinstance(variantes, list):
                    continue                     # claves de comentario / metadatos
                for v in variantes:
                    _alias_idx[normalizar_cliente(v)] = canonico
    return _alias_idx


def cliente_canonico(nombre):
    """Agrupa por: 1) normalización (quita SA/SRL/etc) y 2) mapa de alias de Cowork."""
    norm = normalizar_cliente(nombre)
    return _cargar_alias().get(norm, norm)

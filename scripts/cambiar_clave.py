# -*- coding: utf-8 -*-
"""Cambia la contraseña de un usuario de Ariana. La tipeas vos, localmente."""
import json, os, sys, getpass
from werkzeug.security import generate_password_hash

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F = os.path.join(BASE, "usuarios.json")

users = json.load(open(F, encoding="utf-8"))
print("users disponibles:", ", ".join(users))
u = input("user a cambiar: ").strip().lower()
if u not in users:
    print("Ese usuario no existe."); sys.exit(1)

p1 = getpass.getpass("Nueva contraseña: ")
p2 = getpass.getpass("Repetir contraseña: ")
if not p1 or p1 != p2:
    print("Vacía o no coinciden. No se cambió nada."); sys.exit(1)

users[u]["hash"] = generate_password_hash(p1)
json.dump(users, open(F, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
print(f"Contraseña de '{u}' actualizada. Toma efecto en el próximo login (no hace falta reiniciar).")

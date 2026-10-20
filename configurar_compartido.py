"""
Se corre UNA SOLA VEZ (en tu laptop) para que lo que se completa a mano en el Reporte diario
(proyectados, objetivos, indicadores de producción, personas planilladas, despachos, inspectores, stock e imágenes)
se guarde para TODOS: cualquiera que entre con su DNI lo ve y lo puede corregir.

Qué hace:
  1. Te pide un «token» de GitHub (lo creas en GitHub, te lo explico abajo) y comprueba que funcione.
  2. Crea una clave para cifrar lo que se escribe a mano.
  3. Guarda el token y esa clave DENTRO de fito.json, cifrados (solo se abren entrando con un DNI con acceso),
     y lo sube a GitHub. El token nunca queda a la vista en el repositorio.

Cómo crear el token (una vez):
  GitHub → tu foto (arriba a la derecha) → Settings → Developer settings → Personal access tokens →
  Fine-grained tokens → Generate new token
    Token name:          reporte-diario
    Expiration:          la más larga que te deje (cuando venza, vuelves a correr este script con uno nuevo)
    Repository access:   Only select repositories → JohanCh01/flujo-cosecha
    Permissions → Repository permissions → Contents: Read and write
  Generate token → cópialo (empieza con github_pat_…) y pégalo cuando este script lo pida.

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python configurar_compartido.py
"""

import base64
import getpass
import json
import os
import sys
import urllib.request
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import recepcion_sql as rs          # misma clave de administrador, mismo cifrado y misma subida a GitHub
af, log, git = rs.af, rs.log, rs.git

REPO = "JohanCh01/flujo-cosecha"


def probar_token(tok):
    req = urllib.request.Request(f"https://api.github.com/repos/{REPO}",
                                 headers={"Authorization": f"Bearer {tok}", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        info = json.loads(r.read().decode("utf-8"))
    permisos = info.get("permissions") or {}
    if not permisos.get("push"):
        sys.exit("El token funciona pero NO puede escribir en el repositorio. Revisa: Contents → Read and write.")
    return info


def main():
    tok = getpass.getpass("Pega el token de GitHub (no se ve al pegar) y presiona Enter: ").strip()
    if not tok:
        sys.exit("No pegaste nada.")
    try:
        probar_token(tok)
    except Exception as e:
        sys.exit(f"El token no funcionó con {REPO}: {e}")
    log("Token comprobado: puede leer y escribir en el repositorio.")

    clave = rs.clave_admin()
    for _ in range(3):
        git("pull", "--ff-only", "-q")
        dnis, it, sal = af.accesos(clave)
        contenido, _ = af.abrir_fito(clave)
        if not contenido:
            sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
        comp = contenido.get("comp") or {}
        if not comp.get("mk"):
            comp["mk"] = base64.b64encode(os.urandom(32)).decode()      # clave de lo que se escribe a mano (no cambia)
        comp["gh"] = tok
        comp["repo"] = REPO
        comp["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
        contenido["comp"] = comp
        af.guardar_fito(contenido, dnis, clave, it, sal)
        git("add", "fito.json")
        git("commit", "-q", "-m", f"Guardado compartido del Reporte diario ({datetime.now():%d/%m/%Y %H:%M})")
        try:
            git("push", "-q")
            log("Listo. En 1–2 minutos, al entrar al Reporte diario, lo que se complete a mano se guarda para todos.")
            log("Abre la página en TU navegador primero: lo que ya tenías escrito se sube solo.")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo; vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


if __name__ == "__main__":
    main()

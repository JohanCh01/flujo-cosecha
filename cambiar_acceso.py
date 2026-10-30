"""
Cambia el acceso del sitio a UN usuario y UNA contraseña comunes, y renueva todas las llaves.

Lo corre la tarea «Cambiar acceso» de GitHub (Actions → Cambiar acceso → Run workflow).
Se usa la primera vez (para dejar de entrar con DNI) y cada vez que quieras cambiar la contraseña.

Qué hace:
  1. Vuelve a cifrar datos.json, fito.json, paletas.json y comparativo.json con una llave NUEVA,
     que solo se abre con el usuario + la contraseña. Los DNIs dejan de servir.
  2. Renueva la clave con la que se guarda lo que se escribe a mano en el Reporte diario
     (manual.json y man_*.json) y vuelve a cifrar esos archivos.
  3. Si le das un token nuevo de GitHub (secreto TOKEN_REPORTE), lo pone en lugar del anterior.
  4. Comprueba que cada archivo se abre con el acceso nuevo antes de guardar nada.
  Si algo falla, no cambia ningún archivo.

Datos que usa (en GitHub son «secrets»; nunca van escritos aquí):
    ADMIN_CLAVE      la clave de administrador de siempre
    ACCESO_USUARIO   el usuario común (en la tarea: Calidadnv5)
    ACCESO_CODIGO    la contraseña común: al menos 12 letras o números (no distingue mayúsculas)
    TOKEN_REPORTE    (opcional) token nuevo de GitHub para el guardado compartido del Reporte diario
"""

import base64
import gzip
import json
import os
import re
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import actualizar_fito as af          # mismo cifrado de siempre: _clave, _llave_dni, _firma, _abrir_adm

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

CARPETA = Path(os.environ.get("CARPETA_SITIO", "."))
PAQUETES = ["datos.json", "fito.json", "paletas.json", "comparativo.json"]   # datos.json primero: de ahí leen los demás scripts
MIN_CLAVE = 12


def log(m):
    print(m, flush=True)


def b64(b):
    return base64.b64encode(b).decode("ascii")


def deb64(t):
    return base64.b64decode(t)


def limpiar(v):
    return re.sub(r"[^0-9A-Z]", "", str(v or "").upper())


def credencial(usuario, codigo):
    """Lo mismo que hace la página (credOf): usuario + contraseña juntos, en mayúsculas y sin signos."""
    u, c = limpiar(usuario), limpiar(codigo)
    if not u:
        raise af.Detener("Falta el usuario (ACCESO_USUARIO).")
    if len(c) < MIN_CLAVE:
        raise af.Detener(f"La contraseña (secreto ACCESO_CODIGO) debe tener al menos {MIN_CLAVE} letras o números.")
    return af._llave_dni(u + c)


def abrir(pk, clave_admin):
    """Devuelve (contenido_cifrado_en_bytes, datos_del_administrador)."""
    a = af._abrir_adm(pk, clave_admin)
    plano = AESGCM(deb64(a["k"])).decrypt(deb64(pk["iv"]), deb64(pk["data"]), None)
    return plano, a


def recifrar(pk, plano, a, cred, clave_admin, sal, it):
    """Llave nueva para los datos, una sola cerradura (el acceso común) y la del administrador."""
    nueva, iv, kiv = os.urandom(32), os.urandom(12), os.urandom(12)
    pk = dict(pk)
    pk["salt"], pk["it"] = b64(sal), it
    pk["iv"], pk["data"] = b64(iv), b64(AESGCM(nueva).encrypt(iv, plano, None))
    pk["keys"] = [{"i": b64(kiv), "c": b64(AESGCM(af._clave(cred, sal, it)).encrypt(kiv, nueva, None))}]
    a = dict(a)
    a["k"] = b64(nueva)
    if "dnis" in a:                       # datos.json: aquí se guarda quién tiene acceso (lo leen los demás scripts)
        a["dnis"] = [cred]
    if "firma" in a:                      # fito.json: marca de la lista de accesos
        a["firma"] = af._firma([cred], sal, it)
    s_adm, i_adm = os.urandom(16), os.urandom(12)
    c_adm = AESGCM(af._clave("adm:" + clave_admin, s_adm, it)).encrypt(i_adm, json.dumps(a, separators=(",", ":")).encode("utf-8"), None)
    pk["adm"] = {"s": b64(s_adm), "i": b64(i_adm), "c": b64(c_adm)}
    return pk


def comprobar(pk, cred):
    """Abre el archivo como lo haría la página con el acceso nuevo."""
    bits = af._clave(cred, deb64(pk["salt"]), int(pk["it"]))
    k = AESGCM(bits).decrypt(deb64(pk["keys"][0]["i"]), deb64(pk["keys"][0]["c"]), None)
    AESGCM(k).decrypt(deb64(pk["iv"]), deb64(pk["data"]), None)


def es_manual(pk):
    return isinstance(pk, dict) and set(pk) == {"v", "iv", "data"}


def main():
    clave_admin = os.environ.get("ADMIN_CLAVE", "").strip()
    if not clave_admin:
        raise af.Detener("Falta el secreto ADMIN_CLAVE.")
    cred = credencial(os.environ.get("ACCESO_USUARIO", ""), os.environ.get("ACCESO_CODIGO", ""))
    token = os.environ.get("TOKEN_REPORTE", "").strip()

    ruta_datos = CARPETA / "datos.json"
    if not ruta_datos.exists():
        raise af.Detener("No existe datos.json.")
    datos = json.loads(ruta_datos.read_text(encoding="utf-8"))
    sal, it = deb64(datos["salt"]), int(datos["it"])

    salida = {}        # nombre de archivo -> texto nuevo (se escribe todo al final, si nada falló)
    for nombre in PAQUETES:
        ruta = CARPETA / nombre
        if not ruta.exists():
            log(f"- {nombre}: no existe, se omite.")
            continue
        pk = json.loads(ruta.read_text(encoding="utf-8"))
        try:
            plano, a = abrir(pk, clave_admin)
        except Exception:
            raise af.Detener(f"No se pudo abrir {nombre} con la clave de administrador (ADMIN_CLAVE). No se cambió nada.")

        if nombre == "fito.json":         # aquí viven la clave de lo escrito a mano y el token
            contenido = json.loads(gzip.decompress(plano) if pk.get("z") else plano)
            comp = contenido.get("comp") or {}
            if comp.get("mk"):
                vieja = AESGCM(deb64(comp["mk"]))
                nueva_mk = os.urandom(32)
                nueva = AESGCM(nueva_mk)
                for r in sorted(CARPETA.glob("*.json")):
                    if r.name in PAQUETES:
                        continue
                    try:
                        m = json.loads(r.read_text(encoding="utf-8"))
                    except Exception:
                        continue
                    if not es_manual(m):
                        continue
                    try:
                        claro = vieja.decrypt(deb64(m["iv"]), deb64(m["data"]), None)
                    except Exception:
                        raise af.Detener(f"No se pudo abrir {r.name} (lo escrito a mano). No se cambió nada.")
                    iv = os.urandom(12)
                    salida[r.name] = json.dumps({"v": m["v"], "iv": b64(iv), "data": b64(nueva.encrypt(iv, claro, None))}, separators=(",", ":"))
                    log(f"- {r.name}: lo escrito a mano, cifrado con clave nueva.")
                comp["mk"] = b64(nueva_mk)
            if token:
                comp["gh"] = token
                log("- Token del Reporte diario: reemplazado por el nuevo.")
            elif comp.get("gh"):
                log("- AVISO: no hay secreto TOKEN_REPORTE; se mantiene el token anterior. Conviene crear uno nuevo y borrar el viejo.")
            if comp:
                contenido["comp"] = comp
            plano = gzip.compress(json.dumps(contenido, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), mtime=0)
            pk["z"] = True

        nuevo = recifrar(pk, plano, a, cred, clave_admin, sal, it)
        comprobar(nuevo, cred)
        salida[nombre] = json.dumps(nuevo, separators=(",", ":"))
        log(f"- {nombre}: cifrado con el acceso nuevo y comprobado.")

    for nombre, texto in salida.items():
        (CARPETA / nombre).write_text(texto, encoding="utf-8")
    log(f"Listo: {len(salida)} archivo(s) actualizados. Desde ahora se entra con el usuario y la contraseña nuevos.")


if __name__ == "__main__":
    try:
        main()
    except af.Detener as e:
        log(f"ERROR: {e}")
        sys.exit(1)

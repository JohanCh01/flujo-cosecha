"""
Kilos recepcionados desde el SQL (para la sección «Kilos de Espárrago Recepcionado» de reporte.html).

Corre en TU LAPTOP (el servidor CEDMARTPRD solo se alcanza desde la red de Danper; GitHub no llega).
Qué hace:
  1. Lee dbo.HECHO_INGRESOMMPP de CEDMARTPRD / BDDMDANPER (pyodbc), igual que tu Power BI:
     (usuario de SQL; se pide la primera vez y se guarda en ~/.flujo_cosecha_sql, fuera del repositorio)
     Proceso = FR, fecha = solo la fecha de HoraCosecha, suma de Peso Planta, sin MODULO = REPROCESO.
     Zona: "OLMOS 2"/"OLMOS 3" -> OLMOS y se quita "ZONA-" (como tu columna Zona Agricola Est).
  2. Guarda SOLO el total por día y zona (del año que elijas, por defecto el actual) dentro de fito.json, cifrado.
  3. Sube fito.json a GitHub (git pull, commit y push). La actualización automática de GitHub lo conserva.

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python recepcion_sql.py                 (año actual, y lo sube a GitHub)
    python recepcion_sql.py --sin-subir     (solo prueba: muestra los totales y no sube nada)
    python recepcion_sql.py --anio 2026

La clave de administrador del Flujo se lee de la variable ADMIN_CLAVE o del archivo  ~/.flujo_cosecha_clave
(fuera del repositorio; la primera vez el script te la pide y te ofrece guardarla ahí, para el Programador de tareas).
"""

import argparse
import getpass
import os
import re
import subprocess
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
os.chdir(AQUI)                      # para que encuentre datos.json y fito.json aunque lo lance el Programador de tareas
sys.path.insert(0, str(AQUI))
import actualizar_fito as af        # mismo cifrado y mismas llaves que la página

SERVIDOR = os.environ.get("SQL_SERVIDOR", "CEDMARTPRD")
BASE = os.environ.get("SQL_BASE", "BDDMDANPER")
TABLA = os.environ.get("SQL_TABLA", "dbo.HECHO_INGRESOMMPP")
ARCHIVO_CLAVE = Path.home() / ".flujo_cosecha_clave"
ARCHIVO_SQL = Path.home() / ".flujo_cosecha_sql"      # usuario y contraseña del SQL (2 líneas), fuera del repositorio

# Posibles nombres de cada columna en la tabla (se comparan sin espacios, guiones ni mayúsculas)
COLUMNAS = {
    "hora": ["HORACOSECHA"],
    "zona": ["ZONAAGRICOLA", "UNIDADAGRICOLA"],
    "proveedor": ["PROVEEDOR"],
    "campo": ["CAMPO"],
    "proceso": ["PROCESO", "CODPROCESO", "PROCESOCOD", "TIPOPROCESO"],
    "peso": ["PESOPLANTA", "PESOPLANTAKG", "KGPLANTA", "PESONETOPLANTA"],
}


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[\s_\-.]", "", s.upper())


def clave_admin():
    c = os.environ.get("ADMIN_CLAVE", "").strip()
    if not c and ARCHIVO_CLAVE.exists():
        c = ARCHIVO_CLAVE.read_text(encoding="utf-8").strip()
    if not c:
        if not sys.stdin.isatty():
            sys.exit("Falta la clave de administrador (variable ADMIN_CLAVE o archivo ~/.flujo_cosecha_clave).")
        c = getpass.getpass("Clave de administrador del Flujo (no se ve al escribir): ").strip()
        if input(f"¿Guardarla en {ARCHIVO_CLAVE} para no volver a escribirla? (s/n): ").strip().lower().startswith("s"):
            ARCHIVO_CLAVE.write_text(c, encoding="utf-8")
            log(f"Clave guardada en {ARCHIVO_CLAVE} (fuera del repositorio; no se sube a GitHub).")
    return c


def credenciales_sql():
    """Usuario y contraseña de SQL: variables SQL_USUARIO / SQL_CLAVE, o el archivo ~/.flujo_cosecha_sql.
    Nunca van escritas en este archivo (el repositorio es público)."""
    u, c = os.environ.get("SQL_USUARIO", "").strip(), os.environ.get("SQL_CLAVE", "")
    if not u and ARCHIVO_SQL.exists():
        lineas = ARCHIVO_SQL.read_text(encoding="utf-8").splitlines()
        if len(lineas) >= 2:
            u, c = lineas[0].strip(), lineas[1]
    if not u:
        if not sys.stdin.isatty():
            sys.exit("Faltan el usuario y la contraseña del SQL (archivo ~/.flujo_cosecha_sql).")
        u = input("Usuario del SQL: ").strip()
        c = getpass.getpass("Contraseña del SQL (no se ve al escribir): ")
        if input(f"¿Guardarlos en {ARCHIVO_SQL} para no volver a escribirlos? (s/n): ").strip().lower().startswith("s"):
            ARCHIVO_SQL.write_text(u + "\n" + c, encoding="utf-8")
            log(f"Guardado en {ARCHIVO_SQL} (fuera del repositorio; no se sube a GitHub).")
    return u, c


def conectar():
    import pyodbc
    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
    if not drivers:
        sys.exit("No encontré un driver ODBC de SQL Server en esta PC.")
    preferido = next((d for d in drivers if "ODBC Driver 18" in d), None) or next((d for d in drivers if "ODBC Driver 17" in d), None) or drivers[0]
    extra = ";TrustServerCertificate=yes" if "18" in preferido else ""
    usuario, clave = credenciales_sql()
    log(f"Conectando a {SERVIDOR} / {BASE} como {usuario} con «{preferido}»…")
    try:
        return pyodbc.connect(f"DRIVER={{{preferido}}};SERVER={SERVIDOR};DATABASE={BASE};UID={usuario};PWD={{{clave}}}{extra}", timeout=30)
    except pyodbc.Error as e:
        if "18456" in str(e) and ARCHIVO_SQL.exists():
            sys.exit(f"El SQL rechazó el usuario o la contraseña. Borra {ARCHIVO_SQL} y vuelve a correr el script para escribirlos de nuevo.")
        raise


def columnas_de(cx):
    esquema, tabla = TABLA.split(".", 1) if "." in TABLA else ("dbo", TABLA)
    cur = cx.cursor()
    cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?", esquema, tabla)
    reales = [r[0] for r in cur.fetchall()]
    if not reales:
        sys.exit(f"No encontré la tabla {TABLA} en {BASE}.")
    por_norma = {norm(c): c for c in reales}
    elegidas, faltan = {}, []
    for k, opciones in COLUMNAS.items():
        real = next((por_norma[o] for o in opciones if o in por_norma), None)
        if real:
            elegidas[k] = real
        else:
            faltan.append(k)
    if faltan:
        print("Columnas de la tabla:", ", ".join(reales))
        sys.exit("No encontré estas columnas: " + ", ".join(faltan) + ". Pásale esta lista a Claude para ajustarlo.")
    return elegidas


def zona_est(z):
    z = " ".join(str(z or "").upper().split())
    if z in ("OLMOS 2", "OLMOS 3"):
        return "OLMOS"
    return z[5:] if z.startswith("ZONA-") else z


def es_reproceso(proveedor, campo):
    """MODULO = Campo antes de "/" cuando el proveedor es DANPER SAC; si no, el Proveedor (igual que tu Power Query)."""
    prov = " ".join(str(proveedor or "").upper().split())
    modulo = str(campo or "").split("/")[0].strip().upper() if prov == "DANPER SAC" else prov
    return modulo == "REPROCESO" or prov == "REPROCESO"


def leer_sql(anio):
    cx = conectar()
    c = columnas_de(cx)
    q = (f"SELECT CAST([{c['hora']}] AS date) AS f, [{c['zona']}] AS z, [{c['proveedor']}] AS p, [{c['campo']}] AS ca, "
         f"SUM(CAST([{c['peso']}] AS float)) AS kg "
         f"FROM {TABLA} WHERE UPPER(LTRIM(RTRIM([{c['proceso']}]))) IN ('FR', 'FRESCO') "
         f"AND [{c['hora']}] >= ? AND [{c['hora']}] < ? "
         f"GROUP BY CAST([{c['hora']}] AS date), [{c['zona']}], [{c['proveedor']}], [{c['campo']}]")
    cur = cx.cursor()
    cur.execute(q, date(anio, 1, 1), date(anio + 1, 1, 1))
    tot, repro = {}, 0.0
    for f, z, p, ca, kg in cur.fetchall():
        if f is None or kg is None:
            continue
        if es_reproceso(p, ca):
            repro += kg
            continue
        ymd = f.year * 10000 + f.month * 100 + f.day
        llave = (ymd, zona_est(z))
        tot[llave] = tot.get(llave, 0.0) + float(kg)
    cx.close()
    if repro:
        log(f"Se excluyeron {repro:,.2f} kg de REPROCESO.")
    return tot


def empacar(tot):
    base = min(k[0] for k in tot)
    zonas, pos, filas = [], {}, []
    for (ymd, z), kg in sorted(tot.items()):
        if z not in pos:
            pos[z] = len(zonas)
            zonas.append(z)
        filas += [af._dias_entre(base, ymd), pos[z], af._redondeo(kg * 100)]
    return {"base": base, "maxYmd": max(k[0] for k in tot), "uas": zonas, "w": 3, "rows": filas,
            "gen": datetime.now().strftime("%Y-%m-%dT%H:%M"), "fuente": f"{SERVIDOR}/{BASE} {TABLA}"}


def resumen(tot):
    meses = {}
    for (ymd, z), kg in tot.items():
        g = "COMPOSITAN" if "COMPOSITAN" in z else "OLMOS" if "OLMOS" in z else "TERCEROS"
        m = meses.setdefault(ymd // 100 % 100, {"COMPOSITAN": 0.0, "OLMOS": 0.0, "TERCEROS": 0.0})
        m[g] += kg
    print(f"{'Mes':>4} {'Olmos':>14} {'Terceros':>14} {'Compositan':>14} {'Total':>15}")
    for m in sorted(meses):
        o = meses[m]
        print(f"{m:>4} {o['OLMOS']:>14,.2f} {o['TERCEROS']:>14,.2f} {o['COMPOSITAN']:>14,.2f} {sum(o.values()):>15,.2f}")


def git(*args):
    r = subprocess.run(["git", *args], capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError((r.stderr or r.stdout).strip())
    return r.stdout.strip()


def guardar(paquete, clave):
    dnis, it, sal = af.accesos(clave)
    contenido, _ = af.abrir_fito(clave)
    if not contenido:
        sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
    contenido["rsql"] = paquete
    contenido["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    af.guardar_fito(contenido, dnis, clave, it, sal)


def main():
    ap = argparse.ArgumentParser(description="Sube a la página los kilos recepcionados (FR, Peso Planta) leídos del SQL")
    ap.add_argument("--anio", type=int, default=date.today().year)
    ap.add_argument("--sin-subir", action="store_true", help="solo muestra los totales; no cambia ni sube nada")
    a = ap.parse_args()

    tot = leer_sql(a.anio)
    if not tot:
        sys.exit(f"El SQL no devolvió kilos FR en {a.anio}.")
    resumen(tot)
    if a.sin_subir:
        return
    clave = clave_admin()
    paquete = empacar(tot)

    for intento in range(3):
        git("pull", "--ff-only", "-q")                 # trae lo último (la actualización automática también escribe fito.json)
        guardar(paquete, clave)
        if not git("status", "--porcelain", "--", "fito.json"):
            log("fito.json ya tenía estos mismos datos: no hay nada que subir.")
            return
        git("add", "fito.json")
        git("commit", "-q", "-m", f"Kilos recepcionados desde SQL ({datetime.now().strftime('%d/%m/%Y %H:%M')})")
        try:
            git("push", "-q")
            log("Listo: subido a GitHub. En 1–2 minutos se ve en la página (botón Actualizar).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo (la actualización automática); vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")     # deshace solo este commit y vuelve a armarlo sobre lo último
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


if __name__ == "__main__":
    main()

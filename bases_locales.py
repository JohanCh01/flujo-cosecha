"""
Sube a la web (Evaluación fitosanitaria y Reporte diario) las 3 bases que tienes en tu laptop, en vez de las de Drive:

    Escritorio\\SEGUIMIENTO DE MMPP\\FOSFINA                el Excel de la base fitosanitaria (BD_FOSFINA)
    Escritorio\\SEGUIMIENTO DE MMPP\\CALIDADES Y CALIBRES   el Excel de calidades y calibres
    Escritorio\\SEGUIMIENTO DE MMPP\\LONGITUDES             el Excel de longitudes

De cada carpeta toma el Excel MÁS RECIENTE, lo lee con las mismas reglas que la actualización automática
y lo guarda en fito.json (cifrado). Luego lo sube a GitHub.

La actualización automática de GitHub (8:32 a. m. y 1:05 p. m.) ya no tapa estas bases con las de Drive:
solo usa la de Drive si alguien la vuelve a actualizar allá DESPUÉS de que tú subiste la tuya.

Si un Excel no cambió desde la última vez que lo subiste, no se vuelve a leer (es rápido).

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python bases_locales.py                         (las 3 bases)
    python bases_locales.py --solo fosfina          (solo una o algunas: fosfina, calidades, longitudes)
    python bases_locales.py --carpeta "C:/Users/jchinguel/Desktop/SEGUIMIENTO DE MMPP"   (si está en otra ruta)
    python bases_locales.py --forzar                (vuelve a leer aunque el Excel no haya cambiado)
    python bases_locales.py --sin-subir             (solo lee y muestra; no sube nada)
Necesita: pip install openpyxl python-calamine
"""

import argparse
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import recepcion_sql as rs          # misma clave de administrador, mismo cifrado y misma subida a GitHub
af, log, git = rs.af, rs.log, rs.git

# clave en fito.json, subcarpeta, función que lee, nombre que se ve
BASES = [
    ("fito", "FOSFINA", af.leer_fosfina, "fosfina"),
    ("cal", "CALIDADES Y CALIBRES", af.leer_calidades, "calidades"),
    ("lon", "LONGITUDES", af.leer_longitudes, "longitudes"),
]


def carpeta_mmpp(dada):
    if dada:
        return Path(dada)
    casa = Path.home()
    candidatas = [casa / n / "SEGUIMIENTO DE MMPP" for n in ("Desktop", "Escritorio")]
    candidatas += [d / n / "SEGUIMIENTO DE MMPP" for d in casa.glob("OneDrive*") for n in ("Desktop", "Escritorio")]
    return next((c for c in candidatas if c.is_dir()), candidatas[0])


def excel_mas_reciente(carpeta):
    if not carpeta.is_dir():
        return None
    hallados = [f for f in carpeta.glob("*.xls*") if not f.name.startswith("~$") and f.is_file()]
    return max(hallados, key=lambda f: f.stat().st_mtime) if hallados else None


def fecha_archivo(f):
    return datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%dT%H:%M")


def main():
    ap = argparse.ArgumentParser(description="Sube a la web las bases de fosfina, calidades y longitudes de tu laptop")
    ap.add_argument("--carpeta", help="carpeta SEGUIMIENTO DE MMPP (si no, la busca en tu Escritorio)")
    ap.add_argument("--solo", nargs="+", choices=[b[3] for b in BASES], help="solo estas bases")
    ap.add_argument("--forzar", action="store_true", help="leer aunque el Excel no haya cambiado")
    ap.add_argument("--sin-subir", action="store_true", help="solo leer y mostrar; no sube nada")
    a = ap.parse_args()

    raiz = carpeta_mmpp(a.carpeta)
    if not raiz.is_dir():
        sys.exit(f"No encuentro la carpeta {raiz}.\n"
                 'Ponla entre comillas, por ejemplo:  python bases_locales.py --carpeta "C:/Users/jchinguel/Desktop/SEGUIMIENTO DE MMPP"')
    log(f"Carpeta: {raiz}")

    clave = None if a.sin_subir else rs.clave_admin()
    if clave:
        git("pull", "--ff-only", "-q")
    anterior = (af.abrir_fito(clave)[0] or {}) if clave else {}
    fuentes_ant = anterior.get("fuentes") or {}

    nuevos, fuentes_nuevas = {}, {}
    for k, sub, leer, nombre in BASES:
        if a.solo and nombre not in a.solo:
            continue
        f = excel_mas_reciente(raiz / sub)
        if not f:
            log(f"[!] {sub}: no hay ningún Excel en {raiz / sub}. Se deja la que ya estaba.")
            continue
        marca = {"origen": "laptop", "archivo": f.name, "modificado": fecha_archivo(f)}
        ya = fuentes_ant.get(k) or {}
        if not a.forzar and ya.get("origen") == "laptop" and ya.get("archivo") == f.name and ya.get("modificado") == marca["modificado"]:
            log(f"{sub}: «{f.name}» ({marca['modificado'].replace('T', ' ')}) ya estaba subido. No cambió.")
            continue
        log(f"{sub}: leyendo «{f.name}» (guardado el {marca['modificado'].replace('T', ' ')})… las bases grandes tardan un poco")
        try:
            datos = leer(f)
        except Exception as e:
            log(f"[!] {sub}: no se pudo leer ({type(e).__name__}: {e}). Se deja la que ya estaba.")
            continue
        if k == "fito":
            m = str(datos.get("maxYmd") or "")
            log(f"   {datos['nRows']:,} evaluaciones de {len(datos['fields'])} campos" + (f", hasta el {m[6:]}/{m[4:6]}/{m[:4]}." if m else "."))
            datos["name"] = f"Base fitosanitaria ({f.name})"
        else:
            log("   leída.")
        nuevos[k], fuentes_nuevas[k] = datos, marca

    if not nuevos:
        log("No hay nada nuevo que subir.")
        return
    if a.sin_subir:
        log("Prueba (--sin-subir): no se subió nada.")
        return

    ahora_utc = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M")
    for _ in range(3):
        git("pull", "--ff-only", "-q")
        dnis, it, sal = af.accesos(clave)
        contenido, _ = af.abrir_fito(clave)
        if not contenido:
            sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
        fuentes = dict(contenido.get("fuentes") or {})
        for k, datos in nuevos.items():
            contenido[k] = datos
            fuentes[k] = dict(fuentes_nuevas[k], utc=ahora_utc)
        contenido["fuentes"] = fuentes
        contenido["gen"] = datetime.now(af.LIMA).strftime("%Y-%m-%dT%H:%M")
        af.guardar_fito(contenido, dnis, clave, it, sal)
        git("add", "fito.json")
        nombres = ", ".join(sub for k, sub, _, _ in BASES if k in nuevos)
        git("commit", "-q", "-m", f"Bases desde la laptop: {nombres} ({datetime.now():%d/%m/%Y %H:%M})")
        try:
            git("push", "-q")
            log(f"Listo: subido ({nombres}). En 1–2 minutos se ve en la página (Ctrl + F5).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo; vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


if __name__ == "__main__":
    main()

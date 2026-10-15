"""
Mano de Obra – Calidad (CMO $/KGDW) desde el Power BI «CONTROL DE ASISTENCIA N05».
Para la sección «Mano de Obra/Personal – Calidad» de reporte.html y su diapositiva en la Vista PPT.

Corre en TU LAPTOP (mismo perfil de Edge que drenados_pbi.py). Qué hace, para cada SEMANA DE CORTE N:
  1. Abre el reporte y pone Año = el año elegido (por defecto, el actual); deja Mes y Fecha en «Todas».
  2. AVANCE DE CAMPAÑA: en «Sem» marca «Seleccionar todo» y desmarca las semanas mayores que N
     → lee la tarjeta CMO ($/KGDW): el valor grande (0,0172) y su «Proy: 0,0162 (+6,4 %)».
  3. SOLO LA SEMANA N: deja marcada solo N → lee lo mismo.
  4. Lo guarda dentro de fito.json (cifrado) y lo sube a GitHub.
Por defecto, N = la semana pasada y la actual.

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python mano_obra_pbi.py                  (semana pasada y actual; sube)
    python mano_obra_pbi.py --semanas 41     (solo esa semana de corte)
    python mano_obra_pbi.py --sin-subir      (solo prueba: muestra los valores y no sube nada)
    python mano_obra_pbi.py --ver            (con la ventana de Edge visible)
    python mano_obra_pbi.py --diagnostico    (si no encuentra el CMO: guarda en tu carpeta de usuario
                                              mano_obra_diagnostico.txt con lo que ve; mándamelo)

No lo corras al mismo tiempo que drenados_pbi.py ni que el script de capturas (usan el mismo perfil de Edge).
"""

import argparse
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import drenados_pbi as dp           # mismo navegador, mismos filtros y misma subida que el de drenados
rs, af, log, git, norm = dp.rs, dp.af, dp.log, dp.git, dp.norm

REPORT_URL = ("https://app.powerbi.com/groups/a5f26412-a444-426d-8728-4382ef3c0ef7/reports/"
              "aa024434-26e6-435c-a6f2-ea43b30544a6/5627b592e2ca0de3f630?experience=power-bi")
FILTRO_SEM = "Sem"
ARCHIVO_DIAG = Path.home() / "mano_obra_diagnostico.txt"

# Cada tarjeta KPI de la fila de arriba tiene <div class="indicatorText"> (el valor) y <div class="goalText"> («Proy: …»).
# Se toma la que tiene el título «CMO» (si no se ve el título, la de más a la izquierda).
JS_KPI = r"""
const out = {tarjetas: []};
for (const ind of document.querySelectorAll('.indicatorText')) {
  const r = ind.getBoundingClientRect(); if (r.width < 1) continue;
  let caja = ind.closest('visual-container') || ind.closest('.visualContainer') || ind.parentElement;
  const tx = [], w = document.createTreeWalker(caja, NodeFilter.SHOW_TEXT);
  while (w.nextNode()) { const t = w.currentNode.nodeValue.trim(); if (t) tx.push(t); }
  const g = caja.querySelector('.goalText');
  const titulo = Array.from(caja.querySelectorAll('[aria-label]')).map(e => e.getAttribute('aria-label')).join(' || ');
  out.tarjetas.push({x: r.x, y: r.y, w: r.width, h: r.height, v: ind.textContent.trim(), goal: g ? g.textContent.trim() : '',
                     t: tx.join(' | ').slice(0, 300), a: titulo.slice(0, 300)});
}
return out;
"""


def decimal(t):
    """«0,0172» → 0.0172 ; «0.0172» → 0.0172 ; «+6,4 %» → 6.4"""
    m = re.search(r"[-+]?\d[\d.,]*", t or "")
    if not m:
        return None
    n = m.group(0)
    if "," in n:
        n = n.replace(".", "").replace(",", ".")
    return float(n)


def leer_cmo(driver):
    d = dp.en_todos_los_frames(driver, JS_KPI)
    ts = [t for t in d.get("tarjetas") or [] if decimal(t["v"]) is not None]
    if not ts:
        return None
    cmo = [t for t in ts if "CMO" in norm(t["t"] + " " + t["a"]) and "TENDENCIA" not in norm(t["t"])]
    t = min(cmo or ts, key=lambda t: (round(t["y"] / 40), t["x"]))    # la de arriba a la izquierda
    g = t["goal"] or ""
    m = re.search(r"Proy\.?\s*:?\s*([-+]?\d[\d.,]*)", g, re.I)
    p = re.search(r"\(\s*([-+]?\s*\d[\d.,]*)\s*%?\s*\)", g)
    return {"v": decimal(t["v"]), "proy": decimal(m.group(1)) if m else None,
            "dif": decimal(p.group(1).replace(" ", "")) if p else None, "texto": g}


def leer_estable(driver, antes=None, espera=30):
    time.sleep(2.5)
    t0, fin, previo = time.time(), time.time() + espera, object()
    while time.time() < fin:
        v = leer_cmo(driver)
        if v is not None and antes is not None and v == antes and time.time() - t0 < 12:
            time.sleep(1); continue
        if v is not None and v == previo:
            return v
        previo = v
        time.sleep(1.5)
    return previo if isinstance(previo, dict) else None


def solo_semana(driver, n):
    """Deja marcada SOLO la semana n: «Seleccionar todo» (todo marcado) → otra vez (nada marcado) → la semana n."""
    s, popup = dp.abrir_slicer(driver, FILTRO_SEM)
    if popup is None:
        log(f"[!] No encontré el filtro «{FILTRO_SEM}».")
        return False
    dp.recorrer(driver, popup, dp.es_todo)                      # todo marcado
    time.sleep(1)
    dp.recorrer(driver, popup, dp.es_todo, solo_marcadas=True)  # todo desmarcado
    time.sleep(1)
    ok = dp.recorrer(driver, popup, lambda t: not dp.es_todo(t) and dp.num_semana(t) == n)
    time.sleep(0.5)
    marc = dp.recorrer(driver, popup, None, con_estado=True)
    dp.cerrar_slicer(driver, s)
    otras = [t for t, sel in marc if sel and not dp.es_todo(t) and dp.num_semana(t) != n]
    if not ok or otras:
        log(f"[!] No quedó marcada solo la semana {n} (también: {otras[:5]}).")
        return False
    return True


def diagnostico(driver):
    d = dp.en_todos_los_frames(driver, JS_KPI)
    with open(ARCHIVO_DIAG, "w", encoding="utf-8") as f:
        f.write(f"{datetime.now():%d/%m/%Y %H:%M}\n\nTARJETAS KPI ({len(d.get('tarjetas') or [])}):\n")
        for t in d.get("tarjetas") or []:
            f.write(f"  x={t['x']:.0f} y={t['y']:.0f}  valor={t['v']}  meta={t['goal']}\n    texto: {t['t']}\n    aria:  {t['a']}\n")
    log(f"Diagnóstico guardado en {ARCHIVO_DIAG}")


def texto(v):
    if not v:
        return "(no se pudo leer)"
    return f"CMO {v['v']:.4f}" + (f" · Proy {v['proy']:.4f}" if v["proy"] is not None else " · (sin Proy)") + \
           (f" ({v['dif']:+.1f} %)" if v["dif"] is not None else "")


def guardar(anio, datos, clave):
    dnis, it, sal = af.accesos(clave)
    contenido, _ = af.abrir_fito(clave)
    if not contenido:
        sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
    mo = contenido.get("mo") or {}
    a = mo.setdefault("anios", {}).setdefault(str(anio), {"acu": {}, "sem": {}})
    for n, par in datos.items():
        for k in ("acu", "sem"):
            if par.get(k):
                a.setdefault(k, {})[str(n)] = {x: par[k][x] for x in ("v", "proy", "dif") if par[k].get(x) is not None}
    mo["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    mo["fuente"] = "Power BI Control de Asistencia N05 · CMO ($/KGDW)"
    contenido["mo"] = mo
    contenido["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    af.guardar_fito(contenido, dnis, clave, it, sal)


def subir(anio, datos, clave):
    for _ in range(3):
        git("pull", "--ff-only", "-q")
        guardar(anio, datos, clave)
        if not git("status", "--porcelain", "--", "fito.json"):
            log("fito.json ya tenía estos mismos valores: no hay nada que subir.")
            return
        git("add", "fito.json")
        git("commit", "-q", "-m", f"Mano de obra (CMO) desde Power BI ({datetime.now():%d/%m/%Y %H:%M})")
        try:
            git("push", "-q")
            log("Listo: subido a GitHub. En 1–2 minutos se ve en la página (botón Actualizar).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo; vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


def main():
    ap = argparse.ArgumentParser(description="Sube a la página el CMO de mano de obra del Power BI")
    ap.add_argument("--anio", type=int, default=date.today().year)
    ap.add_argument("--semanas", type=int, nargs="+", help="semanas de corte (por defecto: la pasada y la actual)")
    ap.add_argument("--sin-subir", action="store_true")
    ap.add_argument("--ver", action="store_true")
    ap.add_argument("--diagnostico", action="store_true")
    a = ap.parse_args()

    clave = None if a.sin_subir else rs.clave_admin()
    driver = dp.iniciar_driver(a.ver or a.diagnostico)
    datos = {}
    try:
        log("Abriendo el Power BI de Control de Asistencia…")
        driver.get(REPORT_URL)
        time.sleep(12)
        dp.cancelar_reautenticacion(driver)
        if a.diagnostico:
            diagnostico(driver)
        if not dp.elegir(driver, "Año", lambda t: norm(t) == str(a.anio)):
            log(f"[!] No pude poner Año = {a.anio} (sigo con el que tenga).")
        time.sleep(2)
        for f in ("Mes", "Fecha"):
            dp.limpiar(driver, f)

        s, popup = dp.abrir_slicer(driver, FILTRO_SEM)
        if popup is None:
            sys.exit(f"No encontré el filtro «{FILTRO_SEM}».")
        disponibles = sorted({n for n in (dp.num_semana(t) for t in dp.recorrer(driver, popup) if not dp.es_todo(t)) if n})
        dp.cerrar_slicer(driver, s)
        if not disponibles:
            sys.exit(f"El filtro «{FILTRO_SEM}» no tiene semanas.")
        if a.semanas:
            cortes = [x for x in a.semanas if x in disponibles]
        else:
            hoy = date.today().isocalendar()[1]
            cortes = [x for x in (hoy - 1, hoy) if x in disponibles]
        log(f"Semanas de corte: {cortes}")

        antes = leer_cmo(driver)
        for n in cortes:
            par = {}
            if dp.marcar_hasta(driver, FILTRO_SEM, n):
                par["acu"] = leer_estable(driver, antes)
                antes = par["acu"] or antes
                log(f"Avance hasta la semana {n}: {texto(par['acu'])}")
            if solo_semana(driver, n):
                par["sem"] = leer_estable(driver, antes)
                antes = par["sem"] or antes
                log(f"Solo la semana {n}: {texto(par['sem'])}")
            par = {k: v for k, v in par.items() if v}
            if not par:
                diagnostico(driver)
                log(f"[!] Semana {n}: no pude leer el CMO. Mándame {ARCHIVO_DIAG.name}.")
                continue
            datos[n] = par
    finally:
        driver.quit()

    if not datos:
        sys.exit("No se leyó ninguna semana.")
    if a.sin_subir:
        return
    subir(a.anio, datos, clave)


if __name__ == "__main__":
    main()

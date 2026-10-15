"""
Kilos DRENADOS desde el Power BI «5.07 Evaluaciones de PT NV05» (hoja Resumen Producto Observado).
Para la sección «Kilogramos drenados» de reporte.html y su diapositiva en la Vista PPT.

Corre en TU LAPTOP (usa tu perfil de Edge, donde ya tienes la sesión de Power BI iniciada).
Qué hace (sin capturas: lee el número):
  1. Abre el reporte, borra todos los filtros y pone Año = el año elegido (por defecto, el actual).
  2. Elige UNA semana a la vez en el filtro «Semana» y lee el «Kg Prod» del lado 2026.
  3. Guarda el Kg Prod de cada semana dentro de fito.json (cifrado) y lo sube a GitHub.
     En la página: Ejecutados hasta S.N = suma de las semanas 1 a N; Solo semana N = la semana N.

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python drenados_pbi.py                  (vuelve a leer las 3 últimas semanas con datos y sube)
    python drenados_pbi.py --todo           (lee TODAS las semanas del año; la primera vez)
    python drenados_pbi.py --semanas 40 41  (solo esas semanas)
    python drenados_pbi.py --sin-subir      (solo prueba: muestra los kilos y no sube nada)
    python drenados_pbi.py --ver            (con la ventana de Edge visible, para mirar qué hace)
    python drenados_pbi.py --diagnostico    (si no encuentra el «Kg Prod»: guarda en tu carpeta de usuario
                                             drenados_diagnostico.txt con lo que ve en la hoja; mándamelo)

Cierra antes cualquier otro script que use el mismo perfil de Edge (por ejemplo, el de las capturas):
dos Edge no pueden abrir el mismo perfil a la vez.
"""

import argparse
import json
import re
import sys
import time
import unicodedata
from datetime import date, datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import recepcion_sql as rs          # misma clave de administrador, mismo cifrado y misma subida a GitHub
af = rs.af
log, git = rs.log, rs.git

EDGE_PROFILE_DIR = r"C:\Users\jchinguel\EdgeAutomationProfile"
REPORT_URL = ("https://app.powerbi.com/groups/a5f26412-a444-426d-8728-4382ef3c0ef7/reports/"
              "71e08bb7-cda5-4937-909b-489b95f4258d/ReportSectionbd5181d2b9608c3c0140?experience=power-bi")
XPATH_LIMPIAR_TODO = "//div[@aria-label='Marcador . Borrar todos los filtros']"
XPATH_MOSTRAR_FILTROS = "//div[@aria-label='Marcador . Mostrar Filtros']"
def xpath_slicer(nombre):
    return f"//div[@aria-label='{nombre}' and contains(@class,'slicer-dropdown-menu')]"
XPATH_HOJA = "//*[@data-automation-type='visualContainerHost']"
ETIQUETA = "KG PROD"                 # el texto que acompaña al número que se lee
ARCHIVO_DIAG = Path.home() / "drenados_diagnostico.txt"


# ---------------------------------------------------------------- navegador (igual que tu script de capturas)
def iniciar_driver(ver):
    o = Options()
    o.add_argument(f"--user-data-dir={EDGE_PROFILE_DIR}")
    if ver:
        o.add_argument("--start-maximized")
    else:
        for a in ("--headless=new", "--window-size=1920,1080", "--disable-gpu", "--log-level=3"):
            o.add_argument(a)
    return webdriver.Edge(options=o)


def _buscar(driver, xpath, espera, profundidad):
    try:
        return WebDriverWait(driver, espera).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except Exception:
        pass
    if profundidad <= 0:
        return None
    for i in range(len(driver.find_elements(By.TAG_NAME, "iframe"))):
        try:
            driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[i])
            r = _buscar(driver, xpath, espera, profundidad - 1)
            if r is not None:
                return r
            driver.switch_to.parent_frame()
        except Exception:
            driver.switch_to.default_content()
    return None


def buscar(driver, xpath, espera=10):
    """Busca en la página y, si no está, dentro de los iframes del reporte (queda dentro del iframe donde lo halló)."""
    driver.switch_to.default_content()
    return _buscar(driver, xpath, espera, 2)


def clic(driver, el):
    try:
        el.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", el)
        except Exception as e:
            print(f"    [!] No se pudo hacer clic: {e}")


def norm(t):
    t = unicodedata.normalize("NFKD", t or "")
    return "".join(c for c in t if not unicodedata.combining(c)).strip().upper()


def cancelar_reautenticacion(driver):
    for _ in range(2):
        if buscar(driver, "//*[contains(text(),'Debe comprobar su identidad')]", 3) is None:
            return
        b = buscar(driver, "//*[@id='cancelButton']", 5)
        if b is None:
            return
        log("Aviso «Debe comprobar su identidad»: lo cancelo.")
        clic(driver, b)
        time.sleep(1.5)


# ---------------------------------------------------------------- filtros
def texto_opcion(o):
    return (o.text or "").strip() or (o.get_attribute("title") or "").strip()


def abrir_slicer(driver, nombre):
    s = buscar(driver, xpath_slicer(nombre), 10)
    if s is None:
        return None, None
    pid = s.get_attribute("aria-controls")
    if s.get_attribute("aria-expanded") != "true":
        clic(driver, s)
        time.sleep(1)
    try:
        popup = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, pid)))
        WebDriverWait(driver, 10).until(lambda _: any(texto_opcion(o) for o in popup.find_elements(By.XPATH, ".//li | .//div[@role='option']")))
    except Exception:
        popup = None
    return s, popup


def cerrar_slicer(driver, s):
    try:
        if s.get_attribute("aria-expanded") == "true":
            clic(driver, s)
            time.sleep(0.5)
    except Exception:
        pass


def recorrer(driver, popup, quiere=None):
    """Recorre la lista (que se carga a medida que se baja). Sin «quiere»: devuelve todos los textos.
    Con «quiere(texto)»: hace clic en la primera opción que cumpla y devuelve True."""
    cont = driver.execute_script("""const r=arguments[0];let m=r,d=0;for(const e of r.querySelectorAll('*')){const x=e.scrollHeight-e.clientHeight;if(x>d){d=x;m=e;}}return m;""", popup)
    driver.execute_script("arguments[0].scrollTop = 0;", cont)
    time.sleep(0.3)
    vistos = []
    for _ in range(60):
        for o in popup.find_elements(By.XPATH, ".//li | .//div[@role='option']"):
            t = texto_opcion(o)
            if not t:
                continue
            if quiere is None:
                if t not in vistos:
                    vistos.append(t)
            elif quiere(t):
                if o.get_attribute("aria-selected") != "true":
                    clic(driver, o)
                return True
        antes = driver.execute_script("return arguments[0].scrollTop;", cont)
        driver.execute_script("arguments[0].scrollTop += arguments[0].clientHeight;", cont)
        time.sleep(0.3)
        if driver.execute_script("return arguments[0].scrollTop;", cont) == antes:
            break
    return vistos if quiere is None else False


def limpiar(driver, nombre):
    """Borra lo que tenga marcado un filtro (botón «Borrar selecciones» junto a su título)."""
    s = buscar(driver, xpath_slicer(nombre), 4)
    if s is None:
        return
    try:
        cont = s.find_element(By.XPATH, "./ancestor::*[contains(@class,'slicer-container')][1]")
        b = cont.find_element(By.XPATH, ".//span[contains(@class,'slicer-header-clear')]")
        if "none" not in (b.get_attribute("style") or "").replace(" ", ""):
            clic(driver, b)
            time.sleep(1)
    except Exception:
        pass


def num_semana(t):
    m = re.search(r"(\d{1,2})\s*$", t.strip()) or re.search(r"(\d{1,2})", t)
    return int(m.group(1)) if m else None


def elegir(driver, nombre, cumple):
    s, popup = abrir_slicer(driver, nombre)
    if popup is None:
        log(f"[!] No encontré el filtro «{nombre}».")
        return False
    ok = recorrer(driver, popup, cumple)
    cerrar_slicer(driver, s)
    return ok


# ---------------------------------------------------------------- lectura del «Kg Prod» de 2026
JS_VISUALES = r"""
const out = [];
let vs = document.querySelectorAll('visual-container');
if (!vs.length) vs = document.querySelectorAll('.visualContainer, [class*="visual-container"]');
for (const v of vs) {
  const r = v.getBoundingClientRect();
  if (r.width < 2 || r.height < 2) continue;
  const tx = [], w = document.createTreeWalker(v, NodeFilter.SHOW_TEXT);
  while (w.nextNode()) { const t = w.currentNode.nodeValue.trim(); if (t) tx.push(t); }
  const ar = [];
  for (const e of v.querySelectorAll('[aria-label]')) { const a = e.getAttribute('aria-label').trim(); if (a && ar.indexOf(a) < 0) ar.push(a); }
  out.push({x: r.x, y: r.y, w: r.width, h: r.height, t: tx.join(' | '), a: ar.slice(0, 6).join(' || ')});
}
return out;
"""
NUM = r"(\(En blanco\)|-?\d[\d.,]*\s*(?:mill|mil|k|m)?\b)"


def a_numero(t):
    t = (t or "").strip()
    if not t or "BLANCO" in norm(t):
        return 0.0
    m = re.match(r"(-?\d[\d.,]*)\s*(mill|mil|k|m)?", t, re.I)
    if not m:
        return None
    n, suf = m.group(1), (m.group(2) or "").lower()
    if "," in n and "." in n:                          # 1.234,5 o 1,234.5: el último separador es el decimal
        dec = "," if n.rfind(",") > n.rfind(".") else "."
        n = n.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif n.count(".") > 1 or n.count(",") > 1:
        n = n.replace(".", "").replace(",", "")
    elif re.fullmatch(r"-?\d{1,3}[.,]\d{3}", n) and not suf:
        n = n.replace(".", "").replace(",", "")        # 30.922 → 30922 (separador de miles)
    else:
        n = n.replace(",", ".")
    v = float(n)
    return v * (1e6 if suf in ("mill", "m") else 1e3 if suf in ("mil", "k") else 1)


def visuales(driver):
    if buscar(driver, XPATH_HOJA, 20) is None:          # entra al iframe donde está la hoja
        return []
    return driver.execute_script(JS_VISUALES) or []


def kg_prod_2026(driver, anio):
    vs = visuales(driver)
    cx = lambda v: v["x"] + v["w"] / 2
    cy = lambda v: v["y"] + v["h"] / 2
    cab = {}
    for a in (str(anio), str(anio - 1)):
        c = [v for v in vs if v["t"].strip() == a]
        if c:
            cab[a] = min(c, key=lambda v: v["y"])       # el título «2026» de más arriba
    def del_anio(v):                                    # ¿está del lado de 2026 (y no del de 2025)?
        if str(anio) not in cab:
            return True
        d = abs(cx(v) - cx(cab[str(anio)]))
        return str(anio - 1) not in cab or d < abs(cx(v) - cx(cab[str(anio - 1)]))
    etq = [v for v in vs if ETIQUETA in norm(v["t"] + " " + v["a"]) and del_anio(v)]
    if not etq:
        return None
    if str(anio) in cab:
        etq.sort(key=lambda v: abs(cx(v) - cx(cab[str(anio)])) + abs(cy(v) - cy(cab[str(anio)])) * 0.1)
    for v in etq:
        for campo in (v["t"], v["a"]):                  # 1) el número en el mismo cuadro: «Kg Prod: | 30.922»
            m = re.search(r"Kg\s*Prod\.?\s*:?\s*(?:\|\s*)?" + NUM, campo, re.I)
            if m:
                return a_numero(m.group(1))
        fila = [w for w in vs if w is not v and abs(cy(w) - cy(v)) < max(v["h"], w["h"]) / 2 + 8 and w["x"] >= v["x"] + v["w"] * 0.5
                and re.fullmatch(r"\s*" + NUM + r"\s*", w["t"] or "", re.I) and del_anio(w)]
        if fila:                                        # 2) el número en el cuadro de al lado
            return a_numero(min(fila, key=lambda w: w["x"])["t"])
    return None


def leer_estable(driver, anio, espera=25):
    """Espera a que el reporte termine de recalcular: el mismo número dos veces seguidas."""
    time.sleep(3)
    previo, fin = object(), time.time() + espera
    while time.time() < fin:
        v = kg_prod_2026(driver, anio)
        if v is not None and v == previo:
            return v
        previo = v
        time.sleep(1.5)
    return previo if isinstance(previo, float) else None


def diagnostico(driver):
    vs = visuales(driver)
    with open(ARCHIVO_DIAG, "w", encoding="utf-8") as f:
        f.write(f"{len(vs)} visuales — {datetime.now():%d/%m/%Y %H:%M}\n")
        for v in sorted(vs, key=lambda v: (round(v["y"] / 20), v["x"])):
            f.write(f"x={v['x']:.0f} y={v['y']:.0f} w={v['w']:.0f} h={v['h']:.0f}\n   texto: {v['t'][:300]}\n   aria:  {v['a'][:300]}\n")
    log(f"Diagnóstico guardado en {ARCHIVO_DIAG}")


# ---------------------------------------------------------------- guardar en fito.json y subir
def guardar(anio, kilos, clave):
    dnis, it, sal = af.accesos(clave)
    contenido, _ = af.abrir_fito(clave)
    if not contenido:
        sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
    dren = contenido.get("dren") or {}
    sem = dren.setdefault("anios", {}).setdefault(str(anio), {})
    for s, kg in kilos.items():
        sem[str(s)] = round(kg, 2)
    dren["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    dren["fuente"] = "Power BI 5.07 Evaluaciones de PT NV05 · Kg Prod"
    contenido["dren"] = dren
    contenido["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    af.guardar_fito(contenido, dnis, clave, it, sal)


def subir(anio, kilos, clave):
    for _ in range(3):
        git("pull", "--ff-only", "-q")
        guardar(anio, kilos, clave)
        if not git("status", "--porcelain", "--", "fito.json"):
            log("fito.json ya tenía estos mismos kilos: no hay nada que subir.")
            return
        git("add", "fito.json")
        git("commit", "-q", "-m", f"Kilos drenados desde Power BI ({datetime.now():%d/%m/%Y %H:%M})")
        try:
            git("push", "-q")
            log("Listo: subido a GitHub. En 1–2 minutos se ve en la página (botón Actualizar).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo; vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


def semanas_guardadas(anio, clave):
    try:
        contenido, _ = af.abrir_fito(clave)
        return {int(k): v for k, v in ((contenido or {}).get("dren", {}).get("anios", {}).get(str(anio), {})).items()}
    except Exception:
        return {}


# ---------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser(description="Sube a la página los kilos drenados (Kg Prod por semana) del Power BI")
    ap.add_argument("--anio", type=int, default=date.today().year)
    ap.add_argument("--todo", action="store_true", help="lee todas las semanas del año")
    ap.add_argument("--semanas", type=int, nargs="+", help="solo estas semanas")
    ap.add_argument("--sin-subir", action="store_true", help="solo muestra los kilos; no cambia ni sube nada")
    ap.add_argument("--ver", action="store_true", help="muestra la ventana de Edge")
    ap.add_argument("--diagnostico", action="store_true", help="guarda lo que ve en la hoja (para revisar)")
    a = ap.parse_args()

    clave = None if a.sin_subir else rs.clave_admin()
    driver = iniciar_driver(a.ver or a.diagnostico)
    kilos = {}
    try:
        log("Abriendo el Power BI…")
        driver.get(REPORT_URL)
        time.sleep(12)
        cancelar_reautenticacion(driver)
        if a.diagnostico:
            diagnostico(driver)
        b = buscar(driver, XPATH_LIMPIAR_TODO, 10)
        if b is not None:
            log("Borrando todos los filtros…")
            clic(driver, b)
            time.sleep(3)
        b = buscar(driver, XPATH_MOSTRAR_FILTROS, 10)
        if b is not None:
            clic(driver, b)
            time.sleep(3)
        if not elegir(driver, "Año", lambda t: norm(t) == str(a.anio)):
            sys.exit(f"No pude poner Año = {a.anio}.")
        time.sleep(2)
        for f in ("Mes", "Fecha", "fecha"):
            limpiar(driver, f)

        s, popup = abrir_slicer(driver, "Semana")
        if popup is None:
            sys.exit("No encontré el filtro «Semana».")
        disponibles = sorted({n for n in (num_semana(t) for t in recorrer(driver, popup)) if n})
        cerrar_slicer(driver, s)
        if not disponibles:
            sys.exit(f"El filtro «Semana» no tiene semanas para {a.anio}.")
        log(f"Semanas con datos en {a.anio}: {disponibles[0]} a {disponibles[-1]}")

        if a.semanas:
            leer = [x for x in a.semanas if x in disponibles]
        elif a.todo or a.sin_subir:
            leer = disponibles
        else:
            hechas = semanas_guardadas(a.anio, clave)
            leer = [x for x in disponibles if x not in hechas] + [x for x in disponibles[-3:] if x in hechas]   # nuevas + las 3 últimas otra vez
            leer = sorted(set(leer))

        for n in leer:
            limpiar(driver, "Semana")
            if not elegir(driver, "Semana", lambda t, n=n: num_semana(t) == n):
                log(f"[!] No pude elegir la semana {n}; la salto.")
                continue
            kg = leer_estable(driver, a.anio)
            if kg is None:
                if a.diagnostico or not kilos:
                    diagnostico(driver)
                log(f"[!] Semana {n}: no encontré el «Kg Prod» de {a.anio}. Corre con --diagnostico y mándame {ARCHIVO_DIAG.name}.")
                if not kilos:
                    sys.exit(1)
                continue
            kilos[n] = kg
            log(f"Semana {n}: {kg:,.2f} kg")
        limpiar(driver, "Semana")
    finally:
        driver.quit()

    if not kilos:
        sys.exit("No se leyó ninguna semana.")
    log(f"Total de las semanas leídas: {sum(kilos.values()):,.2f} kg")
    if a.sin_subir:
        return
    subir(a.anio, kilos, clave)


if __name__ == "__main__":
    main()

"""
Kilos DRENADOS y Producto terminado OBSERVADO desde el Power BI «5.07 Evaluaciones de PT NV05» (hoja Resumen Producto Observado).
Para la sección «Kilogramos drenados» de reporte.html y su diapositiva en la Vista PPT.

Corre en TU LAPTOP (usa tu perfil de Edge, donde ya tienes la sesión de Power BI iniciada).
Qué hace (sin capturas: lee el número), para cada SEMANA DE CORTE N:
  1. Abre el reporte, borra todos los filtros y pone Año = el año elegido (por defecto, el actual).
  2. ACUMULADO: en «Semana» marca «Seleccionar todo» y desmarca las semanas mayores que N
     → lee del lado 2026: Kg Prod (Ejecutados hasta S.N), Kg Obs y % observado (avance de campaña).
  3. SOLO LA SEMANA N: borra el filtro «Semana» y marca solo N → lee los mismos valores (y el objetivo).
  4. Lo guarda dentro de fito.json (cifrado) y lo sube a GitHub.
Por defecto, N = la semana pasada y la actual (así sirve elijas la que elijas en la página).

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python drenados_pbi.py                  (semana pasada y actual; sube)
    python drenados_pbi.py --semanas 41     (solo esa semana de corte)
    python drenados_pbi.py --sin-subir      (solo prueba: muestra los valores y no sube nada)
    python drenados_pbi.py --ver            (con la ventana de Edge visible, para mirar qué hace)
    python drenados_pbi.py --diagnostico    (si no encuentra los valores: guarda en tu carpeta de usuario
                                             drenados_diagnostico.txt con lo que ve en la hoja; mándamelo)

Cierra antes cualquier otro script que use el mismo perfil de Edge (por ejemplo, el de las capturas):
dos Edge no pueden abrir el mismo perfil a la vez.
"""

import argparse
import json
import os
import subprocess
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
ARCHIVO_DIAG = Path.home() / "drenados_diagnostico.txt"


# ---------------------------------------------------------------- navegador (igual que tu script de capturas)
def cerrar_edge_colgado():
    """Si quedó un Edge abierto con el perfil de automatización (p. ej. al cortar el script con Ctrl+C),
    lo cierra. Solo ese: tu Edge normal no se toca."""
    if os.name != "nt":
        return
    carpeta = Path(EDGE_PROFILE_DIR).name
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
          f"Where-Object {{ $_.CommandLine -like '*{carpeta}*' }} | "
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $_.ProcessId }")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True, timeout=30)
        if r.stdout.strip():
            log("Cerré un Edge que había quedado abierto con el perfil de automatización.")
            time.sleep(2)
    except Exception:
        pass


def iniciar_driver(ver):
    cerrar_edge_colgado()
    try:
        return _iniciar_driver(ver)
    except Exception as e:
        if "crashed" not in str(e) and "DevToolsActivePort" not in str(e):
            raise
        log("Edge no arrancó (el perfil estaba ocupado); lo intento otra vez…")
        cerrar_edge_colgado()
        time.sleep(3)
        return _iniciar_driver(ver)


def _iniciar_driver(ver):
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


def recorrer(driver, popup, quiere=None, solo_marcadas=False, con_estado=False):
    """Recorre la lista (que se carga a medida que se baja). Sin «quiere»: devuelve todos los textos
    (con con_estado=True, pares (texto, marcada)). Con «quiere(texto)»: hace clic en la primera opción que cumpla
    (con solo_marcadas=True, solo si está marcada) y devuelve True."""
    cont = driver.execute_script("""const r=arguments[0];let m=r,d=0;for(const e of r.querySelectorAll('*')){const x=e.scrollHeight-e.clientHeight;if(x>d){d=x;m=e;}}return m;""", popup)
    driver.execute_script("arguments[0].scrollTop = 0;", cont)
    time.sleep(0.3)
    vistos = []
    for _ in range(60):
        for o in popup.find_elements(By.XPATH, ".//li | .//div[@role='option']"):
            t = texto_opcion(o)
            if not t:
                continue
            sel = o.get_attribute("aria-selected") == "true" or o.get_attribute("aria-checked") == "true"
            if quiere is None:
                item = (t, sel) if con_estado else t
                if all((v[0] if con_estado else v) != t for v in vistos):
                    vistos.append(item)
            elif quiere(t):
                if solo_marcadas:
                    if sel:
                        clic(driver, o)
                        return True
                    continue
                if not sel:
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


def es_todo(t):
    n = norm(t)
    return n.startswith("SELECCIONAR TODO") or n.startswith("SELECT ALL")


def marcar_hasta(driver, nombre, hasta):
    """Marca «Seleccionar todo» y desmarca las semanas mayores que «hasta». Devuelve True si quedó bien."""
    s, popup = abrir_slicer(driver, nombre)
    if popup is None:
        log(f"[!] No encontré el filtro «{nombre}».")
        return False
    ok = recorrer(driver, popup, es_todo)              # clic en «Seleccionar todo» (si ya estaba marcado, no lo toca)
    if not ok:
        log("[!] El filtro «Semana» no tiene la opción «Seleccionar todo».")
        cerrar_slicer(driver, s)
        return False
    time.sleep(1)
    quitar = lambda t: not es_todo(t) and (num_semana(t) or 0) > hasta
    for _ in range(60):                                 # desmarca una por una las semanas mayores (la lista se vuelve a dibujar)
        if not recorrer(driver, popup, quitar, solo_marcadas=True):
            break
        time.sleep(0.8)
    # comprobación: ninguna semana mayor marcada y alguna menor sí
    marc = recorrer(driver, popup, None, con_estado=True)
    cerrar_slicer(driver, s)
    mal = [t for t, sel in marc if sel and quitar(t)]
    bien = [t for t, sel in marc if sel and not es_todo(t) and (num_semana(t) or 99) <= hasta]
    if mal or not bien:
        log(f"[!] La selección de semanas no quedó bien (marcadas de más: {mal[:5]}).")
        return False
    return True


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


# ---------------------------------------------------------------- lectura de los valores de 2026
# En la hoja, cada valor es una TARJETA cuyo aria-label dice «<nombre>, <valor> tarjeta», p. ej. «Kg dw, 30.922 tarjeta»
# (Kg Prod) y la del Kg Obs. El % observado es el indicador (gauge): su número grande es <text class="mainText"> «2,4 %»
# y el objetivo es <text class="targetText"> «3,5 %». Hay uno de cada por año: se toma el que está debajo del título «2026».
JS_TARJETAS = r"""
const out = {tarjetas: [], titulos: [], gauges: [], etiquetas: []};
const rc = e => { const r = e.getBoundingClientRect(); return {x: r.x, y: r.y, w: r.width, h: r.height}; };
for (const e of document.querySelectorAll('[aria-label]')) {
  const a = (e.getAttribute('aria-label') || '').replace(/ /g, ' ').trim();
  const r = rc(e);
  if (r.w < 2 || r.h < 2) continue;
  if (/^[^,]{1,40},\s*.*tarjeta\s*$/i.test(a)) out.tarjetas.push(Object.assign({a: a}, r));
  if (/\d/.test(a) && out.etiquetas.length < 400) out.etiquetas.push({a: a.slice(0, 200), x: Math.round(r.x), y: Math.round(r.y)});
}
for (const t of document.querySelectorAll('text.mainText')) {
  const r = rc(t); if (r.w < 1) continue;
  const svg = t.closest('svg'), obj = svg ? svg.querySelector('text.targetText') : null;
  const txt = n => n ? (n.childNodes[0] && n.childNodes[0].nodeType === 3 ? n.childNodes[0].nodeValue : n.textContent).trim() : '';
  out.gauges.push(Object.assign({v: txt(t), obj: txt(obj)}, r));
}
const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
while (w.nextNode()) {
  const t = w.currentNode.nodeValue.trim();
  if (/^20\d\d$/.test(t)) { const r = rc(w.currentNode.parentElement); if (r.w > 1) out.titulos.push(Object.assign({t: t}, r)); }
}
return out;
"""


def en_todos_los_frames(driver, js, profundidad=2):
    """Corre el JavaScript en la página y en cada iframe; devuelve el primer resultado que traiga tarjetas."""
    driver.switch_to.default_content()
    vistos = []
    def correr(nivel):
        try:
            r = driver.execute_script(js)
        except Exception:
            r = None
        if r:
            vistos.append(r)
            if r.get("tarjetas"):
                return r
        if nivel <= 0:
            return None
        for k in range(len(driver.find_elements(By.TAG_NAME, "iframe"))):
            try:
                driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[k])
                r = correr(nivel - 1)
                if r:
                    return r
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return None
    r = correr(profundidad)
    driver.switch_to.default_content()
    return r or (vistos[0] if vistos else {"tarjetas": [], "titulos": [], "gauges": [], "etiquetas": []})


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


def nombre_tarjeta(a):
    return norm(a.split(",")[0])


def valor_tarjeta(a):
    """«Kg dw, 30.922 tarjeta» → 30922.0 ; «Kg dw, (En blanco) tarjeta» → 0.0"""
    resto = a.split(",", 1)[1] if "," in a else ""
    resto = re.sub(r"\s*tarjeta\s*$", "", resto, flags=re.I).strip()
    return a_numero(resto)


def porcentaje(t):
    """«2,4 %» → 0.024"""
    v = a_numero((t or "").replace("%", "").strip()) if t and re.search(r"\d", t) else None
    return None if v is None else v / 100


def es_prod(a):
    return nombre_tarjeta(a) in ("KG DW", "KG PROD", "KGDW", "KGPROD")


def es_obs(a):
    n = nombre_tarjeta(a)
    return n.startswith("KG") and "OBS" in n


def leer_2026(driver, anio):
    """{'prod': Kg Prod, 'obs': Kg Obs, 'pct': % observado (fracción), 'obj': objetivo (fracción)} del lado 2026."""
    d = en_todos_los_frames(driver, JS_TARJETAS)
    cx = lambda v: v["x"] + v["w"] / 2
    prods = [t for t in d["tarjetas"] if es_prod(t["a"]) and valor_tarjeta(t["a"]) is not None]
    if not prods:
        return None
    tit, cerca = {}, lambda t: min(abs(cx(t) - cx(c)) for c in prods)
    for t in d["titulos"]:                              # de cada año, el título más cercano a las tarjetas (no el de la tabla)
        if t["t"] in (str(anio), str(anio - 1)) and (t["t"] not in tit or cerca(t) < cerca(tit[t["t"]])):
            tit[t["t"]] = t
    x26 = cx(tit[str(anio)]) if str(anio) in tit else max(cx(t) for t in prods)
    x25 = cx(tit[str(anio - 1)]) if str(anio - 1) in tit else None

    def del_lado(lista):
        lado = [t for t in lista if x25 is None or abs(cx(t) - x26) < abs(cx(t) - x25)]
        return min(lado, key=lambda t: abs(cx(t) - x26)) if lado else None

    out = {"prod": valor_tarjeta(del_lado(prods)["a"]) if del_lado(prods) else None, "obs": None, "pct": None, "obj": None}
    o = del_lado([t for t in d["tarjetas"] if es_obs(t["a"]) and valor_tarjeta(t["a"]) is not None])
    if o:
        out["obs"] = valor_tarjeta(o["a"])
    g = del_lado(d.get("gauges") or [])
    if g:
        out["pct"], out["obj"] = porcentaje(g["v"]), porcentaje(g["obj"])
    return out if out["prod"] is not None else None


def leer_estable(driver, anio, antes=None, espera=30):
    """Espera a que el reporte termine de recalcular: primero que cambie respecto de «antes» (hasta 12 s)
    y luego los mismos valores dos veces seguidas."""
    time.sleep(2.5)
    t0, fin, previo = time.time(), time.time() + espera, object()
    while time.time() < fin:
        v = leer_2026(driver, anio)
        if v is not None and antes is not None and v == antes and time.time() - t0 < 12:
            time.sleep(1); continue
        if v is not None and v == previo:
            return v
        previo = v
        time.sleep(1.5)
    return previo if isinstance(previo, dict) else None


def diagnostico(driver):
    d = en_todos_los_frames(driver, JS_TARJETAS)
    with open(ARCHIVO_DIAG, "w", encoding="utf-8") as f:
        f.write(f"{datetime.now():%d/%m/%Y %H:%M}\n\nTARJETAS ({len(d['tarjetas'])}):\n")
        for t in d["tarjetas"]:
            f.write(f"  x={t['x']:.0f} y={t['y']:.0f}  {t['a']}\n")
        f.write(f"\nINDICADORES (gauge) ({len(d.get('gauges') or [])}):\n")
        for t in d.get("gauges") or []:
            f.write(f"  x={t['x']:.0f} y={t['y']:.0f}  valor={t['v']}  objetivo={t['obj']}\n")
        f.write(f"\nTÍTULOS DE AÑO ({len(d['titulos'])}):\n")
        for t in d["titulos"]:
            f.write(f"  x={t['x']:.0f} y={t['y']:.0f}  {t['t']}\n")
        f.write(f"\nOTRAS ETIQUETAS CON NÚMEROS ({len(d['etiquetas'])}):\n")
        for t in d["etiquetas"]:
            f.write(f"  x={t['x']} y={t['y']}  {t['a']}\n")
    log(f"Diagnóstico guardado en {ARCHIVO_DIAG}")


# ---------------------------------------------------------------- guardar en fito.json y subir
def guardar(anio, kilos, clave):
    """kilos = {N: {'acu': {prod, obs, pct, obj}, 'sem': {...}}}"""
    dnis, it, sal = af.accesos(clave)
    contenido, _ = af.abrir_fito(clave)
    if not contenido:
        sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
    dren = contenido.get("dren") or {}
    # anios/obs/pct = solo esa semana · acu = acumulado de la semana 1 hasta N · obj = objetivo
    prod = dren.setdefault("anios", {}).setdefault(str(anio), {})
    obs = dren.setdefault("obs", {}).setdefault(str(anio), {})
    pct = dren.setdefault("pct", {}).setdefault(str(anio), {})
    acu = dren.setdefault("acu", {}).setdefault(str(anio), {})
    for n, par in kilos.items():
        v = par.get("sem")
        if v:
            prod[str(n)] = round(v["prod"], 2)
            if v.get("obs") is not None:
                obs[str(n)] = round(v["obs"], 2)
            if v.get("pct") is not None:
                pct[str(n)] = round(v["pct"], 6)
            if v.get("obj"):
                dren["obj"] = round(v["obj"], 6)
        w = par.get("acu")
        if w:
            acu[str(n)] = {k: (round(w[k], 6) if k == "pct" else round(w[k], 2)) for k in ("prod", "obs", "pct") if w.get(k) is not None}
            if w.get("obj"):
                dren["obj"] = round(w["obj"], 6)
    dren["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    dren["fuente"] = "Power BI 5.07 Evaluaciones de PT NV05 · Kg Prod, Kg Obs y % observado"
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


# ---------------------------------------------------------------- principal
def texto(v):
    if not v:
        return "(no se pudo leer)"
    return (f"Kg Prod {v['prod']:,.2f}" + (f" · Kg Obs {v['obs']:,.2f}" if v["obs"] is not None else " · (sin Kg Obs)") +
            (f" · Observado {v['obs'] / v['prod'] * 100:.2f} %" if v["obs"] is not None and v["prod"] else
             f" · Observado {v['pct'] * 100:.1f} %" if v["pct"] is not None else "") +
            (f" · objetivo {v['obj'] * 100:.1f} %" if v["obj"] else ""))


def main():
    ap = argparse.ArgumentParser(description="Sube a la página los kilos drenados y el producto observado del Power BI")
    ap.add_argument("--anio", type=int, default=date.today().year)
    ap.add_argument("--semanas", type=int, nargs="+", help="semanas de corte (por defecto: la pasada y la actual)")
    ap.add_argument("--sin-subir", action="store_true", help="solo muestra los valores; no cambia ni sube nada")
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
        disponibles = sorted({n for n in (num_semana(t) for t in recorrer(driver, popup) if not es_todo(t)) if n})
        cerrar_slicer(driver, s)
        if not disponibles:
            sys.exit(f"El filtro «Semana» no tiene semanas para {a.anio}.")
        log(f"Semanas con datos en {a.anio}: {disponibles[0]} a {disponibles[-1]}")

        if a.semanas:
            cortes = [x for x in a.semanas if x in disponibles]
        else:
            hoy = date.today().isocalendar()[1] if a.anio == date.today().year else disponibles[-1]
            cortes = [x for x in (hoy - 1, hoy) if x in disponibles] or disponibles[-1:]
        anterior = leer_2026(driver, a.anio)
        for n in cortes:
            par = {}
            # acumulado: Seleccionar todo y quitar las semanas mayores que n
            limpiar(driver, "Semana")
            if marcar_hasta(driver, "Semana", n):
                par["acu"] = leer_estable(driver, a.anio, anterior)
                anterior = par["acu"] or anterior
                log(f"Hasta la semana {n}: {texto(par['acu'])}")
            # solo la semana n
            limpiar(driver, "Semana")
            if elegir(driver, "Semana", lambda t, n=n: not es_todo(t) and num_semana(t) == n):
                par["sem"] = leer_estable(driver, a.anio, anterior)
                anterior = par["sem"] or anterior
                log(f"Solo la semana {n}: {texto(par['sem'])}")
            par = {k: v for k, v in par.items() if v}
            if not par:
                diagnostico(driver)
                log(f"[!] Semana {n}: no pude leer los valores de {a.anio}. Mándame {ARCHIVO_DIAG.name}.")
                continue
            kilos[n] = par
        limpiar(driver, "Semana")
    finally:
        driver.quit()

    if not kilos:
        sys.exit("No se leyó ninguna semana.")
    if a.sin_subir:
        return
    subir(a.anio, kilos, clave)


if __name__ == "__main__":
    main()

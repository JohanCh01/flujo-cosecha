"""
Ratio de reclamos penalizados desde el Power BI «RECLAMOS PENALIZADOS» (1.3 Informe Gerencial Reclamo de Clientes).
Para la sección «Ratio de reclamos penalizados» de reporte.html y su parte de la diapositiva 6 en la Vista PPT.

Corre en TU LAPTOP (mismo perfil de Edge que drenados_pbi.py). Qué hace:
  1. Abre el reporte. Filtros:
       AÑO       si ya está en el año actual no se toca; si no, se deja solo ese año
       PRODUCTO  solo ESPÁRRAGO VERDE
       MES       no se toca (se deja como esté en el reporte)
  2. Lee: % Penalidad (el indicador) y su objetivo, Penalizado $ y Venta $.
  3. Saca la captura del gráfico «Penalidad ($) por clasificación de reclamo» (va en lugar de «Imagen 2»).
  4. Lo guarda dentro de fito.json (cifrado, también la imagen) y lo sube a GitHub.

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python reclamos_pbi.py                  (lee y sube)
    python reclamos_pbi.py --sin-subir      (solo prueba: muestra los valores y guarda la captura en tu carpeta de usuario)
    python reclamos_pbi.py --ver            (con la ventana de Edge visible)
    python reclamos_pbi.py --hasta-mes 9    (otro mes final; por defecto, el mes actual)
    python reclamos_pbi.py --diagnostico    (si no encuentra algo: guarda reclamos_diagnostico.txt en tu carpeta de usuario)

No lo corras al mismo tiempo que los otros scripts de Power BI (usan el mismo perfil de Edge).
"""

import argparse
import base64
import io
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import drenados_pbi as dp           # mismo navegador, mismos filtros y misma subida que el de drenados
rs, af, log, git, norm = dp.rs, dp.af, dp.log, dp.git, dp.norm
By = dp.By

REPORT_URL = ("https://app.powerbi.com/groups/bb93e98d-f183-452d-a03b-7d12418aeb12/reports/"
              "69520418-699d-446f-a503-9054c787f203/29ae8d7e35638b86d803?experience=power-bi")
TITULO_GRAFICO = "Penalidad ($) por clasificación de reclamo"
PRODUCTO = "ESPÁRRAGO VERDE"
SIN_MESES = {2, 3}                  # febrero y marzo no van
ARCHIVO_DIAG = Path.home() / "reclamos_diagnostico.txt"
ARCHIVO_CAPTURA = Path.home() / "reclamos_grafico.jpg"
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
NOMBRE_MES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]


def num_mes(t):
    n = norm(t)
    if n.startswith("SET"):
        return 9
    for i, m in enumerate(MESES):
        if n.startswith(m):
            return i + 1
    m = re.fullmatch(r"(?:\d{4}\D)?(\d{1,2})", n)
    return int(m.group(1)) if m and 1 <= int(m.group(1)) <= 12 else None


def num_es(t):
    """Números como los muestra el Power BI en español: «220.542» → 220542 ; «0,31 %» → 0.31 ; «71.046.582» → 71046582"""
    m = re.search(r"-?\d[\d.,]*", (t or "").replace(" ", " "))
    if not m:
        return None
    n = m.group(0).rstrip(".,")
    if "," in n:
        n = n.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", n):
        n = n.replace(".", "")
    try:
        return float(n)
    except ValueError:
        return None


# ---------------------------------------------------------------- filtros
def clic_si(driver, popup, cond):
    """Recorre la lista y hace clic en la primera opción con cond(texto, marcada) verdadero. True si hizo clic."""
    cont = driver.execute_script("""const r=arguments[0];let m=r,d=0;for(const e of r.querySelectorAll('*')){const x=e.scrollHeight-e.clientHeight;if(x>d){d=x;m=e;}}return m;""", popup)
    driver.execute_script("arguments[0].scrollTop = 0;", cont)
    time.sleep(0.3)
    for _ in range(60):
        for o in popup.find_elements(By.XPATH, ".//li | .//div[@role='option']"):
            t = dp.texto_opcion(o)
            if not t:
                continue
            sel = o.get_attribute("aria-selected") == "true" or o.get_attribute("aria-checked") == "true"
            if cond(t, sel):
                dp.clic(driver, o)
                return True
        antes = driver.execute_script("return arguments[0].scrollTop;", cont)
        driver.execute_script("arguments[0].scrollTop += arguments[0].clientHeight;", cont)
        time.sleep(0.3)
        if driver.execute_script("return arguments[0].scrollTop;", cont) == antes:
            break
    return False


def ajustar(driver, nombre, quiero):
    """Deja marcadas exactamente las opciones con quiero(texto) verdadero: primero desmarca las que sobran, luego marca las que faltan."""
    s, popup = dp.abrir_slicer(driver, nombre)
    if popup is None:
        log(f"[!] No encontré el filtro «{nombre}».")
        return False
    for _ in range(40):                                  # 1) desmarcar lo que sobra
        if not clic_si(driver, popup, lambda t, sel: sel and not dp.es_todo(t) and not quiero(t)):
            break
        time.sleep(0.8)
    for _ in range(40):                                  # 2) marcar lo que falta
        if not clic_si(driver, popup, lambda t, sel: not sel and not dp.es_todo(t) and quiero(t)):
            break
        time.sleep(0.8)
    estado = dp.recorrer(driver, popup, None, con_estado=True)
    dp.cerrar_slicer(driver, s)
    time.sleep(1.5)
    mal = [t for t, sel in estado if not dp.es_todo(t) and sel != bool(quiero(t))]
    marcadas = [t for t, sel in estado if sel and not dp.es_todo(t)]
    if mal:
        log(f"[!] «{nombre}» no quedó como debía (revisar: {mal[:6]}).")
    return marcadas


def restatement(driver, nombre):
    s = dp.buscar(driver, dp.xpath_slicer(nombre), 10)
    return dp.valor_actual_slicer(s) if s is not None and hasattr(dp, "valor_actual_slicer") else (
        s.find_element(By.XPATH, ".//div[contains(@class,'slicer-restatement')]").text.strip() if s is not None else "")


# ---------------------------------------------------------------- lectura
JS_VALORES = r"""
const out = {tarjetas: []};
for (const v of document.querySelectorAll('visual-container')) {
  // el contenedor puede medir 0 (lo que se ve está en sus hijos): se mide el primer hijo con tamaño
  let r = v.getBoundingClientRect();
  if (r.width < 2) { for (const c of v.querySelectorAll('*')) { const q = c.getBoundingClientRect(); if (q.width >= 2) { r = q; break; } } }
  const tx = [], w = document.createTreeWalker(v, NodeFilter.SHOW_TEXT);
  while (w.nextNode()) { const t = w.currentNode.nodeValue.trim(); if (t) tx.push(t); }
  const valores = Array.from(v.querySelectorAll('text.value')).map(e => e.textContent.trim());
  const main = v.querySelector('text.mainText'), obj = v.querySelector('text.targetText');
  const solo = n => n ? (n.childNodes[0] && n.childNodes[0].nodeType === 3 ? n.childNodes[0].nodeValue : n.textContent).trim() : '';
  if (valores.length || main) out.tarjetas.push({x: r.x, y: r.y, t: tx.join(' | ').slice(0, 300), valores: valores, main: solo(main), obj: solo(obj)});
}
return out;
"""


def leer(driver):
    d = dp.en_todos_los_frames(driver, JS_VALORES)
    ts = d.get("tarjetas") or []
    out = {"pct": None, "obj": None, "pen": None, "venta": None}
    for t in ts:
        txt = norm(t["t"])
        if t["main"] and ("PENALIDAD" in txt or out["pct"] is None):
            out["pct"], out["obj"] = num_es(t["main"]), num_es(t["obj"])
        if t["valores"] and "PENALIZADO" in txt and "VENTA" not in txt:
            out["pen"] = num_es(t["valores"][0])
        if t["valores"] and "VENTA" in txt and "PENALIZADO" not in txt:
            out["venta"] = num_es(t["valores"][0])
    return out if any(v is not None for v in out.values()) else None


def leer_estable(driver, espera=40):
    time.sleep(3)
    fin, previo = time.time() + espera, object()
    while time.time() < fin:
        v = leer(driver)
        if v is not None and v == previo and v["pct"] is not None:
            return v
        previo = v
        time.sleep(2)
    return previo if isinstance(previo, dict) else None


JS_GRAFICO = r"""
const quiero = arguments[0].toLowerCase();
for (const v of document.querySelectorAll('visual-container')) {
  const h = v.querySelector('[data-testid="visual-title"], .visualTitle, h3');
  const t = ((h && (h.getAttribute('title') || h.textContent)) || '').trim().toLowerCase();
  if (t.indexOf(quiero) < 0) continue;
  const caja = v.querySelector('.vcBody') || v;
  caja.scrollIntoView({block: 'center'});
  const r = caja.getBoundingClientRect();          // el rectángulo que SE VE (ya con el achicado del reporte)
  return {x: r.x, y: r.y, w: r.width, h: r.height, dpr: window.devicePixelRatio || 1};
}
return null;
"""



def captura(driver):
    """Foto del gráfico «Penalidad ($) por clasificación de reclamo» → JPG en bytes.
    Se toma la pantalla completa y se recorta justo el rectángulo visible del gráfico (sin pedazos de otros)."""
    driver.switch_to.default_content()
    if driver.execute_script(JS_GRAFICO, TITULO_GRAFICO) is None:
        return None
    time.sleep(1.5)
    r = driver.execute_script(JS_GRAFICO, TITULO_GRAFICO)    # otra vez, ya quieto después de moverse a la vista
    png = driver.get_screenshot_as_png()
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(png)).convert("RGB")
        k = im.width / driver.execute_script("return window.innerWidth;") if r else 1
        caja = (max(0, int(r["x"] * k)), max(0, int(r["y"] * k)), min(im.width, int((r["x"] + r["w"]) * k)), min(im.height, int((r["y"] + r["h"]) * k)))
        im = im.crop(caja)
        b = io.BytesIO()
        im.save(b, "JPEG", quality=90, optimize=True)
        return b.getvalue(), "image/jpeg", im.size
    except Exception as ex:
        log(f"[!] No pude recortar la captura ({ex}); va la pantalla completa.")
        return png, "image/png", None


JS_ESTADO = r"""
return {url: location.href, titulo: document.title,
        visuales: document.querySelectorAll('visual-container').length,
        indicadores: document.querySelectorAll('text.mainText').length,
        valores: document.querySelectorAll('text.value').length,
        filtros: Array.from(document.querySelectorAll('.slicer-dropdown-menu')).map(e => e.getAttribute('aria-label')),
        iframes: document.querySelectorAll('iframe').length,
        texto: (document.body ? document.body.innerText : '').replace(/\s+/g, ' ').slice(0, 600)};
"""


def estado(driver):
    """Lo que hay en la página (y en cada iframe): para saber si el reporte cargó."""
    driver.switch_to.default_content()
    out = []
    try:
        out.append(("página", driver.execute_script(JS_ESTADO)))
    except Exception as e:
        out.append(("página", {"error": str(e)}))
    for k in range(len(driver.find_elements(By.TAG_NAME, "iframe"))):
        try:
            driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[k])
            out.append((f"iframe {k + 1}", driver.execute_script(JS_ESTADO)))
        except Exception as e:
            out.append((f"iframe {k + 1}", {"error": str(e)}))
        driver.switch_to.default_content()
    return out


def esperar_reporte(driver, maximo=90):
    """Espera hasta que aparezcan los gráficos del reporte (visual-container) y los filtros."""
    fin = time.time() + maximo
    while time.time() < fin:
        for _, e in estado(driver):
            if e.get("visuales", 0) >= 5 and e.get("filtros"):
                time.sleep(4)
                return True
        time.sleep(3)
    return False


def diagnostico(driver):
    d = dp.en_todos_los_frames(driver, JS_VALORES)
    with open(ARCHIVO_DIAG, "w", encoding="utf-8") as f:
        f.write(f"{datetime.now():%d/%m/%Y %H:%M}\n\n")
        for donde, e in estado(driver):
            f.write(f"[{donde}]\n")
            for k, v in e.items():
                f.write(f"  {k}: {v}\n")
            f.write("\n")
        f.write(f"TARJETAS CON VALORES ({len(d.get('tarjetas') or [])}):\n")
        for t in d.get("tarjetas") or []:
            f.write(f"x={t['x']:.0f} y={t['y']:.0f}  valores={t['valores']}  indicador={t['main']}  objetivo={t['obj']}\n   texto: {t['t']}\n")
    log(f"Diagnóstico guardado en {ARCHIVO_DIAG}")


def meses_texto(nums):
    """[1, 4, 5, 6, 7, 8, 9, 10] → «Enero y Abril – Octubre»"""
    nums = sorted(set(nums))
    grupos, ini = [], None
    for i, n in enumerate(nums):
        if ini is None:
            ini = n
        if i == len(nums) - 1 or nums[i + 1] != n + 1:
            grupos.append(NOMBRE_MES[ini - 1] if ini == n else f"{NOMBRE_MES[ini - 1]} – {NOMBRE_MES[n - 1]}")
            ini = None
    return grupos[0] if len(grupos) == 1 else ", ".join(grupos[:-1]) + " y " + grupos[-1]


# ---------------------------------------------------------------- guardar y subir
def guardar(paquete, clave):
    dnis, it, sal = af.accesos(clave)
    contenido, _ = af.abrir_fito(clave)
    if not contenido:
        sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
    contenido["recl"] = paquete
    contenido["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
    af.guardar_fito(contenido, dnis, clave, it, sal)


def subir(paquete, clave):
    for _ in range(3):
        git("pull", "--ff-only", "-q")
        guardar(paquete, clave)
        if not git("status", "--porcelain", "--", "fito.json"):
            log("fito.json ya tenía estos mismos valores: no hay nada que subir.")
            return
        git("add", "fito.json")
        git("commit", "-q", "-m", f"Ratio de reclamos desde Power BI ({datetime.now():%d/%m/%Y %H:%M})")
        try:
            git("push", "-q")
            log("Listo: subido a GitHub. En 1–2 minutos se ve en la página (botón Actualizar).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo; vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


def main():
    ap = argparse.ArgumentParser(description="Sube a la página el ratio de reclamos penalizados del Power BI")
    ap.add_argument("--anio", type=int, default=date.today().year)
    ap.add_argument("--hasta-mes", type=int, default=date.today().month)
    ap.add_argument("--sin-subir", action="store_true")
    ap.add_argument("--ver", action="store_true")
    ap.add_argument("--diagnostico", action="store_true")
    a = ap.parse_args()
    quiero_mes = lambda t: (num_mes(t) or 0) <= a.hasta_mes and num_mes(t) not in SIN_MESES and num_mes(t) is not None

    clave = None if a.sin_subir else rs.clave_admin()
    driver = dp.iniciar_driver(a.ver or a.diagnostico)
    try:
        log("Abriendo el Power BI de Reclamos penalizados…")
        driver.get(REPORT_URL)
        time.sleep(10)
        dp.cancelar_reautenticacion(driver)
        log("Esperando a que cargue el reporte…")
        if not esperar_reporte(driver):
            diagnostico(driver)
            sys.exit(f"El reporte no terminó de cargar (o abrió otra página). Mándame {ARCHIVO_DIAG.name}.")
        if a.diagnostico:
            diagnostico(driver)
        # AÑO: si ya dice el año, no se toca
        anio_txt = restatement(driver, "AÑO")
        if norm(anio_txt) == str(a.anio):
            log(f"AÑO ya está en {a.anio}: no lo toco.")
        else:
            log(f"AÑO dice «{anio_txt}»: lo pongo en {a.anio}.")
            ajustar(driver, "AÑO", lambda t: norm(t) == str(a.anio))
        # PRODUCTO: solo espárrago verde
        if norm(restatement(driver, "PRODUCTO")) == norm(PRODUCTO):
            log(f"PRODUCTO ya está en {PRODUCTO}.")
        else:
            log(f"PRODUCTO → solo {PRODUCTO}")
            ajustar(driver, "PRODUCTO", lambda t: norm(t) == norm(PRODUCTO))
        # MES: no se toca (se deja como esté en el reporte)
        meses = []
        log("MES: no lo toco; se deja como está en el reporte.")

        v = leer_estable(driver)
        if not v or v["pct"] is None:
            diagnostico(driver)
            sys.exit(f"No pude leer el % Penalidad. Mándame {ARCHIVO_DIAG.name}.")
        log(f"% Penalidad {v['pct']:.2f} %" + (f" · objetivo {v['obj']:.2f} %" if v["obj"] is not None else " · (sin objetivo)") +
            (f" · Penalizado $ {v['pen']:,.0f}" if v["pen"] is not None else " · (sin Penalizado)") +
            (f" · Venta $ {v['venta']:,.0f}" if v["venta"] is not None else " · (sin Venta)"))
        img = captura(driver)
        if img:
            ARCHIVO_CAPTURA.write_bytes(img[0])
            log(f"Captura del gráfico guardada en {ARCHIVO_CAPTURA} ({len(img[0]) // 1024} KB)")
        else:
            log(f"[!] No encontré el gráfico «{TITULO_GRAFICO}» para la captura.")
    finally:
        driver.quit()

    if a.sin_subir:
        return
    paquete = {"anio": a.anio, "meses": sorted(meses), "mesesTxt": meses_texto(meses) if meses else f"la campaña {a.anio}",
               "pct": v["pct"] / 100, "obj": v["obj"] / 100 if v["obj"] is not None else None, "pen": v["pen"], "venta": v["venta"],
               "gen": datetime.now().strftime("%Y-%m-%dT%H:%M"), "fuente": "Power BI Reclamos penalizados (1.3 Informe Gerencial Reclamo de Clientes)"}
    if img:
        paquete["img"] = "data:" + img[1] + ";base64," + base64.b64encode(img[0]).decode()
        if img[2]:
            paquete["imgW"], paquete["imgH"] = img[2]
    subir(paquete, clave)


if __name__ == "__main__":
    main()

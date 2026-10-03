"""
Actualización automática del Flujo Comparativo de la Cosecha.

Qué hace, en orden:
1. Entra a la intranet de Danper, abre el reporte "Recepción de Materia Prima",
   elige ESPÁRRAGO / ESPÁRRAGO VERDE, pone Fecha Inicio = ayer y Fecha Fin = hoy,
   espera a que cargue y exporta el Excel. Esta parte es el script
   recepcion_materia_prima.py que ya funcionaba, con los mismos pasos.
2. Lee el Excel igual que la página web (mismas columnas y mismas reglas).
3. Comprueba que las fechas del Excel correspondan a lo que se pidió. Si no
   coinciden (día y mes invertidos), repite la descarga escribiendo la fecha
   en el otro orden; si aun así no coinciden, se detiene SIN publicar.
4. Abre datos.json (los datos cifrados del enlace), reemplaza los días
   descargados, vuelve a cifrar con la misma lista de DNIs y guarda.
   En "Bases cargadas" queda una sola base por mes (p. ej. "Octubre 2026
   (automático)"), que va creciendo con cada descarga.

Pensado para correr en GitHub Actions (ver .github/workflows/actualizar.yml),
pero también corre en una PC con Python:

    python actualizar.py --solo-descarga --visible     (prueba: solo descarga y resume)
    python actualizar.py --excel reporte.xlsx          (usa un Excel ya descargado)

Credenciales: nunca van escritas aquí. Se leen de variables de entorno
(INTRANET_USUARIO, INTRANET_CLAVE, ADMIN_CLAVE); en GitHub son "secrets".
"""

import argparse
import base64
import gzip
import json
import math
import os
import re
import sys
import time
import unicodedata
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

LIMA = timezone(timedelta(hours=-5))
TIPO_CULTIVO = "ESPÁRRAGO"
MATERIA_PRIMA = "ESPÁRRAGO VERDE"
RUTA_DATOS = Path(os.environ.get("DATOS_JSON", "datos.json"))
EPOCA = datetime(1970, 1, 1)


def log(msg):
    print(f"[{datetime.now(LIMA).strftime('%H:%M:%S')}] {msg}", flush=True)


class Detener(Exception):
    """Error esperado: se muestra el mensaje y no se publica nada."""


# ============================================================
# 1. LECTURA DEL EXCEL (mismas reglas que la página web)
# ============================================================

COLS = ["ua", "prov", "campo", "cant", "c", "r", "s", "p", "placa", "guia", "ticket", "alm", "pr", "kp", "kc", "kb"]
CAMPOS = {
    "r": ["fecha recojo", "fecha de recojo", "hora recojo"],
    "c": ["hora cosecha", "inicio de cosecha", "inicio cosecha"],
    "s": ["hora salida", "salida"],
    "p": ["hora recepcion", "recepcion"],
    "ua": ["unidad agricola"], "prov": ["proveedor"], "campo": ["campo"], "cant": ["cantidad", "cant"],
    "placa": ["placa"], "guia": ["guia remision", "guia de remision", "guia"],
    "ticket": ["nro ticket", "ticket", "numero ticket"], "alm": ["almacen"], "pr": ["proceso"],
    "kp": ["peso planta"], "kc": ["peso campo"], "kb": ["peso bruto"],
}
OBLIGATORIAS = {"r": "Fecha Recojo", "ua": "Unidad Agricola", "prov": "Proveedor", "campo": "Campo", "cant": "Cantidad"}
RE_FECHA = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})(?:[ T]+(\d{1,2}):(\d{2})(?::\d{2})?\s*(?:([ap])\.?\s*m\.?)?)?$", re.I)
RE_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T]+(\d{1,2}):(\d{2}))?")


def _norm(s):
    s = "" if s is None else str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def _minutos(y, mo, d, h=0, mi=0):
    return int((datetime(y, mo, d, h, mi) - EPOCA).total_seconds() // 60)


def _fecha_hora(v):
    """Devuelve minutos desde 1970 (hora local del reporte, sin zona) o None."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return _minutos(v.year, v.month, v.day, v.hour, v.minute)
    if isinstance(v, date):
        return _minutos(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(math.floor((v - 25569) * 1440 + 0.5)) if 20000 < v < 90000 else None
    s = str(v).strip()
    m = RE_FECHA.match(s)
    h = mi = 0
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if m.group(4) is not None:
            h, mi = int(m.group(4)), int(m.group(5))
            if m.group(6):
                pm = m.group(6).lower() == "p"
                if pm and h < 12:
                    h += 12
                if not pm and h == 12:
                    h = 0
    else:
        m = RE_ISO.match(s)
        if not m:
            return None
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if m.group(4) is not None:
            h, mi = int(m.group(4)), int(m.group(5))
    if not (1 <= mo <= 12 and 1 <= d <= 31 and h <= 23 and mi <= 59):
        return None
    try:
        return _minutos(y, mo, d, h, mi)
    except ValueError:
        return None


def _txt(v):
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def _entero_si_puede(n):
    return int(n) if float(n).is_integer() else n


def _num(v):
    if v is None or v == "":
        return None
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(n):
        return None
    return _entero_si_puede(math.floor(n * 100 + 0.5) / 100)


def _cant(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return 0
    return _entero_si_puede(n) if math.isfinite(n) else 0


def dia_de(minutos):
    return (EPOCA + timedelta(minutes=minutos)).strftime("%Y-%m-%d")


def leer_excel(ruta):
    """Devuelve la lista de recojos del Excel (cada uno un dict con las claves de COLS)."""
    from openpyxl import load_workbook

    wb = load_workbook(ruta, read_only=True, data_only=True)
    faltan = None
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(values_only=True))
        for hi in range(min(len(filas), 25)):
            cab = [_norm(x) for x in (filas[hi] or [])]
            idx = {}
            for k, alias in CAMPOS.items():
                for a in alias:
                    if a in cab:
                        idx[k] = cab.index(a)
                        break
            if "r" not in idx:
                continue
            f = [OBLIGATORIAS[k] for k in OBLIGATORIAS if k not in idx]
            if f:
                faltan = f
                continue
            salida = []
            for w in filas[hi + 1:]:
                w = w or []

                def g(k):
                    return w[idx[k]] if k in idx and idx[k] < len(w) else None

                r, ua = _fecha_hora(g("r")), _txt(g("ua"))
                if r is None or not ua:
                    continue   # fila de totales o vacía
                salida.append({
                    "ua": ua, "prov": _txt(g("prov")) or "Sin proveedor", "campo": _txt(g("campo")) or "Sin campo",
                    "cant": _cant(g("cant")), "c": _fecha_hora(g("c")), "r": r, "s": _fecha_hora(g("s")), "p": _fecha_hora(g("p")),
                    "placa": _txt(g("placa")), "guia": _txt(g("guia")), "ticket": _txt(g("ticket")), "alm": _txt(g("alm")),
                    "pr": None if "pr" not in idx else (_txt(g("pr")) or "(sin proceso)"),
                    "kp": _num(g("kp")), "kc": _num(g("kc")), "kb": _num(g("kb")),
                })
            return salida
    raise Detener("Al Excel le faltan las columnas: " + ", ".join(faltan) if faltan else "El Excel no trae la columna Fecha Recojo.")


def empacar(dia, filas):
    doc = {"fecha": dia, "n": len(filas)}
    for c in COLS:
        doc[c] = [f.get(c) for f in filas]
    return doc


def desempacar(doc):
    n = len(doc.get("r") or [])
    filas = []
    for i in range(n):
        f = {c: (doc[c][i] if doc.get(c) else None) for c in COLS}
        if isinstance(f["r"], (int, float)):
            filas.append(f)
    return filas


# ============================================================
# 2. datos.json: abrir, combinar y volver a cifrar
#    (mismo formato que usa la página: PBKDF2-SHA256 + AES-GCM)
# ============================================================

def _b64(b):
    return base64.b64encode(b).decode("ascii")


def _clave(texto, sal, iteraciones):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=iteraciones).derive(texto.encode("utf-8"))


def _llave_dni(v):
    d = re.sub(r"[^0-9A-Z]", "", str(v or "").upper())
    return re.sub(r"^0+(?=\d)", "", d) if d.isdigit() else d


def abrir_datos(clave_admin):
    """Descifra datos.json con la clave de administrador. Devuelve (contenido, lista_dnis, iteraciones, sal)."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.exceptions import InvalidTag

    if not RUTA_DATOS.exists():
        raise Detener("No existe datos.json en el repositorio. Primero entra al enlace como administrador y carga al menos un Excel.")
    pk = json.loads(RUTA_DATOS.read_text(encoding="utf-8"))
    if not pk.get("adm"):
        raise Detener("datos.json no tiene sección de administrador. Vuelve a publicar desde el enlace (Accesos → Guardar y publicar).")
    it = int(pk["it"])
    adm = pk["adm"]
    try:
        k_adm = _clave("adm:" + clave_admin, base64.b64decode(adm["s"]), it)
        a = json.loads(AESGCM(k_adm).decrypt(base64.b64decode(adm["i"]), base64.b64decode(adm["c"]), None))
    except InvalidTag:
        raise Detener("La clave de administrador (secreto ADMIN_CLAVE) no coincide con la que usas en el enlace.")
    plano = AESGCM(base64.b64decode(a["k"])).decrypt(base64.b64decode(pk["iv"]), base64.b64decode(pk["data"]), None)
    if pk.get("z"):
        plano = gzip.decompress(plano)
    return json.loads(plano), list(a.get("dnis") or []), it, base64.b64decode(pk["salt"])


def guardar_datos(contenido, dnis, clave_admin, iteraciones, hoy, sal=None):
    """Vuelve a cifrar. Se conserva la misma 'sal' para que las sesiones recordadas
    en los equipos (Mantener sesión iniciada) sigan valiendo después de cada actualización."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    plano = gzip.compress(json.dumps(contenido, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), mtime=0)
    cruda = os.urandom(32)
    iv = os.urandom(12)
    datos = AESGCM(cruda).encrypt(iv, plano, None)
    sal = sal or os.urandom(16)
    llaves = []
    for dni in dnis:   # una "cerradura" por cada DNI autorizado
        kiv = os.urandom(12)
        llaves.append({"i": _b64(kiv), "c": _b64(AESGCM(_clave(_llave_dni(dni), sal, iteraciones)).encrypt(kiv, cruda, None))})
    s_adm, i_adm = os.urandom(16), os.urandom(12)
    c_adm = AESGCM(_clave("adm:" + clave_admin, s_adm, iteraciones)).encrypt(
        i_adm, json.dumps({"k": _b64(cruda), "dnis": dnis}, separators=(",", ":")).encode("utf-8"), None)
    dias = contenido["index"]["dias"]
    pk = {"v": 1, "gen": hoy.strftime("%Y-%m-%d"), "desde": dias[0] if dias else None, "hasta": dias[-1] if dias else None,
          "salt": _b64(sal), "it": iteraciones, "keys": llaves, "iv": _b64(iv), "z": True, "data": _b64(datos),
          "adm": {"s": _b64(s_adm), "i": _b64(i_adm), "c": _b64(c_adm)}}
    RUTA_DATOS.write_text(json.dumps(pk, separators=(",", ":")), encoding="utf-8")


MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
ETIQUETA_ANTIGUA = "Automático (intranet)"


def nombre_base(mes):
    """'2026-10' -> 'Octubre 2026 (automático)': el nombre de la base consolidada de ese mes."""
    return f"{MESES[int(mes[5:7]) - 1]} {mes[:4]} (automático)"


def combinar(contenido, por_dia):
    """
    Reemplaza, en cada día descargado, las unidades agrícolas que trae el Excel (igual que Cargar Excel)
    y deja UNA sola base por mes en "Bases cargadas": cada descarga se suma a la del mes en curso, y
    cuando empieza un mes nuevo se abre otra con el nombre de ese mes.
    """
    dias = contenido.setdefault("days", {})
    indice = contenido.setdefault("index", {"dias": [], "uas": [], "cargas": []})
    total = 0
    for dia, nuevas in sorted(por_dia.items()):
        previas = desempacar(dias[dia]) if dia in dias else []
        uas = {f["ua"] for f in nuevas}
        dias[dia] = empacar(dia, [f for f in previas if f["ua"] not in uas] + nuevas)
        total += len(nuevas)
    indice["dias"] = sorted(set(indice.get("dias", [])) | set(por_dia))
    indice["uas"] = sorted(set(indice.get("uas", [])) | {f["ua"] for fs in por_dia.values() for f in fs})

    ahora = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    cargas = list(indice.get("cargas", []))
    # las entradas sueltas de versiones anteriores (una por descarga) pasan a su base mensual
    meses = {d[:7] for d in por_dia}
    for c in cargas:
        if c.get("nombre") == ETIQUETA_ANTIGUA:
            for d in (c.get("desde"), c.get("hasta")):
                if d:
                    meses.add(str(d)[:7])
    cargas = [c for c in cargas if c.get("nombre") != ETIQUETA_ANTIGUA]
    for mes in sorted(meses):
        del_mes = [d for d in indice["dias"] if d[:7] == mes and d in dias]
        if not del_mes:
            continue
        nueva = {"nombre": nombre_base(mes), "filas": sum(len(dias[d].get("r") or []) for d in del_mes),
                 "desde": del_mes[0], "hasta": del_mes[-1], "cuando": ahora}
        cargas = [c for c in cargas if c.get("nombre") != nueva["nombre"]] + [nueva]
    indice["cargas"] = cargas[-40:]
    return total


# ============================================================
# 3. DESCARGA DESDE LA INTRANET (Selenium)
#
#    Esta parte es tu script recepcion_materia_prima.py, el que ya te
#    funciona: mismo ingreso, mismo recorrido por el menú, mismas listas,
#    mismas fechas y la misma espera de la tabla. Solo se cambió lo
#    necesario para que corra dentro de este programa (credenciales por
#    variables de entorno, carpeta de descarga y navegador oculto).
# ============================================================

INTRANET_URL = "https://intranet.danper.com/#/"


def swap_dia_mes(fecha_ddmmyyyy):
    d, m, y = fecha_ddmmyyyy.split("/")
    return f"{m}/{d}/{y}"


def click_robusto(driver, elemento):
    from selenium.common.exceptions import ElementClickInterceptedException
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
    time.sleep(0.3)
    try:
        elemento.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", elemento)


def click_super_robusto(driver, elemento):
    from selenium.webdriver.common.action_chains import ActionChains
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
    time.sleep(0.3)
    try:
        elemento.click()
        return
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", elemento)
        return
    except Exception:
        pass
    try:
        ActionChains(driver).move_to_element(elemento).pause(0.2).click().perform()
        return
    except Exception:
        pass
    driver.execute_script("""
        var el = arguments[0];
        ['mousedown','mouseup','click'].forEach(function(tipo){
            var evento = new MouseEvent(tipo, {bubbles: true, cancelable: true, view: window});
            el.dispatchEvent(evento);
        });
    """, elemento)


def click_menu(driver, wait, texto, espera=1.5):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    wait.until(EC.presence_of_element_located((By.XPATH, f"//*[contains(text(),'{texto}')]")))
    elementos = driver.find_elements(By.XPATH, f"//*[contains(text(),'{texto}')]")
    visibles = [e for e in elementos if e.is_displayed()]
    if not visibles:
        raise Detener(f"No se encontró ningún elemento VISIBLE con texto: {texto}")
    click_robusto(driver, visibles[0])
    time.sleep(espera)


def buscar_boton_por_texto(driver, wait, texto):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    xpath = f"//button[contains(., '{texto}')]"
    elementos = wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    visibles = [e for e in elementos if e.is_displayed()]
    if not visibles:
        raise Detener(f"No se encontró ningún botón VISIBLE con texto: {texto}")
    return visibles[0]


def llenar_fecha(driver, wait, xpath, valor_deseado_ddmmyyyy, invertido=True):
    """Mismo truco que en Cajas Confirmadas: el campo internamente espera
    mm/dd/yyyy aunque se vea dd/mm/yyyy, por eso se hace swap_dia_mes.
    (invertido=False escribe tal cual; solo se usa si el Excel sale con
    fechas que no se pidieron, ver el programa principal)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support import expected_conditions as EC
    valor_a_escribir = swap_dia_mes(valor_deseado_ddmmyyyy) if invertido else valor_deseado_ddmmyyyy
    campo = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", campo)
    campo.click()
    time.sleep(0.2)
    campo.send_keys(Keys.CONTROL, "a")
    campo.send_keys(Keys.DELETE)
    time.sleep(0.15)
    for caracter in valor_a_escribir:
        campo.send_keys(caracter)
        time.sleep(0.05)
    time.sleep(0.2)
    valor_actual = campo.get_attribute("value")
    driver.execute_script("document.activeElement.blur();")
    time.sleep(0.2)
    log(f"   -> Fecha pedida: {valor_deseado_ddmmyyyy} | se escribió: {valor_actual} | el campo quedó: {campo.get_attribute('value')}")


def seleccionar_dropdown(driver, wait, label_texto, opcion_texto, espera=1.0):
    """
    Abre un dropdown (mat-select / combo) ubicado cerca de una etiqueta
    de texto (ej. 'Tipo Cultivo') y elige la opción indicada.
    """
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    xpath_control = (
        f"//*[contains(text(),'{label_texto}')]"
        f"/ancestor::*[self::div or self::mat-form-field][1]"
        f"//*[self::mat-select or self::select or @role='combobox']"
    )
    control = wait.until(EC.presence_of_element_located((By.XPATH, xpath_control)))
    click_super_robusto(driver, control)
    time.sleep(espera)

    xpath_opcion = (
        f"//mat-option[contains(.,'{opcion_texto}')] | "
        f"//li[contains(@class,'option') and contains(.,'{opcion_texto}')] | "
        f"//option[contains(.,'{opcion_texto}')]"
    )
    try:
        opcion = wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcion)))
    except TimeoutException:
        raise Detener(
            f"No se pudo seleccionar '{opcion_texto}' en el dropdown '{label_texto}'. "
            "Es posible que el clic no haya abierto el desplegable, o que el texto de "
            "la opción no coincida exactamente."
        )
    click_super_robusto(driver, opcion)
    time.sleep(0.5)
    try:   # solo informa lo que quedó a la vista; no cambia nada
        muestra = driver.execute_script(JS_VALOR, control) or ""
    except Exception:
        muestra = "?"
    log(f"   -> Dropdown '{label_texto}' -> '{opcion_texto}' seleccionado (la lista muestra: «{' '.join(muestra.split())}»)")


def esperar_datos_en_tabla(driver, timeout=180, espera_minima=170, confirmaciones_necesarias=3):
    """
    Este reporte no muestra indicador de 'cargando'. Aquí se espera
    activamente a que APAREZCAN filas de datos en la tabla de resultados,
    revisando cada 2 segundos.

    El '0 of 0' puede aparecer momentáneamente mientras la búsqueda todavía
    está cargando. Por eso: (1) no se acepta como 'sin resultados' antes de
    'espera_minima' segundos desde el clic en Buscar, y (2) debe verse ese
    mismo estado varias veces seguidas antes de darlo por confirmado. Si en
    cualquier momento aparecen filas de datos reales, se retorna de inmediato.
    """
    from selenium.webdriver.common.by import By
    t0 = time.time()
    xpath_filas = "//table//tbody//tr | //mat-row"
    xpath_sin_resultados = (
        "//*[contains(text(),'No se encontraron registros')] | "
        "//*[contains(text(),'0 of 0')]"
    )
    confirmaciones_seguidas = 0

    while time.time() - t0 < timeout:
        try:
            filas = driver.find_elements(By.XPATH, xpath_filas)
            filas_con_texto = [f for f in filas if f.is_displayed() and f.text.strip()]
        except Exception:   # la tabla se estaba redibujando justo al revisarla
            filas_con_texto = []
        if filas_con_texto:
            transcurrido = round(time.time() - t0, 1)
            log(f"   -> Datos detectados en la tabla (tardó {transcurrido}s, {len(filas_con_texto)} filas visibles).")
            time.sleep(1)
            return True

        transcurrido = time.time() - t0
        try:
            sin_resultados = any(e.is_displayed() for e in driver.find_elements(By.XPATH, xpath_sin_resultados))
        except Exception:
            sin_resultados = False
        if transcurrido >= espera_minima and sin_resultados:
            confirmaciones_seguidas += 1
            log(f"   -> Posible 'sin resultados' ({confirmaciones_seguidas}/{confirmaciones_necesarias} confirmaciones, {round(transcurrido,1)}s transcurridos)...")
            if confirmaciones_seguidas >= confirmaciones_necesarias:
                log("   -> Confirmado: la búsqueda terminó sin resultados.")
                return False
        else:
            confirmaciones_seguidas = 0  # se reinicia si en algún momento deja de verse

        time.sleep(2)

    log(f"   -> Aviso: tras {timeout}s no se detectaron filas de datos ni se confirmó 'sin registros'.")
    return False


def esperar_descarga_y_obtener_archivo(carpeta, archivos_antes, timeout=60):
    t0 = time.time()
    while time.time() - t0 < timeout:
        actuales = set(os.listdir(carpeta))
        nuevos = actuales - archivos_antes
        nuevos_validos = [f for f in nuevos if not f.endswith(".crdownload") and not f.endswith(".tmp")]
        if nuevos_validos:
            time.sleep(1.5)
            actuales2 = set(os.listdir(carpeta))
            if any(f + ".crdownload" in actuales2 for f in nuevos_validos):
                time.sleep(1)
                continue
            nuevos_validos.sort(key=lambda f: os.path.getmtime(os.path.join(carpeta, f)), reverse=True)
            ruta = os.path.join(carpeta, nuevos_validos[0])
            log(f"   -> Descarga detectada ({os.path.getsize(ruta) // 1024} KB).")
            return ruta
        time.sleep(0.5)
    log("   -> Aviso: no se detectó ningún archivo nuevo dentro del tiempo de espera.")
    return None


# Lo que muestra una lista desplegable (solo para informar en el registro).
JS_VALOR = """
var e = arguments[0];
if (e.tagName === 'SELECT') return e.selectedIndex >= 0 ? e.options[e.selectedIndex].text : '';
var v = e.querySelector('.mat-select-value-text, .mat-mdc-select-value-text');
return v ? (v.innerText || '') : '';
"""

# Pistas de la pantalla cuando algo falla, sin datos del reporte (el registro de GitHub es público).
JS_PISTAS = """
function corto(s){ return (s||'').replace(/\\s+/g,' ').trim().slice(0,70); }
function visible(e){ return !!e.getClientRects().length; }
var r = {listas:[], casillas:[], opciones:[], paginador:''};
Array.prototype.slice.call(document.querySelectorAll('mat-select, select, [role="combobox"]')).filter(visible).slice(0,10).forEach(function(e){
  var c = e.closest('mat-form-field') || e.parentElement; r.listas.push(corto(c ? c.innerText : '')); });
Array.prototype.slice.call(document.querySelectorAll('input')).filter(visible).slice(0,12).forEach(function(e){
  r.casillas.push(corto([e.getAttribute('placeholder'), e.getAttribute('aria-label')].filter(Boolean).join(' | '))); });
Array.prototype.slice.call(document.querySelectorAll('mat-option')).filter(visible).slice(0,25).forEach(function(e){ r.opciones.push(corto(e.innerText)); });
var p = document.querySelector('.mat-paginator-range-label, .mat-mdc-paginator-range-label'); r.paginador = p ? corto(p.innerText) : '';
return r;
"""


def _diagnostico(driver):
    try:
        log(f"   Pantalla al fallar: {driver.current_url.split('#')[-1]} | título «{driver.title}»")
        p = driver.execute_script(JS_PISTAS) or {}
        log(f"   Listas: {p.get('listas')}")
        log(f"   Casillas: {p.get('casillas')}")
        log(f"   Opciones desplegadas: {p.get('opciones')} | pie de tabla: «{p.get('paginador')}»")
        carpeta = os.environ.get("DIAGNOSTICO_DIR")   # solo para pruebas en tu PC; en GitHub no se define
        if carpeta:
            Path(carpeta).mkdir(parents=True, exist_ok=True)
            driver.save_screenshot(str(Path(carpeta) / "fallo.png"))
            (Path(carpeta) / "fallo.html").write_text(driver.page_source, encoding="utf-8")
            log(f"   Captura y HTML guardados en {carpeta}")
    except Exception:
        pass


def _crear_navegador(carpeta, headless):
    """Edge, como en tu script. Si Edge no está en la máquina, se usa Chrome con las mismas opciones."""
    from selenium import webdriver

    def opciones(Options):
        options = Options()
        options.add_experimental_option("prefs", {
            "download.default_directory": str(carpeta),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })
        if headless:
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--log-level=3")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        return options

    orden = ["chrome", "edge"] if os.environ.get("NAVEGADOR", "edge").lower() == "chrome" else ["edge", "chrome"]
    error = None
    for nombre in orden:
        try:
            if nombre == "edge":
                from selenium.webdriver.edge.options import Options
                from selenium.webdriver.edge.service import Service
                driver = webdriver.Edge(service=Service(), options=opciones(Options))
            else:
                from selenium.webdriver.chrome.options import Options
                from selenium.webdriver.chrome.service import Service
                driver = webdriver.Chrome(service=Service(), options=opciones(Options))
            log(f"Navegador: {nombre.capitalize()}{' (oculto)' if headless else ''}.")
            return driver
        except Exception as e:
            error = e
            log(f"Aviso: no se pudo abrir {nombre.capitalize()} ({type(e).__name__}).")
    raise Detener(f"No se pudo abrir ningún navegador ({type(error).__name__}).")


def _abrir_navegador_y_login_recepcion(carpeta, headless=True, intentos=2):
    """
    Reintenta hasta 'intentos' veces (cerrando y volviendo a abrir el
    navegador cada vez) si la intranet no responde a tiempo en el primer
    intento. Devuelve (driver, wait) con la sesión ya lista.
    """
    from selenium.common.exceptions import WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    usuario = os.environ.get("INTRANET_USUARIO", "").strip()
    clave = os.environ.get("INTRANET_CLAVE", "").strip() or usuario
    if not usuario:
        raise Detener("Falta el secreto INTRANET_USUARIO.")
    ultimo_error = None

    for intento in range(1, intentos + 1):
        driver = None
        try:
            driver = _crear_navegador(carpeta, headless)
            try:
                driver.command_executor.set_timeout(300)
            except Exception:
                pass
            driver.set_page_load_timeout(120)
            wait = WebDriverWait(driver, 20)

            log(f"Abriendo intranet (intento {intento}/{intentos})...")
            driver.get(INTRANET_URL)
            time.sleep(3)

            log("Iniciando sesión en la intranet...")
            campo_usuario = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[contains(@placeholder,'Usuario')]")
            ))
            campo_usuario.click()
            campo_usuario.send_keys(usuario)

            campo_password = driver.find_element(By.XPATH, "//input[contains(@placeholder,'Contraseña')]")
            campo_password.click()
            campo_password.send_keys(clave)

            driver.find_element(By.XPATH, "//button[contains(.,'Acceder')]").click()

            wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Aplicaciones Danper')]")))
            log("Login exitoso, dashboard cargado.")
            time.sleep(1.5)
            return driver, wait

        except WebDriverException as e:
            ultimo_error = e
            log(f"Aviso: la intranet falló en el intento {intento}/{intentos} ({type(e).__name__}). "
                f"{'Reintentando...' if intento < intentos else 'Sin más intentos.'}")
            if driver is not None:
                _diagnostico(driver)
                try:
                    driver.quit()
                except Exception:
                    pass
            time.sleep(3)

    raise Detener(f"No se pudo iniciar sesión en la intranet ({type(ultimo_error).__name__}). "
                  "Revisa los secretos INTRANET_USUARIO e INTRANET_CLAVE, o si la intranet estaba caída.")


def descargar(inicio, fin, carpeta, visible, invertido=True):
    """
    Login a la intranet, navega hasta Recepción de Materia Prima, llena los
    filtros, espera datos y exporta a Excel. Devuelve la ruta del archivo
    descargado, o None si la búsqueda no trajo datos.
    """
    carpeta.mkdir(parents=True, exist_ok=True)
    for archivo_viejo in carpeta.glob("*.xls*"):
        archivo_viejo.unlink()

    driver, wait = _abrir_navegador_y_login_recepcion(carpeta, headless=not visible, intentos=2)
    try:
        log("Navegando: Gestión Agrícola -> Recojo Materia Prima -> Procesos y Operaciones -> Reportes -> Recepción Materia Prima")
        click_menu(driver, wait, "Gestión Agrícola")
        click_menu(driver, wait, "Recojo Materia Prima")
        click_menu(driver, wait, "Procesos y Operaciones")
        click_menu(driver, wait, "Reportes")
        click_menu(driver, wait, "Recepción Materia Prima")
        log("Página de Recepción de Materia Prima cargada.")
        time.sleep(1.5)

        log("Configurando filtros (Tipo Cultivo, Materia Prima, Origen, Fechas)...")
        seleccionar_dropdown(driver, wait, "Tipo Cultivo", TIPO_CULTIVO)
        seleccionar_dropdown(driver, wait, "Materia Prima", MATERIA_PRIMA)
        # 'Origen' se deja tal cual (por defecto "Todos"), no se toca

        xpath_fecha_inicio = "//input[contains(@aria-label,'Fecha Inicio') or contains(@placeholder,'Fecha Inicio')]"
        xpath_fecha_fin = "//input[contains(@aria-label,'Fecha Fin') or contains(@placeholder,'Fecha Fin')]"
        llenar_fecha(driver, wait, xpath_fecha_inicio, inicio.strftime("%d/%m/%Y"), invertido)
        llenar_fecha(driver, wait, xpath_fecha_fin, fin.strftime("%d/%m/%Y"), invertido)

        hay_datos = False
        intentos_busqueda = 3
        for intento in range(1, intentos_busqueda + 1):
            boton_buscar = buscar_boton_por_texto(driver, wait, "Buscar")
            click_super_robusto(driver, boton_buscar)
            log(f"Clic en Buscar enviado (intento {intento}/{intentos_busqueda}), esperando hasta 1 minuto...")

            hay_datos = esperar_datos_en_tabla(driver, timeout=60, espera_minima=50, confirmaciones_necesarias=3)
            if hay_datos:
                break
            log(f"   -> Sin datos aún tras el intento {intento}/{intentos_busqueda}.")

        if not hay_datos:
            log("No se encontraron datos en Recepción de Materia Prima para ese rango de fechas "
                f"(se reintentó el clic en Buscar {intentos_busqueda} veces).")
            _diagnostico(driver)
            return None

        archivos_antes = set(os.listdir(carpeta))
        boton_exportar = buscar_boton_por_texto(driver, wait, "Exportar a Excel")
        click_super_robusto(driver, boton_exportar)
        log("Clic en Exportar a Excel OK, esperando la descarga...")

        ruta_descargada = esperar_descarga_y_obtener_archivo(str(carpeta), archivos_antes)
        if not ruta_descargada:
            raise Detener("La descarga de Recepción de Materia Prima no se completó.")
        return Path(ruta_descargada)

    except Detener:
        _diagnostico(driver)
        raise
    except Exception as e:
        _diagnostico(driver)
        raise Detener(f"Fallo en el navegador ({type(e).__name__}). Revisa las pistas de arriba.")
    finally:
        log("Cerrando navegador de la intranet...")
        try:
            driver.quit()
        except Exception:
            pass


# ============================================================
# 4. PROGRAMA PRINCIPAL
# ============================================================

def main():
    ap = argparse.ArgumentParser(description="Descarga Recepción de Materia Prima y actualiza datos.json")
    ap.add_argument("--dias", type=int, default=int(os.environ.get("DIAS_ATRAS") or 1), help="cuántos días hacia atrás traer (1 = solo ayer)")
    ap.add_argument("--excel", help="usar este Excel en vez de descargarlo de la intranet")
    ap.add_argument("--visible", action="store_true", help="mostrar el navegador (para pruebas en tu PC)")
    ap.add_argument("--solo-descarga", action="store_true", help="descargar y resumir, sin tocar datos.json")
    ap.add_argument("--hoy", help="fecha de hoy AAAA-MM-DD (solo para pruebas)")
    ap.add_argument("--incluir-hoy", action="store_true", help="publicar también lo que va de hoy (día incompleto)")
    a = ap.parse_args()

    hoy = datetime.strptime(a.hoy, "%Y-%m-%d").date() if a.hoy else datetime.now(LIMA).date()
    dias = max(1, min(a.dias, 31))
    inicio, fin = hoy - timedelta(days=dias), hoy
    log(f"Hoy es {hoy.strftime('%d/%m/%Y')} (hora de Lima). Fecha Inicio {inicio.strftime('%d/%m/%Y')}, Fecha Fin {fin.strftime('%d/%m/%Y')}.")

    # Comprobación de fechas: lo descargado tiene que corresponder al rango pedido.
    d_ini, d_fin = (inicio - timedelta(days=1)).strftime("%Y-%m-%d"), (fin + timedelta(days=1)).strftime("%Y-%m-%d")

    def fechas_bien(fs):
        return sum(1 for f in fs if not (d_ini <= dia_de(f["r"]) <= d_fin)) <= 0.2 * len(fs)

    def rango(fs):
        vistos = sorted({dia_de(f["r"]) for f in fs})
        return f"{vistos[0]} a {vistos[-1]}"

    filas, mal = None, None
    if a.excel:
        filas = leer_excel(Path(a.excel))
        if filas and not fechas_bien(filas):
            mal, filas = rango(filas), None
    else:
        carpeta = Path(os.environ.get("DESCARGAS_DIR", "descargas")).resolve()
        # Primero se teclea la fecha como en las automatizaciones anteriores (día y mes cambiados).
        # Si el Excel no trae las fechas pedidas, se repite tecleando en el otro orden.
        for vuelta, invertido in enumerate((True, False)):
            if vuelta:
                log("Se repite la descarga escribiendo las fechas en el otro orden...")
            ruta = descargar(inicio, fin, carpeta, a.visible, invertido)
            leidas = leer_excel(ruta) if ruta else []
            if not leidas:
                continue
            if fechas_bien(leidas):
                filas, mal = leidas, None
                break
            mal = rango(leidas)
            log(f"El Excel trajo fechas que no se pidieron ({mal}): no se usa.")
    if mal:
        raise Detener(f"Las fechas del Excel no corresponden a lo pedido ({mal}). "
                      "El portal invirtió día y mes. No se publica nada.")
    if not filas:
        log("Sin datos nuevos: no se cambia nada.")
        return 0
    log(f"El Excel trae {len(filas)} recojos.")

    # Normalmente solo se publican días completos: desde Fecha Inicio hasta ayer.
    # En la corrida de la tarde se agrega lo que va de hoy; la corrida de la mañana
    # siguiente lo reemplaza por el día completo.
    incluir_hoy = a.incluir_hoy or os.environ.get("INCLUIR_HOY", "").strip().lower() in ("1", "true", "si", "sí")
    objetivo = {(inicio + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(dias + (1 if incluir_hoy else 0))}
    if incluir_hoy:
        log("Esta corrida incluye lo que va de hoy (día todavía incompleto).")
    por_dia = {}
    for f in filas:
        d = dia_de(f["r"])
        if d in objetivo:
            por_dia.setdefault(d, []).append(f)
    for d in sorted(por_dia):
        log(f"   {d[8:10]}/{d[5:7]}/{d[0:4]}: {len(por_dia[d])} recojos")
    if not por_dia:
        log("El Excel no trae recojos de los días pedidos: no se cambia nada.")
        return 0
    if a.solo_descarga:
        log("Prueba de descarga terminada (no se tocó datos.json).")
        return 0

    clave_admin = os.environ.get("ADMIN_CLAVE", "")
    if not clave_admin:
        raise Detener("Falta el secreto ADMIN_CLAVE.")
    contenido, dnis, it, sal = abrir_datos(clave_admin)
    total = combinar(contenido, por_dia)
    guardar_datos(contenido, dnis, clave_admin, it, hoy, sal)
    log(f"datos.json actualizado: {total} recojos en {len(por_dia)} día(s); {len(contenido['index']['dias'])} días en total; {len(dnis)} DNIs con acceso.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Detener as e:
        log(f"DETENIDO: {e}")
        sys.exit(1)

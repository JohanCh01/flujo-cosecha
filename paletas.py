"""
Stock de paletas (Reporte Lanzado de Paletas) — se corre al presionar «Generar tabla» en el Reporte diario
(sección Kilos en stock) y solo a las 3:02 y 4:05 a. m.

Qué hace, en orden:
1. Entra a la intranet de Danper (mismo inicio de sesión que actualizar.py) y abre
   Gestión Agrícola → Recojo Materia Prima → Reportes → Reporte Lanzado de Paletas.
2. Elige Tipo Cultivo = ESPÁRRAGO, Materia Prima = ESPÁRRAGO VERDE, Recepción = PLANTA MUCHIK,
   Desde = ayer y Hasta = hoy (hora de Lima), presiona Buscar, espera los datos y exporta el Excel.
3. Del Excel se queda solo con Stock = SI y Proceso = FRESCO, y suma el Peso Neto Paleta
   por Unidad Agrícola y Categoría (la misma tabla dinámica que haces en Excel).
4. Guarda el resultado cifrado en paletas.json, con los mismos DNIs del Flujo de cosecha,
   para que solo lo vean quienes tienen acceso.

Pensado para GitHub Actions (.github/workflows/paletas.yml), pero también corre en tu PC:

    python paletas.py --visible                 (muestra el navegador; prueba)
    python paletas.py --excel reporte.xlsx      (usa un Excel ya descargado, sin entrar a la intranet)

Credenciales: igual que actualizar.py, de las variables INTRANET_USUARIO, INTRANET_CLAVE y ADMIN_CLAVE.
"""

import argparse
import base64
import gzip
import json
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import actualizar as ac          # inicio de sesión, clics, fechas y espera de la tabla (lo que ya funciona)
import actualizar_fito as af     # lista de DNIs y cifrado (la misma sesión abre esta página)

LIMA, Detener, log, _norm = ac.LIMA, ac.Detener, ac.log, ac._norm

TIPO_CULTIVO = "ESPÁRRAGO"
MATERIA_PRIMA = "ESPÁRRAGO VERDE"
RECEPCION = "PLANTA MUCHIK"
REPORTE = "Reporte Lanzado de Paletas"
RUTA_SALIDA = Path(os.environ.get("PALETAS_JSON", "paletas.json"))
RUTA_SOLICITUD = Path("solicitud_paletas.json")

# Columnas del Excel (se aceptan variantes del nombre, sin tildes ni mayúsculas)
COLUMNAS = {
    "stock": ["stock"],
    "proceso": ["proceso"],
    "ua": ["unidad agricola", "unidad agricola descripcion", "u agricola"],
    "cat": ["categoria", "categoria campo", "categoria proveedor"],
    "peso": ["peso neto paleta", "peso neto", "neto paleta"],
}
NOMBRES = {"stock": "Stock", "proceso": "Proceso", "ua": "Unidad Agricola", "cat": "Categoria", "peso": "Peso Neto Paleta"}


# ============================================================
# 1. DESCARGA DESDE LA INTRANET
# ============================================================

def elegir(driver, wait, etiqueta, opcion):
    """Elige una opción de la lista que tiene esa etiqueta. Busca primero el control dentro del formulario
    (así «Recepción» no se confunde con los ítems del menú lateral) y, si no, usa la búsqueda de actualizar.py."""
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    xpath = (f"//mat-form-field[.//mat-label[contains(.,'{etiqueta}')] or .//label[contains(.,'{etiqueta}')]]"
             f"//*[self::mat-select or @role='combobox' or self::select]")
    try:
        control = WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except TimeoutException:
        return ac.seleccionar_dropdown(driver, wait, etiqueta, opcion)
    ac.click_super_robusto(driver, control)
    time.sleep(1.0)
    xpath_op = (f"//mat-option[normalize-space(.)='{opcion}'] | //mat-option[contains(.,'{opcion}')] | "
                f"//li[contains(@class,'option') and contains(.,'{opcion}')] | //option[contains(.,'{opcion}')]")
    try:
        op = wait.until(EC.presence_of_element_located((By.XPATH, xpath_op)))
    except TimeoutException:
        raise Detener(f"No apareció la opción «{opcion}» en la lista «{etiqueta}».")
    ac.click_super_robusto(driver, op)
    time.sleep(0.6)
    try:
        muestra = driver.execute_script(ac.JS_VALOR, control) or ""
    except Exception:
        muestra = "?"
    log(f"   -> {etiqueta}: {opcion} (la lista muestra «{' '.join(muestra.split())}»)")


def xpath_fecha(nombre):
    return (f"//input[contains(@placeholder,'{nombre}') or contains(@aria-label,'{nombre}')] | "
            f"//mat-form-field[.//mat-label[contains(.,'{nombre}')] or .//label[contains(.,'{nombre}')]]//input")


def descargar(desde, hasta, carpeta, visible, invertido=True):
    carpeta.mkdir(parents=True, exist_ok=True)
    for viejo in carpeta.glob("*.xls*"):
        viejo.unlink()
    driver, wait = ac._abrir_navegador_y_login_recepcion(carpeta, headless=not visible, intentos=2)
    try:
        log(f"Navegando: Gestión Agrícola -> Recojo Materia Prima -> Reportes -> {REPORTE}")
        ac.click_menu(driver, wait, "Gestión Agrícola")
        ac.click_menu(driver, wait, "Recojo Materia Prima")
        ac.click_menu(driver, wait, "Procesos y Operaciones")
        ac.click_menu(driver, wait, "Reportes")
        ac.click_menu(driver, wait, REPORTE)
        log("Página del reporte cargada.")
        time.sleep(1.5)

        log("Configurando filtros (Tipo Cultivo, Materia Prima, Recepción, Desde, Hasta)...")
        elegir(driver, wait, "Tipo Cultivo", TIPO_CULTIVO)
        elegir(driver, wait, "Materia Prima", MATERIA_PRIMA)
        elegir(driver, wait, "Recepción", RECEPCION)
        ac.llenar_fecha(driver, wait, xpath_fecha("Desde"), desde.strftime("%d/%m/%Y"), invertido)
        ac.llenar_fecha(driver, wait, xpath_fecha("Hasta"), hasta.strftime("%d/%m/%Y"), invertido)

        hay = False
        for intento in range(1, 4):
            ac.click_super_robusto(driver, ac.buscar_boton_por_texto(driver, wait, "Buscar"))
            log(f"Clic en Buscar (intento {intento}/3), esperando los datos...")
            hay = ac.esperar_datos_en_tabla(driver, timeout=90, espera_minima=60, confirmaciones_necesarias=3)
            if hay:
                break
        if not hay:
            log("La búsqueda no trajo datos.")
            ac._diagnostico(driver)
            return None

        antes = set(os.listdir(carpeta))
        ac.click_super_robusto(driver, ac.buscar_boton_por_texto(driver, wait, "Exportar a Excel"))
        log("Clic en Exportar a Excel, esperando la descarga...")
        ruta = ac.esperar_descarga_y_obtener_archivo(str(carpeta), antes, timeout=90)
        if not ruta:
            raise Detener("La descarga del Excel no se completó.")
        return Path(ruta)
    except Detener:
        ac._diagnostico(driver)
        raise
    except Exception as e:
        ac._diagnostico(driver)
        raise Detener(f"Fallo en el navegador ({type(e).__name__}). Revisa las pistas de arriba.")
    finally:
        log("Cerrando navegador...")
        try:
            driver.quit()
        except Exception:
            pass


# ============================================================
# 2. LECTURA DEL EXCEL Y TABLA DINÁMICA
# ============================================================

def _hojas(ruta):
    """Filas de cada hoja (xlsx con openpyxl; si no se puede, con calamine, que también lee .xls)."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(ruta, read_only=True, data_only=True)
        return [list(ws.iter_rows(values_only=True)) for ws in wb.worksheets]
    except Exception as e:
        try:
            from python_calamine import CalamineWorkbook
        except ImportError:
            raise Detener(f"No se pudo abrir el Excel ({type(e).__name__}).")
        wb = CalamineWorkbook.from_path(str(ruta))
        return [wb.get_sheet_by_name(n).to_python() for n in wb.sheet_names]


def _indice(cab):
    idx = {}
    for k, alias in COLUMNAS.items():
        for a in alias:                      # primero el nombre exacto
            if a in cab:
                idx[k] = cab.index(a)
                break
        else:
            for i, c in enumerate(cab):      # luego uno que lo contenga (p. ej. «Peso Neto Paleta (kg)»)
                if any(a in c for a in alias):
                    idx[k] = i
                    break
    return idx


def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(" ", "")
    if "," in t and "." in t:
        t = t.replace(",", "") if t.rfind(".") > t.rfind(",") else t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def leer_excel(ruta):
    """Devuelve (filas, columna_fecha) — cada fila: dict con stock, proceso, ua, cat, peso y fecha (minutos o None)."""
    faltan = None
    for filas in _hojas(ruta):
        for hi in range(min(len(filas), 25)):
            cab = [_norm(x) for x in (filas[hi] or [])]
            idx = _indice(cab)
            if "peso" not in idx and "stock" not in idx:
                continue
            f = [NOMBRES[k] for k in COLUMNAS if k not in idx]
            if f:
                faltan = f
                continue
            col_fecha = next((i for i, c in enumerate(cab) if c.startswith("fecha")), None)
            salida = []
            for w in filas[hi + 1:]:
                w = list(w or [])

                def g(k):
                    i = idx[k]
                    return w[i] if i < len(w) else None

                ua = ac._txt(g("ua"))
                if not ua:
                    continue   # fila vacía o de totales
                fecha = ac._fecha_hora(w[col_fecha]) if col_fecha is not None and col_fecha < len(w) else None
                salida.append({"stock": _norm(g("stock")).upper(), "proceso": _norm(g("proceso")).upper(),
                               "ua": ua or "(sin unidad)", "cat": ac._txt(g("cat")) or "(sin categoría)",
                               "peso": _num(g("peso")), "fecha": fecha})
            return salida, (str(filas[hi][col_fecha]) if col_fecha is not None else None)
    raise Detener("Al Excel le faltan las columnas: " + ", ".join(faltan) if faltan
                  else "No se encontró en el Excel la fila de títulos (Stock, Proceso, Unidad Agricola, Categoria, Peso Neto Paleta).")


def tabla(filas):
    """Stock = SI y Proceso = FRESCO; suma de Peso Neto Paleta por Unidad Agrícola y Categoría."""
    suma, cuenta = {}, {}
    for f in filas:
        if f["stock"] != "SI" or f["proceso"] != "FRESCO":
            continue
        k = (f["ua"], f["cat"])
        suma[k] = suma.get(k, 0.0) + f["peso"]
        cuenta[k] = cuenta.get(k, 0) + 1
    return [{"ua": ua, "cat": cat, "kg": round(suma[(ua, cat)], 2), "n": cuenta[(ua, cat)]} for ua, cat in sorted(suma)]


# ============================================================
# 3. GUARDADO CIFRADO (mismos DNIs y misma sal que datos.json)
# ============================================================

def guardar(contenido, cabecera, clave_admin):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    dnis, it, sal = af.accesos(clave_admin)
    plano = gzip.compress(json.dumps(contenido, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), mtime=0)
    cruda, iv = os.urandom(32), os.urandom(12)
    llaves = []
    for dni in dnis:   # una «cerradura» por DNI autorizado
        kiv = os.urandom(12)
        llaves.append({"i": af._b64(kiv), "c": af._b64(AESGCM(af._clave(af._llave_dni(dni), sal, it)).encrypt(kiv, cruda, None))})
    s_adm, i_adm = os.urandom(16), os.urandom(12)
    c_adm = AESGCM(af._clave("adm:" + clave_admin, s_adm, it)).encrypt(
        i_adm, json.dumps({"k": af._b64(cruda)}, separators=(",", ":")).encode("utf-8"), None)
    pk = dict(cabecera)
    pk.update({"v": 1, "salt": af._b64(sal), "it": it, "keys": llaves, "iv": af._b64(iv), "z": True,
               "data": af._b64(AESGCM(cruda).encrypt(iv, plano, None)),
               "adm": {"s": af._b64(s_adm), "i": af._b64(i_adm), "c": af._b64(c_adm)}})
    RUTA_SALIDA.write_text(json.dumps(pk, separators=(",", ":")), encoding="utf-8")
    return len(dnis)


def anotar_error(cabecera, mensaje):
    """Deja el aviso de error para la página, sin borrar la última tabla buena."""
    try:
        pk = json.loads(RUTA_SALIDA.read_text(encoding="utf-8"))
    except Exception:
        pk = {"v": 1}
    pk.update(cabecera)
    pk.update({"estado": "error", "msg": mensaje[:300]})
    RUTA_SALIDA.write_text(json.dumps(pk, separators=(",", ":")), encoding="utf-8")


# ============================================================
# 4. PROGRAMA PRINCIPAL
# ============================================================

def main():
    ap = argparse.ArgumentParser(description="Descarga el Reporte Lanzado de Paletas y arma la tabla de stock")
    ap.add_argument("--excel", help="usar este Excel en vez de descargarlo")
    ap.add_argument("--visible", action="store_true", help="mostrar el navegador (pruebas en tu PC)")
    ap.add_argument("--hoy", help="fecha de hoy AAAA-MM-DD (solo para pruebas)")
    ap.add_argument("--sin-guardar", action="store_true", help="solo mostrar la tabla, sin escribir paletas.json")
    a = ap.parse_args()

    ahora = datetime.now(LIMA)
    hoy = datetime.strptime(a.hoy, "%Y-%m-%d").date() if a.hoy else ahora.date()
    desde, hasta = hoy - timedelta(days=1), hoy
    try:
        sol = json.loads(RUTA_SOLICITUD.read_text(encoding="utf-8")).get("id")
    except Exception:
        sol = None
    cabecera = {"gen": ahora.strftime("%Y-%m-%dT%H:%M"), "sol": sol, "desde": desde.isoformat(), "hasta": hasta.isoformat()}
    log(f"Desde {desde:%d/%m/%Y} hasta {hasta:%d/%m/%Y} (hora de Lima).")

    clave_admin = os.environ.get("ADMIN_CLAVE", "")
    if not clave_admin and not a.sin_guardar:
        raise Detener("Falta el secreto ADMIN_CLAVE.")

    try:
        d_ini = (desde - timedelta(days=1)).isoformat()
        d_fin = (hasta + timedelta(days=1)).isoformat()

        def fechas_bien(fs):
            con = [f for f in fs if f["fecha"] is not None]
            return not con or sum(1 for f in con if not (d_ini <= ac.dia_de(f["fecha"]) <= d_fin)) <= 0.2 * len(con)

        filas, mal = None, None
        if a.excel:
            filas, _ = leer_excel(Path(a.excel))
        else:
            carpeta = Path(os.environ.get("DESCARGAS_DIR", "descargas")).resolve()
            # Primero se teclea la fecha como en la recepción (día y mes cambiados); si no sale bien, al revés.
            for vuelta, invertido in enumerate((True, False)):
                if vuelta:
                    log("Se repite escribiendo las fechas en el otro orden...")
                ruta = descargar(desde, hasta, carpeta, a.visible, invertido)
                leidas, col = leer_excel(ruta) if ruta else ([], None)
                if not leidas:
                    continue
                if fechas_bien(leidas):
                    filas, mal = leidas, None
                    log(f"Fechas comprobadas con la columna «{col}»." if col else "El Excel no trae columna de fecha para comprobar.")
                    break
                mal = "el portal invirtió día y mes"
                log("El Excel trae fechas que no se pidieron: no se usa.")
        if mal:
            raise Detener(f"Las fechas del Excel no corresponden a lo pedido ({mal}).")
        filas = filas or []
        t = tabla(filas)
        total = round(sum(x["kg"] for x in t), 2)
        log(f"El Excel trae {len(filas)} paletas; con Stock = SI y Proceso = FRESCO quedan {sum(x['n'] for x in t)} "
            f"({total:,.1f} kg) en {len({x['ua'] for x in t})} unidades agrícolas.")
        if a.sin_guardar:
            for x in t:
                log(f"   {x['ua']:<28} {x['cat']:<14} {x['kg']:>10,.1f}")
            return 0
        contenido = {"gen": cabecera["gen"], "desde": cabecera["desde"], "hasta": cabecera["hasta"],
                     "filasExcel": len(filas), "tabla": t, "total": total,
                     "otros": {"stock": sorted({f["stock"] for f in filas}), "proceso": sorted({f["proceso"] for f in filas})}}
        cabecera.update({"estado": "ok", "msg": "" if filas else "La búsqueda no trajo datos para esas fechas."})
        n = guardar(contenido, cabecera, clave_admin)
        log(f"paletas.json guardado ({n} DNIs con acceso).")
        return 0
    except Detener as e:
        if not a.sin_guardar:
            anotar_error(cabecera, str(e))
        raise


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Detener as e:
        log(f"DETENIDO: {e}")
        sys.exit(1)

"""
Despachos · Mercado (diapositiva 5) desde el Excel de ventas (hoja «Data»).

Corre en TU LAPTOP. Lee el Excel, se queda solo con las filas de ESPÁRRAGO VERDE y guarda de cada fila lo necesario
(Año Embarque, envío AE/MA, mercado o cliente, Fcls) dentro de fito.json (cifrado). Así la página ya lo tiene cargado
para todos, sin arrastrar nada. Si lo corres con un Excel más nuevo, compara fila por fila con lo que ya estaba
y te dice cuántas filas son nuevas, cuántas cambiaron y cuántas ya no están; luego lo actualiza.

Reglas (iguales a tus columnas):
  ENV  = Tipo Envio, con MU → AE
  MERC = América del Norte / América del Sur → USA · Asia · Cen.America&el Carib → LATAM · Europa
  Cada fila se reconoce por: Contrato + Posición Contrato + Plan Entrega + Posición Plan Entrega + Pedido Vta. + Posición Ped.Vta

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python ventas_excel.py                      (usa el Excel de ventas más reciente de Descargas/Escritorio/Documentos)
    python ventas_excel.py "C:/Users/jchinguel/Downloads/vetnas_15.10.xlsx"   (o uno en particular, entre comillas)
    python ventas_excel.py "…xlsx" --sin-subir      (solo muestra las tablas y las diferencias; no sube nada)
Necesita: pip install python-calamine   (si no lo tienes, usa openpyxl, más lento)
"""

import argparse
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import recepcion_sql as rs          # misma clave de administrador, mismo cifrado y misma subida a GitHub
af, log, git = rs.af, rs.log, rs.git

LLAVE = ["CONTRATO", "POSICIONCONTRATO", "PLANENTREGA", "POSICIONPLANENTREGA", "PEDIDOVTA", "POSICIONPEDVTA"]
MERC_ORDEN = ["Asia", "Europa", "LATAM", "USA"]


def norm(t):
    t = unicodedata.normalize("NFD", str(t if t is not None else ""))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").strip().upper()


def compacto(t):
    return "".join(c for c in norm(t) if c.isalnum())


def texto(v):
    """Igual en Python y en la página: 1.0 → «1», textos sin espacios a los lados."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v if v is not None else "").strip()


def merc_de(m):
    n = norm(m)
    if "NORTE" in n or "SUR" in n:
        return "USA"
    if "ASIA" in n:
        return "Asia"
    if n.startswith("CEN") or "CARIB" in n:
        return "LATAM"
    if "EUROPA" in n:
        return "Europa"
    return ""


def leer_hoja(ruta):
    try:
        from python_calamine import CalamineWorkbook
        wb = CalamineWorkbook.from_path(str(ruta))
        hoja = next((n for n in wb.sheet_names if norm(n) == "DATA"), None)
        if not hoja:
            sys.exit("El Excel no tiene la hoja «Data».")
        return wb.get_sheet_by_name(hoja).to_python()
    except ImportError:
        import openpyxl
        log("No tienes python-calamine (pip install python-calamine); leo con openpyxl, tarda un poco más…")
        wb = openpyxl.load_workbook(str(ruta), read_only=True, data_only=True)
        hoja = next((n for n in wb.sheetnames if norm(n) == "DATA"), None)
        if not hoja:
            sys.exit("El Excel no tiene la hoja «Data».")
        return [list(f) for f in wb[hoja].iter_rows(values_only=True)]


def leer_filas(ruta):
    """{llave: [año, 'AE'/'MA', mercado o cliente, fcls]} solo de ESPÁRRAGO VERDE."""
    d = leer_hoja(ruta)
    enc = [compacto(h) for h in d[0]]
    def col(*nombres):
        for n in nombres:
            if n in enc:
                return enc.index(n)
        return -1
    c = {"anio": col("ANOEMBARQUE"), "grupo": col("GRUPOARTICULO"), "env": col("TIPOENVIO"), "merc": col("MERCADO"),
         "cli": col("CLIENTEDANPER"), "fcls": col("FCLS")}
    llave = [col(n) for n in LLAVE]
    falta = [k for k, v in c.items() if v < 0] + [LLAVE[i] for i, v in enumerate(llave) if v < 0]
    if falta:
        sys.exit(f"Faltan columnas en la hoja Data: {', '.join(falta)}")
    filas = {}
    for r in d[1:]:
        if norm(r[c["grupo"]]) != "ESPARRAGO VERDE":
            continue
        anio = texto(r[c["anio"]])
        if not (len(anio) == 4 and anio.isdigit()):
            continue
        env = texto(r[c["env"]]).upper()
        env = "AE" if env == "MU" else env
        if env not in ("AE", "MA"):
            continue
        grp = merc_de(r[c["merc"]]) if env == "AE" else (texto(r[c["cli"]]) or "(sin cliente)")
        if not grp:
            continue
        try:
            fc = float(str(r[c["fcls"]]).replace(",", ".") or 0)
        except ValueError:
            fc = 0.0
        k = "|".join(texto(r[i]) for i in llave)
        if k in filas:                                   # por si acaso: si una llave se repite, se suma
            filas[k][3] = round(filas[k][3] + fc, 6)
        else:
            filas[k] = [int(anio), env, grp, round(fc, 6)]
    return filas


def tablas(filas, anio):
    ae, ma = {}, {}
    for a, env, grp, fc in filas.values():
        if a != anio:
            continue
        (ae if env == "AE" else ma)[grp] = (ae if env == "AE" else ma).get(grp, 0) + fc
    return ae, ma


def mostrar(filas, anio):
    ae, ma = tablas(filas, anio)
    print(f"\nAÉREO {anio}  (Mercado · N° Fcls)")
    for m in MERC_ORDEN:
        print(f"   {m:<10} {round(ae.get(m, 0)):>6}")
    print(f"   {'Total':<10} {round(sum(ae.values())):>6}")
    print(f"\nMARÍTIMO {anio}  (Cliente · N° Fcls)")
    for k in sorted(ma):
        print(f"   {k:<10} {round(ma[k]):>6}")
    print(f"   {'Total':<10} {round(sum(ma.values())):>6}\n")


def comparar(viejas, nuevas):
    nuevas_k = [k for k in nuevas if k not in viejas]
    cambiadas = [k for k in nuevas if k in viejas and viejas[k] != nuevas[k]]
    fuera = [k for k in viejas if k not in nuevas]
    return nuevas_k, cambiadas, fuera


def buscar_excel():
    """El .xlsx más reciente cuyo nombre diga «venta» o «vetna» en Descargas, Escritorio o Documentos (y OneDrive)."""
    casa = Path.home()
    carpetas = [casa / n for n in ("Downloads", "Descargas", "Desktop", "Escritorio", "Documents", "Documentos")]
    carpetas += [d / n for d in casa.glob("OneDrive*") for n in ("Desktop", "Escritorio", "Documents", "Documentos")]
    hallados = []
    for c in carpetas:
        if c.is_dir():
            for f in c.glob("*.xls*"):
                n = f.name.lower()
                if ("venta" in n or "vetna" in n) and not n.startswith("~$"):
                    hallados.append(f)
    return max(hallados, key=lambda f: f.stat().st_mtime) if hallados else None


def main():
    ap = argparse.ArgumentParser(description="Carga en la página el Excel de ventas (Despachos · Mercado)")
    ap.add_argument("excel", nargs="?", help="ruta del Excel; si no la pones, busca el más reciente en Descargas, Escritorio y Documentos")
    ap.add_argument("--anio", type=int, default=datetime.now().year)
    ap.add_argument("--sin-subir", action="store_true")
    a = ap.parse_args()
    ruta = Path(a.excel) if a.excel else buscar_excel()
    if ruta is None:
        sys.exit("No encontré ningún Excel de ventas en Descargas, Escritorio ni Documentos.\n"
                 'Ponle la ruta entre comillas, por ejemplo:  python ventas_excel.py "C:/Users/jchinguel/Downloads/vetnas_15.10_1.xlsx"')
    if not ruta.exists():
        sys.exit(f"No encuentro el archivo: {ruta}")
    if not a.excel:
        log(f"Uso el Excel de ventas más reciente que encontré: {ruta}")

    log(f"Leyendo {ruta.name}…")
    filas = leer_filas(ruta)
    log(f"{len(filas):,} filas de ESPÁRRAGO VERDE.")
    mostrar(filas, a.anio)

    clave = None if a.sin_subir else rs.clave_admin()
    if clave:
        git("pull", "--ff-only", "-q")
    contenido = af.abrir_fito(clave)[0] if clave else None
    viejas = ((contenido or {}).get("desp") or {}).get("filas") or {}
    if viejas:
        viejas = {k: list(v) for k, v in viejas.items()}
        nu, ca, fu = comparar(viejas, filas)
        log(f"Comparado con lo que ya estaba cargado: {len(nu)} filas nuevas, {len(ca)} cambiadas, {len(fu)} que ya no están.")
    if a.sin_subir:
        return

    paquete = {"archivo": ruta.name, "gen": datetime.now().strftime("%Y-%m-%dT%H:%M"), "filas": filas}
    for _ in range(3):
        git("pull", "--ff-only", "-q")
        dnis, it, sal = af.accesos(clave)
        contenido, _ = af.abrir_fito(clave)
        if not contenido:
            sys.exit("No pude abrir fito.json con esa clave (¿es la clave de administrador correcta?).")
        contenido["desp"] = paquete
        contenido["gen"] = datetime.now().strftime("%Y-%m-%dT%H:%M")
        af.guardar_fito(contenido, dnis, clave, it, sal)
        if not git("status", "--porcelain", "--", "fito.json"):
            log("fito.json ya tenía estos mismos datos: no hay nada que subir.")
            return
        git("add", "fito.json")
        git("commit", "-q", "-m", f"Despachos por mercado desde el Excel de ventas ({datetime.now():%d/%m/%Y %H:%M})")
        try:
            git("push", "-q")
            log("Listo: subido a GitHub. En 1–2 minutos ya sale cargado en la página para todos (botón Actualizar).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo; vuelvo a intentarlo…")
            git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
COMPARATIVO - Espárrago Verde NV05  (para comparativo.html)
===========================================================
Junta todo lo que se cruza por Pedido de Venta y lo publica en la página «Comparativo»:

  - SAP (carpeta SAPP)                         -> datos del pedido
  - TRACKING marítimo + TRACKING AÉREOS        -> N° FCL/AWB, ETA, tipo de envío
  - Matrices de despacho (marítimo / aéreo)    -> resultado de inspección, status, reclamo
  - SQL HECHO_RESUMEN_PRODUCCION               -> procedencia y días cosecha-despacho
  - SQL QM Cloud (Inspección PT y Empaque)     -> C-% prom / máx / mín

Lo que otras personas suben desde la página («Cargar archivos») llega a esas mismas carpetas
del Drive, así que entra en la siguiente corrida.

Corre en TU LAPTOP (el SQL y las matrices solo se alcanzan desde la red de Danper; GitHub no llega).

Regla única del Pedido de Venta (PV): solo dígitos, como texto, y si viene
corto (ej. 69294) se completa a 10 dígitos con "11" adelante -> 1100069294.

Qué deja:
  1. comparativo.json en GitHub, cifrado con los mismos DNIs del Flujo de cosecha (solo lo abre quien entra).
  2. BASE_COMPARATIVO.xlsx en tu Escritorio (se reemplaza en cada corrida), con estas hojas:
       PEDIDOS              una fila por PV con cabecera + todos los indicadores
       INDICADORES_LARGO    lo mismo en formato largo (Grupo / Medida / Valor) para la matriz
       PRODUCCION, INSP_PT, INSP_EM, TRACKING   detalle
       REVISAR_PV           PV que no cruzan entre fuentes
       REVISAR_PROCEDENCIA  cada Razón Social y cómo se clasificó

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python comparativo.py                 (arma el Excel y lo publica en la página)
    python comparativo.py --sin-subir     (solo el Excel del Escritorio; no publica nada)

Credenciales (nunca van escritas aquí: el repositorio es público):
    SQL                ~/.flujo_cosecha_sql    (el mismo de recepcion_sql.py; se pide la primera vez)
    Administrador      ~/.flujo_cosecha_clave  (la misma clave del Flujo de cosecha)
    Clave de subida    ~/.flujo_cosecha_subida (la CLAVE del Apps Script «Subida comparativo»)

Requisitos:  pip install pandas openpyxl pyodbc cryptography
"""

import argparse
import getpass
import os
import re
import subprocess
import sys
import warnings
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

AQUI = Path(__file__).resolve().parent
os.chdir(AQUI)                      # para encontrar datos.json aunque lo lance el Programador de tareas
sys.path.insert(0, str(AQUI))
import actualizar_fito as af        # mismo cifrado que las otras páginas  # noqa: E402
import recepcion_sql as rs          # misma conexión y credenciales del SQL  # noqa: E402

LIMA = timezone(timedelta(hours=-5))

# Avisos inofensivos: formato condicional de los Excel y pandas usando pyodbc directo
warnings.filterwarnings("ignore", message=".*Conditional Formatting extension.*")
warnings.filterwarnings("ignore", message=".*pandas only supports SQLAlchemy.*")

# =============================================================================
# CONFIGURACIÓN  (lo único que normalmente hay que tocar)
# =============================================================================
BASE = Path(r"G:\.shortcut-targets-by-id\18cXmoy85ttslPU8TI-aFQtBDWb3jT33t\NV05"
            r"\SEGUIMIENTO DE CONTENEDORES ESPARRAGO VERDE   2026")

CARPETA_SAP = BASE / "SAPP"
CARPETA_TRACKING_MAR = BASE / "TRACKING"
CARPETA_TRACKING_AER = BASE / "TRACKING AÉREOS"

# Hoja a leer en cada Excel de carpeta: 0 = primera hoja, o el nombre entre comillas
HOJA_SAP = 0
HOJA_TRACKING_MAR = 0
HOJA_TRACKING_AER = 0
FILAS_TITULO_TRACKING = 3          # filas de título antes del encabezado

MATRIZ_DIR = Path(r"\\danper.net\fs\TRU\DAAC\ACTUAL\AC FRESCO\Informacion A.C. P.F"
                  r"\ESPÁRRAGO\4. MATRIZ DESPACHO")
MATRICES = [
    MATRIZ_DIR / "ESPÁRRAGO VERDE_ MATRIZ DE DESPACHOS MARITIMOS (ACT.2026).xlsx",
    MATRIZ_DIR / "ESPÁRRAGO VERDE_ AÉREOS (ACT.2026).xlsx",
    # MATRIZ_DIR / "NOMBRE DE LA MATRIZ 2025.xlsx",   # <- agregar la de 2025
]
HOJA_MATRIZ = "Hoja1"

# Llaves de SAP que se excluyen a mano (fecha|pedido|material|PO)
LLAVES_EXCLUIDAS = ["28/07/2026|1100069535|70016251|4000004382"]

# Servicio de subida (Apps Script «Subida comparativo» en autom.danper@gmail.com).
# La URL va dentro de comparativo.json cifrado; la clave se lee de ~/.flujo_cosecha_subida.
SUBIDA_URL = ("https://script.google.com/macros/s/AKfycbzJSkzV-GjobwTxIAAWuwks6BqNyUoAa1Ib0YIcgsjEpd"
              "yEYnk-bm-Kb75CjTBQkXXK/exec")
ARCHIVO_SUBIDA = Path.home() / ".flujo_cosecha_subida"
RUTA_JSON = Path("comparativo.json")

TIPO_CARTILLA = "Aseguramiento de la Calidad Fresco Espárrago"
CARTILLA_PT = "01 INSPECCIÓN PT - NAVE 05 EVD"
CARTILLA_EM = "02 CALIFICACIÓN PRODUCTO EMPACADO"
ANIO_DESDE_PT = 2025
ANIO_DESDE_EM = 2026

# Clasificación de procedencia por Razón Social (lo que no esté aquí = TERCEROS).
# Revisa la hoja REVISAR_PROCEDENCIA y agrega aquí cualquier nombre que falte.
PROCEDENCIA = {
    "COMPOSITAN": ["COMPOSITAN 1"],
    "OLMOS": ["OLMOS 1", "OLMOS 2", "OLMOS 3"],
    "RE-PROCESO": ["RE-PROCESO"],
}

# Códigos de "Tipo Envio" de SAP -> nombre (se usa cuando el pedido no está en el tracking)
TIPO_ENVIO_SAP = {
    "AE": "Aéreo",
    "MA": "Marítimo",
    "MU": "Multimodal",     # <- confirmar qué significa MU
}

# Días máximos válidos (igual que en las columnas DAX)
DIAS_MAX = 60

# Primer día de la semana para la columna Semana (0 = lunes, 6 = domingo)
PRIMER_DIA_SEMANA = 6

def _escritorio():
    """Busca el Escritorio, también si Windows lo tiene dentro de OneDrive."""
    home = Path.home()
    candidatos = [home / "Desktop", home / "Escritorio"]
    candidatos += sorted(home.glob("OneDrive*/Desktop")) + sorted(home.glob("OneDrive*/Escritorio"))
    for c in candidatos:
        if c.is_dir():
            return c
    return home / "Desktop"


# Por ahora se guarda en el Escritorio y se reescribe en cada corrida.
# Para volver a la carpeta del Drive:  SALIDA = BASE / "CONSOLIDADO" / "BASE_COMPARATIVO.xlsx"
SALIDA = _escritorio() / "BASE_COMPARATIVO.xlsx"

# =============================================================================
# FUNCIONES DE APOYO
# =============================================================================
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
         "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def limpiar_columnas(df):
    """'PV \\nSAP' -> 'PV SAP', 'FECHA  DE DESPACHO' -> 'FECHA DE DESPACHO'."""
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    return df


def texto_entero(v):
    """1100069294.0 -> '1100069294'; textos quedan igual sin espacios extremos."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s or None


def normalizar_pv(v):
    s = texto_entero(v)
    if not s:
        return None
    s = re.sub(r"\D", "", s)
    if not s:
        return None
    if len(s) <= 8:                       # PV corto del tracking -> 11 + 8 dígitos
        s = "11" + s.zfill(8)
    return s


def a_fecha_valor(v):
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).normalize()
    if isinstance(v, (int, float)) and 20000 < v < 80000:     # fecha serial de Excel
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v))
    s = str(v).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):                 # formato SQL: 2026-09-29
        t = pd.to_datetime(s[:10], format="%Y-%m-%d", errors="coerce")
    else:                                                  # formato peruano: 29/09/2026
        t = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return t.normalize() if pd.notna(t) else pd.NaT


def a_fecha(serie):
    return pd.to_datetime(serie.map(a_fecha_valor), errors="coerce")


def a_numero(serie):
    return pd.to_numeric(serie.astype(str).str.strip().str.replace(",", ".", regex=False),
                         errors="coerce")


def semana(f):
    if pd.isna(f):
        return None
    ene1 = pd.Timestamp(f.year, 1, 1)
    desfase = (ene1.weekday() - PRIMER_DIA_SEMANA) % 7
    return f"S{(f.dayofyear - 1 + desfase) // 7 + 1:02d}"


FUENTES = []   # archivos usados en la corrida (se muestran en el panel «Fuentes» de la página)


def anotar_fuente(grupo, ruta, leido=True):
    try:
        mod = datetime.fromtimestamp(Path(ruta).stat().st_mtime, LIMA).isoformat(timespec="seconds")
    except OSError:
        mod = None
    FUENTES.append({"g": grupo, "a": Path(ruta).name, "m": mod, "ok": bool(leido)})


def leer_carpeta(carpeta, hoja, filas_titulo=0, grupo=""):
    """Lee todos los Excel de una carpeta (el más reciente primero) y los une por NOMBRE de columna."""
    carpeta = Path(carpeta)
    archivos = [p for p in carpeta.glob("*")
                if p.suffix.lower() in (".xlsx", ".xlsm", ".xls") and not p.name.startswith("~$")]
    if not archivos:
        raise FileNotFoundError(f"No hay archivos Excel en {carpeta}")
    archivos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    partes = []
    for p in archivos:
        try:
            df = pd.read_excel(p, sheet_name=hoja, header=filas_titulo)
        except Exception as e:
            log(f"  ! No se pudo leer {p.name}: {e}")
            anotar_fuente(grupo, p, leido=False)
            continue
        anotar_fuente(grupo, p)
        df = limpiar_columnas(df)
        df.insert(0, "Archivo", p.name)
        partes.append(df)
    if not partes:
        raise RuntimeError(f"Ningún archivo de {carpeta} se pudo leer")
    return pd.concat(partes, ignore_index=True)


def exigir_columnas(df, columnas, fuente):
    faltan = [c for c in columnas if c not in df.columns]
    if faltan:
        raise KeyError(f"En {fuente} no se encontraron las columnas: {faltan}\n"
                       f"Columnas disponibles: {list(df.columns)}")


# =============================================================================
# CARGA DE FUENTES
# =============================================================================
def cargar_sap():
    log("Leyendo SAP...")
    df = leer_carpeta(CARPETA_SAP, HOJA_SAP, grupo="SAP")
    exigir_columnas(df, ["Grupo Material", "Fecha Despacho PL", "Fecha ETD", "Pedido Vta",
                         "Material", "PO CLIENTE", "Cliente Danper", "Tipo Envio"], "SAP")
    if "Conteo" not in df.columns:
        df["Conteo"] = None
    df = df[df["Grupo Material"].astype(str).str.strip() == "ESPÁRRAGO VERDE"].copy()
    df["Fecha Despacho PL"] = a_fecha(df["Fecha Despacho PL"])
    df["Fecha ETD"] = a_fecha(df["Fecha ETD"])
    df = df[df["Fecha Despacho PL"].notna()].copy()
    df["PV"] = df["Pedido Vta"].map(normalizar_pv)
    df["Material"] = pd.to_numeric(df["Material"], errors="coerce")
    df["LlaveUnica"] = (df["Fecha Despacho PL"].dt.strftime("%d/%m/%Y") + "|"
                        + df["PV"].fillna("") + "|"
                        + df["Material"].map(texto_entero).fillna("") + "|"
                        + df["PO CLIENTE"].map(texto_entero).fillna(""))
    df = df.drop_duplicates("LlaveUnica")
    df = df[~df["LlaveUnica"].isin(LLAVES_EXCLUIDAS)]
    log(f"  SAP: {len(df)} filas, {df['PV'].nunique()} pedidos")
    return df


def cargar_tracking():
    log("Leyendo TRACKING marítimo...")
    mar = leer_carpeta(CARPETA_TRACKING_MAR, HOJA_TRACKING_MAR, FILAS_TITULO_TRACKING, grupo="Tracking marítimo")
    exigir_columnas(mar, ["MES", "CORR.", "PV SAP", "N° FCL", "NAVIERA",
                          "FECHA INICIAL DE LLEGADA A DESTINO"], "TRACKING marítimo")
    mar = mar[mar["MES"].notna() & (mar["MES"].astype(str).str.strip() != "MES")].copy()
    mar["CORR."] = pd.to_numeric(mar["CORR."], errors="coerce")
    mar = mar.drop_duplicates("CORR.")
    mar = mar.rename(columns={"N° FCL": "N° FCL/AWB", "NAVIERA": "LINEA AEREA/NAVIERA"})
    mar["Tipo de Envío"] = "Marítimo"

    log("Leyendo TRACKING aéreos...")
    aer = leer_carpeta(CARPETA_TRACKING_AER, HOJA_TRACKING_AER, FILAS_TITULO_TRACKING, grupo="Tracking aéreo")
    exigir_columnas(aer, ["MES", "CORR.", "PV SAP", "N° AWB", "LINEA AEREA",
                          "FECHA INICIAL DE LLEGADA A DESTINO",
                          "FECHA REAL DE LLEGADA A DESTINO"], "TRACKING aéreos")
    aer = aer[aer["MES"].notna() & (aer["MES"].astype(str).str.strip() != "MES")].copy()
    aer = aer[aer["FECHA REAL DE LLEGADA A DESTINO"].astype(str).str.strip() != "-"]
    aer["CORR."] = pd.to_numeric(aer["CORR."], errors="coerce")
    aer = aer.drop_duplicates("CORR.")
    aer = aer.rename(columns={"N° AWB": "N° FCL/AWB", "LINEA AEREA": "LINEA AEREA/NAVIERA"})
    aer["Tipo de Envío"] = "Aéreo"

    trk = pd.concat([mar, aer], ignore_index=True)
    trk["PV"] = trk["PV SAP"].map(normalizar_pv)
    for c in trk.columns:
        if "FECHA" in c.upper():
            trk[c] = a_fecha(trk[c])            # textos como "LIMA" o "-" quedan vacíos
    trk["N° FCL/AWB"] = trk["N° FCL/AWB"].map(texto_entero)
    log(f"  TRACKING: {len(trk)} filas, {trk['PV'].nunique()} pedidos")
    return trk


def cargar_matrices():
    log("Leyendo matrices de despacho...")
    partes = []
    for ruta in MATRICES:
        if not Path(ruta).exists():
            log(f"  ! No existe {ruta}")
            continue
        df = limpiar_columnas(pd.read_excel(ruta, sheet_name=HOJA_MATRIZ))
        df["Archivo"] = Path(ruta).name
        anotar_fuente("Matriz de despachos", ruta)
        partes.append(df)
    if not partes:
        log("  ! Ninguna matriz leída: Resultado/Status quedarán vacíos")
        return pd.DataFrame(columns=["PV_N", "RESULTADO DE INSPECCIÓN", "STATUS FINAL", "Reclamo"])
    mat = pd.concat(partes, ignore_index=True)
    exigir_columnas(mat, ["PV"], "matrices")
    mat["PV_N"] = mat["PV"].map(normalizar_pv)
    mat = mat[mat["PV_N"].notna()].copy()
    for c in ["RESULTADO DE INSPECCIÓN", "STATUS FINAL", "TIPO DE RECLAMO"]:
        if c not in mat.columns:
            mat[c] = None
        mat[c] = mat[c].map(lambda v: None if v is None or (isinstance(v, float) and pd.isna(v))
                            or str(v).strip() in ("", "-") else str(v).strip())
    mat["Reclamo"] = mat["TIPO DE RECLAMO"].map(
        lambda v: None if v is None else ("Sin reclamo" if v == "Sin reclamo" else "Con Reclamo"))
    log(f"  Matrices: {len(mat)} filas")
    return mat


def conectar_sql():
    """La misma conexión de recepcion_sql.py: usuario y contraseña de ~/.flujo_cosecha_sql."""
    return rs.conectar()


def leer_produccion(conn):
    q = ("SELECT PedidoVenta, RazonSocial, Cantidad, FechaCosecha, ZFER "
         "FROM HECHO_RESUMEN_PRODUCCION")
    return pd.read_sql(q, conn)


VAR_PEDIDO = "Cabecera#Pedido de Venta"
VAR_SEMILLADO = "Calidad de Punta#Semillado (C-)"
VAR_TAMANO = "Tamaño de Muestra#Tamaño de Muestra"


def leer_qm(conn, cartilla, anio_desde):
    # Se filtra por fecha y por cada variable por separado (sin YEAR() ni concatenar en el WHERE),
    # así el servidor puede usar sus índices y la consulta es más rápida. El resultado es el mismo.
    variables = [v.split("#", 1) for v in (VAR_PEDIDO, VAR_SEMILLADO, VAR_TAMANO)]
    q = ("SELECT InspeccionID, GrupoVariable + '#' + VariableNombre AS Combinada, "
         "VariableValor, NroMuestra "
         "FROM HECHO_DATOS_QMCLOUD_CALIDAD "
         "WHERE NombreCartilla = ? AND NombreTipoCartilla = ? "
         "AND FechaHoraMuestra >= ? "
         "AND (" + " OR ".join(["(GrupoVariable = ? AND VariableNombre = ?)"] * len(variables)) + ")")
    params = [cartilla, TIPO_CARTILLA, datetime(anio_desde, 1, 1)] + [x for g in variables for x in g]
    return pd.read_sql(q, conn, params=params)


SQL_TIEMPO_MAX = 900   # segundos por consulta; si el servidor tarda más, se detiene con un aviso


def consulta(nombre, funcion, *args):
    log(f"  {nombre}…")
    t0 = datetime.now()
    try:
        df = funcion(*args)
    except Exception as e:
        if "HYT00" in str(e) or "timeout" in str(e).lower():
            raise RuntimeError(f"El SQL tardó más de {SQL_TIEMPO_MAX // 60} minutos en «{nombre}». "
                               "Puede estar muy ocupado o la conexión está lenta: inténtalo más tarde, conectado al wifi de Danper.")
        raise
    log(f"    listo: {len(df):,} filas en {(datetime.now() - t0).seconds} s")
    return df


def armar_inspeccion(df, decimales):
    """Una fila por inspección: PV y C-% = Semillado / Tamaño de muestra (muestra 1)."""
    df = df.copy()
    df["InspeccionID"] = df["InspeccionID"].map(texto_entero)
    df = df[df["InspeccionID"].notna()]
    df["NroMuestra"] = df["NroMuestra"].map(texto_entero)
    pedido = df[df["Combinada"] == VAR_PEDIDO].groupby("InspeccionID")["VariableValor"].first()
    m1 = df[df["NroMuestra"] == "1"]
    sem = m1[m1["Combinada"] == VAR_SEMILLADO].groupby("InspeccionID")["VariableValor"].first()
    tam = m1[m1["Combinada"] == VAR_TAMANO].groupby("InspeccionID")["VariableValor"].first()
    res = pd.DataFrame({"PedidoVenta": pedido, "Semillado": sem, "TamanoMuestra": tam})
    res.index.name = "InspeccionID"
    res = res.reset_index()
    res["PV"] = res["PedidoVenta"].map(normalizar_pv)
    res["Semillado"] = a_numero(res["Semillado"])
    res["TamanoMuestra"] = a_numero(res["TamanoMuestra"])
    res["C-%"] = (res["Semillado"] / res["TamanoMuestra"].where(res["TamanoMuestra"] != 0)).round(decimales)
    return res


def clasificar_procedencia(razon):
    s = " ".join(str(razon).upper().split()) if razon is not None and not pd.isna(razon) else ""
    for grupo, nombres in PROCEDENCIA.items():
        if s in {" ".join(n.upper().split()) for n in nombres}:
            return grupo
    return "TERCEROS"


# =============================================================================
# CONSOLIDACIÓN
# =============================================================================
GRUPOS = [  # (Grupo, Medida, OrdenGrupo, OrdenSubgrupo, columna en PEDIDOS)
    ("Procedencia de MMPP", "% Compositan", 1, 1, "% Compositan"),
    ("Procedencia de MMPP", "% Olmos", 1, 2, "% Olmos"),
    ("Procedencia de MMPP", "% Terceros", 1, 3, "% Terceros"),
    ("Procedencia de MMPP", "% re-proceso", 1, 4, "% Re-proceso"),
    *[("Dias-Cos-Desp", f"% {d}D", 2, d, f"% {d}D") for d in range(1, 9)],
    ("Inspección PT", "Prom C-% PT", 3, 1, "Prom C-% PT"),
    ("Inspección PT", "Máx C-% PT", 3, 2, "Máx C-% PT"),
    ("Inspección PT", "Mín C-% PT", 3, 3, "Mín C-% PT"),
    ("Inspección Empaque", "Prom C-% EM", 4, 1, "Prom C-% EM"),
    ("Inspección Empaque", "Máx C-% EM", 4, 2, "Máx C-% EM"),
    ("Inspección Empaque", "Mín C-% EM", 4, 3, "Mín C-% EM"),
    ("Indicadores", "Resultados Inspección", 5, 1, "Resultados Inspección"),
    ("Indicadores", "Vida Útil", 5, 2, "Vida Útil"),
    ("Indicadores", "Con Tolerancia", 5, 3, "Con Tolerancia"),
    ("Indicadores", "status", 5, 4, "Status Final"),
]
COLS_TEXTO = {"Resultados Inspección", "Vida Útil", "Con Tolerancia", "Status Final"}


def consolidar(sap, trk, mat, prod, insp_pt, insp_em):
    log("Consolidando por pedido...")
    # --- Cabecera desde SAP (equivale a BASE_SAP_UNICO) ---
    ped = sap.groupby("PV").agg(
        **{"FechaDespacho": ("Fecha Despacho PL", "max"),
           "FechaETD": ("Fecha ETD", "max"),
           "ZFER": ("Material", "max"),
           "Cliente Danper": ("Cliente Danper", "max"),
           "TipoEnvioSAP": ("Tipo Envio", "max"),
           "Conteo": ("Conteo", lambda s: " / ".join(sorted({str(x).strip() for x in s.dropna() if str(x).strip()})))}
    ).reset_index()
    ped["Año"] = ped["FechaDespacho"].dt.year
    ped["Orden Mes"] = ped["FechaDespacho"].dt.month
    ped["Mes"] = ped["Orden Mes"].map(lambda m: MESES[m - 1])
    ped["Semana"] = ped["FechaDespacho"].map(semana)

    # --- Tracking: el registro más reciente por PV ---
    lk = (trk[trk["PV"].notna()]
          .drop_duplicates("PV")[["PV", "N° FCL/AWB", "FECHA INICIAL DE LLEGADA A DESTINO", "Tipo de Envío"]]
          .rename(columns={"FECHA INICIAL DE LLEGADA A DESTINO": "ETA", "Tipo de Envío": "TipoEnvioTRK"}))
    ped = ped.merge(lk, on="PV", how="left")

    def tipo_envio(r):
        if isinstance(r["TipoEnvioTRK"], str):
            return r["TipoEnvioTRK"]
        s = str(r["TipoEnvioSAP"] or "").strip().upper()
        if s in TIPO_ENVIO_SAP:
            return TIPO_ENVIO_SAP[s]
        if "REO" in s:
            return "Aéreo"
        if "MAR" in s:
            return "Marítimo"
        return r["TipoEnvioSAP"]
    ped["Tipo Envío"] = ped.apply(tipo_envio, axis=1)

    # --- Matrices ---
    m = mat.groupby("PV_N").agg(**{
        "Resultados Inspección": ("RESULTADO DE INSPECCIÓN", "first"),
        "Status Final": ("STATUS FINAL", "first"),
        "Reclamo": ("Reclamo", "first")}).reset_index().rename(columns={"PV_N": "PV"})
    ped = ped.merge(m, on="PV", how="left")

    # --- Producción: procedencia y días ---
    prod = prod.merge(ped[["PV", "FechaDespacho", "ETA"]], on="PV", how="left")
    d_desp = (prod["FechaDespacho"] - prod["FechaCosecha"]).dt.days + 1
    prod["Dias Cos-Desp"] = d_desp.where((d_desp >= 0) & (d_desp <= DIAS_MAX))
    d_lleg = (prod["ETA"] - prod["FechaCosecha"]).dt.days
    prod["Dias Cos-Llegada"] = d_lleg.where((d_lleg >= 0) & (d_lleg <= DIAS_MAX))

    total = prod.groupby("PV")["Cantidad"].sum()
    proc = (prod.pivot_table(index="PV", columns="Procedencia", values="Cantidad", aggfunc="sum")
            .reindex(columns=["COMPOSITAN", "OLMOS", "TERCEROS", "RE-PROCESO"]).fillna(0))
    proc = proc.div(total.where(total != 0), axis=0)
    proc.columns = ["% Compositan", "% Olmos", "% Terceros", "% Re-proceso"]

    con_dias = prod[prod["Dias Cos-Desp"].notna()]
    den = con_dias.groupby("PV")["Cantidad"].sum()
    dias = (con_dias.pivot_table(index="PV", columns="Dias Cos-Desp", values="Cantidad", aggfunc="sum")
            .reindex(columns=range(1, 9)).fillna(0))
    dias = dias.div(den.where(den != 0), axis=0)
    dias.columns = [f"% {d}D" for d in range(1, 9)]

    tol = prod.groupby("PV")["ZFER"].apply(
        lambda s: "Si" if s.map(lambda z: (texto_entero(z) or "").startswith("70017")).any() else "--"
    ).rename("Con Tolerancia")

    def vida(s):
        s = s.dropna()
        return f"{int(s.min())} - {int(s.max())}" if len(s) else None
    vida_util = prod.groupby("PV")["Dias Cos-Llegada"].apply(vida).rename("Vida Útil")

    # --- Inspecciones ---
    pt = insp_pt.groupby("PV")["C-%"].agg(["mean", "max", "min"])
    pt.columns = ["Prom C-% PT", "Máx C-% PT", "Mín C-% PT"]
    em = insp_em.groupby("PV")["C-%"].agg(["mean", "max", "min"])
    em.columns = ["Prom C-% EM", "Máx C-% EM", "Mín C-% EM"]

    for extra in [proc, dias, tol, vida_util, pt, em]:
        ped = ped.merge(extra, left_on="PV", right_index=True, how="left")

    # Igual que COALESCE(...,0) del Valor_Switcher para los numéricos
    numericas = [g[4] for g in GRUPOS if g[4] not in COLS_TEXTO]
    ped[numericas] = ped[numericas].fillna(0)
    ped["Con Tolerancia"] = ped["Con Tolerancia"].fillna("--")

    orden = ["PV", "FechaDespacho", "FechaETD", "Año", "Mes", "Orden Mes", "Semana",
             "Cliente Danper", "ZFER", "Conteo", "Tipo Envío", "TipoEnvioSAP", "N° FCL/AWB", "ETA",
             *[g[4] for g in GRUPOS], "Reclamo"]
    ped = ped[orden].sort_values("FechaDespacho", ascending=False)

    # --- Formato largo para la matriz (reemplaza a DIM_GRUPOS + Valor_Switcher) ---
    filas = []
    for _, r in ped.iterrows():
        for grupo, medida, og, osg, col in GRUPOS:
            v = r[col]
            es_txt = col in COLS_TEXTO
            filas.append({"PV": r["PV"], "Grupo": grupo, "Medida_Nombre": medida,
                          "OrdenGrupo": og, "OrdenSubgrupo": osg,
                          "Valor": None if es_txt else v,
                          "ValorTexto": (v if isinstance(v, str) else None) if es_txt
                          else (f"{v:.0%}" if pd.notna(v) else None)})
    largo = pd.DataFrame(filas)

    # --- Revisión ---
    pv_sap, pv_trk = set(ped["PV"]), set(trk["PV"].dropna())
    pv_prod = set(prod["PV"].dropna())
    rev = ([{"PV": p, "Motivo": "En SAP sin tracking"} for p in sorted(pv_sap - pv_trk)]
           + [{"PV": p, "Motivo": "En tracking sin SAP"} for p in sorted(pv_trk - pv_sap)]
           + [{"PV": p, "Motivo": "En SAP sin producción"} for p in sorted(pv_sap - pv_prod)])
    revisar_pv = pd.DataFrame(rev, columns=["PV", "Motivo"])
    revisar_proc = (prod.groupby(["RazonSocial", "Procedencia"], dropna=False)["Cantidad"].sum()
                    .reset_index().sort_values(["Procedencia", "Cantidad"], ascending=[True, False]))

    return ped, largo, prod, revisar_pv, revisar_proc


def guardar(hojas):
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    tmp = SALIDA.with_name(SALIDA.stem + "_tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl", datetime_format="DD/MM/YYYY",
                        date_format="DD/MM/YYYY") as xw:
        for nombre, df in hojas.items():
            df = df.copy()
            fechas = [i for i, c in enumerate(df.columns, start=1)
                      if pd.api.types.is_datetime64_any_dtype(df[c])]
            for i in fechas:                          # solo fecha, sin 00:00:00
                df.iloc[:, i - 1] = df.iloc[:, i - 1].dt.date
            df.to_excel(xw, sheet_name=nombre, index=False)
            ws = xw.sheets[nombre]
            for i in fechas:
                for (celda,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    celda.number_format = "DD/MM/YYYY"
    try:
        os.replace(tmp, SALIDA)
    except PermissionError:
        log(f"! {SALIDA.name} está abierto. El resultado quedó en {tmp.name}")
        return tmp
    return SALIDA


# =============================================================================
# PUBLICACIÓN EN LA PÁGINA (comparativo.json cifrado)
# =============================================================================
COLS_WEB = [  # (clave en la página, columna de PEDIDOS)
    ("pv", "PV"), ("fd", "FechaDespacho"), ("fetd", "FechaETD"), ("eta", "ETA"),
    ("anio", "Año"), ("om", "Orden Mes"), ("sem", "Semana"), ("cli", "Cliente Danper"),
    ("conteo", "Conteo"), ("tipo", "Tipo Envío"), ("fcl", "N° FCL/AWB"),
    *[(g[4], g[4]) for g in GRUPOS], ("reclamo", "Reclamo"),
]


def valor_web(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    if isinstance(v, (bool, int)):
        return v
    if isinstance(v, float):
        if pd.isna(v):
            return None
        return int(v) if v.is_integer() and abs(v) > 1 else round(v, 4)
    if hasattr(v, "item"):                      # números de numpy
        return valor_web(v.item())
    t = str(v).strip()
    return t or None


def paquete_web(ped, clave_subida):
    filas = [[valor_web(r[c]) for _, c in COLS_WEB] for _, r in ped.iterrows()]
    return {
        "gen": datetime.now(LIMA).isoformat(timespec="seconds"),
        "cols": [k for k, _ in COLS_WEB],
        "grupos": [[g[0], g[1], g[4]] for g in GRUPOS],
        "rows": filas,
        "fuentes": FUENTES,
        "subida": {"url": SUBIDA_URL, "clave": clave_subida} if clave_subida else None,
    }


def clave_subida():
    c = os.environ.get("SUBIDA_CLAVE", "").strip()
    if not c and ARCHIVO_SUBIDA.exists():
        c = ARCHIVO_SUBIDA.read_text(encoding="utf-8").strip()
    if not c and sys.stdin.isatty():
        c = getpass.getpass("Clave del servicio de subida (la CLAVE del Apps Script; Enter para omitir): ").strip()
        if c and input(f"¿Guardarla en {ARCHIVO_SUBIDA}? (s/n): ").strip().lower().startswith("s"):
            ARCHIVO_SUBIDA.write_text(c, encoding="utf-8")
            log(f"Guardada en {ARCHIVO_SUBIDA} (fuera del repositorio; no se sube a GitHub).")
    if not c:
        log("Aviso: sin clave de subida; en la página no aparecerá «Cargar archivos».")
    return c


def cifrar(contenido, clave_admin):
    """Mismo cifrado que fito.json (una cerradura por DNI + la del administrador, misma sal que datos.json)."""
    dnis, it, sal = af.accesos(clave_admin)
    anterior = af.RUTA_FITO
    af.RUTA_FITO = RUTA_JSON
    try:
        af.guardar_fito(contenido, dnis, clave_admin, it, sal)
    finally:
        af.RUTA_FITO = anterior
    return len(dnis)


def publicar(contenido, clave_admin):
    for intento in range(3):
        rs.git("pull", "--ff-only", "-q")          # trae lo último (la actualización automática también sube cambios)
        n = cifrar(contenido, clave_admin)
        rs.git("add", RUTA_JSON.name)
        if not rs.git("status", "--porcelain", "--", RUTA_JSON.name):
            log("comparativo.json no cambió: no hay nada que subir.")
            return
        rs.git("commit", "-q", "-m", f"Comparativo ({datetime.now(LIMA).strftime('%d/%m/%Y %H:%M')})")
        try:
            rs.git("push", "-q")
            log(f"Publicado en la página para {n} DNIs. En 1–2 minutos se ve (botón Actualizar de la página).")
            return
        except RuntimeError:
            log("GitHub tenía un cambio más nuevo (la actualización automática); vuelvo a intentarlo…")
            rs.git("reset", "-q", "--keep", "HEAD~1")
    sys.exit("No se pudo subir después de 3 intentos. Prueba otra vez en unos minutos.")


def main():
    ap = argparse.ArgumentParser(description="Arma el comparativo por pedido de venta y lo publica en la página")
    ap.add_argument("--sin-subir", action="store_true", help="solo deja el Excel en el Escritorio; no publica")
    a = ap.parse_args()

    # las claves se piden al inicio, para no tener que esperar la corrida completa
    clave_admin = csub = None
    if not a.sin_subir:
        clave_admin = rs.clave_admin()
        csub = clave_subida()

    sap = cargar_sap()
    trk = cargar_tracking()
    mat = cargar_matrices()

    log("Consultando SQL...")
    conn = conectar_sql()
    try:
        conn.timeout = SQL_TIEMPO_MAX
    except Exception:
        pass
    try:
        prod = consulta("Producción (HECHO_RESUMEN_PRODUCCION)", leer_produccion, conn)
        qm_pt = consulta("Inspección PT (QM Cloud)", leer_qm, conn, CARTILLA_PT, ANIO_DESDE_PT)
        qm_em = consulta("Inspección de empaque (QM Cloud)", leer_qm, conn, CARTILLA_EM, ANIO_DESDE_EM)
    finally:
        conn.close()

    pvs = set(sap["PV"].dropna())
    prod["PV"] = prod["PedidoVenta"].map(normalizar_pv)
    prod = prod[prod["PV"].isin(pvs)].copy()
    prod["Cantidad"] = a_numero(prod["Cantidad"]).fillna(0)
    prod["FechaCosecha"] = a_fecha(prod["FechaCosecha"])
    prod["Procedencia"] = prod["RazonSocial"].map(clasificar_procedencia)
    insp_pt = armar_inspeccion(qm_pt, 3)
    insp_em = armar_inspeccion(qm_em, 2)
    log(f"  Producción: {len(prod)} filas | Insp. PT: {len(insp_pt)} | Insp. EM: {len(insp_em)}")

    ped, largo, prod, rev_pv, rev_proc = consolidar(sap, trk, mat, prod, insp_pt, insp_em)

    ruta = guardar({
        "PEDIDOS": ped,
        "INDICADORES_LARGO": largo,
        "PRODUCCION": prod,
        "INSP_PT": insp_pt,
        "INSP_EM": insp_em,
        "TRACKING": trk,
        "REVISAR_PV": rev_pv,
        "REVISAR_PROCEDENCIA": rev_proc,
    })
    sin_trk = (rev_pv["Motivo"] == "En SAP sin tracking").sum()
    log(f"Excel: {len(ped)} pedidos ({sin_trk} aún sin tracking) -> {ruta}")
    if not a.sin_subir:
        publicar(paquete_web(ped, csub), clave_admin)


if __name__ == "__main__":
    try:
        main()
    except af.Detener as e:
        log(f"ERROR: {e}")
        sys.exit(1)
    except Exception as e:
        log(f"ERROR: {type(e).__name__}: {e}")
        sys.exit(1)

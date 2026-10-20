"""
Actualiza TODO lo que se sube desde tu laptop al Reporte diario, uno detrás de otro:

  1. recepcion_sql.py   Kilos recepcionados (SQL)
  2. drenados_pbi.py    Kilos drenados y Producto terminado observado (Power BI)
  3. mano_obra_pbi.py   Mano de obra – CMO (Power BI)
  4. reclamos_pbi.py    Ratio de reclamos penalizados (Power BI)
  5. ventas_excel.py    Despachos · Mercado (Excel de ventas más reciente de Descargas/Escritorio/Documentos)

Si uno falla, sigue con los demás y al final te dice cuáles salieron bien y cuáles no.

Uso (en Git Bash, dentro de la carpeta flujo-cosecha):
    python actualizar_todo.py
    python actualizar_todo.py --sin-ventas      (sin el Excel de ventas)
    python actualizar_todo.py --solo drenados mano_obra   (solo esos)
"""

import argparse
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
PASOS = [
    ("recepcion", "recepcion_sql.py", "Kilos recepcionados (SQL)"),
    ("drenados", "drenados_pbi.py", "Kilos drenados y producto observado"),
    ("mano_obra", "mano_obra_pbi.py", "Mano de obra – CMO"),
    ("reclamos", "reclamos_pbi.py", "Ratio de reclamos"),
    ("ventas", "ventas_excel.py", "Despachos · Mercado (Excel de ventas)"),
]


def main():
    ap = argparse.ArgumentParser(description="Corre todos los scripts del Reporte diario, uno detrás de otro")
    ap.add_argument("--sin-ventas", action="store_true", help="no corre el del Excel de ventas")
    ap.add_argument("--solo", nargs="+", choices=[p[0] for p in PASOS], help="solo estos pasos")
    a = ap.parse_args()

    pasos = [p for p in PASOS if (not a.solo or p[0] in a.solo) and not (a.sin_ventas and p[0] == "ventas")]
    resultado = []
    inicio = time.time()
    for i, (clave, archivo, nombre) in enumerate(pasos, 1):
        print(f"\n{'=' * 70}\n[{datetime.now():%H:%M:%S}] {i}/{len(pasos)} · {nombre}  ({archivo})\n{'=' * 70}", flush=True)
        t0 = time.time()
        r = subprocess.run([sys.executable, str(AQUI / archivo)], cwd=str(AQUI))
        resultado.append((nombre, r.returncode == 0, time.time() - t0))

    print(f"\n{'=' * 70}\nRESUMEN ({(time.time() - inicio) / 60:.0f} min)")
    for nombre, ok, seg in resultado:
        print(f"  {'OK   ' if ok else 'FALLÓ'}  {nombre}  ({seg / 60:.1f} min)")
    if any(not ok for _, ok, _ in resultado):
        print("\nLo que falló puedes volver a correrlo solo; arriba está el mensaje de cada uno.")
        sys.exit(1)
    print("\nTodo subido. En 1–2 minutos se ve en la página (Ctrl + F5).")


if __name__ == "__main__":
    main()

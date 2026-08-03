"""
Orquesta la actualizacion mensual del dashboard:

  1. Pregunta a la API cual es el corte mas reciente publicado.
  2. Si ese corte ya esta en historial_portafolio.json, no hace nada.
  3. Si es nuevo: lo descarga, lo procesa y lo agrega al historial.
  4. Regenera index.html con los datos actualizados.

Pensado para correr solo (GitHub Actions), pero tambien sirve a mano:
    python3 actualizar_dashboard.py
    python3 actualizar_dashboard.py --forzar          # reprocesa aunque ya exista
    python3 actualizar_dashboard.py --fecha 2026-07-31
"""
import argparse
import json
import re
import sys
from pathlib import Path

import descargar_api
import procesar_portafolio as pp

HISTORIAL = Path("historial_portafolio.json")
PLANTILLA = Path("plantilla_dashboard.html")
SALIDA = Path("index.html")


def cortes_ya_procesados() -> set:
    if not HISTORIAL.exists():
        return set()
    datos = json.loads(HISTORIAL.read_text(encoding="utf-8"))
    return {s["fecha_corte"] for s in datos.get("snapshots", [])}


def regenerar_html():
    """Inyecta el historial en la plantilla para producir el index.html final."""
    if not PLANTILLA.exists():
        raise FileNotFoundError(
            f"Falta {PLANTILLA}. Debe ser el index.html actual pero con la linea de datos "
            f"reemplazada por el marcador __DATOS_PORTAFOLIO__."
        )

    historial = json.loads(HISTORIAL.read_text(encoding="utf-8"))
    html = PLANTILLA.read_text(encoding="utf-8")

    if "__DATOS_PORTAFOLIO__" not in html:
        raise ValueError("La plantilla no contiene el marcador __DATOS_PORTAFOLIO__")

    datos_js = json.dumps(historial, ensure_ascii=False, separators=(",", ":"))
    html = html.replace("__DATOS_PORTAFOLIO__", datos_js)

    SALIDA.write_text(html, encoding="utf-8")
    tam = SALIDA.stat().st_size / 1024 / 1024
    print(f"index.html regenerado ({tam:.2f} MB, {len(historial['snapshots'])} cortes)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fecha", help="procesar un corte especifico (YYYY-MM-DD)")
    ap.add_argument("--forzar", action="store_true", help="reprocesar aunque ya exista")
    ap.add_argument("--solo-revisar", action="store_true",
                    help="solo dice si hay corte nuevo, sin procesarlo")
    args = ap.parse_args()

    # --- 1. que corte hay disponible ---
    if args.fecha:
        fecha = args.fecha
    else:
        print("Consultando la API por el corte mas reciente...")
        disponibles = descargar_api.listar_cortes(1)
        if not disponibles:
            print("ERROR: la API no reporto cortes disponibles")
            sys.exit(1)
        fecha = disponibles[0]

    ya_estan = cortes_ya_procesados()
    print(f"Corte mas reciente en la API : {fecha}")
    print(f"Cortes ya en el historial    : {len(ya_estan)}")

    if fecha in ya_estan and not args.forzar:
        print(f"\nSin novedades: el corte {fecha} ya estaba procesado. No hay nada que hacer.")
        # codigo 0 y sin cambios: el workflow sabra que no debe publicar nada
        return

    if args.solo_revisar:
        print(f"\nHAY CORTE NUEVO: {fecha}")
        return

    # --- 2. descargar y procesar ---
    print(f"\nDescargando corte {fecha} desde la API...")
    df = descargar_api.descargar_corte(fecha)
    print(f"  {len(df):,} filas")

    print("Procesando...")
    snapshot = pp.procesar_dataframe(df)
    pp.actualizar_historial(snapshot, str(HISTORIAL))

    # --- 3. regenerar el dashboard ---
    print("Regenerando dashboard...")
    regenerar_html()

    print(f"\nListo: corte {fecha} agregado y dashboard actualizado.")


if __name__ == "__main__":
    main()

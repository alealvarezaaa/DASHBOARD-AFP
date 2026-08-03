"""
Descarga el Formato 351 (portafolio de AFPs) desde la API de datos abiertos
de la Superintendencia Financiera, en lugar de leer el Excel manualmente.

Dataset: https://www.datos.gov.co/Hacienda-y-Cr-dito-P-blico/Formato-351-Composici-n-del-portafolio-de-inversio/ur2p-h4yf

Uso:
    python3 descargar_api.py --ultimo        # trae el corte mas reciente disponible
    python3 descargar_api.py --fecha 2026-06-30
    python3 descargar_api.py --listar        # solo muestra que cortes hay
"""
import argparse
import sys
import time
from urllib.parse import urlencode
from urllib.request import urlopen, Request
import json

import pandas as pd

API_BASE = "https://www.datos.gov.co/resource/ur2p-h4yf"
PAGINA = 50000  # Socrata permite hasta 50k por pagina con $limit

# Columnas de la API que necesita el pipeline, con el nombre que espera
# procesar_portafolio.py (el mismo que traia el Excel original).
COLUMNAS = {
    "nombre_de_entidad": "Nombre de Entidad",
    "fecha_de_corte": "Fecha de Corte",
    "c_digo_tipo_patrimonio": "Código Tipo Patrimonio",
    "c_digo_patrimonio": "Código Patrimonio",
    "clase_inversion": "Clase de Inversión",
    "valor_mercado_o_pres_pesos": "Vr. mercado o Vr presente en $",
    "nro_id_emisor": "No. ID Emisor",
    "razon_social_emisor": "Razon Social Emisor",
    "no_acciones": "No. Acciones",
    "nemotecnico": "Nemotecnico",
}

# Columnas que deben quedar numericas (la API las entrega como texto)
NUMERICAS = [
    "Código Tipo Patrimonio",
    "Código Patrimonio",
    "Vr. mercado o Vr presente en $",
    "No. Acciones",
]


def _pedir(url: str, reintentos: int = 3, espera: float = 3.0):
    """GET con reintentos: la API a veces responde lento o con 503 transitorio."""
    ultimo_error = None
    for intento in range(1, reintentos + 1):
        try:
            req = Request(url, headers={"User-Agent": "dashboard-afp/1.0"})
            with urlopen(req, timeout=180) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001
            ultimo_error = e
            if intento < reintentos:
                print(f"    reintento {intento}/{reintentos - 1} tras error: {e}")
                time.sleep(espera * intento)
    raise RuntimeError(f"La API no respondio tras {reintentos} intentos: {ultimo_error}")


def listar_cortes(n: int = 12):
    """Devuelve las fechas de corte disponibles, de la mas reciente hacia atras."""
    params = {
        "$select": "fecha_de_corte",
        "$group": "fecha_de_corte",
        "$order": "fecha_de_corte DESC",
        "$limit": n,
    }
    filas = _pedir(f"{API_BASE}.json?{urlencode(params)}")
    return [f["fecha_de_corte"][:10] for f in filas if f.get("fecha_de_corte")]


def descargar_corte(fecha: str) -> pd.DataFrame:
    """
    Trae todas las filas de un corte (fecha en formato YYYY-MM-DD).
    Pagina automaticamente: un corte completo son ~20-25 mil filas.
    """
    columnas_api = ",".join(COLUMNAS.keys())
    todas = []
    offset = 0

    while True:
        params = {
            "$select": columnas_api,
            "$where": f"fecha_de_corte = '{fecha}T00:00:00.000'",
            "$limit": PAGINA,
            "$offset": offset,
            "$order": ":id",  # orden estable, necesario para paginar sin repetir/saltar
        }
        lote = _pedir(f"{API_BASE}.json?{urlencode(params)}")
        if not lote:
            break
        todas.extend(lote)
        print(f"    {len(todas):,} filas descargadas...")
        if len(lote) < PAGINA:
            break
        offset += PAGINA

    if not todas:
        raise RuntimeError(f"La API no devolvio filas para el corte {fecha}")

    df = pd.DataFrame(todas)
    faltantes = [c for c in COLUMNAS if c not in df.columns]
    if faltantes:
        raise RuntimeError(f"La API no trajo estas columnas esperadas: {faltantes}")

    df = df.rename(columns=COLUMNAS)

    for col in NUMERICAS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # procesar_portafolio.py espera texto en estas
    for col in ["Nemotecnico", "Razon Social Emisor", "Nombre de Entidad", "Clase de Inversión"]:
        df[col] = df[col].astype("string")

    df["Fecha de Corte"] = pd.to_datetime(df["Fecha de Corte"])

    return df


def main():
    ap = argparse.ArgumentParser(description="Descarga el Formato 351 desde datos.gov.co")
    ap.add_argument("--ultimo", action="store_true", help="descarga el corte mas reciente")
    ap.add_argument("--fecha", help="descarga un corte especifico (YYYY-MM-DD)")
    ap.add_argument("--listar", action="store_true", help="solo lista los cortes disponibles")
    ap.add_argument("--salida", default="corte_api.parquet", help="archivo donde guardar")
    args = ap.parse_args()

    if args.listar:
        print("Cortes disponibles (mas recientes primero):")
        for f in listar_cortes(24):
            print("  ", f)
        return

    if args.ultimo:
        cortes = listar_cortes(1)
        if not cortes:
            print("ERROR: la API no reporto ningun corte disponible")
            sys.exit(1)
        fecha = cortes[0]
        print(f"Corte mas reciente en la API: {fecha}")
    elif args.fecha:
        fecha = args.fecha
    else:
        ap.error("indica --ultimo, --fecha YYYY-MM-DD o --listar")

    print(f"Descargando corte {fecha} ...")
    df = descargar_corte(fecha)
    print(f"OK: {len(df):,} filas, {df['Nombre de Entidad'].nunique()} entidades")

    df.to_parquet(args.salida, index=False)
    print(f"Guardado en {args.salida}")


if __name__ == "__main__":
    main()

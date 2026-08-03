"""
Procesa el archivo mensual "Formato 351" (detalle de portafolio de AFPs, SFC Colombia)
y genera un snapshot agregado en JSON con foco en acciones locales vs internacionales.

Uso:
    python3 procesar_portafolio.py /ruta/al/archivo.xls
"""
import sys
import json
import pandas as pd
from pathlib import Path

# --- Clasificación de "Clase de Inversión" ---
ACCIONES_LOCALES = {
    "AAEVS", "AAENVS",   # alta liquidez, entidades vig./no vig. SFC
    "ABEVS", "ABENVS",   # baja liquidez
    "AMEVS", "AMENVS",   # media liquidez
}
ACCIONES_INTERNACIONALES = {
    "AEE",               # acciones directas emitidas por entidades del exterior
    "CRAAAVS", "CRAAANVS",  # ADRs (representan acciones colombianas negociadas afuera)
    "PFMUIA",            # fondos internacionales enfocados en acciones
    "FINDI",             # fondos/ETFs que replican índices accionarios (S&P500, MSCI, etc.)
}

# Clases de inversión de renta fija (deuda pública y privada, titularizaciones, CDTs, etc.)
# Todo lo que no caiga en acciones ni en esta lista se agrupa como "Otros Activos"
# (fondos de capital privado, inmobiliarios, depósitos, productos estructurados, etc.)
RENTA_FIJA = {
    "TSTF", "TSUV", "TDPE", "TDPIT", "TSEGE", "TDEPT",           # deuda pública (TES, externa)
    "BOENVS", "BOEVS", "BOEBE", "BOEGE", "BOEEDB", "BGEVS",       # bonos corporativos/bancarios
    "CDT", "CDEBE",                                                # CDTs
    "BHIP", "BPEN",                                                # bonos hipotecarios/pensionales
    "TCCENVS", "TCCEVS", "TCCH",                                   # titularizaciones de contenido crediticio
    "ORFENVS", "ORFEBE",                                           # otros títulos de renta fija
    "FINDIRF",                                                     # fondos índice de renta fija
    "PFMUITD",                                                     # fondos internacionales de deuda
    "TPMTIN",                                                      # titularización inmobiliaria (mixta/participativa)
}

# Reclasificación manual por nemotécnico: casos donde la Clase de Inversión
# reportada (AEE = "emisor del exterior") no refleja que ese ticker específico
# cotiza localmente y el usuario lo quiere tratado como renta variable local.
# --- Tickers de los ADR ---
# Los ADR llegan del Formato 351 SIN nemotécnico, así que por defecto se
# identificaban con el nombre del emisor. Aquí se les asigna su ticker real de
# NYSE. Se mapea por palabra clave y no por NIT porque (a) el NIT viene a veces
# truncado y (b) Bancolombia aparece con dos razones sociales distintas en el
# histórico (BANCOLOMBIA antes del cambio de nombre, GRUPO CIBEST después) y así
# ambas quedan unificadas bajo el mismo ticker.
ADR_TICKERS = {
    "CIBEST": "CIB",
    "BANCOLOMBIA": "CIB",
    "ECOPETROL": "EC",
    "AVAL": "AVAL",
}


def ticker_adr(nombre_emisor: str):
    n = (nombre_emisor or "").upper()
    for palabra, ticker in ADR_TICKERS.items():
        if palabra in n:
            return ticker
    return None


NEMOTECNICO_RECLASIFICACION = {
    "NUAMCO": {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},
    # NUAM (sin sufijo CO) se mantiene como internacional: es el mismo emisor
    # pero el listado que no cotiza en Colombia.

    # Estos tres venían con Clase de Inversión de "otros activos" (titularización
    # inmobiliaria TPMTIN, o participación en fondo bursátil PCCBCTP) y por eso no
    # aparecían como acción. Se reclasifican a peición del usuario:
    "ICOLCAP": {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},  # ETF que replica el COLCAP (PCCBCTP)
    "HCOLSEL": {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},  # ETF Horizons Colombia Select S&P (PCCBCTP)
    "PEI":     {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},  # Patrimonio Estrategias Inmobiliarias (TPMTIN)
    "TIN":     {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},  # Titularizadora Colombiana - Hitos (TPMTIN)

    # Davivienda Group se redomicilió a Panamá (Pais_Emisor 591) y las AFP lo reportan
    # de forma inconsistente entre sí: Porvenir y Protección con clase AMEVS (acción
    # local), Colfondos y Skandia con AEE (acción del exterior). Se fuerza a local en
    # todas para que no quede partido en dos categorías dentro del mismo nemotécnico.
    "PFDAVIGRP": {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},
    "PFDAVVNDA": {"categoria_amplia": "Renta variable local", "categoria_fina": "Acción local"},
    # NOTA: bajo esas mismas dos clases hay otros instrumentos que NO se tocaron
    # porque no son equivalentes: GXTESCOL (PCCBCTP) es un ETF de TES = renta fija,
    # y ESTRATEINTP (TPMTIN) es otra tranche de la fiduciaria de PEI sin confirmar.
}

# --- Acciones en circulación por compañía (para calcular % de participación de cada AFP) ---
# Fuente: Bloomberg (archivo XML5ZW0K.xlsx, columna "Curr Shares Out"), descargado jul-2026.
# IMPORTANTE:
#  1. Esto SOLO se aplica al corte más reciente del historial, nunca a meses pasados,
#     porque el número de acciones en circulación cambia con el tiempo (escisiones,
#     recompras, emisiones) y no tenemos ese dato mes a mes histórico.
#  2. Para emisores con acción ordinaria Y preferencial (Cibest, Grupo Aval, Grupo Sura,
#     Grupo Argos, Corficolombiana), Bloomberg solo trae en este archivo la ordinaria —
#     la preferencial (tickers PFCIBEST, PFAVAL, PFGRUPSURA, PFGRUPOARG, PFCORFICOL)
#     queda sin dato hasta conseguir esa cifra por separado.
ACCIONES_CIRCULACION = {
    "ECOPETROL":  {"acciones": 41_116_696_576, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "BBVACOL":    {"acciones": 17_308_966_912, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "GRUPOAVAL":  {"acciones": 16_178_324_480, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — solo acción ordinaria"},
    "GEB":        {"acciones":  9_181_176_832, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "POPULAR":    {"acciones":  7_886_104_064, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "EXITO":      {"acciones":  1_297_864_320, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "CEMARGOS":   {"acciones":  1_220_886_400, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "PROMIGAS":   {"acciones":  1_134_848_000, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "CONCONCRET": {"acciones":  1_134_254_976, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "ISA":        {"acciones":  1_107_677_952, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "CELSIA":     {"acciones":  1_032_915_648, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "CIBEST":     {"acciones":    509_103_136, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — solo acción ordinaria"},
    "GRUPOARGOS": {"acciones":    397_523_168, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — solo acción ordinaria"},
    "BOGOTA":     {"acciones":    355_251_072, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "CORFICOLCF": {"acciones":    346_403_776, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — solo acción ordinaria"},
    "MINEROS":    {"acciones":    295_780_512, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "VILLAS":     {"acciones":    222_974_688, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "TERPEL":     {"acciones":    181_424_512, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "PFDAVIGRP":  {"acciones":    169_443_760, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — acción preferencial"},
    "GRUPOSURA":  {"acciones":    165_834_032, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — solo acción ordinaria"},
    "OCCIDENTE":  {"acciones":    155_899_712, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    "PFDAVVNDA":  {"acciones":    116_601_008, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out) — acción preferencial"},
    "GRUBOLIVAR": {"acciones":     81_668_728, "fecha_dato": "2026-07", "fuente": "Bloomberg (Curr Shares Out)"},
    # Pendientes: BANCOLDEX, PBBANCOLDE, ETERCOL, INTERBOLSA (no vienen en el pull de
    # Bloomberg, probablemente por baja liquidez), y las preferenciales PFAVAL,
    # PFCIBEST, PFCORFICOL, PFGRUPOARG, PFGRUPSURA (necesitan su propio ticker Bloomberg).
}

# Alias de nemotécnicos: casos confirmados donde el mismo instrumento se registra
# con un ticker distinto según el custodio/AFP que reporta. Se consolidan al
# nemotécnico canónico (columna derecha) para no duplicarlo en el explorador.
# Verificado manualmente contra el No. ID Emisor antes de agregar un alias aquí.
#
# NOTA: NUAM vs NUAMCO se evaluaron y NO son el mismo instrumento — NUAMCO es
# el listado de Holding Bursátil Regional S.A. en la bolsa de Colombia, y NUAM
# corresponde a otro mercado de la operación integrada. Se mantienen separados.
NEMOTECNICO_ALIAS = {}
CATEGORIA_FINA = {
    "AAEVS": "Acción local", "AAENVS": "Acción local",
    "ABEVS": "Acción local", "ABENVS": "Acción local",
    "AMEVS": "Acción local", "AMENVS": "Acción local",
    "AEE": "Acción internacional",
    "CRAAAVS": "ADR", "CRAAANVS": "ADR",
    "PFMUIA": "Fondo internacional (acciones)",
    "FINDI": "Fondo / ETF índice",
}






# Fuente: posiciones oficiales del ETF iShares MSCI COLCAP (ICOLCAP) al 11-jun-2026,
# más clasificación razonable para instrumentos locales que no pertenecen al índice.
SECTOR_LOCAL = {
    "PFCIBEST": "Servicios Financieros", "CIBEST": "Servicios Financieros",
    "ECOPETROL": "Energía y Petróleo",
    "ISA": "Servicios Públicos", "GEB": "Servicios Públicos", "CELSIA": "Servicios Públicos",
    "PROMIGAS": "Servicios Públicos",
    "GRUPOSURA": "Servicios Financieros", "PFGRUPSURA": "Servicios Financieros",
    "CEMARGOS": "Materiales", "GRUPOARGOS": "Materiales", "PFGRUPOARG": "Materiales",
    "MINEROS": "Materiales", "CONCONCRET": "Materiales", "ETERCOL": "Materiales",
    "PFAVAL": "Servicios Financieros", "GRUPOAVAL": "Servicios Financieros",
    "PFDAVIGRP": "Servicios Financieros", "PFDAVVNDA": "Servicios Financieros",
    "CORFICOLCF": "Servicios Financieros", "PFCORFICOL": "Servicios Financieros",
    "GRUBOLIVAR": "Servicios Financieros", "BOGOTA": "Servicios Financieros",
    "OCCIDENTE": "Servicios Financieros", "POPULAR": "Servicios Financieros",
    "VILLAS": "Servicios Financieros", "BBVACOL": "Servicios Financieros",
    "BANCOLDEX": "Servicios Financieros", "PBBANCOLDE": "Servicios Financieros",
    "NUAMCO": "Servicios Financieros", "INTERBOLSA": "Servicios Financieros",
    "TERPEL": "Consumo",
    "EXITO": "Consumo",
    "PEI": "Inmobiliario",
}

# Tickers que forman parte del índice ICOLCAP/COLCAP al corte más reciente disponible.
COLCAP_TICKERS = {
    "PFCIBEST", "CIBEST", "ECOPETROL", "ISA", "GEB", "GRUPOSURA", "PFGRUPSURA",
    "CEMARGOS", "GRUPOARGOS", "PFAVAL", "PFGRUPOARG", "PFDAVIGRP", "CELSIA",
    "PEI", "CORFICOLCF", "MINEROS", "GRUBOLIVAR", "BOGOTA", "PROMIGAS", "TERPEL",
    "EXITO", "PFCORFICOL", "CNEC", "AXL",
}

# Instrumentos internacionales directos (acciones/ADR) con sector conocido.
SECTOR_INTERNACIONAL_DIRECTO = {
    "GEOPARK LIMITED": "Energía y Petróleo",
    "CNEC": "Energía y Petróleo",
    "ALVOPETRO ENERGY LTDA - ALV": "Energía y Petróleo",
    "ECOPETROL S.A.": "Energía y Petróleo",  # ADR
    "NUAM": "Servicios Financieros",
    "GRUPO CIBEST SA": "Servicios Financieros",  # ADR
    "AXL": "Energía y Petróleo",
}

# Palabras clave para fondos/ETFs internacionales cuyo nombre ya indica el sector
# (ej. "Financial Select Sector SPDR", "Energy Select Sector SPDR").
SECTOR_KEYWORDS_FONDOS = [
    ("FINANCIAL", "Servicios Financieros"), ("BANK", "Servicios Financieros"),
    ("ENERGY", "Energía y Petróleo"), ("OIL", "Energía y Petróleo"),
    ("CONSUMER DISCRETIONARY", "Consumo"), ("CONSUMER STAPLES", "Consumo"),
    ("COMMUNICATION", "Comunicaciones"),
    ("TECHNOLOGY", "Tecnología"), ("TECH ", "Tecnología"),
    ("HEALTH", "Salud"),
    ("INDUSTRIAL", "Industrial"),
    ("MATERIALS", "Materiales"),
    ("UTILITIES", "Servicios Públicos"),
    ("REAL ESTATE", "Inmobiliario"), ("REIT", "Inmobiliario"),
]


def sector_de(nemotecnico: str, nombre: str, categoria_amplia: str) -> str:
    if categoria_amplia == "Renta variable local":
        return SECTOR_LOCAL.get(nemotecnico, "Otro / sin clasificar")
    # internacional
    if nemotecnico in SECTOR_INTERNACIONAL_DIRECTO:
        return SECTOR_INTERNACIONAL_DIRECTO[nemotecnico]
    nombre_u = (nombre or "").upper()
    for kw, sector in SECTOR_KEYWORDS_FONDOS:
        if kw in nombre_u:
            return sector
    return "Fondos diversificados / geográficos"


AFP_NOMBRES = {
    "SKANDIA AFP - ACCAI S.A.": "Skandia",
    '"PORVENIR"': "Porvenir",
    '"PROTECCION"': "Protección",
    '"COLFONDOS S.A." Y "COLFONDOS"': "Colfondos",
}


def normalizar_afp(nombre_crudo: str) -> str:
    """
    Normaliza el nombre de la entidad a partir de palabras clave, en vez de una
    coincidencia exacta: el nombre legal reportado varía levemente entre meses
    (comillas, razón social completa, etc.) y una coincidencia exacta deja
    duplicados como AFPs distintas.
    """
    if nombre_crudo in AFP_NOMBRES:
        return AFP_NOMBRES[nombre_crudo]
    n = nombre_crudo.upper()
    if "COLFONDOS" in n:
        return "Colfondos"
    if "PORVENIR" in n:
        return "Porvenir"
    if "PROTECCION" in n or "PROTECCIÓN" in n:
        return "Protección"
    if "SKANDIA" in n:
        return "Skandia"
    return nombre_crudo  # sin coincidencia: se deja tal cual y quedará visible para revisar

# --- Mapeo de perfil de fondo por código de patrimonio ---
PERFIL_POR_CODIGO = {
    1000: "Moderado",
    5000: "Conservador",
    6000: "Mayor Riesgo",
    7000: "Retiro Programado",
    8000: "Alternativo",
}


def clasificar(clase_inversion: str, nemotecnico: str = None) -> str:
    if nemotecnico is not None and nemotecnico in NEMOTECNICO_RECLASIFICACION:
        cat_amplia = NEMOTECNICO_RECLASIFICACION[nemotecnico]["categoria_amplia"]
        if cat_amplia == "Renta variable local":
            return "Acciones Locales"
        if cat_amplia == "Renta variable internacional":
            return "Acciones Internacionales"
    if clase_inversion in ACCIONES_LOCALES:
        return "Acciones Locales"
    if clase_inversion in ACCIONES_INTERNACIONALES:
        return "Acciones Internacionales"
    return "Otro"


def clasificar_general(clase_inversion: str, nemotecnico: str = None) -> str:
    if nemotecnico is not None and nemotecnico in NEMOTECNICO_RECLASIFICACION:
        cat_amplia = NEMOTECNICO_RECLASIFICACION[nemotecnico]["categoria_amplia"]
        if cat_amplia in ("Renta variable local", "Renta variable internacional"):
            return "Renta Variable"
    if clase_inversion in ACCIONES_LOCALES or clase_inversion in ACCIONES_INTERNACIONALES:
        return "Renta Variable"
    if clase_inversion in RENTA_FIJA:
        return "Renta Fija"
    return "Otros Activos"


def perfil_de(row) -> str:
    if row["Código Tipo Patrimonio"] == 5:
        return "Cesantías"
    return PERFIL_POR_CODIGO.get(row["Código Patrimonio"], f"Otro ({row['Código Patrimonio']})")


def construir_detalle_acciones(df_raw: pd.DataFrame) -> list:
    """
    Detalle por instrumento (nemotécnico) y AFP, sumando todos los perfiles/fondos
    de esa AFP (Moderado, Conservador, etc. quedan consolidados en un solo total).

    Importante: se agrupa por Nemotécnico real cuando existe, NO por emisor,
    porque un mismo emisor puede tener varios instrumentos distintos
    (ej. CIBEST = acción ordinaria, PFCIBEST = acción preferencial).
    Solo cuando el instrumento no trae nemotécnico (típico en fondos/ETFs
    internacionales) se usa el nombre del emisor como identificador de respaldo.
    """
    cols_extra = ["No. ID Emisor", "Razon Social Emisor", "No. Acciones", "Nemotecnico"]
    faltantes = [c for c in cols_extra if c not in df_raw.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas para el detalle de acciones: {faltantes}")

    d = df_raw.copy()
    d["categoria_fina"] = d["Clase de Inversión"].map(CATEGORIA_FINA)
    # Incluir también los nemotécnicos con reclasificación manual (ej. ICOLCAP, PEI, TIN)
    # aunque su Clase de Inversión original no sea una de las clases de acciones
    # reconocidas — si no, se descartan aquí antes de llegar a aplicarse el override.
    d["_nemo_tmp"] = d["Nemotecnico"].astype("string").str.strip()
    mask_reclasificado = d["_nemo_tmp"].isin(NEMOTECNICO_RECLASIFICACION.keys())
    d.loc[mask_reclasificado & d["categoria_fina"].isna(), "categoria_fina"] = "Acción local"
    d = d.drop(columns=["_nemo_tmp"])
    d = d[d["categoria_fina"].notna()].copy()
    d["No. Acciones"] = d["No. Acciones"].fillna(0)
    d["categoria_amplia"] = d["categoria_fina"].map(
        lambda c: "Renta variable local" if c == "Acción local" else "Renta variable internacional"
    )
    d["Nemotecnico"] = d["Nemotecnico"].astype("string").str.strip()
    d.loc[d["Nemotecnico"] == "", "Nemotecnico"] = None
    d["Nemotecnico"] = d["Nemotecnico"].replace(NEMOTECNICO_ALIAS)

    con_nemo = d[d["Nemotecnico"].notna()].copy()
    sin_nemo = d[d["Nemotecnico"].isna()].copy()

    # --- Instrumentos CON nemotécnico: se agrupan por el nemotécnico mismo ---
    nombre_por_nemo = (
        con_nemo.groupby("Nemotecnico")["Razon Social Emisor"]
        .agg(lambda s: s.str.strip().mode().iloc[0])
    )
    con_nemo["nombre"] = con_nemo["Nemotecnico"].map(nombre_por_nemo)
    con_nemo["clave"] = con_nemo["Nemotecnico"]

    # --- Instrumentos SIN nemotécnico (fondos/ETFs del exterior): se agrupan
    #     por emisor, usando su nombre canónico como identificador ---
    if len(sin_nemo):
        nombre_por_emisor = (
            sin_nemo.groupby("No. ID Emisor")["Razon Social Emisor"]
            .agg(lambda s: s.str.strip().mode().iloc[0])
        )
        sin_nemo["nombre"] = sin_nemo["No. ID Emisor"].map(nombre_por_emisor)
        sin_nemo["clave"] = sin_nemo["nombre"]
        # Los ADR se identifican con su ticker real de NYSE (CIB, EC, AVAL) en vez
        # del nombre del emisor, que es lo único que trae el archivo.
        es_adr = sin_nemo["categoria_fina"] == "ADR"
        tickers = sin_nemo.loc[es_adr, "nombre"].map(ticker_adr)
        sin_nemo.loc[es_adr, "clave"] = tickers.fillna(sin_nemo.loc[es_adr, "nombre"])

    d = pd.concat([con_nemo, sin_nemo], ignore_index=True)

    # Reclasificación manual (ej. NUAMCO -> renta variable local)
    for nemo, override in NEMOTECNICO_RECLASIFICACION.items():
        mask = d["clave"] == nemo
        for campo, valor in override.items():
            d.loc[mask, campo] = valor

    agg = (
        d.groupby(["afp", "clave", "nombre", "categoria_amplia", "categoria_fina"])
        .agg(no_acciones=("No. Acciones", "sum"), valor_mercado=("Vr. mercado o Vr presente en $", "sum"))
        .reset_index()
        .rename(columns={"clave": "nemotecnico"})
    )
    agg = agg[agg["valor_mercado"] > 0]
    agg["no_acciones"] = agg["no_acciones"].round(2)
    agg["valor_mercado"] = agg["valor_mercado"].round(2)
    agg["sector"] = agg.apply(lambda r: sector_de(r["nemotecnico"], r["nombre"], r["categoria_amplia"]), axis=1)
    agg["colcap"] = agg["nemotecnico"].isin(COLCAP_TICKERS)
    return agg.sort_values("valor_mercado", ascending=False).to_dict(orient="records")


COLUMNAS_REQUERIDAS = [
    "Nombre de Entidad", "Fecha de Corte", "Código Tipo Patrimonio",
    "Código Patrimonio", "Clase de Inversión",
    "Vr. mercado o Vr presente en $",
    "No. ID Emisor", "Razon Social Emisor", "No. Acciones", "Nemotecnico",
]


def leer_excel(path_excel: str) -> pd.DataFrame:
    """Lee el Formato 351 desde el Excel mensual de la Superintendencia."""
    print(f"Leyendo {path_excel} ...")
    engine = "xlrd" if str(path_excel).lower().endswith(".xls") else "openpyxl"
    hojas_candidatas = ["Formato_351", "Fmto-351", "Fmto_351", "FORMATO_351"]
    hojas_disponibles = pd.ExcelFile(path_excel, engine=engine).sheet_names
    hoja = next((h for h in hojas_candidatas if h in hojas_disponibles), None)
    if hoja is None:
        raise ValueError(f"No se encontró la hoja del Formato 351 en {path_excel}. Hojas disponibles: {hojas_disponibles}")
    return pd.read_excel(path_excel, engine=engine, sheet_name=hoja, usecols=COLUMNAS_REQUERIDAS)


def procesar(path_excel: str) -> dict:
    """Procesa un corte a partir del archivo Excel (uso manual)."""
    return procesar_dataframe(leer_excel(path_excel))


def procesar_dataframe(df: pd.DataFrame) -> dict:
    """
    Procesa un corte ya cargado en memoria. Sirve tanto para el Excel como para
    los datos que llegan de la API de datos.gov.co: lo importante es que traiga
    las COLUMNAS_REQUERIDAS con esos nombres.
    """
    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas requeridas: {faltantes}")

    df = df.copy()
    df["Vr. mercado o Vr presente en $"] = df["Vr. mercado o Vr presente en $"].fillna(0)

    df["afp"] = df["Nombre de Entidad"].apply(normalizar_afp)
    df["perfil"] = df.apply(perfil_de, axis=1)
    df["categoria"] = df.apply(
        lambda r: clasificar(
            r["Clase de Inversión"],
            r["Nemotecnico"].strip() if isinstance(r["Nemotecnico"], str) else None,
        ),
        axis=1,
    )

    fecha_corte = pd.to_datetime(df["Fecha de Corte"].iloc[0]).strftime("%Y-%m-%d")

    # Total de portafolio por AFP+perfil (para calcular %)
    total_por_fondo = (
        df.groupby(["afp", "perfil"])["Vr. mercado o Vr presente en $"].sum().rename("total").reset_index()
    )

    # Valor de acciones locales/internacionales por AFP+perfil
    acc = df[df["categoria"] != "Otro"]
    acc_por_fondo = (
        acc.groupby(["afp", "perfil", "categoria"])["Vr. mercado o Vr presente en $"]
        .sum()
        .rename("valor")
        .reset_index()
    )

    detalle = []
    for _, r in total_por_fondo.iterrows():
        fila = {"afp": r["afp"], "perfil": r["perfil"], "total_portafolio": round(r["total"], 2)}
        sub = acc_por_fondo[(acc_por_fondo["afp"] == r["afp"]) & (acc_por_fondo["perfil"] == r["perfil"])]
        fila["acciones_locales"] = round(sub[sub["categoria"] == "Acciones Locales"]["valor"].sum(), 2)
        fila["acciones_internacionales"] = round(sub[sub["categoria"] == "Acciones Internacionales"]["valor"].sum(), 2)
        detalle.append(fila)

    # Totales consolidados por AFP (todos los perfiles sumados)
    resumen_afp = (
        pd.DataFrame(detalle)
        .groupby("afp")[["total_portafolio", "acciones_locales", "acciones_internacionales"]]
        .sum()
        .reset_index()
        .to_dict(orient="records")
    )

    detalle_acciones = construir_detalle_acciones(df)

    df["categoria_general"] = df.apply(
        lambda r: clasificar_general(
            r["Clase de Inversión"],
            r["Nemotecnico"].strip() if isinstance(r["Nemotecnico"], str) else None,
        ),
        axis=1,
    )
    comp = df.groupby("categoria_general")["Vr. mercado o Vr presente en $"].sum()
    composicion_general = {
        "renta_fija": round(comp.get("Renta Fija", 0.0), 2),
        "renta_variable": round(comp.get("Renta Variable", 0.0), 2),
        "otros_activos": round(comp.get("Otros Activos", 0.0), 2),
    }

    snapshot = {
        "fecha_corte": fecha_corte,
        "detalle_por_fondo": detalle,
        "resumen_por_afp": resumen_afp,
        "total_sistema": {
            "total_portafolio": round(df["Vr. mercado o Vr presente en $"].sum(), 2),
            "acciones_locales": round(acc[acc["categoria"] == "Acciones Locales"]["Vr. mercado o Vr presente en $"].sum(), 2),
            "acciones_internacionales": round(acc[acc["categoria"] == "Acciones Internacionales"]["Vr. mercado o Vr presente en $"].sum(), 2),
        },
        "composicion_general": composicion_general,
        "detalle_acciones": detalle_acciones,
    }
    return snapshot


def actualizar_historial(snapshot: dict, historial_path: str = "historial_portafolio.json"):
    p = Path(historial_path)
    if p.exists():
        historial = json.loads(p.read_text())
    else:
        historial = {"snapshots": []}

    # Reemplaza si ya existe un snapshot para esa fecha (reprocesos)
    historial["snapshots"] = [s for s in historial["snapshots"] if s["fecha_corte"] != snapshot["fecha_corte"]]
    historial["snapshots"].append(snapshot)
    historial["snapshots"].sort(key=lambda s: s["fecha_corte"])
    historial["acciones_circulacion"] = ACCIONES_CIRCULACION

    p.write_text(json.dumps(historial, indent=2, ensure_ascii=False))
    print(f"Historial actualizado: {p} ({len(historial['snapshots'])} snapshot(s))")
    return historial


if __name__ == "__main__":
    import glob

    if len(sys.argv) < 2:
        print("Uso:")
        print("  python3 procesar_portafolio.py <archivo.xls> [archivo2.xls ...]")
        print("  python3 procesar_portafolio.py <carpeta_con_xls>/")
        sys.exit(1)

    # Si el argumento es una carpeta, procesa todos los .xls que contenga.
    if len(sys.argv) == 2 and Path(sys.argv[1]).is_dir():
        archivos = sorted(glob.glob(str(Path(sys.argv[1]) / "*.xls")) + glob.glob(str(Path(sys.argv[1]) / "*.xlsx")))
    else:
        archivos = sys.argv[1:]

    if not archivos:
        print("No se encontraron archivos .xls para procesar.")
        sys.exit(1)

    print(f"Se procesarán {len(archivos)} archivo(s).")
    historial = None
    for i, archivo in enumerate(archivos, 1):
        try:
            snap = procesar(archivo)
            historial = actualizar_historial(snap)
            print(f"[{i}/{len(archivos)}] OK — {Path(archivo).name} -> corte {snap['fecha_corte']}")
        except Exception as e:
            print(f"[{i}/{len(archivos)}] ERROR en {Path(archivo).name}: {e}")

    if historial:
        fechas = [s["fecha_corte"] for s in historial["snapshots"]]
        print(f"\nHistorial final: {len(fechas)} corte(s) -> {fechas}")

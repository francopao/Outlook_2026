"""
================================================================================
GLOBAL MARKETS & MACRO INTELLIGENCE TERMINAL
Python + Streamlit | Institutional Investment Research Tool
================================================================================

QUE CONTIENE ESTE ARCHIVO
--------------------------------------------------------------------------------
Una aplicacion Streamlit de un solo archivo con cinco pestanas (COMMODITIES,
PERU, USA, WORLD, FFMM) que combina monitoreo de mercados, investigacion
macro, analisis cross-asset y "data storytelling" (interpretacion automatica
tipo Portfolio Manager / CIO, inspirada en el formato de weekly reports
institucionales como el "CIO Market Pulse" de DBS, y en general el formato de
HSBC/BlackRock/Goldman Sachs: bloques cortos de texto en negrita + 2-3
oraciones con el mecanismo, tablas de performance por periodo con codigo de
color, y graficos de atribucion).

UNICO INPUT LOCAL: BD_Consolidada_Dash_Macro.xlsx
--------------------------------------------------------------------------------
Toda la data que antes vivia en dos archivos separados (ZODA.xlsx y
Portafolio_FFMM.xlsx) y se resolvia via Yahoo Finance ahora proviene de UN
SOLO Excel consolidado, con estas hojas (ver seccion 2.5 CONFIG/EXCEL
CONSOLIDADO):
  - "Precios_ZODA": precios de TODO el equity que transa en la BVL + equity
    de empresas peruanas que transa en USA (ADRs), mas los 6 benchmarks
    regionales (S&P/BVL Peru General, IPSA, IBOVESPA, MEXBOL, COLCAP, MSCI
    Latam) como INDICES REALES -ya no proxies de ETF-.
  - "FFMM_Precios_Retornos" + "Parametria_FFMM": precio y retorno diario YA
    calculados (por ISIN) de los 45 fondos/ETFs tematicos del portafolio, y
    su metadata (Nombre/AssetClass/Mercado/Cat/SubCat/SubCat2/ISIN). Ya NO se
    usa Yahoo Finance para resolver tickers de fondos.
  - "Precios_ZODA_CicloIA": 46 companias de la cadena de valor de IA/
    semiconductores, con su capa (columna CapaIA: Compute, Semicap Equipment,
    Analog / Power Management, Networking, Datacenter, Cloud, Applications,
    Power, Grid, Cooling, Electrical Equipment, All). Se usa en dos sitios:
    (a) como universo PROXY equal-weight de holdings para el performance
    attribution de los fondos FFMM de tecnologia/IA (ver mas abajo), y (b)
    como base del chart didactico "Ciclo de Valor de IA" en la pestana FFMM,
    que muestra como performea cada una de las 12 capas, con un glosario de
    que es cada una.
  - "Produccion_Mineria" e "Indicadores_SBS_Bancos": data de produccion
    fisica mensual de las mineras y de indicadores financieros SBS (Morosidad,
    ROE, Liquidez) de los bancos que cotizan, usados en la seccion
    "Fundamentals" de la pestana PERU junto al precio de la accion.

Yahoo Finance se sigue usando SOLO para: el tipo de cambio USD/PEN, las
series macro de USA/WORLD y los dos pares mineros internacionales de
referencia (Southern Copper, Hudbay Minerals) que no estan en el Excel.

PERFORMANCE ATTRIBUTION DE FFMM (estilo CFA, proxy equal-weight)
--------------------------------------------------------------------------------
El Excel consolidado NO trae los holdings oficiales de cada fondo/ETF (eso
requeriria una fuente paga tipo Morningstar Direct o el factsheet PDF de
cada proveedor). En su lugar, para los fondos de tecnologia/IA se usa el
universo de "Precios_ZODA_CicloIA" como proxy EQUAL-WEIGHT:
  - SubCat2 == "Semiconductors": solo las 3 capas que son, en conjunto, la
    cadena de valor angosta de semiconductores (Compute + Semicap Equipment +
    Analog / Power Management).
  - SubCat2 en {"Core", "Software & Internet", "Thematic"}: el universo
    COMPLETO de las 12 capas, como proxy de la cadena de valor de IA en
    general.
  - Cualquier otro fondo (regionales, de commodities, etc.): NO se construye
    atribucion -no hay ningun universo proxy razonable para ellos, y se
    prefiere declarar la limitacion a fabricar una contribucion ficticia.
Cada lectura de atribucion declara explicitamente que es una aproximacion
equal-weight con pesos vigentes (no los holdings reales ni sus pesos reales
del fondo), nunca presentada como una atribucion Brinson exacta.

COMO EJECUTARLA
--------------------------------------------------------------------------------
1) Instalar dependencias (un solo comando):

   pip install streamlit pandas numpy plotly yfinance requests scipy openpyxl fredapi

2) Ejecutar:

   streamlit run global_macro_terminal.py

No se requiere ninguna API key para que la app funcione: FRED se consulta
primero via la libreria oficial `fredapi` (mas robusta) y, si la API key no
es valida o fredapi no esta instalado, cae automaticamente al endpoint
publico de CSV de FRED (fredgraph.csv), que no requiere autenticacion. Para
usar fredapi con tu propia cuenta, reemplaza la constante FRED_API_KEY (mas
abajo, seccion 1 CONFIG) por tu API key real y gratuita, obtenida en
https://fred.stlouisfed.org/docs/api/api_key.html — la que trae el archivo
por defecto es un placeholder de ejemplo y no funcionara hasta que la
reemplaces.

COMO SUBIR BD_Consolidada_Dash_Macro.xlsx A GITHUB (para no subirlo cada sesion)
--------------------------------------------------------------------------------
Si despliegas esta app en Streamlit Cloud (o cualquier otro host que lea el
codigo desde un repo de GitHub), NO es necesario usar el uploader de la
barra lateral en cada sesion -eso seria inestable e implica volver a
subirlo cada vez que alguien abre la app-. En su lugar:

1) En tu repo de GitHub, crea una carpeta "data/" junto a este archivo
   (ej. si tu app esta en "app.py" en la raiz del repo, crea "data/").
2) Sube ahi el archivo con ese nombre exacto:
     data/BD_Consolidada_Dash_Macro.xlsx
3) Vuelve a desplegar (o espera el redeploy automatico de Streamlit Cloud).

La app busca automaticamente ese archivo al iniciar (ver DATA_DIR y
_resolve_data_file() en la seccion 1 CONFIG) y lo usa como fuente por
defecto -sin pedir nada al usuario-. Si por alguna razon prefieres no tener
la carpeta "data/", tambien funciona si colocas el xlsx directamente en la
raiz del repo (junto a este .py); la app revisa ambas ubicaciones en ese
orden. El uploader de la barra lateral SIGUE disponible pero ahora es un
override opcional de sesion (por ejemplo, para probar una version corregida
del Excel sin tener que hacer commit todavia): si subes un archivo
manualmente, ese tiene prioridad sobre el que esta en el repo; si no subes
nada, se usa el del repo; y si ninguno de los dos existe, las pestanas PERU
y FFMM quedan sin datos locales (declarado explicitamente en la UI), pero
COMMODITIES/USA/WORLD siguen funcionando porque dependen solo de fuentes
externas gratuitas.

LIMITACIONES CONOCIDAS (documentadas segun lo solicitado, no se fabrican datos)
--------------------------------------------------------------------------------
- Todo instrumento (equity Peru, fondo FFMM, compania del Ciclo de Valor de
  IA) sin ISIN/Ticker en su hoja de parametria, o cuyo ISIN no matcheo
  contra la base de precios de origen del Excel, se muestra como "N/D" en
  vez de inventar un precio o adivinar un ticker.
- El performance attribution de FFMM es un proxy equal-weight con el
  universo de Precios_ZODA_CicloIA (ver seccion anterior) solo para fondos
  de tecnologia/IA; el resto de fondos no tiene atribucion por falta de un
  proxy razonable. Se declara en cada lectura.
- La produccion minera se agrega por empresa+mineral principal restringiendo
  a Etapa='CONCENTRACION'/Proceso='FLOTACION' (la salida comparable entre
  plantas de una misma empresa); otras etapas (lixiviacion, refinacion) no
  se incluyen para evitar mezclar unidades de medida distintas en una misma
  serie.
- Los indicadores SBS (Morosidad/ROE/Liquidez) vienen en el Excel como
  fraccion (ej. 0.0264) y se convierten a puntos porcentuales (x100) en el
  punto de entrada de los datos (sbs_indicator_series), para que toda la app
  trabaje consistentemente en pp.
- COLCAP y el MSCI Latam del benchmark regional ahora son los indices reales
  del Excel (ya no proxies de ETF como antes).
- Taiwan no es miembro del Banco Mundial, por lo que la serie de crecimiento
  del PBI (World Bank, via FRED) no existe para Taiwan; se omite con una
  nota en vez de mostrar un dato falso.
- Series de crecimiento del PBI (WORLD_COUNTRIES, columna "PBI real YoY %"):
  se verifico y corrigio el ID de FRED para Peru, Brasil y Colombia (el ID
  original "NYGDPMKTPKDZG<pais>", mirror de World Bank, esta discontinuado
  en FRED al menos para Peru). Para Japon, Zona Euro, Reino Unido, Mexico y
  Chile NO se encontro un reemplazo verificado con las fuentes disponibles;
  si esa fila muestra "N/D" en vez de un numero, probablemente el ID
  original tambien fue discontinuado y se necesita investigar el ID
  correcto en https://fred.stlouisfed.org (buscar "Gross Domestic Product
  for <pais>, Percent Change") en vez de asumir un valor.
- Todas las descargas externas (FRED, Yahoo Finance FX/macro/pares mineros)
  usan cache con TTL y manejo de errores: si una fuente falla, el componente
  correspondiente se marca como "dato no disponible" en lugar de
  interrumpir la aplicacion o mostrar un valor inventado.
- El storytelling es generado por un motor de reglas (patrones
  cuantitativos: direccion, magnitud, percentil historico, aceleracion de
  tendencia, divergencias entre indicadores) a partir de los datos
  observados en tiempo de ejecucion. No es un modelo de lenguaje ni usa
  IA generativa: es determinista y reproducible, y evita afirmar
  causalidad a partir de correlacion.

ARQUITECTURA (resumen)
--------------------------------------------------------------------------------
1. CONFIG            -> catalogos de series FRED / tickers yfinance por pestana,
                         metadata de equity/benchmarks Peru, glosario de las
                         12 capas del Ciclo de Valor de IA
2. DATA FETCH LAYER   -> fetch_fred() (fredapi + fallback CSV), fetch_yf(),
                         fetch_stooq() con cache y manejo de errores,
                         devuelven un FetchResult uniforme
2.5 EXCEL CONSOLIDADO -> _read_consolidated_sheet() (cacheada) + loaders por
                         hoja (load_precios_zoda, load_ffmm_precios,
                         load_ciclo_ia, load_produccion_mineria,
                         load_sbs_bancos, load_parametria_ffmm) del UNICO
                         input BD_Consolidada_Dash_Macro.xlsx
3. ANALYTICS LAYER    -> retornos, volatilidad, drawdown, z-score, percentil,
                         correlaciones, pendiente de curva de rendimientos,
                         retornos WTD/MTD/YTD/1Y/3Y y CAGR, indice equal-
                         weight por capa del Ciclo de Valor de IA
4. STORYTELLING ENGINE-> biblioteca de patrones interpretativos reutilizables
                         + interpretes compuestos (labor, inflacion, mineria,
                         SBS, capas IA, etc.) + performance attribution
                         estilo CFA (proxy equal-weight para FFMM)
5. UI HELPERS         -> tarjetas de metricas, badges de frescura, graficos
                         Plotly estandarizados, tablas con codigo de color,
                         heatmap de correlaciones, scatter riesgo/retorno
6. RESOLUCION DEL INPUT -> resolve_consolidated_source(): repo primero,
                         upload manual como override de sesion
7. TABS               -> render_commodities_tab / render_peru_tab /
                         render_usa_tab / render_world_tab / render_ffmm_tab
8. MAIN               -> page config, sidebar, orquestacion de pestanas
================================================================================
"""

from __future__ import annotations

import io
import logging
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Sequence

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

warnings.filterwarnings("ignore")

try:
    import yfinance as yf
    # yfinance escribe sus propios logs de ERROR a stderr (ej. "symbol may be
    # delisted") incluso cuando esta app ya captura la excepcion y la maneja
    # con gracia (muestra "N/D"). Esto es ruido esperado para los tickers BVL
    # sin cobertura (ver limitaciones): se silencia para que los logs de
    # produccion (ej. Streamlit Cloud) solo muestren errores reales de la app.
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)
except ImportError:  # pragma: no cover
    yf = None

try:
    from fredapi import Fred
except ImportError:  # pragma: no cover
    Fred = None

# ==============================================================================
# 1. CONFIG / CONSTANTS
# ==============================================================================

APP_TITLE = "Global Markets & Macro Intelligence Terminal"
CACHE_TTL_INTRADAY = 60 * 30        # 30 min para precios de mercado
CACHE_TTL_MACRO = 60 * 60 * 6       # 6 horas para series macro (mensuales/trimestrales)
REQUEST_TIMEOUT = 15
HTTP_HEADERS = {"User-Agent": "Mozilla/5.0 (MacroTerminal/1.0; research use)"}

# Carpeta de datos del repo (ver docstring del modulo, "COMO SUBIR EL EXCEL
# CONSOLIDADO A GITHUB"): si existe, la app lee el UNICO input de datos
# locales (BD_Consolidada_Dash_Macro.xlsx) desde aqui al iniciar, sin
# pedirle nada al usuario.
DATA_DIR = Path(__file__).resolve().parent / "data"
APP_DIR = Path(__file__).resolve().parent

# Nombre exacto del archivo Excel consolidado (UNICO input local de esta
# app: equity peruano, benchmarks regionales, FFMM y Ciclo IA -ver seccion
# 2.5-). Se busca primero en DATA_DIR, luego en la raiz del repo.
CONSOLIDATED_XLSX_NAME = "BD_Consolidada_Dash_Macro.xlsx"


def _resolve_data_file(filename: str) -> Optional[Path]:
    """Busca `filename` primero en DATA_DIR (<repo>/data/<filename>) y luego
    en la raiz del repo (junto a este .py), para aceptar ambas convenciones.
    Devuelve el primer Path que exista como archivo, o None si no se
    encuentra en ninguna ubicacion -en ese caso la app sigue funcionando
    normalmente con los uploaders de la barra lateral y/o los universos por
    defecto, sin romperse-."""
    for candidate in (DATA_DIR / filename, APP_DIR / filename):
        if candidate.is_file():
            return candidate
    return None

DEFAULT_LOOKBACK_YEARS = 5

# <<< REEMPLAZA ESTO POR TU API KEY REAL Y GRATUITA DE FRED >>>
# Se obtiene en https://fred.stlouisfed.org/docs/api/api_key.html
# El valor de abajo es un PLACEHOLDER de ejemplo: con el no funcionara
# fredapi, pero la app sigue operando igual gracias al fallback automatico
# al endpoint publico de CSV de FRED (ver fetch_fred() mas abajo).
FRED_API_KEY = "762e2ee1c8fab5d038ce317929d47226"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("macro_terminal")


# ------------------------------------------------------------------------------
# 1.1 COMMODITIES catalog (Yahoo Finance front-month continuous futures)
# ------------------------------------------------------------------------------
COMMODITIES = {
    "WTI Crude Oil":   {"ticker": "CL=F", "unit": "USD/bbl", "asset_class": "energy",
                          "peru_relevance": "baja"},
    "Brent Crude Oil": {"ticker": "BZ=F", "unit": "USD/bbl", "asset_class": "energy",
                          "peru_relevance": "baja"},
    "Natural Gas":     {"ticker": "NG=F", "unit": "USD/MMBtu", "asset_class": "energy",
                          "peru_relevance": "baja"},
    "Gold":            {"ticker": "GC=F", "unit": "USD/oz", "asset_class": "precious_metal",
                          "peru_relevance": "alta"},
    "Silver":          {"ticker": "SI=F", "unit": "USD/oz", "asset_class": "precious_metal",
                          "peru_relevance": "alta"},
    "Copper":          {"ticker": "HG=F", "unit": "USD/lb", "asset_class": "industrial_metal",
                          "peru_relevance": "muy alta"},
}

# ------------------------------------------------------------------------------
# 1.2 PERU: universo de equity local + benchmarks, derivado del UNICO INPUT
#     de esta app (BD_Consolidada_Dash_Macro.xlsx, hoja "Precios_ZODA"), no
#     de Yahoo Finance. El Excel ya trae Ticker (BBG) + Sector + Precio por
#     fecha para TODO equity que transa en la BVL y ademas para las empresas
#     peruanas que transan en USA/Londres (ADRs), mas los INDICES REALES de
#     referencia (ya no proxies de ETF como en la version anterior de esta
#     app: SPBLPGPT Index es el S&P/BVL Peru General autentico, no el ETF
#     EPU). Ver seccion 2.5 (EXCEL CONSOLIDADO) para la logica de carga y
#     seccion 6 (LOADERS) para como se resuelve el archivo (repo o subida).
#
#     Los diccionarios de abajo son METADATA que NO esta en el Excel (nombre
#     comercial corto, tipo de listado) y sirven solo para presentacion;
#     el dato numerico (precio, sector) siempre viene del Excel. La llave es
#     el Ticker BBG exacto tal como aparece en la columna "Ticker" del Excel.
# ------------------------------------------------------------------------------
PERU_EQUITY_META = {
    "BAP US Equity":      {"nombre": "Credicorp (BAP)",            "listing": "ADR / NYSE"},
    "IFS PE Equity":      {"nombre": "Intercorp Financial Serv.",   "listing": "NYSE / BVL"},
    "BBVAC1 PE Equity":   {"nombre": "BBVA Peru",                   "listing": "BVL (local)"},
    "CPACASC1 PE Equity": {"nombre": "Cementos Pacasmayo",          "listing": "BVL (local)"},
    "AENZAC1 PE Equity":  {"nombre": "AENZA",                       "listing": "BVL (local)"},
    "FERREYC1 PE Equity": {"nombre": "Ferreycorp",                  "listing": "BVL (local)"},
    "CORAREC1 PE Equity": {"nombre": "Aceros Arequipa",             "listing": "BVL (local)"},
    "BVN US Equity":      {"nombre": "Buenaventura (BVN)",          "listing": "ADR / NYSE"},
    "CVERDEC1 PE Equity": {"nombre": "Cerro Verde",                 "listing": "BVL (local)"},
    "NEXAPEC1 PE Equity": {"nombre": "Nexa Peru (local)",           "listing": "BVL (local)"},
    "NEXA US Equity":     {"nombre": "Nexa Resources (ADR)",        "listing": "NYSE"},
    "ATACOBC1 PE Equity": {"nombre": "Nexa Atacocha",               "listing": "BVL (local)"},
    "MINSURI1 PE Equity": {"nombre": "Minsur",                      "listing": "BVL (local)"},
    "VOLCABC1 PE Equity": {"nombre": "Volcan",                      "listing": "BVL (local)"},
    "HOC LN Equity":      {"nombre": "Hochschild Mining (Ares)",    "listing": "ADR / LSE"},
    "ORYGENC1 PE Equity": {"nombre": "Orygen",                      "listing": "BVL (local)"},
    "PLUZENC1 PE Equity": {"nombre": "Luz del Sur (Pluzen)",        "listing": "BVL (local)"},
    "ENGIEC1 PE Equity":  {"nombre": "Engie Peru",                  "listing": "BVL (local)"},
    "ALICORC1 PE Equity": {"nombre": "Alicorp",                     "listing": "BVL (local)"},
    "INRETC1 PE Equity":  {"nombre": "InRetail",                    "listing": "BVL (local)"},
    "BACKUSI1 PE Equity": {"nombre": "Backus (AB InBev Peru)",      "listing": "BVL (local)"},
    "AUNA US Equity":     {"nombre": "Auna",                        "listing": "NYSE"},
    "ELCOMEI1 PE Equity": {"nombre": "El Comercio",                 "listing": "BVL (local)"},
}

PERU_BENCHMARK_META = {
    "SPBLPGPT Index": {"nombre": "S&P/BVL Peru General",  "nota": "Indice real de la BVL (ya no se usa el ETF EPU como proxy)."},
    "IPSA Index":     {"nombre": "IPSA (Chile)",          "nota": ""},
    "IBOV Index":     {"nombre": "IBOVESPA (Brasil)",     "nota": ""},
    "MEXBOL Index":   {"nombre": "MEXBOL / S&P BMV IPC (Mexico)", "nota": ""},
    "COLCAP Index":   {"nombre": "COLCAP (Colombia)",     "nota": "Indice real (ya no se usa un ETF proxy como GXG/ICOL)."},
    "MXLA Index":     {"nombre": "MSCI Latam (MXLA)",     "nota": "Benchmark regional de referencia."},
}

# Pares mineros internacionales que SI siguen via Yahoo Finance (no estan en
# el Excel consolidado: son comparables externos, no equity local peruano).
PERU_MINING_PEERS = [
    {"nombre": "Southern Copper", "bbg": "SCCO US Equity", "yahoo_candidates": ["SCCO"]},
    {"nombre": "Hudbay Minerals", "bbg": "HBM CN Equity",  "yahoo_candidates": ["HBM"]},
]

# Mineral principal por ticker minero (para el bloque de "Fundamentals":
# produccion fisica vs. desempeno de la accion). Se usa UN mineral por
# empresa -el de mayor relevancia de ingresos- para evitar mezclar etapas de
# proceso/unidades de medida distintas en un mismo grafico. Fuente: hoja
# "Produccion_Mineria" del Excel consolidado, agregando SOLO las filas de
# la etapa "CONCENTRACION"/proceso "FLOTACION" (la salida fisica principal,
# comparable entre plantas de una misma empresa) por Ticker+Mineral+Fecha.
MINING_PRIMARY_METAL = {
    "CVERDEC1 PE Equity": "COBRE",
    "MINSURI1 PE Equity": "ESTAÑO",
    "BVN US Equity":      "ORO",
    "HOC LN Equity":      "ORO",
    "VOLCABC1 PE Equity": "ZINC",
    "ATACOBC1 PE Equity": "ZINC",
    "NEXAPEC1 PE Equity": "ZINC",
    "NEXA US Equity":     "ZINC",
}

# Mapeo Ticker BVL/ADR -> nombre de "Entidad" en la hoja Indicadores_SBS_Bancos
# (ya viene mapeado en el pipeline de origen, ver README del Excel).
SBS_BANK_TICKERS = {
    "BAP US Equity":    "BCP",
    "IFS PE Equity":    "INTERBANK",
    "BBVAC1 PE Equity": "BBVA",
}
SBS_SYSTEM_LABEL = "SISTEMA BANCARIO"
# Indicadores SBS headline que se muestran como "fundamentals" del banco
# (Codigo exacto de la hoja Indicadores_SBS_Bancos; el prefijo PERxxx_BAN_
# varia por banco asi que se matchea por el sufijo).
SBS_HEADLINE_INDICATORS = {
    "Morosidad Total":          "MOR_MTO",
    "ROE (anualizado)":         "SBS_REN_ROE",
    "Liquidez MN (prom. mes)":  "SBS_LIQ_LMN",
}

PERU_FX = {"ticker": "PEN=X", "nombre": "USD/PEN"}

# Series macro de Peru disponibles gratuitamente via FRED (World Bank mirror /
# IMF WEO). Son de baja frecuencia (anual); se documentan como tales. Cada
# valor es una LISTA de IDs candidatos (se intenta en orden, via
# fetch_fred_with_candidates) porque distintas fuentes mirror en FRED publican
# el mismo dato bajo convenciones de ID distintas y no todas estan siempre
# vigentes. "PERNGDPRPCH" (IMF WEO, Western Hemisphere REO: "Gross Domestic
# Product for Peru, Percent Change") es la serie verificada y correcta al
# momento de escribir esto; "NYGDPMKTPKDZGPER" (World Bank mirror) se deja
# como segundo candidato por si en el futuro vuelve a publicarse, pero NO
# existe actualmente en FRED (causaba "Bad Request. The series does not
# exist." en produccion).
PERU_MACRO_FRED = {
    "Inflacion (CPI, var. % anual)": ["FPCPITOTLZGPER"],
    "Crecimiento del PBI real (var. % anual)": ["PERNGDPRPCH", "NYGDPMKTPKDZGPER"],
}

# ------------------------------------------------------------------------------
# 1.3 USA macro catalog (FRED series IDs, todas de acceso publico sin API key)
# ------------------------------------------------------------------------------
USA_LABOR = {
    "Nonfarm Payrolls (miles, nivel)":   "PAYEMS",
    "Tasa de desempleo (%)":             "UNRATE",
    "Solicitudes iniciales de desempleo":"ICSA",
    "Tasa de participacion laboral (%)": "CIVPART",
    "Salario promedio por hora (USD)":   "CES0500000003",
}

USA_INFLATION = {
    "CPI headline (indice)":      "CPIAUCSL",
    "CPI core (indice)":          "CPILFESL",
    "PCE headline (indice)":      "PCEPI",
    "PCE core (indice)":          "PCEPILFE",
    "Breakeven inflacionario 5Y (%)": "T5YIE",
}

USA_FED_POLICY = {
    "Fed Funds Rate (%)":          "FEDFUNDS",
    "Balance de la Fed (USD mill)":"WALCL",
    "Reverse Repo overnight (USD mill)": "RRPONTSYD",
}

USA_RATES = {
    "UST 2 anos (%)":  "DGS2",
    "UST 10 anos (%)": "DGS10",
    "UST 30 anos (%)": "DGS30",
    "Spread 10Y-2Y (pp)": "T10Y2Y",
}

USA_FISCAL = {
    "Deuda publica federal (USD mill)":        "GFDEBTN",
    "Deuda publica / PBI (%)":                 "GFDEGDQ188S",
    "Resultado fiscal mensual (USD mill)":      "MTSDS133FMS",
}

USA_STRESS = {
    "Chicago Fed Financial Conditions (NFCI)": "NFCI",
    "St. Louis Fed Financial Stress Index":    "STLFSI4",
}
USA_VIX_TICKER = "^VIX"

# ------------------------------------------------------------------------------
# 1.4 WORLD country scorecards
#     cpi: OECD MEI "CPI, all items, YoY %"; rate: central bank policy proxy
#     y10: long-term govt bond yield (OECD MEI); gdp: World Bank GDP growth
#     fx: par Yahoo vs USD; equity: indice bursatil Yahoo
# ------------------------------------------------------------------------------
# NOTA sobre "gdp": cada valor es una LISTA de IDs candidatos de FRED (se
# intenta en orden via fetch_fred_with_candidates), igual que en
# PERU_MACRO_FRED. El patron World Bank-mirror "NYGDPMKTPKDZG<ISO3>" que
# usaba esta app para TODOS los paises resulto estar descontinuado en FRED al
# menos para Peru (confirmado en produccion: "Bad Request. The series does
# not exist."). Se verifico que el patron alterno IMF WEO "<ISO3>NGDPRPCH"
# ("Gross Domestic Product for <Pais>, Percent Change") SI existe y esta
# vigente para Brasil y Colombia (al igual que para Peru, ver
# PERU_MACRO_FRED), asi que se uso como candidato principal para esos dos.
# Para Japon, China, Zona Euro, Reino Unido, Mexico y Chile NO se pudo
# verificar un reemplazo confiable (el patron IMF WEO no aplica de forma
# uniforme fuera de la Region Hemisferio Occidental del FMI, y no se
# encontraron series alternativas confirmadas sin acceso a mas fuentes) -
# se deja el ID original como unico candidato, documentado aqui como
# "no verificado": si FRED lo discontinuo tambien, esa fila mostrara "N/D"
# en vez de un numero fabricado, que es el comportamiento correcto para
# este caso (ver limitaciones en el panel lateral de la app).
WORLD_COUNTRIES = {
    "Japon": {
        "cpi": "CPALTT01JPM657N", "rate": "IRSTCB01JPM156N", "y10": "IRLTLT01JPM156N",
        "gdp": ["NYGDPMKTPKDZGJPN"], "fx": "JPY=X", "equity": "^N225",
        "nota": "PBI real (YoY %): serie no re-verificada tras el bug de FRED detectado en la serie de Peru; si aparece N/D, es posible que FRED haya descontinuado tambien este ID.",
    },
    "China": {
        "cpi": "CHNCPIALLMINMEI", "rate": None, "y10": None,
        "gdp": ["NYGDPMKTPKDZGCHN"], "fx": "CNY=X", "equity": "^HSI",
        "nota": "Profundidad limitada: China publica con menor frecuencia/transparencia que economias OECD; se usa Hang Seng (^HSI) como proxy bursatil de referencia regional y CPI via fuente alterna. PBI real (YoY %): serie no re-verificada, puede mostrar N/D (ver nota general sobre series de PBI).",
    },
    "Zona Euro": {
        "cpi": "CP0000EZ19M086NEST", "rate": "IRSTCB01EZM156N", "y10": "IRLTLT01EZM156N",
        "gdp": ["NYGDPMKTPKDZGEMU"], "fx": "EURUSD=X", "equity": "^STOXX50E",
        "nota": "PBI real (YoY %): serie no re-verificada tras el bug de FRED detectado en la serie de Peru; si aparece N/D, es posible que FRED haya descontinuado tambien este ID.",
    },
    "Reino Unido": {
        "cpi": "CPALTT01GBM657N", "rate": "IRSTCB01GBM156N", "y10": "IRLTLT01GBM156N",
        "gdp": ["NYGDPMKTPKDZGGBR"], "fx": "GBPUSD=X", "equity": "^FTSE",
        "nota": "PBI real (YoY %): serie no re-verificada tras el bug de FRED detectado en la serie de Peru; si aparece N/D, es posible que FRED haya descontinuado tambien este ID.",
    },
    "Brasil": {
        "cpi": "CPALTT01BRM657N", "rate": "IRSTCB01BRM156N", "y10": None,
        "gdp": ["BRANGDPRPCH", "NYGDPMKTPKDZGBRA"], "fx": "BRL=X", "equity": "^BVSP",
    },
    "Mexico": {
        "cpi": "CPALTT01MXM657N", "rate": "IRSTCB01MXM156N", "y10": "IRLTLT01MXM156N",
        "gdp": ["NYGDPMKTPKDZGMEX"], "fx": "MXN=X", "equity": "^MXX",
        "nota": "PBI real (YoY %): serie no re-verificada tras el bug de FRED detectado en la serie de Peru; si aparece N/D, es posible que FRED haya descontinuado tambien este ID.",
    },
    "Chile": {
        "cpi": "CPALTT01CLM657N", "rate": "IRSTCB01CLM156N", "y10": "IRLTLT01CLM156N",
        "gdp": ["NYGDPMKTPKDZGCHL"], "fx": "CLP=X", "equity": "^IPSA",
        "nota": "PBI real (YoY %): serie no re-verificada tras el bug de FRED detectado en la serie de Peru; si aparece N/D, es posible que FRED haya descontinuado tambien este ID.",
    },
    "Taiwan": {
        "cpi": "CPALTT01TWM657N", "rate": None, "y10": None,
        "gdp": None, "fx": "TWD=X", "equity": "^TWII",
        "nota": "Taiwan no es miembro del Banco Mundial: no existe serie oficial de crecimiento del PBI en esta fuente.",
    },
    "Colombia": {
        "cpi": "CPALTT01COM657N", "rate": "IRSTCB01COM156N", "y10": None,
        "gdp": ["COLNGDPRPCH", "NYGDPMKTPKDZGCOL"], "fx": "COP=X", "equity": "GXG",
    },
}

ASSET_CLASS_LABELS = {
    "energy": "energia", "precious_metal": "metal precioso",
    "industrial_metal": "metal industrial", "equity": "renta variable",
    "fx": "tipo de cambio", "rate": "tasa de interes", "macro": "indicador macro",
}

# ------------------------------------------------------------------------------
# 1.5 FFMM: universo de fondos mutuos / ETFs tematicos. Metadata (Nombre,
#     AssetClass, Mercado, Cat, SubCat, SubCat2, ISIN) e historia de precios
#     ya NO se adivinan por ticker de Yahoo Finance: vienen directo del
#     UNICO INPUT de esta app (BD_Consolidada_Dash_Macro.xlsx), hojas
#     "Parametria_FFMM" (metadata) y "FFMM_Precios_Retornos" (precio/retorno
#     diario por ISIN, ya calculado en el pipeline de origen). Ver seccion
#     2.5 (EXCEL CONSOLIDADO). Si un fondo no tiene ISIN en la parametria, o
#     su ISIN no matcheo contra BDPrecios_VB_IN.xlsx en el pipeline de
#     origen (ver hoja README del Excel), queda sin precio y se muestra
#     "N/D" -nunca se inventa ni se adivina un ticker.
# ------------------------------------------------------------------------------

# Capas del ciclo de valor de IA (hoja "Precios_ZODA_CicloIA", columna
# CapaIA) que se usan como universo PROXY de holdings para el performance
# attribution de los fondos/ETFs tematicos de tecnologia (ver seccion 4.2 y
# render_ffmm_tab). "Compute"+"Semicap Equipment"+"Analog / Power Management"
# son, en conjunto, la cadena de valor de SEMICONDUCTORES propiamente dicha
# (fabricacion de chips + las maquinas que los fabrican); se usa este
# subconjunto mas angosto para los ETFs de SubCat2=="Semiconductors", y el
# universo COMPLETO de las 12 capas para los fondos de tecnologia/IA mas
# amplios (Core/Software & Internet/Thematic), como proxy de la cadena de
# valor de IA en general. IMPORTANTE: esto es una aproximacion EQUAL-WEIGHT
# (no son los holdings oficiales reportados por el proveedor del ETF, que
# esta base de datos no trae) y se declara explicitamente en cada lectura.
FFMM_SEMI_CAPAS = {"Compute", "Semicap Equipment", "Analog / Power Management"}

# Definiciones concisas (estilo glosario, en Spanglish institucional) de
# cada una de las 12 capas del ciclo de valor de IA, para memorizarlas.
AI_CYCLE_LAYER_DEFS = {
    "Compute": (
        "Los chips que procesan la carga de trabajo de IA: GPUs/aceleradores (NVIDIA, AMD), el foundry que los "
        "fabrica (TSMC, Samsung), memoria (Micron) y networking silicon (Broadcom, Marvell, Qualcomm). Es el "
        "corazon del ciclo: sin Compute no hay training ni inference."
    ),
    "Semicap Equipment": (
        "Las maquinas que fabrican los chips -litografia, deposicion, etch-: ASML, Applied Materials, Lam "
        "Research, KLA, Tokyo Electron. Son un leading indicator del ciclo: sus ordenes anticipan el capex de "
        "los fabricantes de chips varios quarters antes."
    ),
    "Analog / Power Management": (
        "Semiconductores analogicos que gestionan voltaje y power delivery dentro de cada chip/servidor (Texas "
        "Instruments). Menos glamoroso que el Compute puro, pero critico: cada GPU necesita power management "
        "de precision para no degradarse termicamente."
    ),
    "Networking": (
        "El 'sistema nervioso' que conecta miles de GPUs entre si y con el resto del datacenter -switches, "
        "routers-: Arista, Cisco, Juniper. Mas Compute instalado implica, casi mecanicamente, mas ancho de "
        "banda de red requerido."
    ),
    "Datacenter": (
        "Los REITs/operadores que arriendan el espacio fisico, energia y conectividad donde vive el Compute: "
        "Equinix, Digital Realty. Monetizan la demanda de IA via leasing de largo plazo, no via venta de chips."
    ),
    "Cloud": (
        "Hyperscalers que arriendan Compute como servicio (Oracle Cloud Infrastructure es el nombre 'puro' en "
        "esta canasta; el resto de hyperscalers -Amazon, Microsoft, Alphabet- aparece en 'All' por ser multi-capa)."
    ),
    "Applications": (
        "El software que corre SOBRE la infraestructura de IA y monetiza directamente el caso de uso final: "
        "Salesforce, ServiceNow, Adobe, Palantir, Snowflake. Capturan valor del lado de la demanda (enterprise "
        "software + IA), no del lado del hardware."
    ),
    "Power": (
        "Generadoras/utilities que venden la electricidad que consume el Compute -los datacenters de IA son "
        "extremadamente energo-intensivos-: Constellation, NextEra, Duke Energy, GE Vernova, Vistra. El cuello "
        "de botella energetico del ciclo de IA pasa, en gran medida, por aqui."
    ),
    "Grid": (
        "Transmision y distribucion de esa electricidad hasta el datacenter: American Electric Power, National "
        "Grid, PG&E. Si la Power no llega por falta de capacidad de Grid, el Compute simplemente no puede operar."
    ),
    "Cooling": (
        "Sistemas de enfriamiento -aire, liquido- que evitan que los racks de GPU se sobrecalienten: Vertiv, "
        "Trane Technologies, Carrier Global, Johnson Controls. A mayor densidad de Compute por rack, mayor "
        "intensidad de cooling requerida (liquid cooling es la tendencia dominante)."
    ),
    "Electrical Equipment": (
        "Equipos electricos -tableros, UPS, distribucion de potencia- dentro del datacenter mismo: Legrand. "
        "Conecta la Power/Grid externa con el Compute interno de forma segura y redundante."
    ),
    "All": (
        "Hyperscalers y conglomerados industriales que participan en MAS DE UNA capa a la vez: Amazon, "
        "Alphabet, Meta y Microsoft como compradores masivos de Compute y, a la vez, operadores de Cloud/"
        "Datacenter propios; ABB, Siemens y Schneider Electric como proveedores diversificados de Power/Grid/"
        "Electrical Equipment. Se agrupan aparte para no distorsionar el analisis de una capa 'pura'."
    ),
}


# ==============================================================================
# 2. DATA FETCH LAYER
# ==============================================================================

@dataclass
class FetchResult:
    label: str
    series: Optional[pd.Series] = None
    source: str = ""
    status: str = "error"          # ok | stale | error | empty
    asof: Optional[pd.Timestamp] = None
    error: Optional[str] = None
    frequency_days: Optional[int] = None  # cadencia tipica esperada, para flag de "stale"
    resolved_ticker: Optional[str] = None  # cual candidato de la lista funciono (fetch_with_candidates)

    @property
    def ok(self) -> bool:
        return self.status in ("ok", "stale") and self.series is not None and len(self.series) > 0


def _infer_staleness(last_date: pd.Timestamp, expected_gap_days: int) -> str:
    if last_date is None or pd.isna(last_date):
        return "error"
    gap = (pd.Timestamp(datetime.utcnow().date()) - pd.Timestamp(last_date).normalize()).days
    return "stale" if gap > expected_gap_days else "ok"


@st.cache_resource(show_spinner=False)
def _get_fred_client():
    """Crea (una sola vez, cacheado) el cliente fredapi si la libreria esta
    instalada y hay una API key configurada. Si falla, devuelve None y el
    resto de la app sigue funcionando via el fallback de CSV publico.
    Usa st.cache_resource (no st.cache_data) a proposito: el cliente de
    fredapi no es serializable y no necesita serlo, solo vivir una vez en
    memoria."""
    if Fred is None or not FRED_API_KEY or not FRED_API_KEY.strip():
        return None
    try:
        return Fred(api_key=FRED_API_KEY.strip())
    except Exception as exc:  # noqa: BLE001
        logger.info("No se pudo inicializar fredapi: %s", exc)
        return None


# ------------------------------------------------------------------------------
# NOTA DE DISENO IMPORTANTE (evita un bug real visto en produccion/Streamlit
# Cloud): las funciones decoradas con @st.cache_data NUNCA deben devolver
# FetchResult (ni ninguna otra clase definida en este mismo script). Streamlit
# serializa (pickle) lo que cachea con @st.cache_data, y cuando una clase vive
# en el modulo "__main__" (que es como Streamlit ejecuta este archivo), cada
# re-ejecucion del script puede dejar en memoria una version "vieja" de esa
# clase; al intentar cachear una instancia construida con esa version vieja,
# pickle falla con "PicklingError: ... it's not the same object as
# __main__.FetchResult". Por eso, las funciones *_raw de abajo (las que SI
# estan cacheadas) devuelven solo dict/pd.Series/str, tipos que no sufren
# este problema; y fetch_fred()/fetch_yf()/fetch_stooq() (sin cache propio)
# arman el FetchResult fresco en cada llamada a partir de ese payload ya
# cacheado, lo cual es practicamente gratis (no repite la descarga).
# ------------------------------------------------------------------------------

def _fetch_fred_via_api_raw(series_id: str, expected_gap_days: int) -> Optional[dict]:
    """Intenta la descarga via fredapi (requiere API key valida). Devuelve
    None si fredapi no esta disponible o la consulta falla, para que el
    llamador pueda caer al metodo de CSV publico."""
    client = _get_fred_client()
    if client is None:
        return None
    try:
        raw = client.get_series(series_id)
        series = pd.Series(raw).dropna()
        series.index = pd.to_datetime(series.index)
        if series.empty:
            return None
        asof = series.index.max()
        status = _infer_staleness(asof, expected_gap_days)
        return {"series": series, "source": f"FRED (fredapi):{series_id}", "status": status,
                "asof": asof, "error": None}
    except Exception as exc:  # noqa: BLE001
        logger.info("fredapi fallo para %s (se usara el fallback de CSV publico): %s", series_id, exc)
        return None


def _fetch_fred_via_csv_raw(series_id: str, expected_gap_days: int) -> dict:
    """Descarga una serie de FRED via su endpoint publico de CSV (sin API key).
    Es el metodo original de esta app y sirve de respaldo cuando fredapi no
    esta disponible o la API key no es valida."""
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
    try:
        resp = requests.get(url, headers=HTTP_HEADERS, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        if df.shape[1] < 2:
            raise ValueError("Respuesta de FRED sin columnas esperadas")
        date_col, val_col = df.columns[0], df.columns[1]
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        # FRED usa "." como marcador de dato faltante; errors="coerce" lo convierte en NaN.
        df[val_col] = pd.to_numeric(df[val_col], errors="coerce")
        df = df.dropna(subset=[date_col]).set_index(date_col)[val_col].dropna()
        if df.empty:
            return {"series": None, "source": "FRED (CSV)", "status": "empty", "asof": None,
                    "error": "Sin observaciones"}
        asof = df.index.max()
        status = _infer_staleness(asof, expected_gap_days)
        return {"series": df, "source": f"FRED (CSV):{series_id}", "status": status, "asof": asof, "error": None}
    except Exception as exc:  # noqa: BLE001
        logger.warning("FRED (CSV) fetch failed for %s: %s", series_id, exc)
        return {"series": None, "source": "FRED (CSV)", "status": "error", "asof": None, "error": str(exc)}


@st.cache_data(ttl=CACHE_TTL_MACRO, show_spinner=False)
def _fetch_fred_cached(series_id: str, expected_gap_days: int) -> dict:
    """Logica de descarga de FRED (fredapi + fallback CSV), cacheada. Devuelve
    un dict de tipos primitivos/pandas -nunca FetchResult- (ver nota de diseno
    arriba)."""
    via_api = _fetch_fred_via_api_raw(series_id, expected_gap_days)
    if via_api is not None:
        return via_api
    return _fetch_fred_via_csv_raw(series_id, expected_gap_days)


def fetch_fred(series_id: str, label: str = "", expected_gap_days: int = 60) -> FetchResult:
    """Descarga una serie de FRED. Intenta primero fredapi (requiere
    FRED_API_KEY valida, ver seccion 1 CONFIG); si no esta disponible o
    falla, cae automaticamente al endpoint publico de CSV (sin API key).
    Esta funcion NO esta cacheada directamente (ver nota de diseno arriba);
    delega la parte costosa (la descarga) a _fetch_fred_cached, que si lo
    esta, y solo arma el objeto FetchResult con el resultado."""
    label = label or series_id
    if not series_id:
        return FetchResult(label=label, status="error", error="Serie no configurada", source="FRED")
    payload = _fetch_fred_cached(series_id, expected_gap_days)
    return FetchResult(label=label, series=payload.get("series"), source=payload.get("source", "FRED"),
                        status=payload.get("status", "error"), asof=payload.get("asof"),
                        error=payload.get("error"), frequency_days=expected_gap_days)


def fetch_fred_with_candidates(candidates: Sequence[str], label: str,
                                expected_gap_days: int = 60) -> FetchResult:
    """Igual que fetch_with_candidates pero para series de FRED: prueba una
    lista de IDs candidatos en orden (util cuando una fuente publica el mismo
    dato bajo distintas convenciones de ID, ej. World Bank vs. IMF WEO) y
    devuelve el primero que funcione."""
    last_result = FetchResult(label=label, status="error", error="Sin candidatos configurados", source="FRED")
    for series_id in candidates:
        res = fetch_fred(series_id, label=label, expected_gap_days=expected_gap_days)
        if res.ok:
            return res
        last_result = res
    return last_result


def _fetch_yf_raw(ticker: str, period: str, interval: str, expected_gap_days: int) -> dict:
    if yf is None:
        return {"series": None, "source": "Yahoo Finance", "status": "error",
                "asof": None, "error": "yfinance no instalado", "resolved_ticker": None}
    if not ticker:
        return {"series": None, "source": "Yahoo Finance", "status": "error",
                "asof": None, "error": "Ticker no configurado", "resolved_ticker": None}
    try:
        df = yf.download(ticker, period=period, interval=interval, progress=False, auto_adjust=True)
        if df is None or df.empty:
            return {"series": None, "source": f"Yahoo Finance:{ticker}", "status": "empty",
                    "asof": None, "error": "Sin datos en Yahoo Finance", "resolved_ticker": None}
        close_col = "Close"
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [c[0] for c in df.columns]
        series = df[close_col].dropna()
        series.index = pd.to_datetime(series.index)
        if series.empty:
            return {"series": None, "source": f"Yahoo Finance:{ticker}", "status": "empty",
                    "asof": None, "error": "Serie vacia tras limpieza", "resolved_ticker": None}
        asof = series.index.max()
        status = _infer_staleness(asof, expected_gap_days)
        return {"series": series, "source": f"Yahoo Finance:{ticker}", "status": status,
                "asof": asof, "error": None, "resolved_ticker": ticker}
    except Exception as exc:  # noqa: BLE001
        logger.warning("Yahoo Finance fetch failed for %s: %s", ticker, exc)
        return {"series": None, "source": f"Yahoo Finance:{ticker}", "status": "error",
                "asof": None, "error": str(exc), "resolved_ticker": None}


@st.cache_data(ttl=CACHE_TTL_INTRADAY, show_spinner=False)
def _fetch_yf_cached(ticker: str, period: str, interval: str, expected_gap_days: int) -> dict:
    """Descarga de Yahoo Finance, cacheada. Devuelve un dict (ver nota de
    diseno arriba), nunca FetchResult."""
    return _fetch_yf_raw(ticker, period, interval, expected_gap_days)


def fetch_yf(ticker: str, label: str = "", period: str = "5y",
             interval: str = "1d", expected_gap_days: int = 7) -> FetchResult:
    """Descarga precios de Yahoo Finance via yfinance, con manejo de errores.
    No esta cacheada directamente (ver nota de diseno arriba)."""
    label = label or ticker
    payload = _fetch_yf_cached(ticker, period, interval, expected_gap_days)
    return FetchResult(label=label, series=payload.get("series"), source=payload.get("source", "Yahoo Finance"),
                        status=payload.get("status", "error"), asof=payload.get("asof"),
                        error=payload.get("error"), frequency_days=expected_gap_days,
                        resolved_ticker=payload.get("resolved_ticker"))


def _fetch_stooq_raw(symbol: str, expected_gap_days: int) -> dict:
    url = f"https://stooq.com/q/d/l/?s={symbol}&i=d"
    try:
        resp = requests.get(url, headers=HTTP_HEADERS, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        if "Close" not in df.columns or "Date" not in df.columns or df.empty:
            return {"series": None, "source": f"Stooq:{symbol}", "status": "empty",
                    "asof": None, "error": "Sin datos en Stooq"}
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        series = df.dropna(subset=["Date"]).set_index("Date")["Close"].dropna()
        if series.empty:
            return {"series": None, "source": f"Stooq:{symbol}", "status": "empty",
                    "asof": None, "error": "Serie vacia tras limpieza"}
        asof = series.index.max()
        status = _infer_staleness(asof, expected_gap_days)
        return {"series": series, "source": f"Stooq:{symbol}", "status": status, "asof": asof, "error": None}
    except Exception as exc:  # noqa: BLE001
        logger.warning("Stooq fetch failed for %s: %s", symbol, exc)
        return {"series": None, "source": f"Stooq:{symbol}", "status": "error", "asof": None, "error": str(exc)}


@st.cache_data(ttl=CACHE_TTL_INTRADAY, show_spinner=False)
def _fetch_stooq_cached(symbol: str, expected_gap_days: int) -> dict:
    """Respaldo via Stooq, cacheado. Devuelve un dict (ver nota de diseno
    arriba), nunca FetchResult."""
    return _fetch_stooq_raw(symbol, expected_gap_days)


def fetch_stooq(symbol: str, label: str = "", expected_gap_days: int = 7) -> FetchResult:
    """Respaldo via Stooq (CSV publico) para simbolos sin cobertura en Yahoo
    Finance. No esta cacheada directamente (ver nota de diseno arriba)."""
    label = label or symbol
    payload = _fetch_stooq_cached(symbol, expected_gap_days)
    return FetchResult(label=label, series=payload.get("series"), source=payload.get("source", "Stooq"),
                        status=payload.get("status", "error"), asof=payload.get("asof"),
                        error=payload.get("error"), frequency_days=expected_gap_days)


def fetch_with_candidates(candidates: Sequence[str], label: str, period: str = "5y",
                           expected_gap_days: int = 7) -> FetchResult:
    """Intenta Yahoo Finance con una lista de tickers candidatos en orden; si todos
    fallan, intenta Stooq con el primer candidato como ultimo recurso."""
    last_result = FetchResult(label=label, status="error", error="Sin candidatos configurados")
    for cand in candidates:
        res = fetch_yf(cand, label=label, period=period, expected_gap_days=expected_gap_days)
        if res.ok:
            return res
        last_result = res
    if candidates:
        stooq_symbol = candidates[0].lower().replace("=x", "").replace("^", "")
        res = fetch_stooq(stooq_symbol, label=label, expected_gap_days=expected_gap_days)
        if res.ok:
            return res
    return last_result


# ------------------------------------------------------------------------------
# 2.5 EXCEL CONSOLIDADO (BD_Consolidada_Dash_Macro.xlsx): UNICO INPUT local
#     para equity peruano (BVL + ADRs), benchmarks regionales REALES, el
#     universo FFMM completo (precio y retorno YA calculados), el universo
#     del ciclo de valor de IA (Precios_ZODA_CicloIA) y los "fundamentals"
#     de mineras/bancos (produccion fisica, indicadores SBS). Reemplaza a
#     Yahoo Finance para todo lo anterior: no se hace ninguna llamada de red
#     para estos datos, solo lectura local del Excel (subido o del repo).
#
#     Diseno (misma disciplina de cache que la seccion 2 de arriba): la
#     UNICA funcion cacheada con @st.cache_data es _read_consolidated_sheet,
#     que devuelve un pd.DataFrame -tipo de pandas, no definido en este
#     script, por lo que NO sufre el bug de pickling de clases en __main__
#     (ver nota de diseno de la seccion 2)-. Todas las demas funciones de
#     esta seccion son wrappers SIN cache propio que arman dict/pd.Series a
#     partir de ese DataFrame ya cacheado; como no estan decoradas con
#     @st.cache_data, pueden devolver sin problema dict de FetchResult o
#     cualquier otra estructura, porque nunca cruzan la frontera de pickle
#     de Streamlit.
# ------------------------------------------------------------------------------

SHEET_PRECIOS_ZODA = "Precios_ZODA"
SHEET_PRECIOS_CICLO_IA = "Precios_ZODA_CicloIA"
SHEET_PRODUCCION_MINERIA = "Produccion_Mineria"
SHEET_SBS_BANCOS = "Indicadores_SBS_Bancos"
SHEET_FFMM_PRECIOS = "FFMM_Precios_Retornos"
SHEET_PARAMETRIA_TICKERS = "Parametria_Tickers"
SHEET_PARAMETRIA_FFMM = "Parametria_FFMM"


@st.cache_data(ttl=CACHE_TTL_MACRO, show_spinner=False)
def _read_consolidated_sheet(file_source, sheet_name: str) -> pd.DataFrame:
    """Lee una hoja especifica del Excel consolidado. file_source puede ser
    un UploadedFile de Streamlit o un Path/str (archivo ya detectado en el
    repo). Devuelve un DataFrame VACIO (nunca None) si la hoja no existe o
    la lectura falla, para simplificar el manejo de errores en los
    llamadores (se evalua con `.empty`, igual que el resto de la app evalua
    `FetchResult.ok`)."""
    if file_source is None:
        return pd.DataFrame()
    try:
        df = pd.read_excel(file_source, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as exc:  # noqa: BLE001
        logger.warning("No se pudo leer la hoja '%s' del Excel consolidado: %s", sheet_name, exc)
        return pd.DataFrame()


# Hojas que la app espera encontrar en el Excel consolidado (nombre EXACTO,
# sensible a mayusculas/espacios) -usado solo por _consolidated_file_diagnostics
# para explicarle al usuario, hoja por hoja, por que algo no se pudo leer.
_EXPECTED_SHEETS = [SHEET_PRECIOS_ZODA, SHEET_FFMM_PRECIOS, SHEET_PARAMETRIA_FFMM, SHEET_PRECIOS_CICLO_IA,
                     SHEET_PRODUCCION_MINERIA, SHEET_SBS_BANCOS]


def _consolidated_file_diagnostics(file_source, loaded_rows: dict) -> str:
    """Diagnostico legible, hoja por hoja, de por que el Excel consolidado no
    trajo los datos esperados. No se llama en el camino feliz (cuando todo
    carga bien): solo cuando el usuario necesita entender por que algo salio
    vacio, para no dejarlo adivinando. Compara los nombres de hoja ESPERADOS
    (_EXPECTED_SHEETS) contra los REALES del workbook (si se puede abrir) y
    cuantas filas trajo cada loader (`loaded_rows`: {nombre_hoja: n_filas}).
    Tambien detecta el caso mas comun en despliegues a GitHub: un archivo
    grande subido via Git LFS cuyo "pointer" de texto (no el binario real)
    es lo que termina en el repo si el host no resuelve LFS -sintoma tipico:
    'se detecto el archivo' (existe la ruta) pero ninguna hoja se puede leer."""
    if file_source is None:
        return "No hay ningun archivo resuelto (ni en el repo ni subido)."
    try:
        if hasattr(file_source, "seek"):
            file_source.seek(0)
            head = file_source.read(200)
            file_source.seek(0)
        else:
            path = Path(file_source)
            with open(path, "rb") as fh:
                head = fh.read(200)
        if isinstance(head, bytes) and head.startswith(b"version https://git-lfs"):
            return ("El archivo detectado es un PUNTERO de Git LFS (texto plano), no el Excel real: "
                     "probablemente se subio a GitHub via Git LFS y el host (ej. Streamlit Community "
                     "Cloud) no lo resuelve automaticamente. Solucion: revisa que tu .gitattributes NO "
                     "tenga una regla '*.xlsx filter=lfs' y vuelve a subir el .xlsx binario real sin LFS.")
    except Exception:
        pass  # si ni siquiera se pudo leer la cabecera, el intento de abrir como Excel (abajo) ya lo explica

    try:
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        real_sheets = list(pd.ExcelFile(file_source).sheet_names)
    except Exception as exc:  # noqa: BLE001
        return f"El archivo se detecto pero no se pudo abrir como un workbook de Excel valido: {exc}"

    lines = [f"Hojas encontradas en el archivo: {real_sheets}"]
    for expected in _EXPECTED_SHEETS:
        n = loaded_rows.get(expected, 0)
        if n > 0:
            lines.append(f"✅ '{expected}': {n} fila(s) leida(s) correctamente.")
        elif expected in real_sheets:
            lines.append(f"🟡 '{expected}': la hoja EXISTE con ese nombre exacto, pero no trajo filas "
                          "(revisar que las columnas esperadas dentro de esa hoja no hayan cambiado).")
        else:
            close = [s for s in real_sheets if s.strip().lower() == expected.strip().lower()]
            if close:
                lines.append(f"🔴 '{expected}': no existe con ese nombre EXACTO, pero si existe "
                              f"'{close[0]}' en el archivo -revisa mayusculas/minusculas o espacios de "
                              "mas/menos, son sensibles a eso-.")
            else:
                lines.append(f"🔴 '{expected}': no se encontro ninguna hoja con ese nombre (ni similar) "
                              "en el archivo.")
    return "\n\n".join(lines)


def _series_to_fetch_result(label: str, series: Optional[pd.Series], source: str,
                             expected_gap_days: int = 7) -> FetchResult:
    """Envuelve una pd.Series YA resuelta (leida del Excel consolidado, sin
    I/O adicional) en un FetchResult, para reusar freshness_badge/
    source_caption/plot_line/interpret_market_series etc. sin duplicar esa
    logica. No hace red ni usa cache propio (la lectura cacheada ya ocurrio
    en _read_consolidated_sheet)."""
    if series is None or series.empty:
        return FetchResult(label=label, status="empty",
                            error="No disponible en el Excel consolidado", source=source)
    asof = series.index.max()
    status = _infer_staleness(asof, expected_gap_days)
    return FetchResult(label=label, series=series, source=source, status=status,
                        asof=asof, frequency_days=expected_gap_days)


def load_precios_zoda(file_source) -> dict:
    """{Ticker BBG: pd.Series de precio} desde la hoja Precios_ZODA: equity
    local BVL + ADRs peruanos + los 6 benchmarks regionales (indices REALES,
    ya no proxies de ETF). Devuelve {} si la hoja no esta disponible."""
    df = _read_consolidated_sheet(file_source, SHEET_PRECIOS_ZODA)
    if df.empty or not {"Fecha", "Ticker", "Precio"}.issubset(df.columns):
        return {}
    out = {}
    for ticker, g in df.groupby("Ticker"):
        s = g.set_index("Fecha")["Precio"].sort_index().dropna()
        if not s.empty:
            out[ticker] = s
    return out


def load_ffmm_precios(file_source) -> dict:
    """{ISIN: pd.Series de precio} desde la hoja FFMM_Precios_Retornos (precio
    y retorno diario YA calculados en el pipeline de origen; aqui solo se
    usa Precio, porque toda la capa de analytics -period_return, fund_metrics,
    etc.- ya parte de precios, no de retornos)."""
    df = _read_consolidated_sheet(file_source, SHEET_FFMM_PRECIOS)
    if df.empty or not {"Fecha", "ISIN", "Precio"}.issubset(df.columns):
        return {}
    out = {}
    for isin, g in df.groupby("ISIN"):
        s = g.set_index("Fecha")["Precio"].sort_index().dropna()
        if not s.empty:
            out[isin] = s
    return out


def load_ciclo_ia(file_source) -> dict:
    """{Ticker: {"nombre", "capa", "series"}} desde Precios_ZODA_CicloIA: el
    universo de 46 empresas de la cadena de valor de semiconductores/IA,
    usado como proxy de holdings para el performance attribution de los
    fondos tecnologicos de FFMM, y para el chart didactico de las 12 capas."""
    df = _read_consolidated_sheet(file_source, SHEET_PRECIOS_CICLO_IA)
    if df.empty or not {"Fecha", "Ticker", "NombreFormal", "CapaIA", "Precio"}.issubset(df.columns):
        return {}
    out = {}
    for ticker, g in df.groupby("Ticker"):
        g = g.sort_values("Fecha")
        s = g.set_index("Fecha")["Precio"].dropna()
        if s.empty:
            continue
        out[ticker] = {"nombre": g["NombreFormal"].iloc[0], "capa": g["CapaIA"].iloc[0], "series": s}
    return out


def load_produccion_mineria(file_source) -> pd.DataFrame:
    """DataFrame crudo de la hoja Produccion_Mineria (produccion mensual por
    Ticker/Empresa/Mineral/Etapa/Proceso/Unidad). Ver mining_production_series()
    para la agregacion por empresa+mineral principal usada en el tab PERU."""
    return _read_consolidated_sheet(file_source, SHEET_PRODUCCION_MINERIA)


def load_sbs_bancos(file_source) -> pd.DataFrame:
    """DataFrame crudo de la hoja Indicadores_SBS_Bancos (indicadores SBS
    mensuales por banco/Sistema Bancario). Ver sbs_indicator_series()."""
    return _read_consolidated_sheet(file_source, SHEET_SBS_BANCOS)


def load_parametria_tickers(file_source) -> pd.DataFrame:
    return _read_consolidated_sheet(file_source, SHEET_PARAMETRIA_TICKERS)


def load_parametria_ffmm(file_source) -> pd.DataFrame:
    return _read_consolidated_sheet(file_source, SHEET_PARAMETRIA_FFMM)


def mining_production_series(produccion_df: pd.DataFrame, ticker: str) -> Optional[pd.Series]:
    """Serie mensual agregada de produccion fisica para el mineral PRINCIPAL
    de `ticker` (ver MINING_PRIMARY_METAL). Se suma el Valor de TODAS las
    plantas/unidades de la empresa para ese mineral, restringiendo a
    Etapa=='CONCENTRACION' y Proceso=='FLOTACION' -la salida fisica
    principal y comparable entre plantas de una misma empresa-, para evitar
    mezclar etapas de proceso (lixiviacion/refinacion) con unidades de
    medida distintas en una misma serie. Devuelve None si no hay mineral
    principal mapeado o no hay datos."""
    if produccion_df is None or produccion_df.empty:
        return None
    mineral = MINING_PRIMARY_METAL.get(ticker)
    if mineral is None:
        return None
    mask = (
        (produccion_df["Ticker"] == ticker)
        & (produccion_df["Mineral"].astype(str).str.upper() == mineral)
        & (produccion_df["Etapa"].astype(str).str.upper() == "CONCENTRACIÓN")
        & (produccion_df["Proceso"].astype(str).str.upper() == "FLOTACIÓN")
    )
    sub = produccion_df.loc[mask]
    if sub.empty:
        return None
    s = sub.groupby("Fecha")["Valor"].sum().sort_index()
    s = s[s > 0]
    return s if not s.empty else None


def sbs_indicator_series(sbs_df: pd.DataFrame, entidad: str, indicator_suffix: str) -> Optional[pd.Series]:
    """Serie mensual de un indicador SBS para `entidad` (BCP/INTERBANK/BBVA/
    SISTEMA BANCARIO), matcheando por el SUFIJO del Codigo (ej. 'SBS_REN_ROE')
    para no depender del prefijo especifico de cada banco (PERBBV_/PERBCP_/...).
    La columna 'Valor' del Excel viene como fraccion (ej. 0.0264 = 2.64% de
    morosidad); se multiplica por 100 aqui -una sola vez, en el punto de
    entrada- para que todo consumidor (tabla, chart, storytelling) trabaje
    en puntos porcentuales directamente, igual que el resto de la app."""
    if sbs_df is None or sbs_df.empty:
        return None
    mask = (sbs_df["Entidad"] == entidad) & (sbs_df["Codigo"].astype(str).str.endswith(indicator_suffix))
    sub = sbs_df.loc[mask]
    if sub.empty:
        return None
    s = (sub.set_index("Fecha")["Valor"] * 100.0).sort_index().dropna()
    return s if not s.empty else None


# ==============================================================================
# 3. ANALYTICS LAYER
# ==============================================================================

def pct_change_over(series: pd.Series, periods: int) -> Optional[float]:
    if series is None or len(series) <= periods:
        return None
    base = series.iloc[-1 - periods]
    last = series.iloc[-1]
    if base == 0 or pd.isna(base) or pd.isna(last):
        return None
    return (last / base - 1.0) * 100.0


def level_change_over(series: pd.Series, periods: int) -> Optional[float]:
    if series is None or len(series) <= periods:
        return None
    base = series.iloc[-1 - periods]
    last = series.iloc[-1]
    if pd.isna(base) or pd.isna(last):
        return None
    return last - base


def rolling_zscore(series: pd.Series, window: int = 252) -> Optional[float]:
    if series is None or len(series) < max(20, window // 5):
        return None
    w = min(window, len(series))
    recent = series.iloc[-w:]
    mu, sigma = recent.mean(), recent.std()
    if sigma == 0 or pd.isna(sigma):
        return None
    return float((series.iloc[-1] - mu) / sigma)


def percentile_rank(series: pd.Series, window: Optional[int] = None) -> Optional[float]:
    if series is None or series.empty:
        return None
    hist = series if window is None else series.iloc[-window:]
    if len(hist) < 10:
        return None
    return float((hist <= hist.iloc[-1]).mean() * 100.0)


def annualized_vol(series: pd.Series, window: int = 63, trading_days: int = 252) -> Optional[float]:
    if series is None or len(series) < window + 1:
        return None
    rets = series.pct_change().dropna().iloc[-window:]
    if rets.empty:
        return None
    return float(rets.std() * np.sqrt(trading_days) * 100.0)


def max_drawdown(series: pd.Series, window: Optional[int] = None) -> Optional[float]:
    if series is None or series.empty:
        return None
    hist = series if window is None else series.iloc[-window:]
    running_max = hist.cummax()
    dd = (hist / running_max - 1.0) * 100.0
    return float(dd.min())


def trend_acceleration(series: pd.Series, short: int = 3, long_: int = 6) -> str:
    """Compara el cambio reciente (short) vs el cambio previo (long_-short) para
    detectar aceleracion/desaceleracion de una tendencia. Devuelve 'acelerando',
    'desacelerando' o 'estable'."""
    if series is None or len(series) < long_ + 1:
        return "sin_datos"
    recent_chg = series.iloc[-1] - series.iloc[-1 - short]
    prior_chg = series.iloc[-1 - short] - series.iloc[-1 - long_]
    if pd.isna(recent_chg) or pd.isna(prior_chg):
        return "sin_datos"
    if abs(prior_chg) < 1e-9:
        return "estable"
    same_sign = np.sign(recent_chg) == np.sign(prior_chg)
    if same_sign and abs(recent_chg) > abs(prior_chg) * 1.1:
        return "acelerando"
    if not same_sign or abs(recent_chg) < abs(prior_chg) * 0.6:
        return "desacelerando"
    return "estable"


def correlation_matrix(series_dict: dict) -> pd.DataFrame:
    df = pd.DataFrame({k: v for k, v in series_dict.items() if v is not None})
    df = df.dropna(how="all")
    if df.shape[1] < 2:
        return pd.DataFrame()
    return df.pct_change().dropna().corr()


def resample_monthly(series: pd.Series) -> pd.Series:
    if series is None or series.empty:
        return series
    return series.resample("MS").last().dropna()


def ai_cycle_layer_index(ciclo_ia: dict, capa: str) -> Optional[pd.Series]:
    """Indice equal-weight (base 100 en la primera fecha de cada empresa) de
    todas las companias de `capa` dentro del universo Precios_ZODA_CicloIA.
    Cada empresa se normaliza INDIVIDUALMENTE a base 100 en su propia
    primera fecha disponible (asi una incorporacion tardia de un ticker no
    distorsiona el nivel del indice) y luego se promedia equal-weight por
    fecha entre las empresas con dato ese dia. Usado para el chart didactico
    de 'como van las 12 capas del ciclo de IA' en la pestana FFMM."""
    series_list = [v["series"].dropna() for v in ciclo_ia.values() if v["capa"] == capa]
    series_list = [s for s in series_list if not s.empty]
    if not series_list:
        return None
    normalized = [(s / s.iloc[0]) * 100.0 for s in series_list]
    combined = pd.concat(normalized, axis=1)
    idx = combined.mean(axis=1, skipna=True).dropna()
    return idx if not idx.empty else None


# ------------------------------------------------------------------------------
# 3.1 Retornos por periodo calendario (WTD/MTD/YTD/1Y/3Y) y CAGR, usados
#     intensivamente en la pestana FFMM pero reutilizables en toda la app.
# ------------------------------------------------------------------------------

def period_return(series: pd.Series, anchor: str) -> Optional[float]:
    """Retorno (%) desde el cierre de referencia de 'anchor' hasta el ultimo
    dato disponible. anchor in {'WTD','MTD','YTD','1Y','3Y'}.
    - WTD/MTD/YTD: toman como base el ULTIMO cierre ESTRICTAMENTE ANTERIOR al
      inicio de la semana/mes/anio en curso (convencion estandar: 'cambio
      desde el cierre del periodo anterior'), evaluado sobre la ultima fecha
      disponible en la serie (no sobre la fecha de hoy), para que funcione
      igual de bien con datos de cierre de ayer que con datos de hace dias.
    - 1Y/3Y: toman como base el ultimo cierre disponible en o antes de
      (ultima_fecha - 1 o 3 anios).
    Devuelve None si no hay suficiente historia para construir la base.
    """
    if series is None or series.empty:
        return None
    series = series.dropna()
    if series.empty:
        return None
    last_date = series.index.max()
    last_val = series.iloc[-1]

    if anchor == "WTD":
        period_start = (last_date - pd.Timedelta(days=int(last_date.weekday()))).normalize()
        base_pool = series[series.index < period_start]
    elif anchor == "MTD":
        period_start = last_date.replace(day=1).normalize()
        base_pool = series[series.index < period_start]
    elif anchor == "YTD":
        period_start = last_date.replace(month=1, day=1).normalize()
        base_pool = series[series.index < period_start]
    elif anchor == "1Y":
        cutoff = last_date - pd.DateOffset(years=1)
        base_pool = series[series.index <= cutoff]
    elif anchor == "3Y":
        cutoff = last_date - pd.DateOffset(years=3)
        base_pool = series[series.index <= cutoff]
    else:
        raise ValueError(f"anchor desconocido: {anchor}")

    if base_pool.empty:
        return None
    base_val = base_pool.iloc[-1]
    if base_val == 0 or pd.isna(base_val) or pd.isna(last_val):
        return None
    return float((last_val / base_val - 1.0) * 100.0)


def cagr_from_cumulative(cum_return_pct: Optional[float], years: float) -> Optional[float]:
    """Convierte un retorno acumulado (%) en una tasa anualizada compuesta (%)."""
    if cum_return_pct is None or pd.isna(cum_return_pct) or years <= 0:
        return None
    growth = 1.0 + cum_return_pct / 100.0
    if growth <= 0:
        return None
    return float((growth ** (1.0 / years) - 1.0) * 100.0)


def fund_metrics(series: pd.Series) -> dict:
    """Bloque estandar de metricas de performance para un fondo/ETF/accion:
    ultimo precio, ultimo retorno diario y retornos WTD/MTD/YTD/1Y/3Y
    (acumulado y anualizado), mas volatilidad anualizada 1Y, usado tanto en
    la tabla resumen como en la ficha de detalle de la pestana FFMM."""
    if series is None or series.empty:
        return {}
    series = series.dropna()
    last_price = float(series.iloc[-1])
    daily_ret = None
    if len(series) >= 2:
        prev = series.iloc[-2]
        if prev != 0 and not pd.isna(prev):
            daily_ret = float((series.iloc[-1] / prev - 1.0) * 100.0)
    r3y_cum = period_return(series, "3Y")
    return {
        "ultimo_precio": last_price,
        "ultimo_retorno_%": daily_ret,
        "wtd_%": period_return(series, "WTD"),
        "mtd_%": period_return(series, "MTD"),
        "ytd_%": period_return(series, "YTD"),
        "r1y_%": period_return(series, "1Y"),
        "r3y_acum_%": r3y_cum,
        "r3y_cagr_%": cagr_from_cumulative(r3y_cum, 3.0),
        "vol_1y_%": annualized_vol(series, window=min(len(series) - 1, 252)) if len(series) > 20 else None,
        "fecha_ultimo_dato": series.index.max(),
    }


def dispersion_stats(returns: Sequence[float]) -> dict:
    """Dispersion simple (max-min, desv. estandar) de un conjunto de retornos
    de pares/categoria: util para destacar que tan homogeneo o heterogeneo
    se movio un grupo tematico en la pestana FFMM."""
    vals = [v for v in returns if v is not None and not pd.isna(v)]
    if len(vals) < 2:
        return {}
    arr = np.array(vals, dtype=float)
    return {"max": float(arr.max()), "min": float(arr.min()), "spread": float(arr.max() - arr.min()),
            "std": float(arr.std()), "mean": float(arr.mean()), "n": len(arr)}


# ==============================================================================
# 4. STORYTELLING ENGINE  (biblioteca de patrones interpretativos reutilizables)
# ==============================================================================

def _direction_word(delta: Optional[float], flat_eps: float = 0.05) -> str:
    if delta is None or pd.isna(delta):
        return "sin variacion clara"
    if delta > flat_eps:
        return "al alza"
    if delta < -flat_eps:
        return "a la baja"
    return "lateral"


def _magnitude_word(abs_pct: Optional[float]) -> str:
    if abs_pct is None or pd.isna(abs_pct):
        return ""
    a = abs(abs_pct)
    if a < 1:
        return "marginal"
    if a < 3:
        return "moderado"
    if a < 8:
        return "significativo"
    return "muy marcado"


def _percentile_phrase(pct: Optional[float]) -> str:
    if pct is None or pd.isna(pct):
        return ""
    if pct >= 90:
        return "en zona de maximos historicos recientes (percentil >90)"
    if pct >= 70:
        return "en la parte alta de su rango reciente"
    if pct <= 10:
        return "en zona de minimos historicos recientes (percentil <10)"
    if pct <= 30:
        return "en la parte baja de su rango reciente"
    return "cerca de niveles medios de su rango reciente"


def _trend_phrase(trend: str) -> str:
    return {
        "acelerando": "el movimiento se esta acelerando",
        "desacelerando": "el movimiento muestra senales de desaceleracion",
        "estable": "el ritmo de cambio se mantiene relativamente estable",
        "sin_datos": "",
    }.get(trend, "")


def interpret_market_series(label: str, series: pd.Series, asset_class: str,
                             extra_note: str = "") -> str:
    """Interprete generico para un instrumento de mercado (commodity, equity,
    FX, yield) a partir de direccion, magnitud, percentil historico y
    aceleracion de tendencia. Satisface el patron What/Why/Investment implication
    de forma natural, sin forzar las tres preguntas en cada texto."""
    if series is None or series.empty or len(series) < 10:
        return f"**{label}**: datos insuficientes para generar una interpretacion confiable."

    chg_1m = pct_change_over(series, 21)
    chg_3m = pct_change_over(series, 63)
    chg_12m = pct_change_over(series, 252)
    z = rolling_zscore(series, window=min(len(series), 252))
    pct_rank = percentile_rank(series, window=min(len(series), 252 * 2))
    trend = trend_acceleration(series)
    dd = max_drawdown(series, window=min(len(series), 252))
    vol = annualized_vol(series)

    ref_chg = chg_3m if chg_3m is not None else chg_1m
    direction = _direction_word(ref_chg)
    magnitude = _magnitude_word(ref_chg)
    pct_phrase = _percentile_phrase(pct_rank)
    trend_phrase = _trend_phrase(trend)

    cls_label = ASSET_CLASS_LABELS.get(asset_class, "activo")
    parts = []

    headline = f"**{label}** se mueve {direction}"
    if magnitude:
        headline += f" de forma {magnitude}"
    if chg_3m is not None:
        headline += f" en los ultimos 3 meses ({chg_3m:+.1f}%)"
    headline += "."
    parts.append(headline)

    if chg_12m is not None and chg_3m is not None and np.sign(chg_12m) != np.sign(chg_3m) and abs(chg_12m) > 2:
        parts.append(f"Esto contrasta con la variacion de los ultimos 12 meses ({chg_12m:+.1f}%), lo que sugiere un "
                     f"posible cambio de regimen en curso mas que una simple continuacion de la tendencia anual.")
    elif chg_12m is not None:
        parts.append(f"En el acumulado de 12 meses la variacion es de {chg_12m:+.1f}%.")

    if pct_phrase:
        s = f"El nivel actual se ubica {pct_phrase}"
        if z is not None and abs(z) >= 1.5:
            s += f", con una desviacion de {z:+.1f} veces su volatilidad reciente respecto al promedio"
        s += "."
        parts.append(s)

    if trend_phrase and trend != "sin_datos":
        parts.append(f"En el margen, {trend_phrase}, lo que sugiere vigilar si el patron se mantiene en las proximas lecturas antes de extrapolarlo.")

    if vol is not None and asset_class in ("equity", "fx", "energy", "precious_metal", "industrial_metal"):
        if vol > 30:
            parts.append(f"La volatilidad anualizada reciente ({vol:.0f}%) es elevada, lo que exige posiciones mas conservadoras en terminos de tamano de riesgo.")
        elif vol < 10:
            parts.append(f"La volatilidad anualizada reciente ({vol:.0f}%) es baja para un {cls_label}, lo que puede estar comprimiendo las primas de riesgo asociadas.")

    if dd is not None and dd < -15:
        parts.append(f"El drawdown desde el maximo de los ultimos 12 meses es de {dd:.1f}%, una correccion relevante que condiciona el punto de entrada actual.")

    if extra_note:
        parts.append(extra_note)

    return " ".join(parts)


def interpret_labor_market(unrate: pd.Series, payrolls: pd.Series,
                            claims: pd.Series) -> str:
    """Interprete compuesto siguiendo el ejemplo del punto 3.C: combina 3
    senales (tasa de desempleo, creacion de empleo, solicitudes de desempleo)
    en lugar de leerlas de forma aislada, y evita concluir deterioro cuando
    las senales son contradictorias."""
    signals = []

    unrate_chg = level_change_over(unrate, 6) if unrate is not None else None   # ~6 meses
    payrolls_trend_now = (payrolls.diff().iloc[-3:].mean() if payrolls is not None and len(payrolls) > 4 else None)
    payrolls_trend_prior = (payrolls.diff().iloc[-9:-3].mean() if payrolls is not None and len(payrolls) > 9 else None)
    claims_chg = pct_change_over(claims, 8) if claims is not None else None     # ~2 meses de datos semanales

    cooling_votes = 0
    resilient_votes = 0

    if unrate_chg is not None:
        if unrate_chg > 0.1:
            cooling_votes += 1
            signals.append(f"la tasa de desempleo subio {unrate_chg:+.1f} pp en los ultimos ~6 meses")
        elif unrate_chg < -0.1:
            resilient_votes += 1
            signals.append(f"la tasa de desempleo bajo {unrate_chg:+.1f} pp en los ultimos ~6 meses")

    if payrolls_trend_now is not None and payrolls_trend_prior is not None:
        if payrolls_trend_now < payrolls_trend_prior * 0.7:
            cooling_votes += 1
            signals.append("el ritmo de creacion de empleo se ha desacelerado frente al trimestre previo")
        elif payrolls_trend_now > payrolls_trend_prior * 1.1:
            resilient_votes += 1
            signals.append("la creacion de empleo se mantiene solida o incluso se acelera")

    if claims_chg is not None:
        if claims_chg > 10:
            cooling_votes += 1
            signals.append(f"las solicitudes de desempleo subieron {claims_chg:+.0f}% en ~2 meses")
        elif claims_chg < -10:
            resilient_votes += 1
            signals.append(f"las solicitudes de desempleo cayeron {claims_chg:+.0f}% en ~2 meses")

    if not signals:
        return "**Mercado laboral (EE.UU.)**: datos insuficientes o sin variaciones relevantes para construir una lectura integrada."

    detail = "; ".join(signals) + "."

    if cooling_votes >= 2 and resilient_votes == 0:
        verdict = ("El conjunto de senales apunta a un enfriamiento consistente del mercado laboral. "
                   "Esto reduce la presion salarial sobre la inflacion y amplia el margen de la Fed para "
                   "flexibilizar su politica monetaria; el riesgo principal es que el enfriamiento se acelere "
                   "mas alla de un ajuste ordenado.")
    elif resilient_votes >= 2 and cooling_votes == 0:
        verdict = ("El conjunto de senales sugiere que el mercado laboral se mantiene resiliente. "
                   "Esto sostiene el consumo privado pero tambien limita el espacio para recortes de tasas "
                   "en el corto plazo si la inflacion no converge con suficiente rapidez.")
    else:
        verdict = ("Las senales son mixtas o contradictorias entre si, por lo que no corresponde concluir de "
                   "forma automatica que el mercado laboral se esta deteriorando o fortaleciendo: conviene "
                   "esperar mas lecturas antes de ajustar convicciones de asset allocation ligadas al ciclo laboral.")

    return f"**Mercado laboral (EE.UU.)**: {detail} {verdict}"


def interpret_inflation(cpi: pd.Series, core_cpi: pd.Series,
                         breakeven_5y: Optional[pd.Series] = None) -> str:
    """Interprete compuesto de inflacion: nivel YoY, tendencia (acelera/desacelera)
    y, si esta disponible, la expectativa de mercado (breakeven) para contrastar
    inflacion observada vs. inflacion esperada."""
    if cpi is None or len(cpi) < 14:
        return "**Inflacion (EE.UU.)**: datos insuficientes para una lectura confiable."

    cpi_m = resample_monthly(cpi)
    yoy = pct_change_over(cpi_m, 12)
    yoy_prior = pct_change_over(cpi_m.iloc[:-1], 12) if len(cpi_m) > 13 else None
    core_m = resample_monthly(core_cpi) if core_cpi is not None else None

    core_yoy = pct_change_over(core_m, 12) if core_m is not None else None

    parts = []
    if yoy is not None:
        txt = f"La inflacion headline se ubica en {yoy:.1f}% interanual"
        if yoy_prior is not None:
            delta = yoy - yoy_prior
            if delta > 0.05:
                txt += f", acelerando frente al mes previo (+{delta:.1f} pp)"
            elif delta < -0.05:
                txt += f", desacelerando frente al mes previo ({delta:.1f} pp)"
            else:
                txt += ", sin cambios relevantes frente al mes previo"
        txt += "."
        parts.append(txt)

    if core_yoy is not None and yoy is not None:
        gap = core_yoy - yoy
        if abs(gap) > 0.3:
            direction = "por encima" if gap > 0 else "por debajo"
            parts.append(f"La inflacion core ({core_yoy:.1f}%) corre {direction} de la headline, lo que senala que "
                         f"componentes volatiles (energia/alimentos) estan {'atenuando' if gap>0 else 'presionando'} "
                         f"la lectura total, un matiz relevante para la trayectoria de politica monetaria.")

    if breakeven_5y is not None and not breakeven_5y.empty and yoy is not None:
        be = breakeven_5y.iloc[-1]
        if not pd.isna(be):
            gap2 = yoy - be
            if gap2 > 0.5:
                parts.append(f"La inflacion observada (header {yoy:.1f}%) supera a la expectativa de mercado a 5 anos "
                             f"({be:.1f}%), lo que sugiere que el mercado sigue asignando alta probabilidad a una "
                             f"normalizacion gradual sin necesidad de un ciclo de endurecimiento adicional.")
            elif gap2 < -0.5:
                parts.append(f"La inflacion observada ({yoy:.1f}%) esta por debajo de la expectativa de mercado a 5 anos "
                             f"({be:.1f}%), una divergencia que vale la pena monitorear para detectar cambios de regimen.")

    if not parts:
        return "**Inflacion (EE.UU.)**: sin senales suficientemente claras para una interpretacion."

    return "**Inflacion (EE.UU.)**: " + " ".join(parts)


def interpret_fed_policy(fed_funds: pd.Series, balance_sheet: pd.Series) -> str:
    if fed_funds is None or fed_funds.empty:
        return "**Politica de la Fed**: datos insuficientes."
    rate = fed_funds.iloc[-1]
    rate_chg_6m = level_change_over(resample_monthly(fed_funds), 6)
    parts = [f"La Fed Funds Rate se ubica en {rate:.2f}%."]
    if rate_chg_6m is not None:
        if rate_chg_6m > 0.1:
            parts.append(f"En los ultimos ~6 meses la tasa subio {rate_chg_6m:+.2f} pp, consistente con un sesgo de politica restrictivo o de pausa alta.")
        elif rate_chg_6m < -0.1:
            parts.append(f"En los ultimos ~6 meses la tasa bajo {rate_chg_6m:+.2f} pp, consistente con un ciclo de flexibilizacion en curso.")
        else:
            parts.append("La tasa se ha mantenido estable en los ultimos meses, en linea con una postura de 'esperar y ver' (data-dependent).")

    if balance_sheet is not None and not balance_sheet.empty and len(balance_sheet) > 12:
        bs_m = resample_monthly(balance_sheet)
        chg_6m = pct_change_over(bs_m, 6)
        if chg_6m is not None:
            if chg_6m < -1.0:
                parts.append(f"El balance de la Fed se ha contraido {chg_6m:.1f}% en ~6 meses (Quantitative Tightening en curso), lo que retira liquidez del sistema de forma gradual.")
            elif chg_6m > 1.0:
                parts.append(f"El balance de la Fed se ha expandido {chg_6m:+.1f}% en ~6 meses, una senal de provision de liquidez que contrasta con el discurso de normalizacion.")

    return "**Politica de la Fed**: " + " ".join(parts)


def interpret_yield_curve(y2: pd.Series, y10: pd.Series, y30: pd.Series) -> str:
    if y2 is None or y10 is None or y2.empty or y10.empty:
        return "**Curva de rendimientos (UST)**: datos insuficientes."
    spread = y10.iloc[-1] - y2.iloc[-1]
    spread_hist = (y10 - y2).dropna()
    spread_3m_ago = spread_hist.iloc[-63] if len(spread_hist) > 63 else None

    parts = [f"El spread 10Y-2Y se ubica en {spread:+.2f} pp."]
    if spread < 0:
        parts.append("La curva permanece invertida, una senal historicamente asociada a riesgo de recesion dentro de los siguientes 12-24 meses, aunque con rezagos variables.")
    else:
        parts.append("La curva se mantiene en pendiente positiva (no invertida).")

    if spread_3m_ago is not None:
        delta = spread - spread_3m_ago
        if delta > 0.1:
            parts.append(f"El empinamiento de {delta:+.2f} pp en 3 meses es consistente con un proceso de 'bull steepening' o 'bear steepening' segun si domino la baja del tramo corto o la suba del tramo largo; conviene revisar el nivel absoluto de cada tramo para distinguir el tipo de empinamiento.")
        elif delta < -0.1:
            parts.append(f"El aplanamiento de {delta:+.2f} pp en 3 meses reduce el margen de carry positivo en estrategias de pendiente.")

    if y30 is not None and not y30.empty:
        term_premium_proxy = y30.iloc[-1] - y10.iloc[-1]
        parts.append(f"El tramo largo (30Y) cotiza a {y30.iloc[-1]:.2f}%, con un spread de {term_premium_proxy:+.2f} pp sobre el 10Y, relevante para fondos con pasivos de muy largo plazo.")

    return "**Curva de rendimientos (UST)**: " + " ".join(parts)


def interpret_fiscal(debt_gdp: pd.Series, deficit: pd.Series) -> str:
    parts = []
    if debt_gdp is not None and not debt_gdp.empty:
        lvl = debt_gdp.iloc[-1]
        chg = level_change_over(resample_monthly(debt_gdp), 4)
        s = f"La deuda publica federal equivale a {lvl:.0f}% del PBI"
        if chg is not None:
            s += f", {'en ascenso' if chg > 0.5 else 'relativamente estable' if abs(chg) <= 0.5 else 'en descenso'} en el ultimo ano"
        s += "."
        parts.append(s)
    if deficit is not None and not deficit.empty:
        last12 = deficit.iloc[-12:].sum() if len(deficit) >= 12 else None
        if last12 is not None:
            parts.append(f"El resultado fiscal acumulado de los ultimos 12 meses es de {last12:,.0f} USD millones "
                         f"({'deficit' if last12 < 0 else 'superavit'}), un factor que sostiene la oferta neta de Treasuries "
                         f"y, en el margen, presiona al alza el tramo largo de la curva vía prima por plazo.")
    if not parts:
        return "**Situacion fiscal (EE.UU.)**: datos insuficientes."
    return "**Situacion fiscal (EE.UU.)**: " + " ".join(parts)


def interpret_stress_index(series: pd.Series, label: str, zero_centered: bool = True) -> str:
    if series is None or series.empty:
        return f"**{label}**: datos insuficientes."
    last = series.iloc[-1]
    z = rolling_zscore(series, window=min(len(series), 260))
    if zero_centered:
        if last > 0.5:
            level_txt = "en niveles de estres financiero claramente por encima del promedio historico"
        elif last < -0.5:
            level_txt = "en niveles que reflejan condiciones financieras holgadas (por debajo del promedio historico)"
        else:
            level_txt = "cerca de su promedio historico, sin senales de tension inusual"
    else:
        level_txt = "en el nivel mostrado"
    txt = f"**{label}** se ubica {level_txt} ({last:.2f})."
    if z is not None and abs(z) >= 1.5:
        txt += f" La desviacion respecto a su media reciente ({z:+.1f}x) sugiere un movimiento fuera de lo habitual que merece seguimiento."
    return txt


def interpret_commodity_with_context(name: str, series: pd.Series, peru_relevance: str) -> str:
    base = interpret_market_series(name, series, COMMODITIES.get(name, {}).get("asset_class", "energy"))
    if peru_relevance in ("alta", "muy alta"):
        extra = (f" Dada la relevancia {peru_relevance} de este commodity para la economia y el mercado de capitales "
                 f"peruano (mineria), este movimiento tiene traslado directo a los terminos de intercambio, la "
                 f"recaudacion fiscal minera y el desempeno relativo de las mineras listadas en la pestana PERU.")
        return base + extra
    return base


def interpret_fx(label: str, series: pd.Series, is_usd_base: bool = False) -> str:
    """is_usd_base=True cuando el par se cotiza como 'moneda local por USD'
    (ej. USD/PEN): una suba de la serie implica DEPRECIACION de la moneda local."""
    base = interpret_market_series(label, series, "fx")
    if series is None or len(series) < 30:
        return base
    chg_3m = pct_change_over(series, 63)
    if chg_3m is None:
        return base
    if is_usd_base:
        verb = "depreciandose" if chg_3m > 0 else "apreciandose"
    else:
        verb = "apreciandose" if chg_3m > 0 else "depreciandose"
    extra = f" En terminos de moneda local, esto implica que la divisa se viene {verb} frente al dolar en el trimestre, con implicancias directas sobre inflacion importada, costo de deuda en moneda extranjera y competitividad externa."
    return base + extra


def interpret_country_scorecard(country: str, cpi_yoy: Optional[float], rate: Optional[float],
                                 y10: Optional[float], gdp: Optional[float],
                                 fx_chg_3m: Optional[float]) -> str:
    parts = []
    if cpi_yoy is not None and rate is not None:
        real_rate = rate - cpi_yoy
        parts.append(f"La tasa de politica monetaria ({rate:.2f}%) frente a una inflacion de {cpi_yoy:.1f}% interanual "
                     f"implica una tasa real aproximada de {real_rate:+.1f} pp, "
                     f"{'claramente restrictiva' if real_rate > 2 else 'restrictiva' if real_rate > 0 else 'acomodaticia'}.")
    elif cpi_yoy is not None:
        parts.append(f"La inflacion interanual se ubica en {cpi_yoy:.1f}%.")

    if y10 is not None and rate is not None:
        slope = y10 - rate
        if slope < -0.3:
            parts.append(f"La curva local esta invertida ({slope:+.2f} pp entre el 10Y y la tasa de politica), una senal de expectativas de recorte de tasas o de preocupacion por el ciclo.")

    if gdp is not None:
        parts.append(f"El crecimiento del PBI real anual mas reciente disponible es de {gdp:.1f}%.")

    if fx_chg_3m is not None and abs(fx_chg_3m) > 3:
        direction = "depreciado" if fx_chg_3m > 0 else "apreciado"
        parts.append(f"La moneda local se ha {direction} {abs(fx_chg_3m):.1f}% frente al USD en el trimestre, un movimiento relevante para la narrativa de bonos y equity local.")

    if not parts:
        return f"**{country}**: datos insuficientes para un scorecard completo con las fuentes gratuitas disponibles."
    return f"**{country}**: " + " ".join(parts)


# ------------------------------------------------------------------------------
# 4.1 FFMM: resumen de performance por periodo + performance attribution por
#     Top Holdings, estilo CFA (ver nota de aproximacion de pesos en la
#     cabecera del archivo).
# ------------------------------------------------------------------------------

def fund_periods_summary(label: str, metrics: dict) -> str:
    """Resumen corto del desempeno de un fondo en todos los horizontes
    pedidos (WTD/MTD/YTD/1Y/3Y), en tono PM, mas el ultimo precio/retorno."""
    if not metrics:
        return f"**{label}**: datos insuficientes para resumir su performance."
    wtd, mtd, ytd = metrics.get("wtd_%"), metrics.get("mtd_%"), metrics.get("ytd_%")
    r1y, r3y_cum, r3y_cagr = metrics.get("r1y_%"), metrics.get("r3y_acum_%"), metrics.get("r3y_cagr_%")
    last_price, last_ret = metrics.get("ultimo_precio"), metrics.get("ultimo_retorno_%")

    def _fmt(x):
        return f"{x:+.1f}%" if x is not None else "N/D"

    txt = f"**{label}** cotiza en {last_price:,.2f}"
    if last_ret is not None:
        txt += f" ({last_ret:+.1f}% en la ultima sesion)"
    txt += f". Retornos: WTD {_fmt(wtd)} · MTD {_fmt(mtd)} · YTD {_fmt(ytd)} · 1Y {_fmt(r1y)} · " \
           f"3Y acumulado {_fmt(r3y_cum)} (CAGR {_fmt(r3y_cagr)})."

    if wtd is not None and r1y is not None:
        if np.sign(wtd) != 0 and np.sign(r1y) != 0 and np.sign(wtd) != np.sign(r1y) and abs(wtd) > 2:
            txt += (" El movimiento de la ultima semana va en sentido contrario a la tendencia de 12 meses, "
                   "lo que amerita revisar si es ruido de corto plazo o el inicio de un cambio de regimen.")
    return txt


def performance_attribution_story(fund_label: str, period_label: str, fund_return: Optional[float],
                                   attrib_df: pd.DataFrame, residual_pp: Optional[float],
                                   weights_are_equal_proxy: bool = False) -> str:
    """Narrativa de performance attribution estilo CFA: explica el retorno de
    un fondo/ETF en un periodo a partir de la contribucion (peso x retorno)
    de sus principales holdings. attrib_df debe tener columnas 'Name',
    'Weight_%', 'Return_%', 'Contribution_pp', ordenado por contribucion.
    weights_are_equal_proxy=True cuando los "holdings" no son los oficiales
    del fondo sino un universo PROXY equal-weight (ej. las empresas del
    ciclo de valor de IA de Precios_ZODA_CicloIA, usado para los fondos
    tecnologicos de FFMM que no tienen holdings oficiales en esta base de
    datos): cambia la nota metodologica final para no sobre-representar la
    precision de la atribucion."""
    if fund_return is None:
        return f"**{fund_label}**: no fue posible calcular el retorno del periodo {period_label} para construir la atribucion."
    if attrib_df is None or attrib_df.empty:
        return (f"**{fund_label}** registro un retorno {period_label} de {fund_return:+.1f}%, pero no hay un universo "
                f"de holdings (oficial o proxy) disponible para este instrumento, por lo que no es posible "
                f"descomponer el movimiento por activo subyacente.")

    sorted_df = attrib_df.sort_values("Contribution_pp")
    detractors = sorted_df[sorted_df["Contribution_pp"] < 0].head(2)
    contributors = sorted_df[sorted_df["Contribution_pp"] > 0].sort_values("Contribution_pp", ascending=False).head(2)

    def _holding_phrase(row):
        return (f"{row['Name']} ({row.name}, peso {row['Weight_%']:.1f}%, retorno {row['Return_%']:+.1f}%, "
                f"contribucion {row['Contribution_pp']:+.2f} pp)")

    parts = [f"**{fund_label}** registro un retorno {period_label} de {fund_return:+.1f}%."]

    if not detractors.empty and fund_return < 0:
        phrases = "; ".join(_holding_phrase(r) for _, r in detractors.iterrows())
        parts.append(f"Los principales detractores dentro del Top Holdings fueron {phrases}.")
    elif not contributors.empty and fund_return > 0:
        phrases = "; ".join(_holding_phrase(r) for _, r in contributors.iterrows())
        parts.append(f"Los principales contribuyentes dentro del Top Holdings fueron {phrases}.")

    if fund_return < 0 and not contributors.empty:
        phrases = "; ".join(_holding_phrase(r) for _, r in contributors.iterrows())
        parts.append(f"Como contrapeso parcial, {phrases} aportaron positivamente.")
    elif fund_return > 0 and not detractors.empty:
        phrases = "; ".join(_holding_phrase(r) for _, r in detractors.iterrows())
        parts.append(f"En contraste, {phrases} resto(aron) parte de la ganancia.")

    if residual_pp is not None and abs(residual_pp) > 0.3:
        parts.append(f"El resto de holdings fuera del Top 10 (y efectos de caja/rebalanceo) explican "
                     f"aproximadamente {residual_pp:+.2f} pp adicionales del retorno total.")

    if weights_are_equal_proxy:
        parts.append("Nota metodologica: esta base de datos no trae los holdings oficiales del fondo, asi que la "
                     "atribucion usa como PROXY equal-weight el universo de empresas del ciclo de valor de IA "
                     "(ZODA CicloIA) relevante para esta categoria -no son los holdings reales ni sus pesos "
                     "reales-. Es util para entender la direccion y el mecanismo del movimiento (que eslabon de "
                     "la cadena de IA domino), pero NO es una atribucion CFA/Brinson exacta del fondo en cuestion.")
    else:
        parts.append("Nota metodologica: la atribucion usa los pesos VIGENTES mas recientes de cada holding "
                     "(los unicos que expone Yahoo Finance de forma gratuita), no los pesos exactos al inicio "
                     "del periodo analizado, por lo que es una aproximacion razonable del driver principal del "
                     "movimiento, no una atribucion Brinson exacta con pesos diarios.")

    return " ".join(parts)


def category_theme_commentary(theme_label: str, returns_in_group: dict) -> str:
    """Comentario tipo 'Weekly Top of Mind' para un grupo tematico (ej. todos
    los fondos de SubCat2='Semiconductors'): identifica el mejor y peor
    performer del grupo y la dispersion, en 2-3 oraciones."""
    clean = {k: v for k, v in returns_in_group.items() if v is not None and not pd.isna(v)}
    if len(clean) < 2:
        return f"**{theme_label}**: datos insuficientes para comparar el grupo."
    best_name, best_val = max(clean.items(), key=lambda kv: kv[1])
    worst_name, worst_val = min(clean.items(), key=lambda kv: kv[1])
    stats = dispersion_stats(list(clean.values()))

    txt = (f"**{theme_label}**: dentro de este grupo ({stats['n']} fondos con datos), {best_name} lidera con "
           f"{best_val:+.1f}%, mientras que {worst_name} es el rezagado con {worst_val:+.1f}% "
           f"(dispersion de {stats['spread']:.1f} pp entre el mejor y el peor).")

    if stats["spread"] > 15:
        txt += (" Una dispersion tan amplia dentro de un mismo tema sugiere que el driver del movimiento es mas "
               "idiosincratico (fundamentales especificos, flujos, posicionamiento) que puramente macro/sectorial; "
               "conviene mirar cada fondo individualmente antes de generalizar sobre el tema.")
    elif stats["spread"] < 5:
        txt += (" La baja dispersion indica que el grupo se mueve de forma cohesionada, consistente con un driver "
               "macro o sectorial comun dominando por encima de las diferencias idiosincraticas entre fondos.")

    return txt


# ------------------------------------------------------------------------------
# 4.2 PERU "fundamentals" (produccion minera, indicadores SBS bancarios) y el
#     chart didactico de las 12 capas del ciclo de valor de IA (FFMM).
# ------------------------------------------------------------------------------

def interpret_mining_production(empresa: str, mineral: str, series: pd.Series) -> str:
    """Lectura de la tendencia de produccion fisica de una minera (serie
    mensual, agregada por mining_production_series), en tono PM: direccion,
    magnitud YoY, aceleracion/desaceleracion reciente. No cruza con el
    precio de la accion aqui -eso se hace en render_peru_tab, donde se
    tienen ambas series a la vista- para no forzar causalidad en una
    funcion que solo ve el lado de produccion."""
    if series is None or len(series) < 4:
        return f"**{empresa}** ({mineral.title()}): historia de produccion insuficiente para una lectura de tendencia."
    chg_yoy = pct_change_over(series, 12) if len(series) > 12 else None
    chg_3m = pct_change_over(series, 3)
    trend = trend_acceleration(series, short=3, long_=6) if len(series) > 6 else "sin_datos"
    direction = _direction_word(chg_3m)
    parts = [f"**{empresa}** ({mineral.title()}): la produccion viene {direction} en el margen"]
    if chg_3m is not None:
        parts[-1] += f" ({chg_3m:+.1f}% vs. hace 3 meses)"
    parts[-1] += "."
    if chg_yoy is not None:
        parts.append(f"En terminos interanuales (YoY), la variacion es de {chg_yoy:+.1f}%.")
    trend_phrase = _trend_phrase(trend)
    if trend_phrase:
        parts.append(f"En el margen, {trend_phrase}.")
    return " ".join(parts)


def interpret_banking_fundamentals(banco: str, indicador: str, banco_series: pd.Series,
                                    sistema_series: Optional[pd.Series]) -> str:
    """Lectura de un indicador SBS (Morosidad/ROE/Liquidez) de un banco
    peruano vs. el Sistema Bancario de referencia: nivel actual, tendencia
    reciente y si el banco esta mejor o peor posicionado que el sistema."""
    if banco_series is None or banco_series.empty:
        return f"**{banco}** — {indicador}: dato no disponible."
    last_val = banco_series.iloc[-1]
    chg_6m = level_change_over(banco_series, 6) if len(banco_series) > 6 else None
    parts = [f"**{banco}** — {indicador}: {last_val:.2f}% a {banco_series.index.max().strftime('%b-%Y')}"]
    if chg_6m is not None:
        parts[-1] += f" ({chg_6m:+.2f} pp vs. hace 6 meses)"
    parts[-1] += "."
    if sistema_series is not None and not sistema_series.empty:
        sys_last = sistema_series.reindex(banco_series.index, method="ffill").iloc[-1] \
            if not sistema_series.empty else None
        if sys_last is not None and not pd.isna(sys_last):
            gap = last_val - sys_last
            better_worse = "mejor" if (gap < 0 and "Morosidad" in indicador) or (gap > 0 and "Morosidad" not in indicador) else "peor"
            parts.append(f"Esto se compara con {sys_last:.2f}% del Sistema Bancario, es decir, {abs(gap):.2f} pp "
                         f"{better_worse} posicionado que el promedio del sistema.")
    return " ".join(parts)


def interpret_ai_cycle_layers(period_label: str, layer_returns: dict) -> str:
    """Lectura comparativa de las 12 capas del ciclo de valor de IA: lider,
    rezagado, dispersion entre capas, y una senal de divergencia Compute vs.
    Semicap Equipment (leading indicator del ciclo) cuando aplica."""
    valid = {k: v for k, v in layer_returns.items() if v is not None and not pd.isna(v)}
    if len(valid) < 3:
        return "Universo de capas insuficiente para comparar en este periodo."
    ranked = sorted(valid.items(), key=lambda kv: kv[1], reverse=True)
    leader, leader_ret = ranked[0]
    laggard, laggard_ret = ranked[-1]
    disp = dispersion_stats(list(valid.values()))

    parts = [
        f"**{leader}** lidera el ciclo de valor de IA en {period_label} con {leader_ret:+.1f}% "
        f"({_magnitude_word(leader_ret)}), mientras **{laggard}** rezaga con {laggard_ret:+.1f}%."
    ]
    if disp:
        if disp["spread"] > 8:
            parts.append(f"La dispersion entre capas es amplia ({disp['spread']:.1f} pp entre la mejor y la peor): "
                         f"el ciclo de IA se esta moviendo de forma heterogenea entre eslabones de la cadena de "
                         f"valor, no como un bloque monolitico -mas consistente con rotacion/drivers "
                         f"idiosincraticos por capa que con un unico factor macro comun.")
        else:
            parts.append(f"La dispersion entre capas es moderada ({disp['spread']:.1f} pp), consistente con un "
                         f"driver macro o de sentiment hacia la IA en general moviendo a la mayoria de las capas "
                         f"en una direccion parecida.")

    compute_ret, upstream_ret = valid.get("Compute"), valid.get("Semicap Equipment")
    if compute_ret is not None and upstream_ret is not None:
        if (compute_ret > 1 and upstream_ret < compute_ret - 3) or (compute_ret < -1 and upstream_ret > compute_ret + 3):
            parts.append("Semicap Equipment -las maquinas que fabrican los chips, cuyas ordenes suelen "
                         "adelantarse varios quarters al capex real de los fabricantes- no esta confirmando "
                         "del todo el movimiento de Compute en este periodo; es una divergencia a vigilar como "
                         "posible senal temprana de cambio de ciclo, sin afirmar causalidad todavia.")
    return " ".join(parts)


# ==============================================================================
# 5. UI HELPERS
# ==============================================================================

def freshness_badge(result: FetchResult) -> str:
    if result.status == "ok":
        return "🟢 actualizado"
    if result.status == "stale":
        return f"🟡 desactualizado (ult. dato: {result.asof.date() if result.asof is not None else 'N/D'})"
    if result.status == "empty":
        return "⚪ sin observaciones"
    return "🔴 no disponible"


def source_caption(result: FetchResult) -> str:
    cap = f"Fuente: {result.source} · {freshness_badge(result)}"
    if result.error and not result.ok:
        cap += f" · motivo: {result.error}"
    return cap


def plot_line(series: pd.Series, title: str, y_title: str = "", ma_windows: Sequence[int] = ()) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=series.index, y=series.values, mode="lines",
                              name=title, line=dict(width=2)))
    for w in ma_windows:
        if len(series) > w:
            ma = series.rolling(w).mean()
            fig.add_trace(go.Scatter(x=ma.index, y=ma.values, mode="lines",
                                      name=f"MA{w}", line=dict(width=1, dash="dot")))
    fig.update_layout(title=title, yaxis_title=y_title, margin=dict(l=10, r=10, t=40, b=10),
                       height=320, template="plotly_white", legend=dict(orientation="h", y=-0.2))
    return fig


def plot_multi(series_dict: dict, title: str, normalize: bool = True) -> go.Figure:
    fig = go.Figure()
    for name, s in series_dict.items():
        if s is None or s.empty:
            continue
        y = (s / s.iloc[0] * 100.0) if normalize else s
        fig.add_trace(go.Scatter(x=y.index, y=y.values, mode="lines", name=name))
    fig.update_layout(title=title, margin=dict(l=10, r=10, t=40, b=10), height=380,
                       template="plotly_white", yaxis_title="Base 100" if normalize else "",
                       legend=dict(orientation="h", y=-0.25))
    return fig


def metric_row(items: Sequence[tuple]):
    """items: lista de (label, value_str, delta_str_or_None)"""
    cols = st.columns(len(items))
    for col, (label, value, delta) in zip(cols, items):
        col.metric(label, value, delta)


def render_result_block(result: FetchResult, chart: bool = True, ma_windows: Sequence[int] = (),
                         y_title: str = "", story_fn=None, **story_kwargs):
    st.markdown(f"#### {result.label}")
    if not result.ok:
        st.warning(f"Dato no disponible. {source_caption(result)}")
        return
    c1, c2 = st.columns([2, 1])
    with c1:
        if chart:
            st.plotly_chart(plot_line(result.series, result.label, y_title, ma_windows),
                             width="stretch")
    with c2:
        last = result.series.iloc[-1]
        chg_1m = pct_change_over(result.series, 21)
        chg_12m = pct_change_over(result.series, 252)
        st.metric("Ultimo valor", f"{last:,.2f}",
                   f"{chg_1m:+.1f}% (1M)" if chg_1m is not None else None)
        if chg_12m is not None:
            st.caption(f"Var. 12M: {chg_12m:+.1f}%")
        st.caption(source_caption(result))
    if story_fn is not None:
        st.info(story_fn(result.label, result.series, **story_kwargs))


def _sign_color(value: Optional[float], light: bool = False) -> str:
    if value is None or pd.isna(value):
        return "#f0f0f0" if light else "#ffffff"
    if value > 0:
        return "#e6f4ea" if light else "#1e7e34"
    if value < 0:
        return "#fbe9e7" if light else "#c62828"
    return "#f0f0f0" if light else "#555555"


def styled_returns_table(df: pd.DataFrame, pct_cols: Sequence[str], price_cols: Sequence[str] = ()):
    """Renderiza una tabla con celdas coloreadas (verde/rojo segun signo) para
    las columnas de retorno indicadas, al estilo de las tablas de 'Market
    Performance' de los weekly reports institucionales (DBS/HSBC/BlackRock/
    Goldman Sachs: codigo de color + valor numerico, sin adornos extra)."""
    def _color_col(s: pd.Series):
        return [f"background-color: {_sign_color(v, light=True)}" for v in s]

    def _pct_fmt(x):
        return f"{x:+.1f}%" if pd.notna(x) else "N/D"

    def _price_fmt(x):
        return f"{x:,.2f}" if pd.notna(x) else "N/D"

    fmt = {c: _pct_fmt for c in pct_cols}
    fmt.update({c: _price_fmt for c in price_cols})
    styled = df.style.apply(_color_col, subset=list(pct_cols), axis=0).format(fmt)
    st.dataframe(styled, width="stretch", hide_index=True)


def plot_winners_losers(returns: dict, title: str, top_n: int = 5) -> go.Figure:
    """Barra horizontal con los top N ganadores y top N perdedores de un
    conjunto de retornos {nombre: retorno_%}."""
    clean = {k: v for k, v in returns.items() if v is not None and not pd.isna(v)}
    ordered = sorted(clean.items(), key=lambda kv: kv[1])
    losers = ordered[:top_n]
    winners = ordered[-top_n:]
    combo = losers + [w for w in winners if w not in losers]
    combo.sort(key=lambda kv: kv[1])
    names = [c[0] for c in combo]
    vals = [c[1] for c in combo]
    colors = [_sign_color(v) for v in vals]
    fig = go.Figure(go.Bar(x=vals, y=names, orientation="h", marker_color=colors,
                            text=[f"{v:+.1f}%" for v in vals], textposition="outside"))
    fig.update_layout(title=title, height=max(300, 28 * len(names)), template="plotly_white",
                       margin=dict(l=10, r=40, t=40, b=10), xaxis_title="%")
    return fig


def plot_attribution_bars(attrib_df: pd.DataFrame, residual_pp: Optional[float], title: str) -> go.Figure:
    """Barras de contribucion (peso x retorno) por holding, estilo CFA
    performance attribution, mas una barra de 'Resto del portafolio'."""
    df = attrib_df.sort_values("Contribution_pp")
    names = list(df["Name"] + " (" + df.index.astype(str) + ")")
    vals = list(df["Contribution_pp"])
    if residual_pp is not None:
        names = ["Resto del portafolio / caja"] + names
        vals = [residual_pp] + vals
    colors = [_sign_color(v) for v in vals]
    fig = go.Figure(go.Bar(x=vals, y=names, orientation="h", marker_color=colors,
                            text=[f"{v:+.2f} pp" for v in vals], textposition="outside"))
    fig.update_layout(title=title, height=max(320, 32 * len(names)), template="plotly_white",
                       margin=dict(l=10, r=60, t=40, b=10), xaxis_title="Contribucion (pp)")
    return fig


def plot_heatmap(corr_df: pd.DataFrame, title: str) -> go.Figure:
    fig = go.Figure(go.Heatmap(z=corr_df.values, x=list(corr_df.columns), y=list(corr_df.index),
                                colorscale="RdBu", zmid=0, zmin=-1, zmax=1,
                                text=np.round(corr_df.values, 2), texttemplate="%{text}",
                                colorbar=dict(title="corr")))
    fig.update_layout(title=title, height=max(400, 28 * len(corr_df)), template="plotly_white",
                       margin=dict(l=10, r=10, t=40, b=10))
    return fig


def plot_risk_return_scatter(df: pd.DataFrame, title: str) -> go.Figure:
    """Scatter clasico riesgo/retorno: eje X = volatilidad anualizada 1Y,
    eje Y = retorno 1Y, color por categoria, tamano fijo. df debe tener
    columnas 'Nombre', 'vol_1y_%', 'r1y_%', 'cat'."""
    fig = go.Figure()
    for cat, sub in df.groupby("cat"):
        fig.add_trace(go.Scatter(x=sub["vol_1y_%"], y=sub["r1y_%"], mode="markers", name=cat,
                                  text=sub["Nombre"], marker=dict(size=11, opacity=0.85),
                                  hovertemplate="%{text}<br>Vol 1Y: %{x:.1f}%<br>Retorno 1Y: %{y:.1f}%<extra></extra>"))
    fig.update_layout(title=title, xaxis_title="Volatilidad anualizada 1Y (%)",
                       yaxis_title="Retorno 1Y (%)", height=460, template="plotly_white",
                       margin=dict(l=10, r=10, t=40, b=10), legend=dict(orientation="h", y=-0.25))
    return fig


# ==============================================================================
# 6. RESOLUCION DEL UNICO INPUT (BD_Consolidada_Dash_Macro.xlsx): sube/repo
# ==============================================================================

def resolve_consolidated_source(uploaded_file):
    """Resuelve que fuente usar para el Excel consolidado -el UNICO input de
    toda la capa de datos local de esta app (Peru equity, FFMM, Ciclo IA,
    fundamentals)-: el archivo subido manualmente en esta sesion si hay uno,
    o el detectado en el repo (data/BD_Consolidada_Dash_Macro.xlsx o la
    raiz) si no. Devuelve (source, origen_str) donde origen_str describe de
    donde vino para mostrarlo en la UI; source es None si no hay ninguno
    disponible (la app sigue funcionando, mostrando 'N/D' en todo lo que
    depende de este archivo, consistente con 'no fabricar datos')."""
    repo_path = _resolve_data_file(CONSOLIDATED_XLSX_NAME)
    if uploaded_file is not None:
        return uploaded_file, "subido manualmente en esta sesion"
    if repo_path is not None:
        return repo_path, f"detectado en el repo: {repo_path.relative_to(APP_DIR)}"
    return None, "no disponible (ni en el repo ni subido)"


# ==============================================================================
# 7. TAB RENDERERS
# ==============================================================================

def render_commodities_tab():
    st.subheader("Commodities internacionales")
    st.caption("Futuros continuos (front-month) via Yahoo Finance. Relevancia para Peru indicada por metal.")

    results = {name: fetch_yf(meta["ticker"], label=name) for name, meta in COMMODITIES.items()}

    ok_series = {name: r.series for name, r in results.items() if r.ok}
    if ok_series:
        st.plotly_chart(plot_multi(ok_series, "Desempeno normalizado (base 100)"), width="stretch")

    metric_items = []
    for name, r in results.items():
        if r.ok:
            chg = pct_change_over(r.series, 21)
            metric_items.append((name, f"{r.series.iloc[-1]:,.2f}", f"{chg:+.1f}% (1M)" if chg is not None else None))
        else:
            metric_items.append((name, "N/D", None))
    for i in range(0, len(metric_items), 3):
        metric_row(metric_items[i:i + 3])

    st.markdown("---")
    st.markdown("##### Lectura de Portfolio Manager, por commodity")
    for name, meta in COMMODITIES.items():
        r = results[name]
        if not r.ok:
            st.warning(f"**{name}**: {source_caption(r)}")
            continue
        with st.expander(f"{name} ({meta['unit']})", expanded=(meta["peru_relevance"] in ("alta", "muy alta"))):
            st.plotly_chart(plot_line(r.series, name, meta["unit"], ma_windows=(50, 200)), width="stretch")
            st.info(interpret_commodity_with_context(name, r.series, meta["peru_relevance"]))
            st.caption(source_caption(r))

    gold = results.get("Gold")
    copper = results.get("Copper")
    if gold and copper and gold.ok and copper.ok:
        ratio = (gold.series / copper.series).dropna()
        if len(ratio) > 60:
            st.markdown("##### Ratio Oro/Cobre (proxy de apetito por riesgo macro)")
            st.plotly_chart(plot_line(ratio, "Oro/Cobre", ""), width="stretch")
            chg = pct_change_over(ratio, 63)
            if chg is not None:
                direction = "subiendo" if chg > 0 else "bajando"
                st.caption(f"El ratio viene {direction} {abs(chg):.1f}% en 3 meses: una suba suele asociarse a mayor "
                           f"aversion al riesgo / expectativas de desaceleracion global, y una baja a mayor apetito "
                           f"por crecimiento ciclico (positivo para economias de base minera como Peru).")


def render_peru_tab(precios_zoda: dict, produccion_df: pd.DataFrame, sbs_df: pd.DataFrame, source_label: str):
    st.subheader("Mercado de capitales peruano")
    st.caption(f"Equity local (BVL + ADRs peruanos) y benchmarks regionales desde {source_label}. "
               "FX y pares mineros internacionales via Yahoo Finance (no estan en el Excel consolidado).")

    fx_res = fetch_yf(PERU_FX["ticker"], label=PERU_FX["nombre"])
    c1, c2 = st.columns([2, 1])
    with c1:
        if fx_res.ok:
            st.plotly_chart(plot_line(fx_res.series, "USD/PEN", "PEN por USD", ma_windows=(50, 200)),
                             width="stretch")
        else:
            st.warning(source_caption(fx_res))
    with c2:
        if fx_res.ok:
            st.info(interpret_fx("USD/PEN", fx_res.series, is_usd_base=True))
        st.caption(source_caption(fx_res))

    st.markdown("##### Contexto macro de Peru (FRED, frecuencia anual)")
    macro_cols = st.columns(len(PERU_MACRO_FRED))
    for col, (label, series_candidates) in zip(macro_cols, PERU_MACRO_FRED.items()):
        r = fetch_fred_with_candidates(series_candidates, label=label, expected_gap_days=420)
        with col:
            if r.ok:
                col.metric(label, f"{r.series.iloc[-1]:.1f}%", None)
                col.caption(source_caption(r))
            else:
                col.warning(f"{label}: no disponible")

    # --- Equities peruanas: metricas completas WTD/MTD/YTD/1Y/3Y (ya posible
    #     gracias a tener toda la historia de precios en el Excel consolidado,
    #     a diferencia de la version anterior que solo mostraba 1M/YTD) ---
    st.markdown("---")
    st.markdown("##### Equities peruanas (BVL + ADRs)")
    rows, equity_series, equity_wtd = [], {}, {}
    for ticker, meta in PERU_EQUITY_META.items():
        series = precios_zoda.get(ticker)
        r = _series_to_fetch_result(meta["nombre"], series, source_label, expected_gap_days=5)
        m = fund_metrics(r.series) if r.ok else {}
        rows.append({
            "Empresa": meta["nombre"], "Listing": meta["listing"], "BBG": ticker,
            "Ultimo precio": m.get("ultimo_precio"), "WTD %": m.get("wtd_%"), "MTD %": m.get("mtd_%"),
            "YTD %": m.get("ytd_%"), "1Y %": m.get("r1y_%"), "Estado dato": freshness_badge(r),
        })
        if r.ok:
            equity_series[meta["nombre"]] = r.series
            equity_wtd[meta["nombre"]] = m.get("wtd_%")
    df_equity = pd.DataFrame(rows)
    styled_returns_table(df_equity, pct_cols=["WTD %", "MTD %", "YTD %", "1Y %"], price_cols=["Ultimo precio"])
    n_ok = len(equity_series)
    st.caption(f"{n_ok} de {len(PERU_EQUITY_META)} instrumentos con precio disponible en {source_label}. "
               "Los que muestran 'N/D' no tuvieron match de precio en el pipeline de origen para este ticker.")

    if equity_series:
        st.plotly_chart(plot_winners_losers(equity_wtd, "Ganadores / rezagados de la semana, WTD (equity local)"),
                         width="stretch")
        default_pick = list(equity_series.keys())[:6]
        picked = st.multiselect("Comparar desempeno (base 100)", sorted(equity_series.keys()), default=default_pick)
        if picked:
            st.plotly_chart(plot_multi({k: equity_series[k] for k in picked}, "Desempeno normalizado (base 100)"),
                             width="stretch")

    # --- Benchmarks regionales: ahora indices REALES del Excel, ya no
    #     proxies de ETF (mejora clave vs. la version anterior de esta app) ---
    st.markdown("---")
    st.markdown("##### Benchmarks regionales (indices reales)")
    bench_series = {}
    for ticker, meta in PERU_BENCHMARK_META.items():
        series = precios_zoda.get(ticker)
        r = _series_to_fetch_result(meta["nombre"], series, source_label, expected_gap_days=5)
        if r.ok:
            bench_series[meta["nombre"]] = r.series
        cols = st.columns([3, 1, 2])
        cols[0].write(f"**{meta['nombre']}**")
        cols[1].write(freshness_badge(r))
        if meta["nota"]:
            cols[2].caption(meta["nota"])
    if bench_series:
        st.plotly_chart(plot_multi(bench_series, "Indices regionales (base 100)"), width="stretch")

    st.markdown("---")
    st.markdown("##### Pares mineros internacionales (comparables de Peru)")
    peer_series = {}
    for p in PERU_MINING_PEERS:
        r = fetch_with_candidates(p["yahoo_candidates"], label=p["nombre"])
        if r.ok:
            peer_series[p["nombre"]] = r.series
    for local_name in ("Buenaventura (BVN)", "Hochschild Mining (Ares)"):
        if local_name in equity_series:
            peer_series[local_name] = equity_series[local_name]
    if peer_series:
        st.plotly_chart(plot_multi(peer_series, "Mineras: Peru vs. pares internacionales (base 100)"),
                         width="stretch")
        st.caption("Comparar BVN/Hochschild contra Southern Copper y Hudbay ayuda a distinguir si un movimiento "
                   "en una minera peruana responde a un factor sectorial global (commodity, risk-on/off) o a un "
                   "factor idiosincratico local.")

    # --- Fundamentals: produccion minera fisica vs. desempeno de la accion ---
    st.markdown("---")
    st.markdown("##### Fundamentals: produccion minera")
    st.caption(f"Produccion mensual (etapa concentracion/flotacion) desde {source_label}, hoja Produccion_Mineria. "
               "Se usa un (1) mineral principal por empresa para evitar mezclar unidades de medida.")
    mining_tickers = [t for t in MINING_PRIMARY_METAL if t in PERU_EQUITY_META]
    if produccion_df is None or produccion_df.empty or not mining_tickers:
        st.info("Hoja Produccion_Mineria no disponible en el Excel consolidado.")
    else:
        mining_label = st.selectbox("Selecciona una minera", [PERU_EQUITY_META[t]["nombre"] for t in mining_tickers],
                                     key="peru_mining_select")
        sel_ticker = next(t for t in mining_tickers if PERU_EQUITY_META[t]["nombre"] == mining_label)
        mineral = MINING_PRIMARY_METAL[sel_ticker]
        prod_series = mining_production_series(produccion_df, sel_ticker)
        price_series = precios_zoda.get(sel_ticker)
        if prod_series is None:
            st.warning(f"Sin serie de produccion de {mineral.title()} disponible para {mining_label}.")
        else:
            overlay = {f"Precio ({mining_label})": price_series, f"Produccion {mineral.title()} (TMF/KG/G finos)": prod_series}
            st.plotly_chart(plot_multi({k: v for k, v in overlay.items() if v is not None and not v.empty},
                                        f"{mining_label}: precio vs. produccion fisica (base 100)"),
                             width="stretch")
            st.info(interpret_mining_production(mining_label, mineral, prod_series))
            if price_series is not None and not price_series.empty:
                price_chg_3m = pct_change_over(price_series, 63)
                prod_chg_3m = pct_change_over(prod_series, 3)
                if price_chg_3m is not None and prod_chg_3m is not None:
                    if np.sign(price_chg_3m) != np.sign(prod_chg_3m) and abs(price_chg_3m) > 3:
                        st.caption(f"Nota: el precio de la accion ({price_chg_3m:+.1f}% en 3M) se mueve en sentido "
                                   f"contrario a la produccion fisica reciente ({prod_chg_3m:+.1f}%), lo que sugiere "
                                   f"que el mercado esta descontando otro factor (precio del commodity, costos, "
                                   f"flujos) por encima del volumen operativo.")

    # --- Fundamentals: indicadores SBS de los bancos cotizados ---
    st.markdown("---")
    st.markdown("##### Fundamentals: indicadores SBS bancarios")
    st.caption(f"Indicadores mensuales SBS desde {source_label}, hoja Indicadores_SBS_Bancos, vs. Sistema Bancario.")
    if sbs_df is None or sbs_df.empty:
        st.info("Hoja Indicadores_SBS_Bancos no disponible en el Excel consolidado.")
    else:
        bank_label = st.selectbox("Selecciona un banco",
                                   [PERU_EQUITY_META[t]["nombre"] for t in SBS_BANK_TICKERS], key="peru_bank_select")
        sel_bank_ticker = next(t for t in SBS_BANK_TICKERS if PERU_EQUITY_META[t]["nombre"] == bank_label)
        entidad = SBS_BANK_TICKERS[sel_bank_ticker]
        indicator_label = st.radio("Indicador", list(SBS_HEADLINE_INDICATORS.keys()), horizontal=True,
                                    key="peru_bank_indicator")
        suffix = SBS_HEADLINE_INDICATORS[indicator_label]
        banco_series = sbs_indicator_series(sbs_df, entidad, suffix)
        sistema_series = sbs_indicator_series(sbs_df, SBS_SYSTEM_LABEL, suffix)
        if banco_series is None:
            st.warning(f"Sin dato de '{indicator_label}' disponible para {bank_label}.")
        else:
            cmp_series = {f"{bank_label}": banco_series}
            if sistema_series is not None:
                cmp_series["Sistema Bancario"] = sistema_series
            st.plotly_chart(plot_multi(cmp_series, f"{indicator_label} — {bank_label} vs. Sistema Bancario",
                                        normalize=False), width="stretch")
            st.info(interpret_banking_fundamentals(bank_label, indicator_label, banco_series, sistema_series))


def render_usa_tab():
    st.subheader("Estados Unidos: economia real, mercado laboral, inflacion, Fed y fiscal")

    st.markdown("##### Mercado laboral")
    unrate = fetch_fred(USA_LABOR["Tasa de desempleo (%)"], "Tasa de desempleo (%)", expected_gap_days=55)
    payrolls = fetch_fred(USA_LABOR["Nonfarm Payrolls (miles, nivel)"], "Nonfarm Payrolls", expected_gap_days=55)
    claims = fetch_fred(USA_LABOR["Solicitudes iniciales de desempleo"], "Initial Claims", expected_gap_days=14)
    participation = fetch_fred(USA_LABOR["Tasa de participacion laboral (%)"], "Tasa de participacion (%)", expected_gap_days=55)

    c1, c2 = st.columns(2)
    with c1:
        if unrate.ok:
            st.plotly_chart(plot_line(unrate.series.iloc[-120:], "Tasa de desempleo (%)"), width="stretch")
        else:
            st.warning(source_caption(unrate))
    with c2:
        if claims.ok:
            st.plotly_chart(plot_line(claims.series.iloc[-260:], "Solicitudes iniciales de desempleo"), width="stretch")
        else:
            st.warning(source_caption(claims))

    st.info(interpret_labor_market(
        unrate.series if unrate.ok else None,
        payrolls.series if payrolls.ok else None,
        claims.series if claims.ok else None,
    ))
    if participation.ok:
        st.caption(f"Tasa de participacion laboral mas reciente: {participation.series.iloc[-1]:.1f}% "
                   f"({source_caption(participation)}).")

    st.markdown("---")
    st.markdown("##### Inflacion")
    cpi = fetch_fred(USA_INFLATION["CPI headline (indice)"], "CPI headline", expected_gap_days=55)
    core_cpi = fetch_fred(USA_INFLATION["CPI core (indice)"], "CPI core", expected_gap_days=55)
    breakeven = fetch_fred(USA_INFLATION["Breakeven inflacionario 5Y (%)"], "Breakeven 5Y", expected_gap_days=10)

    if cpi.ok:
        yoy_series = resample_monthly(cpi.series).pct_change(12).dropna() * 100.0
        st.plotly_chart(plot_line(yoy_series.iloc[-84:], "CPI headline YoY (%)"), width="stretch")
    st.info(interpret_inflation(cpi.series if cpi.ok else None, core_cpi.series if core_cpi.ok else None,
                                 breakeven.series if breakeven.ok else None))

    st.markdown("---")
    st.markdown("##### Politica de la Fed y balance")
    fed_funds = fetch_fred(USA_FED_POLICY["Fed Funds Rate (%)"], "Fed Funds Rate (%)", expected_gap_days=55)
    balance = fetch_fred(USA_FED_POLICY["Balance de la Fed (USD mill)"], "Balance de la Fed", expected_gap_days=14)
    c1, c2 = st.columns(2)
    with c1:
        if fed_funds.ok:
            st.plotly_chart(plot_line(fed_funds.series.iloc[-260:], "Fed Funds Rate (%)"), width="stretch")
    with c2:
        if balance.ok:
            st.plotly_chart(plot_line(balance.series.iloc[-260:], "Balance de la Fed (USD mill)"), width="stretch")
    st.info(interpret_fed_policy(fed_funds.series if fed_funds.ok else None, balance.series if balance.ok else None))

    st.markdown("---")
    st.markdown("##### Curva de rendimientos del Tesoro")
    y2 = fetch_fred(USA_RATES["UST 2 anos (%)"], "UST 2Y", expected_gap_days=10)
    y10 = fetch_fred(USA_RATES["UST 10 anos (%)"], "UST 10Y", expected_gap_days=10)
    y30 = fetch_fred(USA_RATES["UST 30 anos (%)"], "UST 30Y", expected_gap_days=10)
    if y2.ok and y10.ok:
        curve_df = pd.DataFrame({"2Y": y2.series, "10Y": y10.series})
        if y30.ok:
            curve_df["30Y"] = y30.series
        curve_df = curve_df.dropna().iloc[-260:]
        fig = go.Figure()
        for col in curve_df.columns:
            fig.add_trace(go.Scatter(x=curve_df.index, y=curve_df[col], mode="lines", name=col))
        fig.update_layout(title="Rendimientos UST", height=340, template="plotly_white",
                           margin=dict(l=10, r=10, t=40, b=10), legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig, width="stretch")
    st.info(interpret_yield_curve(y2.series if y2.ok else None, y10.series if y10.ok else None,
                                   y30.series if y30.ok else None))

    st.markdown("---")
    st.markdown("##### Situacion fiscal")
    debt_gdp = fetch_fred(USA_FISCAL["Deuda publica / PBI (%)"], "Deuda/PBI (%)", expected_gap_days=200)
    deficit = fetch_fred(USA_FISCAL["Resultado fiscal mensual (USD mill)"], "Resultado fiscal mensual", expected_gap_days=55)
    c1, c2 = st.columns(2)
    with c1:
        if debt_gdp.ok:
            st.plotly_chart(plot_line(debt_gdp.series, "Deuda publica / PBI (%)"), width="stretch")
    with c2:
        if deficit.ok:
            st.plotly_chart(plot_line(deficit.series.iloc[-120:], "Resultado fiscal mensual (USD mill)"),
                             width="stretch")
    st.info(interpret_fiscal(debt_gdp.series if debt_gdp.ok else None, deficit.series if deficit.ok else None))

    st.markdown("---")
    st.markdown("##### Estres financiero y volatilidad")
    nfci = fetch_fred(USA_STRESS["Chicago Fed Financial Conditions (NFCI)"], "NFCI", expected_gap_days=14)
    vix = fetch_yf(USA_VIX_TICKER, "VIX", expected_gap_days=7)
    c1, c2 = st.columns(2)
    with c1:
        if nfci.ok:
            st.plotly_chart(plot_line(nfci.series.iloc[-260:], "NFCI"), width="stretch")
            st.info(interpret_stress_index(nfci.series, "NFCI (Chicago Fed)"))
    with c2:
        if vix.ok:
            st.plotly_chart(plot_line(vix.series.iloc[-260:], "VIX"), width="stretch")
            st.info(interpret_market_series("VIX", vix.series, "macro",
                     extra_note="Niveles de VIX por debajo de 15 suelen asociarse a complacencia; por encima de 25-30, a estres agudo."))


def render_world_tab():
    st.subheader("Mapa macro global por pais")
    st.caption("CPI interanual, tasa de politica, rendimiento 10Y, crecimiento del PBI y tipo de cambio. "
               "Fuentes: FRED (series OECD/World Bank espejadas) y Yahoo Finance. Ver notas por pais.")

    scorecard_rows = []
    equity_series = {}
    fx_series = {}

    for country, cfg in WORLD_COUNTRIES.items():
        cpi_r = fetch_fred(cfg["cpi"], f"{country} CPI", expected_gap_days=55) if cfg.get("cpi") else None
        rate_r = fetch_fred(cfg["rate"], f"{country} tasa", expected_gap_days=55) if cfg.get("rate") else None
        y10_r = fetch_fred(cfg["y10"], f"{country} 10Y", expected_gap_days=55) if cfg.get("y10") else None
        gdp_r = fetch_fred_with_candidates(cfg["gdp"], f"{country} PBI", expected_gap_days=420) if cfg.get("gdp") else None
        fx_r = fetch_yf(cfg["fx"], f"{country} FX", expected_gap_days=7) if cfg.get("fx") else None
        eq_r = fetch_yf(cfg["equity"], f"{country} equity", expected_gap_days=7) if cfg.get("equity") else None

        cpi_val = cpi_r.series.iloc[-1] if cpi_r and cpi_r.ok else None
        rate_val = rate_r.series.iloc[-1] if rate_r and rate_r.ok else None
        y10_val = y10_r.series.iloc[-1] if y10_r and y10_r.ok else None
        gdp_val = gdp_r.series.iloc[-1] if gdp_r and gdp_r.ok else None
        fx_chg = pct_change_over(fx_r.series, 63) if fx_r and fx_r.ok else None
        eq_chg = pct_change_over(eq_r.series, 63) if eq_r and eq_r.ok else None

        if eq_r and eq_r.ok:
            equity_series[country] = eq_r.series
        if fx_r and fx_r.ok:
            fx_series[country] = fx_r.series

        scorecard_rows.append({
            "Pais": country,
            "CPI YoY %": f"{cpi_val:.1f}" if cpi_val is not None else "N/D",
            "Tasa politica %": f"{rate_val:.2f}" if rate_val is not None else "N/D",
            "Bono 10Y %": f"{y10_val:.2f}" if y10_val is not None else "N/D",
            "PBI real YoY %": f"{gdp_val:.1f}" if gdp_val is not None else "N/D",
            "FX 3M %": f"{fx_chg:+.1f}" if fx_chg is not None else "N/D",
            "Equity 3M %": f"{eq_chg:+.1f}" if eq_chg is not None else "N/D",
        })

    st.dataframe(pd.DataFrame(scorecard_rows), width="stretch", hide_index=True)

    if equity_series:
        st.plotly_chart(plot_multi(equity_series, "Indices bursatiles globales (base 100)"), width="stretch")
    if fx_series:
        st.plotly_chart(plot_multi(fx_series, "Tipos de cambio vs. USD (base 100)"), width="stretch")

    st.markdown("---")
    st.markdown("##### Lectura de PM, por pais")
    for country, cfg in WORLD_COUNTRIES.items():
        with st.expander(country):
            cpi_r = fetch_fred(cfg["cpi"], f"{country} CPI") if cfg.get("cpi") else None
            rate_r = fetch_fred(cfg["rate"], f"{country} tasa") if cfg.get("rate") else None
            y10_r = fetch_fred(cfg["y10"], f"{country} 10Y") if cfg.get("y10") else None
            gdp_r = fetch_fred_with_candidates(cfg["gdp"], f"{country} PBI") if cfg.get("gdp") else None
            fx_r = fetch_yf(cfg["fx"], f"{country} FX") if cfg.get("fx") else None

            cpi_val = cpi_r.series.iloc[-1] if cpi_r and cpi_r.ok else None
            rate_val = rate_r.series.iloc[-1] if rate_r and rate_r.ok else None
            y10_val = y10_r.series.iloc[-1] if y10_r and y10_r.ok else None
            gdp_val = gdp_r.series.iloc[-1] if gdp_r and gdp_r.ok else None
            fx_chg = pct_change_over(fx_r.series, 63) if fx_r and fx_r.ok else None

            st.write(interpret_country_scorecard(country, cpi_val, rate_val, y10_val, gdp_val, fx_chg))
            if cfg.get("nota"):
                st.caption(f"Nota de cobertura de datos: {cfg['nota']}")


def _ffmm_group_label(row: pd.Series) -> str:
    parts = [p for p in [row.get("Mercado"), row.get("Cat"), row.get("SubCat"), row.get("SubCat2")] if p and str(p) != "nan"]
    return " · ".join(parts)


def _ffmm_proxy_universe(subcat2: str, ciclo_ia: dict) -> Optional[dict]:
    """Devuelve el subconjunto de `ciclo_ia` a usar como proxy de holdings
    equal-weight para un fondo FFMM de tecnologia/IA, segun su SubCat2:
    'Semiconductors' -> solo la cadena de valor angosta de semis
    (FFMM_SEMI_CAPAS); 'Core'/'Software & Internet'/'Thematic' -> el
    universo completo de 46 companias de las 12 capas. None para el resto
    (fondos sin ningun proxy razonable: no se fabrica una atribucion)."""
    if subcat2 == "Semiconductors":
        return {k: v for k, v in ciclo_ia.items() if v["capa"] in FFMM_SEMI_CAPAS}
    if subcat2 in ("Core", "Software & Internet", "Thematic"):
        return dict(ciclo_ia)
    return None


def render_ffmm_tab(ffmm_precios: dict, parametria_ffmm: pd.DataFrame, ciclo_ia: dict, source_label: str):
    st.subheader("FFMM: fondos mutuos y ETFs tematicos")
    st.caption(
        f"Universo e historia de precios/retornos desde {CONSOLIDATED_XLSX_NAME} (hojas Parametria_FFMM + "
        f"FFMM_Precios_Retornos, {source_label}) — ya no se consulta Yahoo Finance para este modulo. El formato "
        "de las tablas de performance por periodo y el tono de las lecturas esta inspirado en weekly reports "
        "institucionales tipo 'CIO Market Pulse' y, en general, HSBC/BlackRock/Goldman Sachs: encabezado corto "
        "en negrita + 2-3 oraciones con el mecanismo, no solo el numero."
    )

    if parametria_ffmm.empty:
        st.warning(f"No se pudo leer la hoja Parametria_FFMM de {CONSOLIDATED_XLSX_NAME}. Verifica el archivo "
                   "en la barra lateral.")
        return

    # --- 1) Resolver precios + metricas de TODOS los fondos de la parametria ---
    rows = []
    resolved = {}  # nombre -> (FetchResult, metrics dict, fila parametria)
    for _, p in parametria_ffmm.iterrows():
        nombre, isin = p.get("Nombre"), p.get("ISIN")
        s = ffmm_precios.get(isin) if isin and str(isin) != "nan" else None
        r = _series_to_fetch_result(nombre, s, f"Excel consolidado ({CONSOLIDATED_XLSX_NAME})", expected_gap_days=7)
        m = {}
        if r.ok:
            m = fund_metrics(r.series)
            resolved[nombre] = (r, m, p)
        rows.append({
            "Nombre": nombre, "Mercado": p.get("Mercado", ""), "Cat": p.get("Cat", ""),
            "SubCat": p.get("SubCat", ""), "SubCat2": p.get("SubCat2") or "",
            "ISIN": isin if isin and str(isin) != "nan" else "N/D",
            "Ultimo precio": m.get("ultimo_precio"), "Ultimo ret. %": m.get("ultimo_retorno_%"),
            "WTD %": m.get("wtd_%"), "MTD %": m.get("mtd_%"), "YTD %": m.get("ytd_%"),
            "1Y %": m.get("r1y_%"), "3Y acum %": m.get("r3y_acum_%"), "3Y CAGR %": m.get("r3y_cagr_%"),
        })
    df_all = pd.DataFrame(rows)
    n_ok, n_total = len(resolved), len(parametria_ffmm)

    c1, c2, c3 = st.columns(3)
    c1.metric("Fondos con precio en el Excel", f"{n_ok} / {n_total}")
    c2.metric("Sin ISIN o sin precio (N/D)", n_total - n_ok)
    n_con_attrib = sum(1 for p in parametria_ffmm["SubCat2"] if p in FFMM_SEMI_CAPAS or p in
                        ("Semiconductors", "Core", "Software & Internet", "Thematic"))
    c3.metric("Con proxy de attribution (IA/semis)", n_con_attrib)

    st.markdown("##### Tabla de performance (estilo Market Performance)")
    styled_returns_table(
        df_all.sort_values(["Mercado", "Cat", "SubCat", "SubCat2", "Nombre"]),
        pct_cols=["Ultimo ret. %", "WTD %", "MTD %", "YTD %", "1Y %", "3Y acum %", "3Y CAGR %"],
        price_cols=["Ultimo precio"],
    )
    st.caption("WTD/MTD/YTD: retorno desde el ultimo cierre antes del inicio de la semana/mes/anio en curso. "
               "1Y/3Y: retorno desde el cierre disponible hace 1/3 anios respecto a la ultima fecha con dato. "
               "3Y CAGR: tasa anualizada compuesta del retorno acumulado a 3 anios.")

    sin_dato = parametria_ffmm[~parametria_ffmm["Nombre"].isin(resolved.keys())]
    if not sin_dato.empty:
        with st.expander(f"Fondos sin precio disponible en el Excel consolidado ({len(sin_dato)})"):
            for _, p in sin_dato.iterrows():
                motivo = "sin ISIN en la parametria" if pd.isna(p.get("ISIN")) else \
                    "su ISIN no matcheo contra la base de precios de origen"
                st.write(f"🔴 **{p['Nombre']}** — {motivo}; se muestra 'N/D' en vez de inventar un precio.")

    if not resolved:
        st.warning("No se pudo resolver ningun fondo con precio en el Excel consolidado.")
        return

    # --- 2) Winners / Losers WTD ---
    st.markdown("---")
    st.markdown("##### Ganadores y rezagados de la semana (WTD)")
    wtd_returns = {name: m.get("wtd_%") for name, (_, m, _) in resolved.items()}
    st.plotly_chart(plot_winners_losers(wtd_returns, "Top 5 / Bottom 5 WTD"), width="stretch")

    # --- 3) Grupos tematicos: grafico comparado + comentario tipo CIO ---
    st.markdown("---")
    st.markdown("##### Lectura por grupo tematico")
    groups = {}
    for name, (_, _, p) in resolved.items():
        groups.setdefault(_ffmm_group_label(p), []).append(name)
    groups = {g: names for g, names in groups.items() if len(names) >= 2}
    if groups:
        group_choice = st.selectbox("Selecciona un grupo tematico", sorted(groups.keys()))
        names_in_group = groups[group_choice]
        series_dict = {n: resolved[n][0].series for n in names_in_group}
        st.plotly_chart(plot_multi(series_dict, f"{group_choice} — desempeno normalizado (base 100)"),
                         width="stretch")
        returns_in_group = {n: resolved[n][1].get("wtd_%") for n in names_in_group}
        st.info(category_theme_commentary(group_choice, returns_in_group))
    else:
        st.caption("Aun no hay suficientes fondos resueltos en un mismo grupo para comparar (minimo 2).")

    # --- 4) Ficha de detalle por fondo + performance attribution (proxy Ciclo IA) ---
    st.markdown("---")
    st.markdown("##### Ficha de detalle y performance attribution (estilo CFA)")
    selected = st.selectbox("Selecciona un fondo/ETF", sorted(resolved.keys()), key="ffmm_detail_select")
    r_sel, m_sel, p_sel = resolved[selected]
    subcat2_sel = p_sel.get("SubCat2") or ""

    c1, c2 = st.columns([2, 1])
    with c1:
        st.plotly_chart(plot_line(r_sel.series, selected, "Precio", ma_windows=(50, 200)), width="stretch")
    with c2:
        metric_row([("WTD", f"{m_sel['wtd_%']:+.1f}%" if m_sel.get("wtd_%") is not None else "N/D", None),
                    ("MTD", f"{m_sel['mtd_%']:+.1f}%" if m_sel.get("mtd_%") is not None else "N/D", None)])
        metric_row([("YTD", f"{m_sel['ytd_%']:+.1f}%" if m_sel.get("ytd_%") is not None else "N/D", None),
                    ("1Y", f"{m_sel['r1y_%']:+.1f}%" if m_sel.get("r1y_%") is not None else "N/D", None)])
        metric_row([("3Y acum.", f"{m_sel['r3y_acum_%']:+.1f}%" if m_sel.get("r3y_acum_%") is not None else "N/D", None),
                    ("3Y CAGR", f"{m_sel['r3y_cagr_%']:+.1f}%" if m_sel.get("r3y_cagr_%") is not None else "N/D", None)])
        st.caption(source_caption(r_sel))

    st.info(fund_periods_summary(selected, m_sel))
    st.info(interpret_market_series(selected, r_sel.series, "equity"))

    st.markdown("###### Performance attribution (proxy equal-weight, Ciclo de Valor de IA)")
    proxy_universe = _ffmm_proxy_universe(subcat2_sel, ciclo_ia)
    if proxy_universe is None:
        st.caption(
            f"**{selected}** (SubCat2: {subcat2_sel or 'N/A'}): esta base de datos no trae los holdings oficiales "
            "del fondo, y no es un fondo de tecnologia/semiconductores para el que el universo de "
            "Precios_ZODA_CicloIA sea un proxy razonable, asi que no se construye una atribucion (preferible a "
            "forzar pesos ficticios sin fundamento)."
        )
    else:
        scope_label = "cadena de valor angosta de Semiconductores" if subcat2_sel == "Semiconductors" \
            else "universo completo de las 12 capas del Ciclo de Valor de IA"
        st.caption(f"Proxy: {len(proxy_universe)} companias ({scope_label}, equal-weight) — NO son los holdings "
                   f"oficiales de {selected}.")

        period_map = {"WTD": "wtd_%", "MTD": "mtd_%", "YTD": "ytd_%", "1Y": "r1y_%"}
        period_choice = st.radio("Periodo para la atribucion", list(period_map.keys()), horizontal=True,
                                  key="ffmm_attrib_period")
        fund_period_ret = m_sel.get(period_map[period_choice])
        n_proxy = len(proxy_universe)
        equal_weight_pct = 100.0 / n_proxy if n_proxy else 0.0

        attrib_rows = []
        for ticker, info in proxy_universe.items():
            h_ret = period_return(info["series"], period_choice)
            contribution = (equal_weight_pct / 100.0) * h_ret if h_ret is not None else None
            attrib_rows.append({"Symbol": ticker, "Name": info["nombre"], "Weight_%": equal_weight_pct,
                                 "Return_%": h_ret, "Contribution_pp": contribution})
        attrib_df = pd.DataFrame(attrib_rows).set_index("Symbol")
        attrib_df_valid = attrib_df.dropna(subset=["Contribution_pp"]).sort_values("Contribution_pp")
        # Top/bottom 10 por |contribucion| para que el chart siga siendo legible en mobile.
        attrib_df_top = pd.concat([attrib_df_valid.head(6), attrib_df_valid.tail(6)]).drop_duplicates()

        if attrib_df_valid.empty:
            st.warning("No se pudieron calcular retornos individuales del universo proxy en este periodo.")
        else:
            explained_pp = float(attrib_df_valid["Contribution_pp"].sum())
            residual_pp = (fund_period_ret - explained_pp) if fund_period_ret is not None else None
            st.plotly_chart(
                plot_attribution_bars(attrib_df_top, residual_pp,
                                       f"Performance attribution (proxy) — {selected} ({period_choice})"),
                width="stretch",
            )
            st.info(performance_attribution_story(selected, period_choice, fund_period_ret, attrib_df_valid,
                                                   residual_pp, weights_are_equal_proxy=True))
            n_missing = len(attrib_df) - len(attrib_df_valid)
            if n_missing:
                st.caption(f"{n_missing} compania(s) del universo proxy sin retorno disponible en este periodo "
                           f"(excluidas de la atribucion numerica).")

    # --- 5) Ciclo de Valor de IA: las 12 capas, lado a lado ---
    st.markdown("---")
    st.markdown("##### Ciclo de Valor de IA: performance por capa (Precios_ZODA_CicloIA)")
    st.caption(
        "Cada capa es un indice equal-weight (base 100 en el primer dato de cada compania) de las companias de "
        "esa capa dentro del universo de 46 nombres. Sirve para ver que eslabon de la cadena de IA esta "
        "liderando o rezagando, mas alla del performance agregado de un solo fondo/ETF."
    )
    if not ciclo_ia:
        st.warning(f"No se pudo leer la hoja {SHEET_PRECIOS_CICLO_IA} de {CONSOLIDATED_XLSX_NAME}.")
    else:
        capas_presentes = sorted({v["capa"] for v in ciclo_ia.values()})
        layer_indices = {capa: ai_cycle_layer_index(ciclo_ia, capa) for capa in capas_presentes}
        layer_indices = {k: v for k, v in layer_indices.items() if v is not None}

        default_capas = [c for c in capas_presentes if c != "All"] or capas_presentes
        capas_choice = st.multiselect("Capas a mostrar en el chart (desmarca para comparar un subset)",
                                       capas_presentes, default=default_capas, key="ffmm_capas_multiselect")
        if capas_choice:
            st.plotly_chart(
                plot_multi({c: layer_indices[c] for c in capas_choice if c in layer_indices},
                           "Las 12 capas del Ciclo de Valor de IA — indice equal-weight (base 100)"),
                width="stretch",
            )

        period_map_capas = {"WTD": "WTD", "MTD": "MTD", "YTD": "YTD", "1Y": "1Y"}
        period_choice_capas = st.radio("Periodo para el ranking de capas", list(period_map_capas.keys()),
                                        horizontal=True, key="ffmm_capas_period")
        layer_returns = {capa: period_return(idx, period_choice_capas) for capa, idx in layer_indices.items()}
        layer_returns_valid = {k: v for k, v in layer_returns.items() if v is not None}
        if layer_returns_valid:
            ranked = pd.DataFrame(
                {"Capa": list(layer_returns_valid.keys()), f"Retorno {period_choice_capas} %": list(layer_returns_valid.values())}
            ).sort_values(f"Retorno {period_choice_capas} %", ascending=True)
            fig = go.Figure(go.Bar(
                x=ranked[f"Retorno {period_choice_capas} %"], y=ranked["Capa"], orientation="h",
                marker_color=[_sign_color(v) for v in ranked[f"Retorno {period_choice_capas} %"]],
                text=[f"{v:+.1f}%" for v in ranked[f"Retorno {period_choice_capas} %"]], textposition="outside",
            ))
            fig.update_layout(title=f"Ranking de capas — {period_choice_capas}",
                               height=max(360, 32 * len(ranked)), template="plotly_white",
                               margin=dict(l=10, r=60, t=40, b=10), xaxis_title="%")
            st.plotly_chart(fig, width="stretch")
            st.info(interpret_ai_cycle_layers(period_choice_capas, layer_returns_valid))

        with st.expander("📘 Glosario: que es cada una de las 12 capas del Ciclo de Valor de IA"):
            for capa in ["Compute", "Semicap Equipment", "Analog / Power Management", "Networking", "Datacenter",
                         "Cloud", "Applications", "Power", "Grid", "Cooling", "Electrical Equipment", "All"]:
                definicion = AI_CYCLE_LAYER_DEFS.get(capa)
                if definicion:
                    st.markdown(f"**{capa}** — {definicion}")

    # --- 6) Diversificacion: correlaciones + riesgo/retorno ---
    st.markdown("---")
    st.markdown("##### Diversificacion: correlaciones y riesgo/retorno")
    corr_universe = {n: resolved[n][0].series for n in resolved}
    if len(corr_universe) >= 3:
        corr_df = correlation_matrix(corr_universe)
        if not corr_df.empty:
            st.plotly_chart(plot_heatmap(corr_df, "Correlacion de retornos diarios (todo el universo resuelto)"),
                             width="stretch")
            st.caption("Una correlacion alta entre dos fondos del portafolio implica poca diversificacion efectiva "
                       "entre ellos, incluso si estan en categorias nominalmente distintas.")

    rr_rows = []
    for name, (res_fund, met, p) in resolved.items():
        if met.get("vol_1y_%") is not None and met.get("r1y_%") is not None:
            rr_rows.append({"Nombre": name, "vol_1y_%": met["vol_1y_%"], "r1y_%": met["r1y_%"],
                             "cat": p.get("Mercado", "")})
    if len(rr_rows) >= 3:
        st.plotly_chart(plot_risk_return_scatter(pd.DataFrame(rr_rows), "Riesgo (vol. 1Y) vs. Retorno (1Y)"),
                         width="stretch")

    st.caption(
        "Aviso: las lecturas de esta pestana son generadas por reglas cuantitativas deterministas a partir del "
        "Excel consolidado, no son un consejo de inversion. La atribucion por holdings es una aproximacion "
        "equal-weight con el universo del Ciclo de Valor de IA como proxy -no son los holdings oficiales ni sus "
        "pesos reales-. Verificar siempre contra la ficha oficial del fondo."
    )


# ==============================================================================
# 8. MAIN
# ==============================================================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon="🌐")
    st.title("🌐 " + APP_TITLE)
    st.caption("Terminal personal de research Top-Down · Market Monitoring + Macro Research + "
               "Cross-Asset Analysis + Investment Interpretation. Datos gratuitos, publicos, con cache local.")

    with st.sidebar:
        st.header("Configuracion")

        # --- UNICO input local de toda la app: BD_Consolidada_Dash_Macro.xlsx.
        #     Patron repo-primero con upload manual como override de sesion
        #     (ver docstring del modulo: "COMO SUBIR A GITHUB"). Alimenta
        #     Precios_ZODA (equity Peru + benchmarks), FFMM_Precios_Retornos +
        #     Parametria_FFMM, Precios_ZODA_CicloIA, Produccion_Mineria e
        #     Indicadores_SBS_Bancos -ya no se usa Yahoo Finance para ninguno
        #     de estos bloques, solo para FX, macro FRED/World Bank y los
        #     pares mineros internacionales (SCCO/HBM) en USA/WORLD/PERU. ---
        st.caption(f"Base de datos consolidada ({CONSOLIDATED_XLSX_NAME})")
        repo_path = _resolve_data_file(CONSOLIDATED_XLSX_NAME)
        uploaded = st.file_uploader(
            f"Subir {CONSOLIDATED_XLSX_NAME} para esta sesion (opcional"
            + (" — ya se detecto uno en el repo, ver abajo)" if repo_path else ")"),
            type=["xlsx"], key="consolidado_uploader")
        consolidated_source, source_label = resolve_consolidated_source(uploaded)

        if consolidated_source is not None:
            st.success(f"Excel consolidado {source_label}.")
        else:
            st.error(f"No se encontro {CONSOLIDATED_XLSX_NAME} en el repo ni se subio ninguno. Las pestanas "
                     "PERU y FFMM no tendran datos locales disponibles (COMMODITIES/USA/WORLD si funcionan, "
                     "porque dependen solo de fuentes externas gratuitas).")

        precios_zoda = load_precios_zoda(consolidated_source)
        ffmm_precios = load_ffmm_precios(consolidated_source)
        parametria_ffmm = load_parametria_ffmm(consolidated_source)
        ciclo_ia = load_ciclo_ia(consolidated_source)
        produccion_mineria = load_produccion_mineria(consolidated_source)
        sbs_bancos = load_sbs_bancos(consolidated_source)

        if consolidated_source is not None:
            st.caption(f"{len(precios_zoda)} instrumentos locales (equity Peru + benchmarks) · "
                       f"{len(ffmm_precios)}/{len(parametria_ffmm)} fondos FFMM con precio · "
                       f"{len(ciclo_ia)} companias en el Ciclo de Valor de IA.")

            # --- Diagnostico accionable: si CUALQUIERA de las hojas clave no
            #     trajo datos, no dejar al usuario adivinando por que -esto
            #     cubre tanto "detectado en el repo" pero con hojas rotas
            #     (ej. Git LFS, nombre de hoja distinto) como fallas
            #     parciales en una sola hoja. ---
            loaded_rows = {
                SHEET_PRECIOS_ZODA: sum(len(s) for s in precios_zoda.values()),
                SHEET_FFMM_PRECIOS: sum(len(s) for s in ffmm_precios.values()),
                SHEET_PARAMETRIA_FFMM: len(parametria_ffmm),
                SHEET_PRECIOS_CICLO_IA: sum(len(v["series"]) for v in ciclo_ia.values()),
                SHEET_PRODUCCION_MINERIA: len(produccion_mineria),
                SHEET_SBS_BANCOS: len(sbs_bancos),
            }
            if any(n == 0 for n in loaded_rows.values()):
                st.warning("Una o mas hojas del Excel consolidado no trajeron datos. Ver diagnostico abajo.")
                with st.expander("🔎 Diagnostico detallado del Excel consolidado", expanded=True):
                    st.code(_consolidated_file_diagnostics(consolidated_source, loaded_rows), language=None)

        st.markdown("---")
        if st.button("🔄 Forzar refresco de datos (limpiar cache)"):
            st.cache_data.clear()
            st.rerun()

        st.markdown("---")
        with st.expander("Limitaciones de datos conocidas"):
            st.markdown(
                "- Toda la data local (equity Peru, FFMM, Ciclo de Valor de IA, produccion minera, indicadores "
                "SBS) proviene de un unico archivo -BD_Consolidada_Dash_Macro.xlsx-; si un instrumento no tiene "
                "ISIN/Ticker en su hoja de parametria, o su ISIN no matcheo contra la base de precios de origen, "
                "se muestra 'N/D' en vez de inventar un dato.\n"
                "- Los benchmarks regionales (S&P/BVL Peru, IPSA, IBOVESPA, MEXBOL, COLCAP, MSCI Latam) son "
                "indices REALES provistos en el Excel, ya no proxies de ETF.\n"
                "- El performance attribution de FFMM usa, cuando el fondo es de tecnologia/IA (SubCat2 en "
                "Semiconductors/Core/Software & Internet/Thematic), un universo PROXY equal-weight de "
                "Precios_ZODA_CicloIA -NO son los holdings oficiales del fondo ni sus pesos reales-. Para el "
                "resto de fondos (regionales, de commodities, etc.) no se construye ninguna atribucion, para no "
                "fabricar un proxy sin fundamento.\n"
                "- La produccion minera se agrega por empresa+mineral principal restringiendo a Etapa="
                "'CONCENTRACION'/Proceso='FLOTACION' (la salida comparable entre plantas de una misma empresa); "
                "otras etapas (lixiviacion, refinacion) no se incluyen para evitar mezclar unidades de medida.\n"
                "- Los indicadores SBS (Morosidad/ROE/Liquidez) del Excel vienen como fraccion y se convierten a "
                "puntos porcentuales (x100) en el punto de entrada de los datos.\n"
                "- Las interpretaciones son generadas por reglas cuantitativas deterministas, no por un "
                "modelo de lenguaje; no deben tomarse como recomendacion de inversion."
            )
        st.caption(f"Ultima recarga de la app: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    tab_commodities, tab_peru, tab_usa, tab_world, tab_ffmm = st.tabs(
        ["🛢️ COMMODITIES", "🇵🇪 PERU", "🇺🇸 USA", "🌍 WORLD", "📊 FFMM"]
    )

    with tab_commodities:
        render_commodities_tab()
    with tab_peru:
        render_peru_tab(precios_zoda, produccion_mineria, sbs_bancos, source_label)
    with tab_usa:
        render_usa_tab()
    with tab_world:
        render_world_tab()
    with tab_ffmm:
        render_ffmm_tab(ffmm_precios, parametria_ffmm, ciclo_ia, source_label)

    st.markdown("---")
    st.caption(
        "Aviso: esta herramienta es de uso personal/educativo para desarrollar criterio de inversion Top-Down. "
        "No constituye recomendacion de inversion. Verificar siempre cifras criticas contra la fuente oficial "
        "antes de tomar decisiones de portafolio."
    )


if __name__ == "__main__":
    main()

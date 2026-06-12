"""
extract_browser.py — Automatización: navega URLs, extrae digitaldata + AA beacon, escribe Excel.

Workflow:
  1. URLs desde --urls (JSON) o --input (Excel plano)
  2. Procesa URLs concurrentemente (--workers N):
     - Extrae window.digitaldata (data layer)
     - Captura TODOS los beacons AA (s.t. + s.tl.)
     - Acumula, parsea, escribe
  3. Multi-sheet historial (recomendado, con --urls):
     - Cada corrida crea un sheet con fecha (YYYY-MM-DD)
     - Sheet _control registra metadata de cada auditoria
     - Archivos previos NUNCA se modifican
  4. Guardado incremental cada 5 URLs + guardado final
  5. Muestra metricas al final

Uso:
  # RECOMENDADO: historial multi-sheet
  python extract_browser.py --urls urls.json --output historial.xlsx
  python extract_browser.py --urls urls.json --workers 3

  # Clasico: Excel plano existente
  python extract_browser.py                                       # RevisionManual.xlsx
  python extract_browser.py --input otro.xlsx --output res.xlsx
  python extract_browser.py --row 3 --headed
  python extract_browser.py --resume --discard-cookies
  python extract_browser.py --log-file audit.log --proxy http://proxy:8080
  python extract_browser.py --run-clean

Requiere:
  pip install playwright openpyxl
  python -m playwright install chromium
"""

import asyncio
import json
import os
import re
import sys
import time
import subprocess
import logging
from copy import copy
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs

import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright, TimeoutError as PwTimeout

INPUT_FILE = "RevisionManual.xlsx"

# ── Helpers de formato ──
def _pretty_json(obj):
    return json.dumps(obj, indent=2, ensure_ascii=False)

def _set_col_widths(ws):
    for col, w in [("A", 15), ("B", 15), ("D", 80), ("E", 100), ("F", 80), ("G", 60)]:
        ws.column_dimensions[col].width = w

def _write_cell(ws, row, col, value, wrap=True):
    cell = ws.cell(row, col)
    cell.value = value
    cell.alignment = Alignment(wrap_text=True, vertical="top")
SAVE_EVERY_N = 5  # guardado incremental cada N URLs

AA_DOMAINS = [
    "smetrics.ford.com", "sc.omtrdc.net", "smetrics.omtrdc.net",
    "2o7.net", "data.adobedc.net", "edge.adobedc.net",
]

DATA_LAYER_NAMES = [
    "window.digitaldata", "window.dataLayer",
    "window.digital_data", "window.utag_data",
]


# ═══════════════════════════════════════════════════════════════════════════
# PARSE (100% pura, testeable sin navegador)
# ═══════════════════════════════════════════════════════════════════════════

def parse_aa_beacon(beacon_url: str, page_title: str = "") -> dict:
    """Parsea URL de beacon de Adobe Analytics a JSON estructurado."""
    parsed = urlparse(beacon_url)
    qs = parse_qs(parsed.query)

    def first(key):
        vals = qs.get(key, [])
        return vals[0] if vals else ""

    path_parts = parsed.path.split("/")
    report_suite = ""
    hit_id = ""
    if len(path_parts) >= 4 and path_parts[1] == "b" and path_parts[2] == "ss":
        report_suite = path_parts[3]
        hit_id = path_parts[-1] if len(path_parts) > 4 else ""

    props, evars = {}, {}
    for key, val in qs.items():
        m = re.match(r"^c(\d+)$", key)
        if m:
            props[f"prop{m.group(1)}"] = val[0]
            continue
        m = re.match(r"^v(\d+)$", key)
        if m:
            evars[f"eVar{m.group(1)}"] = val[0]

    events_raw = first("events")
    events = [e.strip() for e in events_raw.split(",") if e.strip()]

    visitor = {}
    mid = first("mid")
    if mid:
        visitor["experienceCloudId"] = mid
    aamlh = first("aamlh")
    if aamlh:
        visitor["audienceManagerHint"] = aamlh

    browser = {}
    res = first("res")
    if res:
        browser["resolution"] = res
    for k, name in [("bw", "browserWidth"), ("bh", "browserHeight")]:
        v = first(k)
        if v and v.isdigit():
            browser[name] = int(v)
    cd = first("cd")
    if cd:
        browser["colorDepth"] = cd
    ce = first("ce")
    if ce:
        browser["charset"] = ce

    result = {
        "solution": "analytics",
        "page": {"title": page_title or "", "url": first("g")},
        "request": {
            "method": "GET",
            "hostname": parsed.hostname,
            "pathname": parsed.path,
        },
        "visitor": visitor,
        "hit": {"id": hit_id, "type": "pageView", "reportSuiteId": report_suite},
        "browser": browser,
        "events": events,
        "eVars": evars,
        "props": props,
        "pageName": first("pageName"),
        "channel": first("ch"),
    }

    ts = first("t") or first("ts")
    if ts:
        result["request"]["collectedTimestamp"] = ts
    products_raw = first("products")
    if products_raw:
        result["products"] = products_raw
    return result


def build_aa_from_s(s_obj: dict, page_title: str = "") -> dict:
    """Convierte window.s a JSON estructurado."""
    props, evars = {}, {}
    for key, val in s_obj.items():
        m = re.match(r"^prop(\d+)$", key, re.IGNORECASE)
        if m:
            props[f"prop{m.group(1)}"] = val
        m = re.match(r"^eVar(\d+)$", key, re.IGNORECASE)
        if m:
            evars[f"eVar{m.group(1)}"] = val
    events_raw = s_obj.get("events", "")
    return {
        "solution": "analytics",
        "page": {"title": page_title, "url": s_obj.get("pageURL", "")},
        "pageName": s_obj.get("pageName", ""),
        "request": {"source": "window.s"},
        "events": [e.strip() for e in events_raw.split(",") if e.strip()],
        "eVars": evars, "props": props,
        "channel": s_obj.get("channel", ""),
        "products": s_obj.get("products", ""),
    }


# ═══════════════════════════════════════════════════════════════════════════
# EXTRACT (async, dependen de page de Playwright)
# ═══════════════════════════════════════════════════════════════════════════

async def extract_s_object(page) -> dict | None:
    """Lee window.s desde el navegador."""
    try:
        s_obj = await page.evaluate("""() => {
            const s = window.s || (window.s_c_il && window.s_c_il[window.s_c_il.length-1]);
            if (!s) return null;
            const vars = {};
            const KEYS = ['pageName','pageURL','channel','server','pageType',
                'events','products','linkTrackVars','linkTrackEvents',
                'charSet','visitorID','currencyCode','transactionID',
            ];
            for (let i = 1; i <= 75; i++) { KEYS.push('prop'+i, 'eVar'+i); }
            for (const p of KEYS) {
                if (s[p] !== undefined && s[p] !== '') vars[p] = s[p];
            }
            return Object.keys(vars).length > 0 ? vars : null;
        }""")
        return s_obj
    except Exception as e:
        logging.debug("window.s extraction failed: %s", e)
        return None


async def extract_digital_data(page) -> dict | None:
    """Extrae data layer probando varios nombres."""
    for var_name in DATA_LAYER_NAMES:
        try:
            dd = await page.evaluate(var_name)
            if dd and isinstance(dd, dict) and len(dd) > 0:
                return dd
        except Exception as e:
            logging.debug("data layer '%s' failed: %s", var_name, e)
    return None


async def extract_title(page) -> str:
    try:
        return (await page.evaluate("document.title") or "").strip()
    except Exception as e:
        logging.debug("title failed: %s", e)
        return ""


async def try_dismiss_cookie_consent(page):
    """Intenta cerrar banners de consentimiento comunes."""
    selectors = [
        "button:has-text('Aceptar')", "button:has-text('Accept')",
        "button:has-text('Aceptar todas')", "button:has-text('Accept All')",
        "#onetrust-accept-btn-handler", ".cookie-accept", ".cc-accept",
        "[aria-label='Accept cookies']", "#cookiescript_accept",
    ]
    for sel in selectors:
        try:
            btn = await page.query_selector(sel)
            if btn:
                await btn.click()
                logging.debug("Cookie consent dismissed: %s", sel)
                return True
        except Exception:
            continue
    return False


# ═══════════════════════════════════════════════════════════════════════════
# EXCEL HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def validate_sheet(ws) -> list:
    errores = []
    header = str(ws.cell(1, 2).value or "").strip().lower()  # col B
    if "pagina" not in header:
        errores.append("Col B debe tener header 'pagina auditada'")
    has_urls = any(ws.cell(row, 2).value for row in range(2, ws.max_row + 1))
    if not has_urls:
        errores.append("No hay URLs en col B")
    return errores


def save_workbook(wb, path):
    """Guarda con fallback si el archivo está bloqueado."""
    try:
        wb.save(path)
        return path
    except PermissionError:
        name, ext = os.path.splitext(path)
        fallback = f"{name}_browser{ext}"
        wb.save(fallback)
        logging.warning("Archivo bloqueado, guardado como %s", fallback)
        return fallback


# ═══════════════════════════════════════════════════════════════════════════
# WORKER — procesa UNA URL, devuelve resultado
# ═══════════════════════════════════════════════════════════════════════════

async def process_url(
    page, row: int, url: str,
    discard_cookies: bool = False,
    timeout_ms: int = 35000,
    max_retry: int = 1,
) -> dict:
    """
    Navega a una URL, captura beacons + data layer.
    Devuelve dict con resultados.
    Cada worker tiene su propia page + handler de beacons.
    """
    beacon_urls = []

    def on_beacon(request):
        url_lower = request.url.lower()
        for domain in AA_DOMAINS:
            if domain in url_lower and "/b/ss/" in url_lower:
                beacon_urls.append(request.url)
                return

    page.on("request", on_beacon)

    result = {"row": row, "url": url, "status": 0, "error": None,
              "title": "", "digitaldata": None, "aa_parsed": None,
              "aa_source": None, "extra_beacons": [], "elapsed_s": 0.0,
              "retries_used": 0}

    t0 = time.time()

    for attempt in range(1 + max_retry):
        try:
            resp = await page.goto(url, wait_until="load", timeout=timeout_ms)
            await asyncio.sleep(2)
            if resp:
                result["status"] = resp.status

            if discard_cookies:
                await try_dismiss_cookie_consent(page)

            result["title"] = await extract_title(page)
            result["digitaldata"] = await extract_digital_data(page)
            break  # éxito

        except PwTimeout:
            result["status"] = -1
            result["error"] = "timeout"
        except Exception as e:
            result["status"] = -2
            result["error"] = str(e)[:120]
            logging.error("Error navegando fila %d: %s", row, e)

        if attempt < max_retry and result["status"] in (-1, -2):
            wait = 5 * (attempt + 1)
            logging.warning("  Fila %d: retry %d en %ds", row, attempt + 1, wait)
            await asyncio.sleep(wait)
            result["retries_used"] += 1
        else:
            break

    # Procesar beacons capturados
    if beacon_urls:
        result["aa_parsed"] = parse_aa_beacon(beacon_urls[0], result["title"])
        result["aa_source"] = "beacon"
        if len(beacon_urls) > 1:
            result["extra_beacons"] = [parse_aa_beacon(u, result["title"])
                                       for u in beacon_urls[1:]]

    # Fallback a window.s
    if not result["aa_parsed"]:
        s_obj = await extract_s_object(page)
        if s_obj and s_obj.get("pageName"):
            result["aa_parsed"] = build_aa_from_s(s_obj, result["title"])
            result["aa_source"] = "window.s"

    if not result["aa_parsed"] and not result["error"]:
        result["error"] = "no AA data captured"

    result["elapsed_s"] = round(time.time() - t0, 1)
    return result


# ═══════════════════════════════════════════════════════════════════════════
# ESCRITURA EXCEL (protegida por lock)
# ═══════════════════════════════════════════════════════════════════════════

async def write_result(
    ws, result: dict, metrics: dict,
    excel_lock: asyncio.Lock, output_path: str,
    saved_count: list,  # [int] mutable para closure
    show_progress: bool = False,
    total_urls: int = 0,
    workers: int = 1,
) -> None:
    """
    Escribe el resultado de una URL en el Excel.
    Usa lock para evitar corrupción entre workers.
    """
    async with excel_lock:
        row = result["row"]
        n_beacons = 1 + len(result.get("extra_beacons", []))

        # digitaldata → col D (rojo: no relevante, sin fill)
        dd_val = result.get("digitaldata")
        if dd_val is not None:
            _write_cell(ws, row, 4, _pretty_json(dd_val))
            metrics["ok_dd"] += 1
        else:
            ws.cell(row, 4).value = "(no digitaldata)"

        # AA → col E (amarillo: media importancia, sin fill)
        if result.get("aa_parsed"):
            _write_cell(ws, row, 5, _pretty_json(result["aa_parsed"]))
            metrics["ok_aa"] += 1
        else:
            ws.cell(row, 5).value = f"({result.get('error', 'no AA')})"

        # Score por URL (0-100)
        url_score = compute_url_score(result)

        # Metadata → col G
        meta = {"score": url_score, "status": result["status"],
                "aa_source": result["aa_source"],
                "beacons": n_beacons, "title": result["title"],
                "error": result.get("error"), "elapsed_s": result["elapsed_s"],
                "url": result["url"][:120]}
        if result.get("extra_beacons"):
            meta["extra_beacons"] = result["extra_beacons"]
        _write_cell(ws, row, 7, _pretty_json(meta))

        if result.get("error") or not result["aa_parsed"]:
            metrics["errors"] += 1
            metrics["errores_detalle"].append({"row": row, "error": result.get("error", "no AA")})

        metrics["total_beacons"] += n_beacons
        metrics["times"].append(result["elapsed_s"])
        metrics["retries"] += result.get("retries_used", 0)

        # Guardado incremental
        saved_count[0] += 1
        if saved_count[0] % SAVE_EVERY_N == 0:
            save_workbook(ws.parent, output_path)
            logging.info("  Guardado incremental (#%d)", saved_count[0])
        if show_progress:
            print_progress(saved_count[0], total_urls, metrics["errors"], workers)


# ═══════════════════════════════════════════════════════════════════════════
# PROGRESS BAR
# ═══════════════════════════════════════════════════════════════════════════

def print_progress(done: int, total: int, errors: int, workers: int):
    """Barra de progreso simple (solo ASCII, funciona en Windows)."""
    pct = done / total * 100 if total else 0
    bar_len = 20
    filled = int(bar_len * done / total) if total else 0
    bar = "#" * filled + "." * (bar_len - filled)
    print(f"\r  [{done}/{total}] [{bar}] {pct:.0f}% | err:{errors} | w:{workers}   ", end="", flush=True)
    if done == total:
        print()  # nueva línea al final


# ═══════════════════════════════════════════════════════════════════════════
# ERROR CLASSIFICATION
# ═══════════════════════════════════════════════════════════════════════════

def classify_errors(errors_detail: list[dict]) -> dict[str, list[int]]:
    """Agrupa errores por categoría para output legible."""
    categories = {
        "HTTP 403 (acceso denegado)": [],
        "Timeout": [],
        "Sin dato AA (no beacon)": [],
        "Error de red/conexión": [],
    }
    for e in errors_detail:
        err = e.get("error", "")
        if "timeout" in err.lower():
            categories["Timeout"].append(e["row"])
        elif "403" in err:
            categories["HTTP 403 (acceso denegado)"].append(e["row"])
        elif "no aa" in err.lower() or "no AA" in err:
            categories["Sin dato AA (no beacon)"].append(e["row"])
        else:
            categories["Error de red/conexión"].append(e["row"])
    # Filtrar vacías
    return {k: v for k, v in categories.items() if v}


# ═══════════════════════════════════════════════════════════════════════════
# SCORE
# ═══════════════════════════════════════════════════════════════════════════

def compute_score(metrics: dict) -> int:
    success_rate = (metrics["ok_aa"] / max(metrics["total"], 1)) * 100
    dd_rate = (metrics["ok_dd"] / max(metrics["total"], 1)) * 100
    avg_time = sum(metrics["times"]) / max(len(metrics["times"]), 1)
    beacons_per_url = metrics["total_beacons"] / max(metrics["total"], 1)
    retry_efficiency = max(0, 1 - metrics["retries"] / max(metrics["total"], 1)) * 100
    return int(
        success_rate * 0.40 +
        dd_rate * 0.25 +
        min(avg_time, 60) / 60 * 100 * 0.15 +
        min(beacons_per_url, 3) / 3 * 100 * 0.10 +
        retry_efficiency * 0.10
    )


def compute_url_score(result: dict) -> int:
    """Score 0-100 por URL individual."""
    s = 0
    if result.get("digitaldata"):
        s += 30
    if result.get("aa_parsed"):
        s += 30
    if result.get("extra_beacons"):
        s += 10
    if not result.get("error") and result.get("status", 0) not in (-1, -2, 403):
        s += 20
    elapsed = result.get("elapsed_s", 99)
    if elapsed < 5:
        s += 10
    elif elapsed < 15:
        s += 5
    return s


# ═══════════════════════════════════════════════════════════════════════════
# MULTI-SHEET HISTORIAL
# ═══════════════════════════════════════════════════════════════════════════

SHEET_HEADERS = [
    "nombre pagina auditada",
    "pagina auditada (URL)",
    "digitaldata (manual)",
    "digitaldata (automatica)",
    "AA analytics (automatico)",
    "AA analytics (estructurado)",
    "metadata / extra beacons",
]

CONTROL_HEADERS = [
    "Fecha", "Source", "URLs", "AA OK", "DD OK", "Errores",
    "Reintentos", "Score", "Tiempo (s)", "Workers",
]

# Colores de columnas (A=azul claro estructura, C/F=verde, E=amarillo, D=rojo, G=sin color)
HEADER_FILLS = {
    "A": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # azul claro
    "B": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # azul claro
    "C": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # verde
    "D": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),  # rojo
    "E": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),  # amarillo
    "F": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # verde
}
# G: sin color

DATA_FILLS = {
    "C": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # verde
    "F": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # verde
}
# D/E/G: sin fill en datos


def setup_multisheet(output_path: str, urls_source: str, resume: bool) -> tuple:
    """
    Carga o crea workbook con sheets _control + fecha actual.
    Retorna (wb, ws, audit_date, skipped).
    skipped=True si resume y el sheet de hoy ya existe.
    """
    audit_date = datetime.now().strftime("%Y-%m-%d")

    if os.path.exists(output_path):
        wb = openpyxl.load_workbook(output_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws_ctrl = wb.create_sheet("_control")
        ws_ctrl.append(CONTROL_HEADERS)
        ws_ctrl.column_dimensions["A"].width = 14
        ws_ctrl.column_dimensions["B"].width = 30

    if audit_date in wb.sheetnames:
        if resume:
            logging.info("Sheet '%s' ya existe, saltando (--resume)", audit_date)
            return wb, None, audit_date, True
        wb.remove(wb[audit_date])
        logging.info("Reemplazando sheet existente: %s", audit_date)

    ws = wb.create_sheet(audit_date)
    ws.append(SHEET_HEADERS)
    # Aplicar colores a headers
    for col_letter, fill in HEADER_FILLS.items():
        col_idx = openpyxl.utils.column_index_from_string(col_letter)
        ws.cell(1, col_idx).fill = fill
    _set_col_widths(ws)
    return wb, ws, audit_date, False


def update_control(wb, audit_date: str, source: str, total: int,
                   ok_aa: int, ok_dd: int, errors: int, retries: int,
                   score: int, elapsed_s: float, workers: int):
    """Agrega fila al sheet _control con metadata de la corrida."""
    ws = wb["_control"]
    next_row = ws.max_row + 1
    vals = [audit_date, source, total, ok_aa, ok_dd, errors,
            retries, score, f"{elapsed_s:.0f}", workers]
    for col, val in enumerate(vals, start=1):
        ws.cell(row=next_row, column=col).value = val


def update_vars_sheet(wb, audit_date: str, rows_aa: list[tuple[int, dict]]):
    """Sheet _vars: rastrea eVars/props/events de AA entre corridas."""
    if "_vars" not in wb.sheetnames:
        vs = wb.create_sheet("_vars")
        vs.cell(1, 1).value = "Variable"
        vs.column_dimensions["A"].width = 20
        vs.cell(1, 2).value = "Tipo"
        vs.cell(1, 3).value = audit_date
        vs.column_dimensions["C"].width = 40
        next_row = 2
    else:
        vs = wb["_vars"]
        # Agregar columna si es nueva fecha
        existing_cols = [str(vs.cell(1, c).value or "") for c in range(1, vs.max_column + 1)]
        if audit_date not in existing_cols:
            new_col = vs.max_column + 1
            vs.cell(1, new_col).value = audit_date
            vs.column_dimensions[get_column_letter(new_col)].width = 40
        next_row = vs.max_row + 1

    seen = {}
    for _, aa in rows_aa:
        for key in ("eVars", "evars", "props", "events", "products"):
            val = aa.get(key)
            if val:
                if isinstance(val, dict):
                    for k in val:
                        seen[f"{key}.{k}"] = key
                elif isinstance(val, list):
                    for item in val:
                        seen[f"{key}.{item}"] = key
                elif isinstance(val, str):
                    seen[f"{key}"] = key

    for var_name, var_type in seen.items():
        # Buscar si ya existe
        found = False
        for r in range(2, vs.max_row + 1):
            if vs.cell(r, 1).value == var_name:
                found = True
                break
        if not found:
            vs.cell(next_row, 1).value = var_name
            vs.cell(next_row, 2).value = var_type
            next_row += 1


# ═══════════════════════════════════════════════════════════════════════════
# MAIN ASYNC
# ═══════════════════════════════════════════════════════════════════════════

async def amain():
    import argparse
    parser = argparse.ArgumentParser(description="Extrae digitaldata + AA beacon desde URLs en Excel")
    parser.add_argument("--row", type=int, help="Procesar solo una fila")
    parser.add_argument("--headed", action="store_true", help="Navegador visible")
    parser.add_argument("--input", default=INPUT_FILE, help="Archivo Excel de entrada")
    parser.add_argument("--output", help="Archivo Excel de salida (default: escribe sobre input)")
    parser.add_argument("--urls", help="Archivo JSON con URLs (crea Excel desde 0). Formato: [{'url':'...','nombre':'...'}, 'url_simple']")
    parser.add_argument("--timeout", type=int, default=35000, help="Timeout por pagina (ms)")
    parser.add_argument("--resume", action="store_true", help="Saltar filas con datos")
    parser.add_argument("--log-file", help="Archivo de log")
    parser.add_argument("--proxy", help="Proxy HTTP (ej: http://proxy:8080)")
    parser.add_argument("--retry", type=int, default=1, help="Reintentos por URL")
    parser.add_argument("--workers", type=int, default=1, help="URLs concurrentes (default: 1)")
    parser.add_argument("--discard-cookies", action="store_true", help="Rechazar banners de cookies")
    parser.add_argument("--progress", action="store_true", help="Mostrar barra de progreso")
    parser.add_argument("--diff", action="store_true", help="Mostrar diferencias entre ultima y penultima auditoria")
    parser.add_argument("--config", help="Archivo JSON con config (workers, proxy, retry, etc)")
    parser.add_argument("--retry-failed", action="store_true", help="Solo URLs con error en corridas anteriores")
    args = parser.parse_args()

    # ── Cargar config desde JSON (si existe) ──
    config_file = args.config or "audit.json"
    if os.path.exists(config_file):
        with open(config_file, encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in cfg.items():
            if not getattr(args, k, None):
                setattr(args, k, v)
        logging.info("Config cargada: %s", config_file)

    if args.row:
        args.workers = 1

    # ── Logging ──
    log_handlers = [logging.StreamHandler(sys.stderr)]
    if args.log_file:
        log_handlers.append(logging.FileHandler(args.log_file, encoding="utf-8"))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                        handlers=log_handlers)

    # ── --diff: comparar ultimas 2 auditorias (sin procesar) ──
    if args.diff and not args.urls:
        diff_path = args.output or "historial.xlsx"
        if not os.path.exists(diff_path):
            logging.error("No existe %s para comparar", diff_path)
            return 1
        wb = openpyxl.load_workbook(diff_path)
        sheets = [s for s in wb.sheetnames if s != "_control"]
        sheets.sort(key=lambda s: s if "-" in s else "0", reverse=True)  # YYYY-MM-DD sort
        if len(sheets) < 2:
            logging.info("Se necesitan al menos 2 auditorias para --diff")
            wb.close()
            return 0
        print(f"\n  DIFERENCIAS: {sheets[1]} → {sheets[0]}")
        print(f"  {'='*50}")
        sa, sb = wb[sheets[0]], wb[sheets[1]]
        for row in range(2, max(sa.max_row or 2, sb.max_row or 2) + 1):
            url_a = sa.cell(row, 2).value or ""
            url_b = sb.cell(row, 2).value or ""
            aa_a = (sa.cell(row, 5).value or "")[:50]
            aa_b = (sb.cell(row, 5).value or "")[:50]
            dd_a = (sa.cell(row, 4).value or "")[:50]
            dd_b = (sb.cell(row, 4).value or "")[:50]
            if str(aa_a).strip() != str(aa_b).strip() or str(dd_a).strip() != str(dd_b).strip():
                print(f"  Fila {row}: {url_a or url_b}")
                if str(aa_a).strip() != str(aa_b).strip():
                    print(f"    AA: {'ANTES' if not sa.cell(row,5).value else 'NUEVO'}")
                if str(dd_a).strip() != str(dd_b).strip():
                    print(f"    DD: {'ANTES' if not sa.cell(row,4).value else 'NUEVO'}")
        wb.close()
        return 0

    # ── Excel: --urls → multi-sheet historial | --input → clasico ──
    audit_date = None  # flag: multi-sheet mode
    if args.urls:
        output_path = args.output or "historial.xlsx"
        wb, ws, audit_date, skipped = setup_multisheet(output_path, args.urls, args.resume)
        if skipped:
            logging.info("Sheet de hoy ya auditado. Usá --diff para ver diferencias.")
            return 0
        with open(args.urls, encoding="utf-8") as f:
            entries = json.load(f)
        for i, entry in enumerate(entries, start=2):
            if isinstance(entry, str):
                name = url = entry
            elif isinstance(entry, dict):
                name = entry.get("nombre", entry.get("url", ""))
                url = entry.get("url", "")
            wrap = Alignment(wrap_text=True, vertical="top")
            ws.cell(i, 1).value = name; ws.cell(i, 1).alignment = wrap
            ws.cell(i, 2).value = url; ws.cell(i, 2).alignment = wrap
        logging.info("Multi-sheet %s: %d URLs", audit_date, len(entries))
    else:
        output_path = args.output or args.input
        # ── Cargar Excel existente ──
        wb = openpyxl.load_workbook(args.input)
        ws = wb.active

    if args.output and not args.urls and not audit_date:
        # Copiamos output sin mutar input (solo modo clasico)
        wb_copy = openpyxl.Workbook()
        wb_copy.remove(wb_copy.active)
        for sheet_name in wb.sheetnames:
            src = wb[sheet_name]
            dst = wb_copy.create_sheet(title=sheet_name)
            for row in src.iter_rows():
                for cell in row:
                    dst_cell = dst.cell(row=cell.row, column=cell.column)
                    dst_cell.value = cell.value
                    if cell.has_style:
                        dst_cell.font = copy(cell.font)
                        dst_cell.alignment = copy(cell.alignment)
                        dst_cell.border = copy(cell.border)
                        dst_cell.fill = copy(cell.fill)
                        dst_cell.number_format = copy(cell.number_format)
                        dst_cell.protection = copy(cell.protection)
        ws = wb_copy.active
        wb = wb_copy

    errs = validate_sheet(ws)
    if errs:
        for e in errs:
            logging.error("Validacion: %s", e)
        wb.close()
        sys.exit(1)

    ws.cell(1, 3).value = ws.cell(1, 3).value or "digitaldata (manual)"
    ws.cell(1, 4).value = ws.cell(1, 4).value or "digitaldata (automatica)"
    ws.cell(1, 7).value = ws.cell(1, 7).value or "metadata / extra beacons"

    rows_to_process = []
    for row in range(2, ws.max_row + 1):
        url = ws.cell(row, 2).value
        if not url:
            continue
        if args.row and row != args.row:
            continue
        if args.resume:
            existing = ws.cell(row, 5).value
            if existing and len(str(existing).strip()) > 30 and "no AA" not in str(existing):
                logging.info("Fila %d: saltando (resume)", row)
                continue
        if args.retry_failed:
            existing = ws.cell(row, 5).value
            if not existing or str(existing).strip().startswith("("):
                logging.info("Fila %d: reintentar (error previo)", row)
            else:
                continue  # saltear las que ya tienen datos OK
        rows_to_process.append((row, str(url).strip()))

    if not rows_to_process:
        logging.info("No hay URLs para procesar.")
        wb.close()
        return

    logging.info("URLs a procesar: %d | Workers: %d", len(rows_to_process), args.workers)

    metrics = {"total": len(rows_to_process), "ok_aa": 0, "ok_dd": 0,
               "errors": 0, "retries": 0, "total_beacons": 0, "times": [],
               "errores_detalle": []}
    aa_data = []  # (row, aa_parsed) para _vars
    start_time = time.time()
    excel_lock = asyncio.Lock()
    saved_count = [0]

    # ── Browser ──
    async with async_playwright() as pw:
        launch_kw = {"headless": not args.headed}
        if args.proxy:
            launch_kw["proxy"] = {"server": args.proxy}

        browser = await pw.chromium.launch(**launch_kw)

        sem = asyncio.Semaphore(args.workers)

        async def worker(row: int, url: str):
            """Crea su propio context + page, procesa, escribe."""
            ctx = await browser.new_context(
                viewport={"width": 1280, "height": 720},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36",
            )
            page = await ctx.new_page()
            try:
                result = await process_url(
                    page, row, url,
                    discard_cookies=args.discard_cookies,
                    timeout_ms=args.timeout,
                    max_retry=args.retry,
                )
                status_str = f"HTTP {result['status']}" if result['status'] > 0 else f"ERR {result['status']}"
                logging.info("Fila %d | %s | AA: %s | %.1fs",
                             row, status_str,
                             result["aa_source"] or "none",
                             result["elapsed_s"])
                await write_result(ws, result, metrics, excel_lock, output_path, saved_count,
                                     show_progress=args.progress,
                                     total_urls=len(rows_to_process),
                                     workers=args.workers)
                if result.get("aa_parsed"):
                    aa_data.append((row, result["aa_parsed"]))
            finally:
                await page.close()
                await ctx.close()

        async def worker_sem(row: int, url: str):
            async with sem:
                await worker(row, url)

        tasks = [worker_sem(row, url) for row, url in rows_to_process]
        await asyncio.gather(*tasks)

        await browser.close()

    # Actualizar _vars con datos AA recolectados
    if audit_date and aa_data:
        update_vars_sheet(wb, audit_date, aa_data)

    # ── Guardado final + _control (multi-sheet) ──
    _set_col_widths(ws)
    total_time = time.time() - start_time
    score = compute_score(metrics)
    if audit_date:
        update_control(wb, audit_date, args.urls or "",
                       metrics["total"],
                       metrics["ok_aa"], metrics["ok_dd"],
                       metrics["errors"], metrics["retries"],
                       score, total_time, args.workers)
    out = save_workbook(wb, output_path)
    wb.close()
    logging.info("Guardado: %s", out)

    # ── Métricas ──
    success_rate = metrics["ok_aa"] / max(metrics["total"], 1) * 100
    dd_rate = metrics["ok_dd"] / max(metrics["total"], 1) * 100
    avg_time = sum(metrics["times"]) / max(len(metrics["times"]), 1)
    beacons_per_url = metrics["total_beacons"] / max(metrics["total"], 1)

    sep = "=" * 55
    dash = "-" * 55
    print(f"""
{sep}
  METRICAS Y SCORE
{sep}
  Config:            {args.workers} worker(s) concurrente(s)
  URLs procesadas:   {metrics['ok_aa']}/{metrics['total']}
  AA capturados:     {success_rate:.0f}%
  digitaldata:       {dd_rate:.0f}%
  Beacons totales:   {metrics['total_beacons']} ({beacons_per_url:.1f}/URL)
  Reintentos:        {metrics['retries']}
  Errores:           {metrics['errors']}
{dash}
  Tiempo total:      {timedelta(seconds=int(total_time))}
  Promedio/URL:      {avg_time:.1f}s
  Guardados incr.:   cada {SAVE_EVERY_N} URLs
{dash}
  SCORE GLOBAL:      {score}/100
{dash}""")

    if metrics["errores_detalle"]:
        print(f"  DETALLE ERRORES ({len(metrics['errores_detalle'])}):")
        categories = classify_errors(metrics["errores_detalle"])
        for cat, rows in categories.items():
            rows_str = ", ".join(str(r) for r in rows)
            print(f"    * {cat}: Fila(s) {rows_str}")

    if score < 60:
        print("  Sugerencias:")
        if success_rate < 50:
            print("    * Verificar VPN / acceso a las URLs")
        if dd_rate < 50:
            print("    * El data layer podria no llamarse window.digitaldata")
        if avg_time > 30:
            print("    * Paginas lentas. Aumentar --timeout o verificar SPAs")

    if os.path.exists(output_path):
        print(f"\n{'='*55}")
        print("  Limpiando AA en Excel...")
        print(f"{'='*55}")
        clean_script = os.path.join(os.path.dirname(__file__) or ".", "extract_aa.py")
        cmd = [sys.executable, clean_script, "--input", output_path]
        if audit_date:
            cmd += ["--sheet", audit_date]
        r = subprocess.run(cmd, capture_output=True, text=True)
        print(r.stdout)
        if r.stderr:
            print(r.stderr, file=sys.stderr)

    return score


def main():
    asyncio.run(amain())


if __name__ == "__main__":
    main()

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
  python extract_browser.py --resume --wait-after 5
  python extract_browser.py --log-file audit.log --proxy http://proxy:8080
  python extract_browser.py --run-clean
  python extract_browser.py --verbose                               # logging debug

Requiere:
  pip install playwright openpyxl
  python -m playwright install chromium
"""

import asyncio
import json
import logging
import os
import re
import signal
import subprocess
import sys
import time
from copy import copy
from datetime import datetime, timedelta
from urllib.parse import parse_qs, urlparse

# ── Graceful shutdown ──
_shutdown_flag = False
"""Global flag: set by signal handler to request graceful shutdown."""

def _request_shutdown(signum=None, frame=None):
    global _shutdown_flag
    if not _shutdown_flag:
        _shutdown_flag = True
        logging.warning("Señal %s recibida. Terminando gracefulmente...", signum)

# ── Validation & sanitization ──

VALID_URL_SCHEMES = ("http", "https")
"""Schemes permitidos para navegación."""

# Dominios conocidos del proyecto Ford preview + Adobe
ALLOWED_HOSTNAME_SUFFIXES = (
    ".ford.com",
    ".brandpr.ford.com",
    ".ford.mx",
    ".ford.com.pr",
    ".omtrdc.net",
    ".adobedc.net",
    "2o7.net",
)

def validate_url(url: str) -> str | None:
    """Valida URL antes de navegar. Retorna mensaje de error o None si OK."""
    if not url or not isinstance(url, str):
        return "URL vacía o inválida"
    url = url.strip()
    if not url:
        return "URL vacía después de trim"
    try:
        parsed = urlparse(url)
    except Exception as e:
        return f"URL no parseable: {e}"
    if parsed.scheme not in VALID_URL_SCHEMES:
        return f"Scheme '{parsed.scheme}' no permitido (solo http/https)"
    if not parsed.netloc:
        return "URL sin hostname"
    # Chequeo de SSRF básico: solo dominios conocidos
    hostname = parsed.netloc.lower()
    # Remover user:password@ si existe
    if "@" in hostname:
        hostname = hostname.split("@")[-1]
    # Remover :puerto si existe
    if ":" in hostname:
        hostname = hostname.split(":")[0]
    if hostname in ("localhost", "127.0.0.1", "0.0.0.0", "::1"):
        return "URL apunta a localhost (posible SSRF)"
    if not hostname.endswith(ALLOWED_HOSTNAME_SUFFIXES):
        return f"Dominio '{hostname}' no está en la whitelist de proyectos"
    return None

def sanitize_url_for_log(url: str, max_len: int = 80) -> str:
    """Limpia URL para logging: trunca y redacta query params sensibles.
    
    Redacta valores de query params que podrían contener PII (email, token, etc).
    """
    if not url:
        return ""
    try:
        parsed = urlparse(url)
        if parsed.query:
            qs = parse_qs(parsed.query)
            sensitive_keys = {"email", "token", "key", "secret", "password", "pass", "auth"}
            cleaned = {k: ("[REDACTED]" if k.lower() in sensitive_keys else v[0][:60])
                       for k, v in qs.items()}
            base = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"[:max_len]
            return f"{base}?{cleaned}" if cleaned else base
        return url[:max_len]
    except Exception:
        return url[:max_len]

import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import TimeoutError as PwTimeout
from playwright.async_api import async_playwright

INPUT_FILE = "RevisionManual.xlsx"

# ── Helpers de formato ──
def _pretty_json(obj):
    return json.dumps(obj, indent=2, ensure_ascii=False)

def _set_col_widths(ws):
    for col, w in [("A", 15), ("B", 15), ("C", 80), ("D", 100), ("E", 80), ("F", 60)]:
        ws.column_dimensions[col].width = w


def _auto_row_height(ws):
    """Ajusta alto de filas segun el maximo de lineas JSON en cols 3-4-5-6."""
    JSON_COLS = [3, 4, 5, 6]
    LINE_HEIGHT = 15
    MAX_HEIGHT = 409
    for row in range(2, ws.max_row + 1):
        max_lines = 1
        for col in JSON_COLS:
            val = ws.cell(row, col).value
            if val:
                lines = str(val).count("\n") + 1
                max_lines = max(max_lines, lines)
        height = min(max_lines * LINE_HEIGHT, MAX_HEIGHT)
        current = ws.row_dimensions[row].height
        if current is None or height > current:
            ws.row_dimensions[row].height = height

def _write_cell(ws, row, col, value, wrap=True):
    cell = ws.cell(row, col)
    cell.value = value
    cell.number_format = "@"  # texto plano — evita que Excel interprete JSON como numero/fecha
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
        try:
            wb.save(fallback)
            logging.warning("Archivo bloqueado, guardado como %s", fallback)
            return fallback
        except PermissionError:
            # Segundo intento con retry
            import time
            for attempt in range(3):
                try:
                    time.sleep(2)
                    wb.save(path)
                    logging.info("Recuperado tras %ds de espera", (attempt + 1) * 2)
                    return path
                except PermissionError:
                    continue
            fallback2 = f"{name}_browser{int(time.time())}{ext}"
            wb.save(fallback2)
            logging.warning("Lock persistente, guardado como %s", fallback2)
            return fallback2


# ═══════════════════════════════════════════════════════════════════════════
# WORKER — procesa UNA URL, devuelve resultado
# ═══════════════════════════════════════════════════════════════════════════

async def process_url(
    page, row: int, url: str,
    wait_after: int = 4,
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

    result = {"row": row, "url": url, "status": 0, "error": None, "code": None,
              "title": "", "digitaldata": None, "aa_parsed": None,
              "aa_source": None, "extra_beacons": [], "elapsed_s": 0.0,
              "retries_used": 0}

    t0 = time.time()

    for attempt in range(1 + max_retry):
        # Limpiar beacons de intentos fallidos previos
        if attempt > 0:
            beacon_urls.clear()
        try:
            resp = await page.goto(url, wait_until="load", timeout=timeout_ms)
            await asyncio.sleep(wait_after)
            if resp:
                result["status"] = resp.status

            await try_dismiss_cookie_consent(page)

            result["title"] = await extract_title(page)
            result["digitaldata"] = await extract_digital_data(page)
            break  # éxito

        except PwTimeout:
            result["status"] = -1
            result["error"] = "timeout"
            result["code"] = "TIMEOUT"
        except Exception as e:
            result["status"] = -2
            result["error"] = str(e)[:120]
            result["code"] = _error_code_from_detail(str(e))
            logging.exception("Error navegando fila %d: %s", row, e)

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
        result["code"] = "NO_AA_DATA"

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

        # digitaldata → col C (rojo: no relevante, sin fill)
        dd_val = result.get("digitaldata")
        if dd_val is not None:
            _write_cell(ws, row, 3, _pretty_json(dd_val))
            metrics["ok_dd"] += 1
        else:
            _write_cell(ws, row, 3, _pretty_json({"error": "no digitaldata", "code": "DD_MISSING"}))

        # AA → col D (amarillo: media importancia, sin fill)
        if result.get("aa_parsed"):
            _write_cell(ws, row, 4, _pretty_json(result["aa_parsed"]))
            metrics["ok_aa"] += 1
        else:
            err_code = _error_code_from_detail(result.get("error", "no AA"))
            _write_cell(ws, row, 4, _pretty_json({"error": result.get("error", "no AA"), "code": err_code}))

        # Score por URL (0-100)
        url_score = compute_url_score(result)

        # Metadata → col F
        meta = {"score": url_score, "status": result["status"],
                "aa_source": result["aa_source"],
                "beacons": n_beacons, "title": result["title"],
                "error": result.get("error"), "code": result.get("code"),
                "elapsed_s": result["elapsed_s"],
                "url": result["url"][:120]}
        if result.get("extra_beacons"):
            meta["extra_beacons"] = result["extra_beacons"]
        _write_cell(ws, row, 6, _pretty_json(meta))

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

# ── Códigos de error estándar ──
ERROR_CODES = {
    "TIMEOUT":       "Tiempo de espera agotado al navegar",
    "HTTP_403":      "Acceso denegado (HTTP 403)",
    "HTTP_ERROR":    "Error HTTP al navegar",
    "NO_AA_DATA":    "No se capturó data de Adobe Analytics",
    "URL_INVALID":   "URL no válida o no permitida",
    "NETWORK_ERROR": "Error de red o conexión fallida",
    "NAV_ERROR":     "Error durante la navegación",
    "UNKNOWN":       "Error desconocido",
}

def _error_code_from_detail(err: str) -> str:
    """Asigna código de error estándar según el texto del error."""
    if not err:
        return "UNKNOWN"
    err_lower = err.lower()
    if "timeout" in err_lower:
        return "TIMEOUT"
    if "403" in err:
        return "HTTP_403"
    if "no aa" in err_lower or "no AA" in err:
        return "NO_AA_DATA"
    if "url_invalid" in err_lower or "url inv" in err_lower:
        return "URL_INVALID"
    if "network" in err_lower or "connection" in err_lower or "dns" in err_lower:
        return "NETWORK_ERROR"
    if "naveg" in err_lower or "nav" in err_lower:
        return "NAV_ERROR"
    return "NETWORK_ERROR"


def classify_errors(errors_detail: list[dict]) -> dict[str, list[int]]:
    """Agrupa errores por categoría para output legible.
    Usa código estándar si está presente, o lo deduce del texto.
    """
    categories = {
        "HTTP 403 (acceso denegado)": [],
        "Timeout": [],
        "Sin dato AA (no beacon)": [],
        "Error de red/conexión": [],
    }
    for e in errors_detail:
        code = e.get("code", _error_code_from_detail(e.get("error", "")))
        err = e.get("error", "")
        if code == "TIMEOUT":
            categories["Timeout"].append(e["row"])
        elif code == "HTTP_403" or "403" in err:
            categories["HTTP 403 (acceso denegado)"].append(e["row"])
        elif code in ("NO_AA_DATA",):
            categories["Sin dato AA (no beacon)"].append(e["row"])
        elif code == "URL_INVALID":
            pass  # ya se reportó en validación
        else:
            categories["Error de red/conexión"].append(e["row"])
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
    time_score = (1 - min(avg_time, 60) / 60) * 100  # más rápido → más score
    return int(
        success_rate * 0.40 +
        dd_rate * 0.25 +
        max(time_score, 0) * 0.15 +
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
    "digitaldata (automatica)",
    "AA analytics (automatico)",
    "AA analytics (estructurado)",
    "metadata / extra beacons",
]

CONTROL_HEADERS = [
    "Fecha", "Source", "URLs", "AA OK", "DD OK", "Errores",
    "Reintentos", "Score", "Tiempo (s)", "Workers",
]

# Colores de columnas (A/B=azul claro, C=rojo, D=amarillo, E=verde, F=sin color)
HEADER_FILLS = {
    "A": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # azul claro
    "B": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # azul claro
    "C": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),  # rojo (DD auto)
    "D": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),  # amarillo (AA auto)
    "E": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # verde (AA estruct)
}

DATA_FILLS = {
    "E": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # verde (AA estruct)
}


def _is_json_error(val: str) -> bool:
    """Check if a JSON string contains an error object."""
    try:
        parsed = json.loads(val)
        return isinstance(parsed, dict) and "error" in parsed and "code" in parsed
    except (json.JSONDecodeError, TypeError):
        return False


def _has_json_data(val: str) -> bool:
    """Check if a JSON string contains real data (not an error)."""
    try:
        parsed = json.loads(val)
        return isinstance(parsed, dict) and "error" not in parsed
    except (json.JSONDecodeError, TypeError):
        return False


def apply_data_fills(ws):
    """Aplica colores de fondo a celdas de datos según su contenido (JSON-aware)."""
    RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for row in range(2, ws.max_row + 1):
        # Col C — DD automático (rojo si es error)
        d = ws.cell(row, 3).value
        if d and isinstance(d, str) and _is_json_error(d):
            ws.cell(row, 3).fill = RED

        # Col D — AA auto (amarillo si tiene datos reales)
        e = ws.cell(row, 4).value
        if e and isinstance(e, str) and _has_json_data(e):
            ws.cell(row, 4).fill = YELLOW

        # Col E — AA estructurado (verde si tiene datos reales)
        f = ws.cell(row, 5).value
        if f and isinstance(f, str) and _has_json_data(f):
            ws.cell(row, 5).fill = DATA_FILLS["E"]


def split_aa_workbooks(wb, audit_date: str, output_dir: str):
    """
    Crea con_aa.xlsx y sin_aa.xlsx a partir del sheet audit_date.
    Cada archivo mantiene headers + colores.
    """
    ws = wb[audit_date]
    con_rows, sin_rows = [], []
    for row in range(2, ws.max_row + 1):
        aa = ws.cell(row, 4).value
        # Real AA data = JSON with "solution" key, no "error" key
        has_aa = False
        if aa and isinstance(aa, str):
            try:
                parsed = json.loads(aa)
                has_aa = isinstance(parsed, dict) and "solution" in parsed and "error" not in parsed
            except (json.JSONDecodeError, TypeError):
                has_aa = False
        if has_aa:
            con_rows.append(row)
        else:
            sin_rows.append(row)

    for suffix, rows in [("con_aa", con_rows), ("sin_aa", sin_rows)]:
        path = os.path.join(output_dir, f"{suffix}.xlsx")
        swb = openpyxl.Workbook()
        swb.remove(swb.active)
        sws = swb.create_sheet(audit_date)
        sws.append(SHEET_HEADERS)
        for col_letter, fill in HEADER_FILLS.items():
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            sws.cell(1, col_idx).fill = fill
        for row_num in rows:
            for col in range(1, 7):
                src = ws.cell(row_num, col)
                dst = sws.cell(row_num, col)
                dst.value = src.value
                if src.has_style:
                    dst.fill = copy(src.fill) if src.fill else None
        _set_col_widths(sws)
        apply_data_fills(sws)
        _auto_row_height(sws)
        save_workbook(swb, path)
        swb.close()
        logging.info("Split %s: %d filas -> %s", suffix, len(rows), path)


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
    parser.add_argument("--discard-cookies", action="store_true", help="[obsoleto] ahora siempre se intenta cerrar el banner de cookies")
    parser.add_argument("--wait-after", type=int, default=4, help="Segundos a esperar tras cargar la pagina (default: 4)")
    parser.add_argument("--market", help="Filtrar por mercado (ES, EN, MX...). Usa campo 'market' en urls.json")
    parser.add_argument("--split-aa", action="store_true", help="Crear con_aa.xlsx y sin_aa.xlsx por separado")
    parser.add_argument("--progress", action="store_true", help="Mostrar barra de progreso")
    parser.add_argument("--diff", action="store_true", help="Mostrar diferencias entre ultima y penultima auditoria")
    parser.add_argument("--config", help="Archivo JSON con config (workers, proxy, retry, etc)")
    parser.add_argument("--backup", action="store_true", help="Crear backup del Excel antes de escribir")
    parser.add_argument("--diagnostic", action="store_true", help="Verificar entorno sin navegar (Python, Playwright, Chromium)")
    parser.add_argument("--retry-failed", action="store_true", help="Solo URLs con error en corridas anteriores")
    parser.add_argument("--verbose", action="store_true", help="Logging detallado (debug)")
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
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format="%(asctime)s | %(levelname)s | %(message)s",
                        handlers=log_handlers)

    # ── --diagnostic: verificar entorno sin navegar ──
    if args.diagnostic:
        print(f"\n{'='*55}")
        print("  DIAGNÓSTICO DEL ENTORNO")
        print(f"{'='*55}")
        ok = True
        # Python
        py_ver = sys.version.split()[0]
        print(f"  Python:           {py_ver}")
        try:
            import playwright
            print(f"  Playwright:       {playwright.__version__}")
        except ImportError:
            print("  Playwright:       ❌ No instalado (pip install playwright)")
            ok = False
        try:
            print(f"  openpyxl:         {openpyxl.__version__}")
        except ImportError:
            print("  openpyxl:         ❌ No instalado (pip install openpyxl)")
            ok = False
        # Chromium
        chromium_path = shutil.which("chromium") or os.environ.get("PLAYWRIGHT_CHROMIUM_PATH")
        if chromium_path:
            print(f"  Chromium:         {chromium_path}")
        else:
            # Revisar ruta por defecto de Playwright en Windows
            default_path = os.path.expandvars(r"%USERPROFILE%\AppData\Local\ms-playwright")
            alt_path = os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright")
            found = None
            for p in (default_path, alt_path):
                if os.path.isdir(p):
                    entries = [d for d in os.listdir(p) if d.startswith("chromium")]
                    if entries:
                        found = os.path.join(p, entries[0])
                        break
            if found:
                print(f"  Chromium:         {found}")
            else:
                print("  Chromium:         ❌ No encontrado (ejecutá: python -m playwright install chromium)")
                ok = False
        # Archivo de entrada
        input_path = args.input or INPUT_FILE
        if os.path.exists(input_path):
            print(f"  Input Excel:      {input_path} ({os.path.getsize(input_path)} bytes)")
        else:
            print(f"  Input Excel:      {input_path} → No encontrado (se creará si usás --urls)")
        # Conectividad proxy
        if args.proxy:
            print(f"  Proxy:            {args.proxy}")
        print(f"\n  {'='*55}")
        if ok:
            print("  ✅ Entorno OK. Listo para navegar.")
        else:
            print("  ⚠️  Hay problemas que resolver antes de navegar.")
        print(f"  {'='*55}\n")
        return 0 if ok else 1

    # ── --diff: comparar ultimas 2 auditorias (sin procesar) ──
    if args.diff:
        if args.urls:
            # Con --urls: buscar en el archivo multi-sheet correspondiente
            if args.market:
                diff_path = args.output or f"{args.market.upper()}/historial.xlsx"
            else:
                diff_path = args.output or "historial.xlsx"
        else:
            diff_path = args.output or "historial.xlsx"
        if not os.path.exists(diff_path):
            logging.error("No existe %s para comparar", diff_path)
            return 1
        wb = openpyxl.load_workbook(diff_path)
        sheets = [s for s in wb.sheetnames if s != "_control" and s != "_vars"]
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
            dd_a = (sa.cell(row, 3).value or "")[:50]
            dd_b = (sb.cell(row, 3).value or "")[:50]
            aa_a = (sa.cell(row, 4).value or "")[:50]
            aa_b = (sb.cell(row, 4).value or "")[:50]
            if str(aa_a).strip() != str(aa_b).strip() or str(dd_a).strip() != str(dd_b).strip():
                print(f"  Fila {row}: {url_a or url_b}")
                if str(aa_a).strip() != str(aa_b).strip():
                    print(f"    AA: {'ANTES' if not sa.cell(row,4).value else 'NUEVO'}")
                if str(dd_a).strip() != str(dd_b).strip():
                    print(f"    DD: {'ANTES' if not sa.cell(row,3).value else 'NUEVO'}")
        wb.close()
        return 0

    # ── Excel: --urls → multi-sheet historial | --input → clasico ──
    audit_date = None  # flag: multi-sheet mode
    if args.urls:
        if args.market:
            market_dir = args.market.upper()
            os.makedirs(market_dir, exist_ok=True)
            output_path = args.output or f"{market_dir}/historial.xlsx"
        else:
            output_path = args.output or "historial.xlsx"
        wb, ws, audit_date, skipped = setup_multisheet(output_path, args.urls, args.resume)
        if skipped:
            logging.info("Sheet de hoy ya auditado. Usá --diff para ver diferencias.")
            return 0
        with open(args.urls, encoding="utf-8") as f:
            entries = json.load(f)
        row = 2
        for entry in entries:
            if isinstance(entry, str):
                name = url = entry
            elif isinstance(entry, dict):
                if args.market and entry.get("market", "").upper() != args.market.upper():
                    continue
                name = entry.get("nombre", entry.get("url", ""))
                url = entry.get("url", "")
            wrap = Alignment(wrap_text=True, vertical="top")
            ws.cell(row, 1).value = name; ws.cell(row, 1).alignment = wrap
            ws.cell(row, 2).value = url; ws.cell(row, 2).alignment = wrap
            row += 1
        total_filtradas = row - 2
        logging.info("Multi-sheet %s: %d URLs (filtradas: %d)", audit_date, total_filtradas, len(entries))
    else:
        output_path = args.output or args.input
        # ── Cargar Excel existente ──
        wb = openpyxl.load_workbook(args.input)
        ws = wb.active

    if args.output and not args.urls and not audit_date:
        # Copia física del input al output, luego trabajamos sobre output.
        # Si el script crashea durante procesamiento, el output queda con los
        # datos originales (no corrupto), y el input jamás se modifica.
        import shutil
        shutil.copy2(args.input, args.output)
        wb = openpyxl.load_workbook(args.output)
        ws = wb.active

    errs = validate_sheet(ws)
    if errs:
        for e in errs:
            logging.error("Validacion: %s", e)
        wb.close()
        sys.exit(1)

    ws.cell(1, 3).value = ws.cell(1, 3).value or "digitaldata (automatica)"
    ws.cell(1, 6).value = ws.cell(1, 6).value or "metadata / extra beacons"

    rows_to_process = []
    for row in range(2, ws.max_row + 1):
        url = ws.cell(row, 2).value
        if not url:
            continue
        if args.row and row != args.row:
            continue
        if args.resume:
            existing = ws.cell(row, 4).value
            if existing and len(str(existing).strip()) > 30 and "no AA" not in str(existing):
                logging.info("Fila %d: saltando (resume)", row)
                continue
        if args.retry_failed:
            existing = ws.cell(row, 4).value
            if not existing or str(existing).strip().startswith("("):
                logging.info("Fila %d: reintentar (error previo)", row)
            else:
                continue  # saltear las que ya tienen datos OK
        rows_to_process.append((row, str(url).strip()))

    if not rows_to_process:
        logging.info("No hay URLs para procesar.")
        wb.close()
        return

    logging.info("URLs a procesar: %d | Workers: %d | Backup: %s",
                 len(rows_to_process), args.workers, getattr(args, "backup", False))

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
                    wait_after=args.wait_after,
                    timeout_ms=args.timeout,
                    max_retry=args.retry,
                )
                status_str = f"HTTP {result['status']}" if result["status"] > 0 else f"ERR {result['status']}"
                log_url = sanitize_url_for_log(url)
                logging.info("Fila %d | %s | AA: %s | %.1fs | %s",
                             row, status_str,
                             result["aa_source"] or "none",
                             result["elapsed_s"], log_url)
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
            if _shutdown_flag:
                return
            async with sem:
                if _shutdown_flag:
                    return
                await worker(row, url)

        # Filtrar URLs que no pasan validación
        valid_rows = []
        for row, url in rows_to_process:
            err = validate_url(url)
            if err:
                logging.error("Fila %d: URL inválida (%s): %s", row, err, sanitize_url_for_log(url))
                metrics["errors"] += 1
                metrics["errores_detalle"].append({"row": row, "error": err})
                _write_cell(ws, row, 4, _pretty_json({"error": f"URL inválida: {err}", "code": "URL_INVALID"}))
                _write_cell(ws, row, 6, _pretty_json({"error": err, "url": url[:80], "code": "URL_INVALID"}))
            else:
                valid_rows.append((row, url))

        # Omitir si todas las URLs son inválidas
        if not valid_rows:
            logging.warning("No hay URLs válidas para procesar.")

        tasks = [worker_sem(row, url) for row, url in valid_rows]
        await asyncio.gather(*tasks)

        await browser.close()

    # Actualizar _vars con datos AA recolectados
    if audit_date and aa_data:
        update_vars_sheet(wb, audit_date, aa_data)

    # ── Guardado final + _control (multi-sheet) ──
    apply_data_fills(ws)
    _set_col_widths(ws)
    total_time = time.time() - start_time
    score = compute_score(metrics)
    if audit_date:
        update_control(wb, audit_date, args.urls or "",
                       metrics["total"],
                       metrics["ok_aa"], metrics["ok_dd"],
                       metrics["errors"], metrics["retries"],
                       score, total_time, args.workers)

    # --backup: copiar el archivo existente antes de sobrescribir
    if args.backup and os.path.exists(output_path):
        backup_name = f"{output_path}.{datetime.now().strftime('%Y%m%d_%H%M%S')}.bak"
        try:
            import shutil
            shutil.copy2(output_path, backup_name)
            logging.info("Backup creado: %s", backup_name)
        except Exception as e:
            logging.warning("No se pudo crear backup: %s", e)

    _auto_row_height(ws)
    out = save_workbook(wb, output_path)

    # ── Split AA (con_aa / sin_aa) ──
    if args.split_aa and audit_date:
        output_dir = os.path.dirname(output_path) or "."
        split_aa_workbooks(wb, audit_date, output_dir)

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
    # Registrar signal handler para graceful shutdown
    try:
        signal.signal(signal.SIGINT, _request_shutdown)
        signal.signal(signal.SIGTERM, _request_shutdown)
    except (ValueError, AttributeError):
        pass  # no siempre disponible (ej: threads)
    try:
        asyncio.run(amain())
    except KeyboardInterrupt:
        _request_shutdown()
        logging.warning("Interrumpido por usuario.")
        sys.exit(130)


if __name__ == "__main__":
    main()

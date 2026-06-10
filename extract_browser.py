"""
extract_browser.py — Automatización: navega URLs, extrae digitalData + AA beacon, escribe Excel.

Workflow:
  1. Lee Excel (--input) → URLs de col B
  2. Procesa URLs concurrentemente (--workers N):
     - Extrae window.digitalData (data layer)
     - Captura TODOS los beacons AA (s.t. + s.tl.)
     - Acumula, parsea, escribe
  3. Escribe a --output (o al input si no se especifica)
     - Col D: Adobe Analytics original (primer beacon)
     - Col F: digitalData (data layer)
     - Col G: metadata / extra beacons
  4. Guardado incremental cada 5 URLs + guardado final
  5. Muestra métricas al final

Uso:
  python extract_browser.py                                   # todo, escribe sobre input
  python extract_browser.py --output resultado.xlsx           # escribe a otro archivo
  python extract_browser.py --workers 3                       # 3 URLs concurrentes
  python extract_browser.py --workers 4 --output res.xlsx     # concurrente + seguro
  python extract_browser.py --row 3                           # solo una fila
  python extract_browser.py --resume                          # saltar procesadas
  python extract_browser.py --headed                          # navegador visible
  python extract_browser.py --input otro.xlsx                 # otro archivo
  python extract_browser.py --log-file resultados.log         # persistir log
  python extract_browser.py --proxy http://proxy:8080         # corporate proxy
  python extract_browser.py --run-clean                       # ejecuta extract_aa.py al final

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
from openpyxl.styles import Alignment
from playwright.async_api import async_playwright, TimeoutError as PwTimeout

INPUT_FILE = "RevisionManual.xlsx"
SAVE_EVERY_N = 5  # guardado incremental cada N URLs

AA_DOMAINS = [
    "smetrics.ford.com", "sc.omtrdc.net", "smetrics.omtrdc.net",
    "2o7.net", "data.adobedc.net", "edge.adobedc.net",
]

DATA_LAYER_NAMES = [
    "window.digitalData", "window.dataLayer", "window.DigitalData",
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

    props = {}
    for key, val in qs.items():
        m = re.match(r"^c(\d+)$", key)
        if m:
            props[f"prop{m.group(1)}"] = val[0]

    evars = {}
    for key, val in qs.items():
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
            props[f"prop{m.group(1)}"] = val; continue
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
              "title": "", "digitalData": None, "aa_parsed": None,
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
            result["digitalData"] = await extract_digital_data(page)
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
) -> None:
    """
    Escribe el resultado de una URL en el Excel.
    Usa lock para evitar corrupción entre workers.
    """
    async with excel_lock:
        row = result["row"]
        n_beacons = 1 + len(result.get("extra_beacons", []))

        # digitalData → col F
        dd_val = result.get("digitalData")
        if dd_val is not None:
            cell = ws.cell(row, 6)
            cell.value = json.dumps(dd_val, indent=2, ensure_ascii=False)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            metrics["ok_dd"] += 1
        else:
            ws.cell(row, 6).value = "(no digitalData)"

        # AA → col D
        if result.get("aa_parsed"):
            cell = ws.cell(row, 4)
            cell.value = json.dumps(result["aa_parsed"], indent=2, ensure_ascii=False)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            metrics["ok_aa"] += 1
        else:
            ws.cell(row, 4).value = f"({result.get('error', 'no AA')})"

        # Metadata → col G
        meta = {"status": result["status"], "aa_source": result["aa_source"],
                "beacons": n_beacons, "title": result["title"],
                "error": result.get("error"), "elapsed_s": result["elapsed_s"],
                "url": result["url"][:120]}
        if result.get("extra_beacons"):
            meta["extra_beacons"] = result["extra_beacons"]
        ws.cell(row, 7).value = json.dumps(meta, indent=2, ensure_ascii=False)

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


# ═══════════════════════════════════════════════════════════════════════════
# MAIN ASYNC
# ═══════════════════════════════════════════════════════════════════════════

async def amain():
    import argparse
    parser = argparse.ArgumentParser(description="Extrae digitalData + AA beacon desde URLs en Excel")
    parser.add_argument("--row", type=int, help="Procesar solo una fila")
    parser.add_argument("--headed", action="store_true", help="Navegador visible")
    parser.add_argument("--input", default=INPUT_FILE, help="Archivo Excel de entrada")
    parser.add_argument("--output", help="Archivo Excel de salida (default: escribe sobre input)")
    parser.add_argument("--timeout", type=int, default=35000, help="Timeout por pagina (ms)")
    parser.add_argument("--resume", action="store_true", help="Saltar filas con datos")
    parser.add_argument("--log-file", help="Archivo de log")
    parser.add_argument("--proxy", help="Proxy HTTP (ej: http://proxy:8080)")
    parser.add_argument("--retry", type=int, default=1, help="Reintentos por URL")
    parser.add_argument("--workers", type=int, default=1, help="URLs concurrentes (default: 1)")
    parser.add_argument("--run-clean", action="store_true", help="Ejecutar extract_aa.py al final")
    parser.add_argument("--discard-cookies", action="store_true", help="Rechazar banners de cookies")
    args = parser.parse_args()

    if args.row:
        args.workers = 1

    # ── Logging ──
    log_handlers = [logging.StreamHandler(sys.stderr)]
    if args.log_file:
        log_handlers.append(logging.FileHandler(args.log_file, encoding="utf-8"))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                        handlers=log_handlers)

    output_path = args.output or args.input

    # ── Excel: SIEMPRE copia de seguridad ──
    wb = openpyxl.load_workbook(args.input)
    ws = wb.active

    if args.output:
        # Si hay --output, copiamos antes de mutar
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
    else:
        # Modo destructivo: trabajamos sobre wb original
        pass

    errs = validate_sheet(ws)
    if errs:
        for e in errs:
            logging.error("Validacion: %s", e)
        wb.close()
        sys.exit(1)

    ws.cell(1, 6).value = ws.cell(1, 6).value or "digitalData (data layer)"
    ws.cell(1, 7).value = ws.cell(1, 7).value or "metadata / extra beacons"

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
        rows_to_process.append((row, str(url).strip()))

    if not rows_to_process:
        logging.info("No hay URLs para procesar.")
        wb.close()
        return

    logging.info("URLs a procesar: %d | Workers: %d", len(rows_to_process), args.workers)

    metrics = {"total": len(rows_to_process), "ok_aa": 0, "ok_dd": 0,
               "errors": 0, "retries": 0, "total_beacons": 0, "times": [],
               "errores_detalle": []}
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
                await write_result(ws, result, metrics, excel_lock, output_path, saved_count)
            finally:
                await page.close()
                await ctx.close()

        async def worker_sem(row: int, url: str):
            async with sem:
                await worker(row, url)

        tasks = [worker_sem(row, url) for row, url in rows_to_process]
        await asyncio.gather(*tasks)

        await browser.close()

    # ── Guardado final ──
    ws.column_dimensions["D"].width = 80
    ws.column_dimensions["F"].width = 60
    ws.column_dimensions["G"].width = 40
    out = save_workbook(wb, output_path)
    wb.close()
    logging.info("Guardado: %s", out)

    # ── Métricas ──
    total_time = time.time() - start_time
    success_rate = metrics["ok_aa"] / max(metrics["total"], 1) * 100
    dd_rate = metrics["ok_dd"] / max(metrics["total"], 1) * 100
    avg_time = sum(metrics["times"]) / max(len(metrics["times"]), 1)
    beacons_per_url = metrics["total_beacons"] / max(metrics["total"], 1)
    score = compute_score(metrics)

    print(f"""
{'='*55}
  MÉTRICAS Y SCORE
{'='*55}
  Config:            {args.workers} worker(s) concurrente(s)
  URLs procesadas:   {metrics['ok_aa']}/{metrics['total']}
  AA capturados:     {success_rate:.0f}%
  digitalData:       {dd_rate:.0f}%
  Beacons totales:   {metrics['total_beacons']} ({beacons_per_url:.1f}/URL)
  Reintentos:        {metrics['retries']}
  Errores:           {metrics['errors']}
{'─'*55}
  Tiempo total:      {timedelta(seconds=int(total_time))}
  Promedio/URL:      {avg_time:.1f}s
  Guardados incr.:   cada {SAVE_EVERY_N} URLs
{'─'*55}
  SCORE GLOBAL:      {score}/100
{'─'*55}""")

    if metrics["errores_detalle"]:
        print("  DETALLE ERRORES:")
        for e in metrics["errores_detalle"]:
            print(f"    Fila {e['row']}: {e['error']}")

    if score < 60:
        print("  Sugerencias:")
        if success_rate < 50:
            print("    • Verificar VPN / acceso a las URLs")
        if dd_rate < 50:
            print("    • El data layer podria no llamarse window.digitalData")
        if avg_time > 30:
            print("    • Paginas lentas. Aumentar --timeout o verificar SPAs")

    if args.run_clean:
        print(f"\n{'='*55}")
        print("  Ejecutando extract_aa.py...")
        print(f"{'='*55}")
        clean_script = os.path.join(os.path.dirname(__file__) or ".", "extract_aa.py")
        inp = args.output or args.input
        r = subprocess.run([sys.executable, clean_script, "--input", inp],
                           capture_output=True, text=True)
        print(r.stdout)
        if r.stderr:
            print(r.stderr, file=sys.stderr)

    return score


def main():
    asyncio.run(amain())


if __name__ == "__main__":
    main()

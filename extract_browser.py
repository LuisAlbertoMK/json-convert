"""
extract_browser.py — Automatización: navega URLs, extrae digitalData + AA beacon, escribe Excel.

Workflow:
  1. Lee RevisionManual.xlsx → URLs de col B
  2. Para cada URL (Playwright headless):
     - Extrae window.digitalData (data layer)
     - Captura beacon AA por 3 métodos:
       a. Network request a smetrics.ford.com / *.omtrdc.net
       b. window.s object (AppMeasurement)
       c. window.__adobe (AEP Debugger)
     - Parsea beacon a JSON estructurado
  3. Escribe:
     - Col D: Adobe Analytics original
     - Col F: digitalData (data layer)
  4. Guarda Excel

Uso:
  python extract_browser.py                        # todo
  python extract_browser.py --row 3                # solo fila 3
  python extract_browser.py --headed               # navegador visible
  python extract_browser.py --input otro.xlsx      # otro archivo

Requiere:
  pip install playwright openpyxl
  python -m playwright install chromium
"""

import json
import os
import re
import time
from urllib.parse import urlparse, parse_qs

import openpyxl
from openpyxl.styles import Alignment
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

INPUT_FILE = "RevisionManual.xlsx"

# Dominios de AA tracking
AA_DOMAINS = [
    "smetrics.ford.com",
    "sc.omtrdc.net",
    "smetrics.omtrdc.net",
    "2o7.net",
    "data.adobedc.net",
    "edge.adobedc.net",
]

# Posibles nombres del data layer
DATA_LAYER_NAMES = [
    "window.digitalData",
    "window.dataLayer",
    "window.DigitalData",
    "window.digital_data",
    "window.utag_data",  # Tealium
]


# ─── helpers ───────────────────────────────────────────────────────────────


def parse_aa_beacon(beacon_url: str, page_title: str = "") -> dict:
    """Parsea URL de beacon de Adobe Analytics a JSON estructurado."""
    parsed = urlparse(beacon_url)
    qs = parse_qs(parsed.query)

    def first(key):
        vals = qs.get(key, [])
        return vals[0] if vals else ""

    # Report suite del path: /b/ss/{rs}/...
    path_parts = parsed.path.split("/")
    report_suite = ""
    hit_id_from_path = ""
    if len(path_parts) >= 4 and path_parts[1] == "b" and path_parts[2] == "ss":
        report_suite = path_parts[3]
        hit_id_from_path = path_parts[-1] if len(path_parts) > 4 else ""

    # Page URL
    page_url = first("g")

    # Props
    props = {}
    for key, val in qs.items():
        m = re.match(r"^c(\d+)$", key)
        if m:
            props[f"prop{m.group(1)}"] = val[0]

    # eVars
    evars = {}
    for key, val in qs.items():
        m = re.match(r"^v(\d+)$", key)
        if m:
            evars[f"eVar{m.group(1)}"] = val[0]

    # Events
    events_raw = first("events")
    events = [e.strip() for e in events_raw.split(",") if e.strip()]

    # Visitor
    visitor = {}
    mid = first("mid")
    if mid:
        visitor["experienceCloudId"] = mid
    aamlh = first("aamlh")
    if aamlh:
        visitor["audienceManagerHint"] = aamlh

    # Browser
    browser = {}
    res = first("res")
    if res:
        browser["resolution"] = res
    bw = first("bw")
    if bw:
        browser["browserWidth"] = int(bw) if bw.isdigit() else bw
    bh = first("bh")
    if bh:
        browser["browserHeight"] = int(bh) if bh.isdigit() else bh
    cd = first("cd")
    if cd:
        browser["colorDepth"] = cd
    ce = first("ce")
    if ce:
        browser["charset"] = ce

    # Hit
    hit = {
        "id": hit_id_from_path,
        "type": "pageView",
        "reportSuiteId": report_suite,
    }

    # Timestamp
    ts = first("t") or first("ts")

    request_obj = {
        "method": "GET",
        "hostname": parsed.hostname,
        "pathname": parsed.path,
    }
    if ts:
        request_obj["collectedTimestamp"] = ts

    result = {
        "solution": "analytics",
        "page": {"title": page_title, "url": page_url},
        "request": request_obj,
        "visitor": visitor,
        "hit": hit,
        "browser": browser,
        "events": events,
        "eVars": evars,
        "props": props,
        "pageName": first("pageName"),
        "channel": first("ch"),
    }

    # Products
    products_raw = first("products")
    if products_raw:
        result["products"] = products_raw

    return result


def extract_s_object(page) -> dict | None:
    """Intenta leer window.s (AppMeasurement). Retorna dict con vars o None."""
    try:
        s_obj = page.evaluate("""
            () => {
                const s = window.s || (window.s_c_il && window.s_c_il[window.s_c_il.length-1]);
                if (!s) return null;
                const vars = {};
                const props = ['pageName','pageURL','channel','server','pageType',
                    'prop1','prop2','prop3','prop4','prop5','prop6','prop7','prop8','prop9','prop10',
                    'prop11','prop12','prop13','prop14','prop15','prop16','prop17','prop18','prop19','prop20',
                    'prop21','prop22','prop23','prop24','prop25','prop26','prop27','prop28','prop29','prop30',
                    'prop31','prop32','prop33','prop34','prop35','prop36','prop37','prop38','prop39','prop40',
                    'prop41','prop42','prop43','prop44','prop45','prop46','prop47','prop48','prop49','prop50',
                    'prop51','prop52','prop53','prop54','prop55','prop56','prop57','prop58','prop59','prop60',
                    'prop61','prop62','prop63','prop64','prop65','prop66','prop67','prop68','prop69','prop70',
                    'prop71','prop72','prop73','prop74','prop75',
                    'eVar1','eVar2','eVar3','eVar4','eVar5','eVar6','eVar7','eVar8','eVar9','eVar10',
                    'eVar11','eVar12','eVar13','eVar14','eVar15','eVar16','eVar17','eVar18','eVar19','eVar20',
                    'eVar21','eVar22','eVar23','eVar24','eVar25','eVar26','eVar27','eVar28','eVar29','eVar30',
                    'eVar31','eVar32','eVar33','eVar34','eVar35','eVar36','eVar37','eVar38','eVar39','eVar40',
                    'eVar41','eVar42','eVar43','eVar44','eVar45','eVar46','eVar47','eVar48','eVar49','eVar50',
                    'eVar51','eVar52','eVar53','eVar54','eVar55','eVar56','eVar57','eVar58','eVar59','eVar60',
                    'eVar61','eVar62','eVar63','eVar64','eVar65','eVar66','eVar67','eVar68','eVar69','eVar70',
                    'eVar71','eVar72','eVar73','eVar74','eVar75',
                    'events','products','linkTrackVars','linkTrackEvents',
                    'charSet','visitorID','visitorMigrationKey','visitorMigrationServer',
                    'currencyCode','transactionID',
                ];
                for (const p of props) {
                    if (s[p] !== undefined && s[p] !== '') {
                        vars[p] = s[p];
                    }
                }
                return Object.keys(vars).length > 0 ? vars : null;
            }
        """)
        return s_obj
    except Exception:
        return None


def extract_adobe_debugger(page) -> dict | None:
    """Intenta leer datos del Adobe Experience Platform Debugger inyectados."""
    try:
        debug_data = page.evaluate("""
            () => {
                // AEP Debugger inyecta datos en window.__adobe o adobe.edge
                if (window.__adobe && window.__adobe.analytics) {
                    return window.__adobe.analytics;
                }
                if (window.adobe && window.adobe.analytics) {
                    return window.adobe.analytics;
                }
                return null;
            }
        """)
        return debug_data
    except Exception:
        return None


def build_aa_parsed_from_s(s_obj: dict, page_title: str = "") -> dict:
    """Convierte window.s vars a JSON estructurado similar al del beacon."""
    props = {}
    evars = {}
    events = []

    for key, val in s_obj.items():
        m = re.match(r"^prop(\d+)$", key, re.IGNORECASE)
        if m:
            props[f"prop{m.group(1)}"] = val
            continue
        m = re.match(r"^eVar(\d+)$", key, re.IGNORECASE)
        if m:
            evars[f"eVar{m.group(1)}"] = val
            continue

    events_raw = s_obj.get("events", "")
    if events_raw:
        events = [e.strip() for e in events_raw.split(",") if e.strip()]

    page_url = s_obj.get("pageURL", "")
    page_name = s_obj.get("pageName", "")

    return {
        "solution": "analytics",
        "page": {"title": page_title, "url": page_url},
        "pageName": page_name,
        "request": {"source": "window.s"},
        "events": events,
        "eVars": evars,
        "props": props,
        "channel": s_obj.get("channel", ""),
        "products": s_obj.get("products", ""),
        "currencyCode": s_obj.get("currencyCode", ""),
    }


# ─── navegación ───────────────────────────────────────────────────────────


def process_page(page, url: str, timeout_ms: int = 35000) -> dict:
    """
    Navega a URL, extrae digitalData + AA por múltiples métodos.
    """
    result = {
        "digitalData": None,
        "aa_parsed": None,
        "aa_source": None,  # 'beacon' | 'window.s' | 'debugger'
        "title": "",
        "status": 0,
        "page_name": "",
    }

    # 1. Interceptar beacons AA
    beacon_url = None

    def on_request(request):
        nonlocal beacon_url
        if beacon_url:
            return
        url_lower = request.url.lower()
        # Coincide con cualquier dominio AA conocido
        for domain in AA_DOMAINS:
            if domain in url_lower and "/b/ss/" in url_lower:
                beacon_url = request.url
                return

    page.on("request", on_request)

    # 2. Navegar
    try:
        resp = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        if resp:
            result["status"] = resp.status
    except PwTimeout:
        result["status"] = -1
        return result
    except Exception as e:
        result["status"] = -2
        return result

    # 3. Esperar carga completa
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PwTimeout:
        pass
    time.sleep(1)  # respiro para beacons async

    # 4. Título de página
    try:
        result["title"] = (page.evaluate("document.title") or "").strip()
    except Exception:
        pass

    # 5. digitalData (probar varios nombres)
    for var_name in DATA_LAYER_NAMES:
        try:
            dd = page.evaluate(var_name)
            if dd and isinstance(dd, dict) and len(dd) > 0:
                result["digitalData"] = dd
                break
        except Exception:
            continue

    # 6. AA: intentar por network beacon
    if beacon_url:
        result["aa_parsed"] = parse_aa_beacon(beacon_url, result["title"])
        result["aa_source"] = "beacon"
        result["page_name"] = result["aa_parsed"].get("pageName", "")
        return result

    # 7. AA: fallback a window.s
    s_obj = extract_s_object(page)
    if s_obj and s_obj.get("pageName"):
        result["aa_parsed"] = build_aa_parsed_from_s(s_obj, result["title"])
        result["aa_source"] = "window.s"
        result["page_name"] = s_obj.get("pageName", "")
        return result

    # 8. AA: fallback a AEP Debugger
    debug_data = extract_adobe_debugger(page)
    if debug_data:
        result["aa_parsed"] = debug_data
        result["aa_source"] = "debugger"

    return result


# ─── Excel ────────────────────────────────────────────────────────────────


def ensure_column(ws, col_letter):
    """Asegura que la columna exista en el worksheet."""
    max_col = ws.max_column or 1
    target_idx = ord(col_letter) - 64
    if target_idx > max_col:
        for c in range(max_col + 1, target_idx + 1):
            ws.cell(1, c)


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Extrae digitalData + AA beacon desde URLs en Excel")
    parser.add_argument("--row", type=int, help="Procesar solo una fila")
    parser.add_argument("--headed", action="store_true", help="Navegador visible")
    parser.add_argument("--input", default=INPUT_FILE, help="Archivo Excel")
    parser.add_argument("--delay", type=float, default=0, help="Segundos de espera entre URLs")
    parser.add_argument("--timeout", type=int, default=35000, help="Timeout por página (ms)")
    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.input)
    ws = wb.active

    # Headers
    ensure_column(ws, "G")
    ws.cell(1, 6).value = ws.cell(1, 6).value or "digitalData (data layer)"
    ws.cell(1, 7).value = ws.cell(1, 7).value or "status / aa_source"

    rows_to_process = []
    for row in range(2, ws.max_row + 1):
        url = ws.cell(row, 2).value
        if not url:
            continue
        if args.row and row != args.row:
            continue
        rows_to_process.append((row, str(url).strip()))

    if not rows_to_process:
        print("No hay URLs para procesar.")
        wb.close()
        return

    print(f"URLs a procesar: {len(rows_to_process)}")
    ok_dd = 0
    ok_aa = 0
    errores = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not args.headed)
        context = browser.new_context(
            viewport={"width": 1280, "height": 720},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36",
        )

        for idx, (row, url) in enumerate(rows_to_process):
            print(f"\n[{idx+1}/{len(rows_to_process)}] Fila {row}", flush=True)
            print(f"    URL: {url[:120]}")

            try:
                page = context.new_page()
                data = process_page(page, url, args.timeout)
                page.close()

                # Estado
                status_str = f"HTTP {data['status']}" if data["status"] > 0 else f"ERROR {data['status']}"
                aa_src = data["aa_source"] or "none"
                print(f"    Status: {status_str} | AA fuente: {aa_src}")

                # digitalData → col F
                if data.get("digitalData") is not None:
                    dd_json = json.dumps(data["digitalData"], indent=2, ensure_ascii=False)
                    cell = ws.cell(row, 6)
                    cell.value = dd_json
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    ok_dd += 1
                    print(f"    digitalData: OK ({len(dd_json)} chars)")
                else:
                    ws.cell(row, 6).value = "(no digitalData)"
                    print(f"    digitalData: NO ENCONTRADO")

                # AA → col D
                if data.get("aa_parsed"):
                    aa_json = json.dumps(data["aa_parsed"], indent=2, ensure_ascii=False)
                    cell = ws.cell(row, 4)
                    cell.value = aa_json
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    ok_aa += 1
                    print(f"    AA (col D): OK ({len(aa_json)} chars, fuente={aa_src})")
                else:
                    ws.cell(row, 4).value = "(no AA data captured)"
                    print(f"    AA (col D): NO CAPTURADO")

                # Metadatos → col G
                meta = json.dumps({"status": data["status"], "aa_source": aa_src, "title": data["title"]})
                ws.cell(row, 7).value = meta

                # Delay entre URLs
                if args.delay and idx < len(rows_to_process) - 1:
                    time.sleep(args.delay)

            except Exception as e:
                errores.append((row, str(e)[:150]))
                print(f"    ERROR: {str(e)[:80]}")
                continue

        context.close()
        browser.close()

    # Anchos
    ws.column_dimensions["D"].width = 80
    ws.column_dimensions["F"].width = 60
    ws.column_dimensions["G"].width = 30

    # Guardar
    out = args.input
    try:
        wb.save(out)
    except PermissionError:
        name, ext = os.path.splitext(out)
        out = f"{name}_browser{ext}"
        wb.save(out)
    wb.close()

    print(f"\n{'='*50}")
    print(f"Guardado: {out}")
    print(f"digitalData: {ok_dd}/{len(rows_to_process)}")
    print(f"AA beacon:  {ok_aa}/{len(rows_to_process)}")
    if errores:
        print(f"Errores: {len(errores)}")
        for r, e in errores:
            print(f"  Fila {r}: {e}")


if __name__ == "__main__":
    main()

"""
json_convert/aa_parser.py — Parseo de beacons Adobe Analytics y
extracción de data layer desde Playwright (async).

Contiene funciones 100% puras (testeables sin navegador) y
funciones async que dependen de page (Playwright).
"""

from __future__ import annotations

import logging
import re
from typing import Any, Protocol, cast

from json_convert.types import BeaconResult
from urllib.parse import parse_qs, urlparse


class PageProtocol(Protocol):
    """Protocolo mínimo para Playwright Page (evita dependencia de stubs externos).

    ponyatil: Solo cubre evaluate + query_selector — lo mínimo que usamos.
    Si se agregan más métodos de Page, expandir aquí.
    """

    async def evaluate(self, expression: str, arg: Any = None) -> Any: ...
    async def query_selector(self, selector: str) -> Any | None: ...

# Dominios conocidos de Adobe Analytics para captura de beacons
AA_DOMAINS = [
    "smetrics.ford.com", "sc.omtrdc.net", "smetrics.omtrdc.net",
    "2o7.net", "data.adobedc.net", "edge.adobedc.net",
]

# Nombres de variables de data layer a probar en el navegador
DATA_LAYER_NAMES = [
    "window.digitalData", "window.digitaldata", "window.dataLayer",
    "window.digital_data", "window.utag_data",
]


def parse_aa_beacon(beacon_url: str, page_title: str = "") -> BeaconResult:
    """Parsea URL de beacon de Adobe Analytics a JSON estructurado."""
    parsed = urlparse(beacon_url)
    qs = parse_qs(parsed.query)

    def first(key: str) -> str:
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

    browser: dict[str, object] = {}
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
        result["request"]["collectedTimestamp"] = ts  # type: ignore[index]
    products_raw = first("products")
    if products_raw:
        result["products"] = products_raw
    return cast(BeaconResult, result)


def build_aa_from_s(s_obj: dict, page_title: str = "") -> BeaconResult:
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


async def extract_s_object(page: PageProtocol) -> dict | None:
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


async def extract_digital_data(page: PageProtocol) -> dict | None:
    """Extrae data layer probando varios nombres con logging detallado."""
    for var_name in DATA_LAYER_NAMES:
        try:
            dd = await page.evaluate(f"(() => {{ try {{ const v = {var_name}; return v !== null && v !== undefined ? v : null; }} catch(e) {{ return null; }} }})()")
            if dd and isinstance(dd, dict) and len(dd) > 0:
                keys = list(dd.keys())[:6]
                logging.info("data layer encontrado en '%s' keys=%s", var_name, keys)
                return dd
            elif dd is None:
                logging.debug("data layer '%s' -> null/undefined (no existe en la pagina)", var_name)
            else:
                tipo = type(dd).__name__
                logging.debug("data layer '%s' -> %s (no es dict: %s)", var_name, repr(dd)[:80], tipo)
        except Exception as e:
            logging.debug("data layer '%s' failed: %s", var_name, e)
    return None


async def debug_dump_available_globals(page: PageProtocol) -> None:
    """Dumpea TODAS las variables globales del window para debug.
    Solo se ejecuta en modo verbose. Busca específicamente objetos
    que parezcan data layers."""
    try:
        js = """
        (() => {
            const candidates = [];
            for (const key of Object.keys(window)) {
                try {
                    const val = window[key];
                    if (val && typeof val === 'object' && !(val instanceof Node) && !(val instanceof Window)) {
                        const keys = Object.keys(val).slice(0, 8);
                        const typeStr = val.constructor ? val.constructor.name : 'Object';
                        candidates.push({key, type: typeStr, sample_keys: keys, len: keys.length});
                    }
                } catch(e) {}
            }
            candidates.sort((a, b) => b.len - a.len);
            return candidates.slice(0, 20);
        })()
        """
        globals_list = await page.evaluate(js)
        if globals_list:
            logging.info("Global objects disponibles (%d encontrados):", len(globals_list))
            for g in globals_list:
                logging.info("  %s -> %s [%s] %s",
                             g.get("key", "?"),
                             g.get("type", "?"),
                             g.get("sample_keys", []),
                             "(data layer candidate)" if any(
                                 dl in g.get("key", "").lower()
                                 for dl in ["digital", "data", "layer", "utag", "analyt"]
                             ) else "")
        else:
            logging.debug("No se encontraron objetos globales (pagina sin JS?)")
    except Exception as e:
        logging.debug("debug_dump_globals falló: %s", e)


async def extract_title(page: PageProtocol) -> str:
    """Extrae el título de la página."""
    try:
        return (await page.evaluate("document.title") or "").strip()
    except Exception as e:
        logging.debug("title failed: %s", e)
        return ""


async def try_dismiss_cookie_consent(page: PageProtocol) -> bool:
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

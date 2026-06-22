"""
browser.py — Lógica de navegación Playwright para extracción de datos.

Separa la interacción con el browser del entry point CLI,
permitiendo testeo independiente de process_url.
"""

from __future__ import annotations

import asyncio
import logging
import random
import time
from typing import cast

from playwright.async_api import Page, TimeoutError as PwTimeout

from json_convert import (
    debug_dump_available_globals,
    extract_digital_data,
    extract_s_object,
    extract_title,
    sanitize_url_for_log,
    try_dismiss_cookie_consent,
    validate_url,
)
from json_convert.types import UrlResult

# ── Graceful shutdown ──
_shutdown_flag = False
"""Global flag: set by signal handler to request graceful shutdown."""


async def _fetch_html_via_http(url: str, timeout_ms: int) -> str:
    """Fetch HTML via stdlib urllib — sin dependencias extra.
    Útil como fallback cuando Playwright goto falla con ERR_ABORTED."""
    import urllib.error
    import urllib.request

    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "es-PR,es;q=0.9,en;q=0.8",
        },
    )
    with urllib.request.urlopen(req, timeout=max(1, timeout_ms // 1000)) as resp:
        raw = resp.read()
    return raw.decode("utf-8", errors="replace")


def _page_name_from_url(url: str) -> str:
    """Deriva un nombre legible desde la URL."""
    try:
        parts = url.rstrip("/").split("/")
        for i, p in enumerate(parts):
            if "ford" in p and i + 2 < len(parts):
                return parts[-1].replace(".html", "").replace("-", " ").title()
        return parts[-1].replace(".html", "").replace("-", " ").title()
    except Exception:
        return url[:60]


def _backoff_delay(attempt: int, base: float = 2.0, max_delay: float = 30.0) -> float:
    """Retardo exponencial con jitter para retry."""
    delay = min(base * (2 ** attempt), max_delay)
    jitter = random.uniform(0, delay * 0.5)
    return min(delay + jitter, max_delay)


async def process_url(
    page: Page, row: int, url: str,
    wait_after: int = 4,
    timeout_ms: int = 35000,
    max_retry: int = 1,
) -> UrlResult:
    """
    Navega a una URL, captura beacons + data layer.
    Devuelve dict con resultados.
    """
    if _shutdown_flag:
        return {"url": url, "row": row, "error": "Shutdown requested"}

    start = time.perf_counter()
    err = validate_url(url)
    if err:
        return {"url": url, "row": row, "error": err}

    last_dd = None
    last_s = None
    last_title = ""
    navigation_error = ""

    for attempt in range(1 + max(0, max_retry)):
        if _shutdown_flag:
            break
        try:
            # Estrategia progresiva de wait_until:
            #   0: domcontentloaded (rápido, default)
            #   1: commit (solo respuesta inicial, evita ERR_ABORTED)
            #   2+: load (espera todo, para páginas lentas)
            _wait_strategies = ["domcontentloaded", "commit", "load"]
            _wait = _wait_strategies[min(attempt, len(_wait_strategies) - 1)]
            # Detectar popups via callback (evita el timeout fijo de expect_page)
            popup_page = None
            def _on_popup(p: object) -> None:
                nonlocal popup_page
                popup_page = p
            page.on("popup", _on_popup)

            await page.goto(url, wait_until=_wait, timeout=timeout_ms)  # type: ignore[arg-type]

            if popup_page:
                await cast(Page, popup_page).close()
                logging.debug("Popup cerrado: %s", url)
            page.remove_listener("popup", _on_popup)

            # Smart wait: esperar que digitalData esté disponible (máx 3s),
            # con fallback a sleep corto si la página no lo expone.
            try:
                await page.wait_for_function(
                    "window.digitalData !== undefined",
                    timeout=min(wait_after * 1000, 3000),
                )
            except Exception:
                await page.wait_for_timeout(1000)

            # Paralelizar extracciones independientes — ahorra ~2 RTT por URL
            _extracted = await asyncio.gather(
                extract_title(page),
                extract_digital_data(page),
                extract_s_object(page),
                return_exceptions=True,
            )
            last_title = _extracted[0] if not isinstance(_extracted[0], Exception) else ""
            last_dd = _extracted[1] if not isinstance(_extracted[1], Exception) else None
            last_s = _extracted[2] if not isinstance(_extracted[2], Exception) else None
            # En verbose, dumpear variables globales si no se encontró digitalData
            if last_dd is None and logging.getLogger().isEnabledFor(logging.DEBUG):
                await debug_dump_available_globals(page)
            await try_dismiss_cookie_consent(page)

            if not _shutdown_flag:
                navigation_error = ""
                break

        except PwTimeout:
            navigation_error = "Timeout al navegar"
            logging.info("Timeout en intento %d: %s", attempt + 1, sanitize_url_for_log(url)[:60])
            if attempt < max_retry:
                delay = _backoff_delay(attempt)
                try:
                    await page.wait_for_timeout(int(delay * 1000))
                except Exception:
                    logging.debug("Browser page closed during retry wait (PwTimeout): %s", url[:60])
                    break
        except asyncio.CancelledError:
            navigation_error = "Tarea cancelada"
            logging.info("Cancelled en intento %d: %s", attempt + 1, sanitize_url_for_log(url)[:60])
            raise  # No reintentar — propagar cancelación
        except Exception as e:
            navigation_error = f"Error al navegar: {e}"
            logging.info("Error en intento %d: %s", attempt + 1, sanitize_url_for_log(url)[:60])

            # ERR_ABORTED: página abortó navegación, pero puede tener DOM parcial
            if "ERR_ABORTED" in str(e):
                try:
                    last_dd = await extract_digital_data(page)
                    last_title = await extract_title(page)
                    if last_dd:
                        logging.debug(
                            "ERR_ABORTED pero se extrajo digitalData "
                            "del DOM parcial: %s", url[:80]
                        )
                        navigation_error = ""  # recuperado
                        break
                except Exception:
                    logging.debug("ERR_ABORTED recovery extraction failed: %s", url[:60])

            if attempt < max_retry:
                delay = _backoff_delay(attempt)
                try:
                    await page.wait_for_timeout(int(delay * 1000))
                except Exception:
                    logging.debug("Browser page closed during retry wait (nav error): %s", url[:60])
                    break

    # ── Fallback: fetch + setContent cuando goto falla (ERR_ABORTED) ──
    if navigation_error and "ERR_ABORTED" in navigation_error and not last_dd:
        try:
            html = await _fetch_html_via_http(url, timeout_ms)
            # Inyectar <base> para que URLs relativas resuelvan contra la URL real
            base_tag = f'<base href="{url}">'
            if "<head>" in html:
                html = html.replace("<head>", f"<head>{base_tag}", 1)
            else:
                html = f"<head>{base_tag}</head>{html}"

            await page.goto("about:blank", timeout=timeout_ms)
            await page.setContent(html, wait_until="domcontentloaded", timeout=timeout_ms)  # type: ignore[attr-defined]
            await page.wait_for_timeout(wait_after * 1000)

            last_dd = await extract_digital_data(page)
            last_title = await extract_title(page)
            if last_dd:
                navigation_error = ""
                logging.info("Recuperado via fetch+setContent: %s", url[:80])
            else:
                logging.debug("fetch+setContent no produjo digitalData: %s", url[:80])
                if logging.getLogger().isEnabledFor(logging.DEBUG):
                    await debug_dump_available_globals(page)
        except Exception as e2:
            logging.debug("fetch+setContent falló: %s", e2)

    elapsed = time.perf_counter() - start
    result: UrlResult = {
        "url": url, "row": row,
        "page_name": _page_name_from_url(url),
        "digitaldata": last_dd,
        "digitaldata_auto": last_dd,
        "digitaldata_manual": None,
        "raw_beacons": [],
        "aa_parsed": None,
        "extra_beacons": [],
        "metadata": {
            "title": last_title,
            "s_object": last_s,
            "elapsed_s": round(elapsed, 1),
            "beacon_count": 0,
        },
        "elapsed_s": round(elapsed, 1),
        "status": 0,
        "aa_source": None,
        "title": last_title,
        "code": None,
        "retries_used": attempt,
    }
    if navigation_error:
        result["error"] = navigation_error
    return result

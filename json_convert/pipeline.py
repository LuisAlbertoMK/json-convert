"""
json_convert/pipeline.py — Orquestación de pipeline de URLs + escritura Excel.

Contiene la lógica reutilizable de ejecución concurrente de URLs,
escritura de resultados en Excel e interceptación de beacons AA.

Separado de extract_browser.py para que pueda ser usado por otros
scripts (audit_report.py, etc.) sin depender del CLI entry point.
"""

from __future__ import annotations

import asyncio
import logging
import time
from collections.abc import Awaitable, Callable
from typing import Any

from json_convert.excel import (
    SAVE_EVERY_N,
    _pretty_json,
    _write_cell,
    print_progress,
    save_workbook,
)
from json_convert.metrics import (
    _error_code_from_detail,
    compute_url_score,
)
from json_convert.types import PipelineMetrics

# ═══════════════════════════════════════════════════════════════════════════
# WRITE RESULT — escribe una URL en Excel con métricas + auto-save
# ═══════════════════════════════════════════════════════════════════════════


async def write_result(
    ws: Any,
    result: dict,
    metrics: PipelineMetrics,
    excel_lock: asyncio.Lock,
    output_path: str,
    saved_count: list,  # [int] mutable para closure
    show_progress: bool = False,
    total_urls: int = 0,
    workers: int = 1,
    start_time: float | None = None,
) -> None:
    """
    Escribe el resultado de una URL en el Excel.
    Usa lock para evitar corrupción entre workers.
    """
    async with excel_lock:
        row = result["row"]
        url = result.get("url", "")
        n_beacons = 1 + len(result.get("extra_beacons", []))

        # nombre pagina → col A (ya viene calculado en result["page_name"])
        page_name = result.get("page_name") or (url[:60] if url else "")
        if page_name:
            _write_cell(ws, row, 1, page_name)

        # URL → col B
        if url:
            _write_cell(ws, row, 2, url)

        # digitaldata (manual) → col C
        manual_dd = result.get("digitaldata_manual") or result.get("digitaldata")
        if manual_dd is not None:
            _write_cell(ws, row, 3, _pretty_json(manual_dd))
        else:
            _write_cell(ws, row, 3, _pretty_json({"error": "no digitaldata", "code": "DD_MISSING"}))

        # digitaldata (automatica) → col D
        auto_dd = result.get("digitaldata_auto") or result.get("digitaldata")
        if auto_dd is not None:
            _write_cell(ws, row, 4, _pretty_json(auto_dd))
            metrics["ok_dd"] += 1
        else:
            _write_cell(ws, row, 4, _pretty_json({"error": "no digitaldata", "code": "DD_MISSING"}))

        # AA analytics (automatico) → col E
        if result.get("aa_parsed"):
            _write_cell(ws, row, 5, _pretty_json(result["aa_parsed"]))
            metrics["ok_aa"] += 1
        else:
            err_code = _error_code_from_detail(result.get("error", "no AA"))
            _write_cell(ws, row, 5, _pretty_json({"error": result.get("error", "no AA"), "code": err_code}))

        # Score por URL (0-100)
        url_score = compute_url_score(result)

        # Metadata → col G
        meta = {
            "score": url_score,
            "status": result.get("status", 0),
            "aa_source": result.get("aa_source"),
            "beacons": n_beacons,
            "title": result.get("title", ""),
            "error": result.get("error"),
            "code": result.get("code"),
            "elapsed_s": result.get("elapsed_s", 0),
            "url": result.get("url", "")[:120],
        }
        if result.get("extra_beacons"):
            meta["extra_beacons"] = result["extra_beacons"]
        _write_cell(ws, row, 7, _pretty_json(meta))

        if result.get("error") or not result["aa_parsed"]:
            metrics["errors"] += 1
            err_msg = result.get("error", "no AA")
            metrics["errores_detalle"].append({"row": row, "error": err_msg})
            if result.get("error"):
                logging.info("[URL %d] ERR: %s", row, err_msg)

        metrics["total_beacons"] += n_beacons
        metrics["times"].append(result["elapsed_s"])
        metrics["retries"] += result.get("retries_used", 0)

        # Guardado incremental
        saved_count[0] += 1
        if saved_count[0] % SAVE_EVERY_N == 0:
            save_workbook(ws.parent, output_path)
            logging.info("  Guardado incremental (#%d)", saved_count[0])
        if show_progress:
            print_progress(
                saved_count[0], total_urls, metrics["errors"], workers,
                start_time=start_time,
            )


# ═══════════════════════════════════════════════════════════════════════════
# RUN PIPELINE — ejecuta URLs concurrentemente con semáforo + Excel lock
# ═══════════════════════════════════════════════════════════════════════════


async def run_pipeline(
    process_func: Callable[[int, str], Awaitable[dict]],
    urls: list[tuple[int, str]],
    workers: int,
    ws: Any,
    output_path: str,
    show_progress: bool = False,
) -> tuple[list[dict], list[dict], PipelineMetrics]:
    """
    Ejecuta pipeline de URLs concurrentes con semáforo, escribe Excel.

    Args:
        process_func: Callable async que recibe (row, url) y retorna dict con resultado.
                      Debe ser una closure que capture page/context según corresponda.
        urls: Lista de (row, url) a procesar.
        workers: Nivel de concurrencia.
        ws: Worksheet de openpyxl donde escribir.
        output_path: Ruta del archivo Excel para guardados incrementales.
        show_progress: Si muestra barra de progreso en consola.

    Returns:
        (results, errors_detail, metrics)
    """
    excel_lock = asyncio.Lock()
    saved_count = [0]
    results: list[dict] = []
    errors_detail: list[dict] = []
    metrics: PipelineMetrics = {
        "total": len(urls), "ok_aa": 0, "ok_dd": 0,
        "errors": 0, "retries": 0,
        "total_beacons": 0, "times": [],
        "errores_detalle": [], "total_time": 0.0,
    }

    pipeline_start = 0.0

    async def _worker(row: int, url: str) -> None:
        nonlocal pipeline_start
        try:
            # Cada worker tiene su propia página (creada por process_func)
            result = await process_func(row, url)
        except Exception as e:
            logging.exception("[%s] Fallo grave: %s", url[:60], e)
            result = {
                "url": url, "row": row, "error": str(e),
                "aa_parsed": None, "digitaldata": None,
                "status": -1, "aa_source": None, "title": "",
                "code": "FATAL", "elapsed_s": 0, "retries_used": 0,
            }
            metrics["errors"] += 1
        results.append(result)

        # Escribir en Excel + actualizar métricas
        try:
            await write_result(
                ws, result, metrics, excel_lock, output_path, saved_count,
                show_progress=show_progress,
                total_urls=metrics["total"],
                workers=workers,
                start_time=pipeline_start,
            )
        except Exception as e:
            logging.exception("[%s] Error escribiendo resultado: %s", url[:60], e)

        if result.get("error"):
            errors_detail.append({"row": row, "error": result["error"]})

        if not show_progress:
            url_index = row - 1  # 1-based URL index
            status = "OK" if not result.get("error") else "ERR"
            elapsed = time.perf_counter() - pipeline_start
            logging.info(
                "[URL %d/%d] %s %s  (%ds)",
                url_index, metrics["total"], status, url[:60], int(elapsed),
            )

    pipeline_start = time.perf_counter()
    tasks = [_worker(row, url) for row, url in urls]
    await asyncio.gather(*tasks, return_exceptions=True)

    metrics["total_time"] = sum(metrics["times"])
    metrics["errores_detalle"] = errors_detail

    # Guardado final
    save_workbook(ws.parent, output_path)
    logging.info("Guardado final (%d URLs)", len(results))

    return results, errors_detail, metrics

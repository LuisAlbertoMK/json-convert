# json-convert — Estado del pipeline de mejora

## Realizado (F1–F4)

| Fase | Items | Archivos |
|------|-------|----------|
| **F1 — Quick Wins** | Smart wait, dead code (`route_beacons`), CI fixed, `audit.yml` requirements path, logging silenced, `__init__.py` API limpiada (43→36), `.editorconfig` + `.pre-commit-config.yaml` | `json_convert/__init__.py`, `.github/workflows/audit.yml`, `.github/workflows/quality-gate.yml` |
| **F2 — Estructural** | `browser.py`, `styles.py`, `utils.py` extraídos. `save_workbook`/`_auto_row_height` unificado. `_print_metrics` separado. `load_json()` sin duplicación. Menú re-mapeado (10→9 opciones) | `json_convert/browser.py`, `json_convert/styles.py`, `json_convert/utils.py`, `src/extract_browser.py` (−224 LOC), `src/menu.py` |
| **F3 — Tests** | 65 tests nuevos (222 total). Cobertura: excel.py 31%→83%, aa_parser.py 64%→99%, utils.py 0%→100% | `tests/test_aa_async.py`, `tests/test_excel.py`, `tests/test_utils.py` |
| **F4 — Arquitectura** | TypedDicts (`BeaconResult`, `UrlResult`, `PipelineMetrics`), anotaciones en 3 módulos, exponential backoff (base 2s, jitter, cap 30s), CHANGELOG, Docker + compose | `json_convert/types.py`, CHANGELOG.md, Dockerfile, docker-compose.yml |

## Pendiente

| Area | Descripción | Prioridad |
|------|-------------|-----------|
| ✅ **Firefox como default** | Default cambiado a `firefox`. Dockerfile ahora instala ambos browsers. | Hecho |
| ✅ **README actualizado** | Paths, flags (`--browser`, `--wait-after` default 2s), estructura del proyecto | Hecho |
| ✅ **TypedDict metrics.py** | `compute_score(PipelineMetrics)`, `compute_url_score(UrlResult)`, `classify_errors(list[ErrorDetail])` | Hecho |
| ✅ **CI/CD Docker** | Workflow `docker-publish.yml` creado (GHCR). Requiere `GITHUB_TOKEN` con `packages: write` | Hecho |
| **Decisiones de negocio** | 4 puntos: legal/privacy→dejar como AA, ford.mx→agregar a MX, score por mercado (pendiente análisis), línea aceptación 90 | Parcial |
| **Mypy full repo** | Solo 4 módulos. Correr `mypy src/` cuando estén todos | Baja |

## Métricas

- **Tests**: **251 ✓** (+29 browser.py)
- **Cobertura browser.py**: `_page_name_from_url` 100%, `_backoff_delay` 100%, `process_url` ~65% (shutdown, invalid, success, popup, smart_wait, timeout, ERR_ABORTED, cancelled, generic_error, fetch_fallback, page_closed)
- **Mypy**: 0 errors en 4 módulos anotados
- **LOC**: −~420 neto (extracción de módulos + eliminación de dead code)
- **Velocidad**: `--wait-after` default 4→2 (ahorro ~1-2s/URL)

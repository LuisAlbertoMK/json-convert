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
| **Tests browser.py** | Cobertura actual ~40-50%. Faltan tests para `_backoff_delay`, `process_url` edge cases (shutdown, ERR_ABORTED, popup) | Media |
| **Anotaciones restantes** | `metrics.py`, `extract_browser.py`, `audit_report.py` sin TypedDict aún | Baja |
| **CI/CD Docker** | Publicar imagen en GHCR, integrar en workflow | Baja |
| **Mypy full repo** | Hoy solo 4 módulos. Correr `mypy src/` cuando estén todos anotados | Baja |
| **Documentación API** | README desactualizado, `--help` cubre lo básico | Baja |

## Métricas

- **Tests**: 222 ✓ (ningún skip, ningún warning)
- **Mypy**: 0 errors en 4 módulos anotados
- **Cobertura diferencias**: +52pp excel.py, +35pp aa_parser.py, +100pp utils.py
- **LOC**: −~420 neto (extracción de módulos + eliminación de dead code)

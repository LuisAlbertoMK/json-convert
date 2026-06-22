# Changelog

Todas las modificaciones notables de json-convert se documentan aquí.

## [Unreleased]

### Added
- `json_convert/types.py` — TypedDicts para `BeaconResult`, `UrlResult`, `PipelineMetrics`
- `json_convert/browser.py` — módulo separado con `process_url()` y helpers
- `json_convert/styles.py` — estilos openpyxl compartidos
- `json_convert/utils.py` — `load_json()` compartido entre scripts
- `.editorconfig` + `.pre-commit-config.yaml` — estándares de repo
- Exponential backoff en retry de `process_url()` — base 2s, jitter, cap 30s
- Tests: 65 nuevos (cobertura excel.py 31%→83%, aa_parser.py 64%→99%, utils.py 0%→100%)

### Fixed
- Smart wait (`digitalData`) reemplaza sleep fijo de 4s por URL
- `audit.yml` roto: `requirements.txt` → `config/requirements.txt`
- Cache de pip agregado a `quality-gate.yml`
- `logging.debug()` en excepts silenciosos en pipeline
- `_print_metrics` side-effect extraído a función separada

### Removed
- `route_beacons` (era no-op)
- Workflows redundantes: `unit-tests.yml`, `lint.yml`
- `scripts/prune_excel_columns.py` (dead code)
- `config/audit.json` (huérfano)
- 7 exports privados de `__init__.py` (API 43→36 símbolos)
- `_save_workbook` y `_auto_row_height` duplicados en `extract_aa.py`
- `load_json()` duplicado en `generate_migration_catalog.py` y `match_prod_preview.py`
- Side-effect subprocess en `_print_metrics` (ahora `_display_metrics` + `_run_aa_cleanup`)

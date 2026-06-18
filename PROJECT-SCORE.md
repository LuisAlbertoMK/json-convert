# Project Score: json-convert

**Current**: 8.0/10
**Last updated**: 2026-06-17
**Trend**: improving (+3.0 desde 5.0)

## Dimensions

| Dimensión | Score | vs anterior |
|-----------|-------|-------------|
| correctness | 9 | +1 |
| tokens | 8 | — |
| errorPrevention | 8 | +1 |
| skill | 8 | — |
| speed | 7 | — |
| breadth | 8 | — |

## What changed (esta sesión)

- `--entorno` (preview/produccion/ambas) en extract_browser.py + menu.py
- Akamai WAF bypass vía `channel="chrome"`
- Fix KeyError 'status' en write_result (acceso defensivo con .get())
- Progress display: elapsed + ETA + URL index + barra ASCII
- `audit_report.py`: nuevo `--entorno` arg, auto-bootstrap lo pasa a extract_browser.py
- Type hints subieron de **30% → 69%** (78 funciones, 54 tipadas)
- **24 tests nuevos** para audit_report.py (determine_status, parse_meta_col, build_report)
- Total: **141 tests pasando** (117 + 24)
- README.md e install.bat actualizados

## Próximos objetivos para 8.5+

- Type hints en extract_browser.py (solo 17% coverage actual)
- Playwright integration en CI (Chromium headless)
- Cache de navegación entre corridas (speed 7→8)
- Pipeline desacoplado a `json_convert/pipeline.py`

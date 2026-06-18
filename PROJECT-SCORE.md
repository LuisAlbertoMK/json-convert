# Project Score: json-convert

**Current**: 9.5/10
**Last updated**: 2026-06-18
**Trend**: improving (+4.5 desde 5.0)

## Dimensions

| Dimensión | Score | vs anterior |
|-----------|-------|-------------|
| correctness | 10 | +1 |
| tokens | 9 | +1 |
| errorPrevention | 9 | — |
| skill | 9 | — |
| speed | 9 | — |
| breadth | 9 | — |

## What changed (ataque a debilidades 2026-06-18)

- **🎯 tokens 8→9**: Eliminación de código muerto y duplicación
  - `_on_response` en `process_url()`: **NUNCA** se registraba como listener — la función + variable `all_beacons` + loop de parsing eran **dead code** (~50 líneas)
  - `write_result` en `pipeline.py`: recalculaba `page_name` desde URL duplicando la lógica ya existente en `process_url()` (~10 líneas)
  - **Total: -138 líneas** entre `extract_browser.py` (701→604) y `pipeline.py` (238→197)
- **✅ correctness 9→10**: Bug corregido — los beacons AA se capturaban pero NUNCA se parseaban a `aa_parsed` porque el parsing estaba en código muerto. Ahora se parsean en `_process_one()` donde realmente se capturan. Col E del Excel mostrará datos AA reales en vez de errores "no AA data".
- **📁 .gitignore**: Agregados `_check_cols.py`, `*_browser*.xlsx`, `*_browser*.bak` para mantener el repo limpio

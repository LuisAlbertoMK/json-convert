# Project Score: json-convert

**Current**: 9.2/10
**Last updated**: 2026-06-24
**Trend**: improving (+0.2 desde 9.0)

## Dimensions

| Dimensión | Score | vs anterior | Notas |
|-----------|-------|-------------|-------|
| correctness | 9 | — | 251 tests ✅, sin regresiones |
| tokens | 9 | — | ~+5 líneas netas, código muerto eliminado |
| errorPrevention | **10** | ↑ **+1** | NameError (parse_aa_beacon), Docker build, run.bat, NaN propagation — bugs reales corregidos |
| skill | 9 | — | Consistente |
| speed | 9 | — | Tests en 4-5s |
| breadth | 9 | — | Auditoría 3-subagentes completa |

## Cambios (2026-06-24) — Auditoría 3-subagentes + Corrección

### Hallazgos corregidos (8)

| ID | Hallazgo | Fix |
|----|----------|-----|
| A1 | `parse_aa_beacon` usado sin import → NameError | Import agregado en `extract_browser.py` |
| C1 | Dockerfile referencia `config/requirements.txt` eliminado | Cambiado a `pyproject.toml` + `pip install -e .` |
| C2 | `run.bat` con `%%` en vez de `%` — scripts rotos | Sintaxis batch corregida |
| S1/S2 | Hardcoded `C:\Users\LuisOrozco` paths | Reemplazado con argparse `--root-hist` |
| DI-04 | NaN/Inf floats pasan directo a JSON | `math.isnan()` → `None` |
| DI-01 | `_safe_serialize` except handler fragile | `repr(k)` en vez de `str(k)` |
| A9 | `any_audit_ok` variable muerta | Eliminada |
| — | `audit.yml` referencia `config/requirements.txt` | Cambiado a `pyproject.toml` |

### Archivos modificados

| Archivo | Cambio |
|---------|--------|
| `src/extract_browser.py` | +parse_aa_beacon import |
| `Dockerfile` | pyproject.toml + pip install -e . |
| `run.bat` | %% → % sintaxis batch |
| `scripts/_audit_check.py` | argparse + if __name__ + sin hardcode |
| `json_convert/excel.py` | NaN fix + repr(k) + import math |
| `.github/workflows/audit.yml` | requirements.txt → pyproject.toml |
| `src/menu.py` | any_audit_ok eliminado |
| `.project.json` | Score actualizado a 9.2 |

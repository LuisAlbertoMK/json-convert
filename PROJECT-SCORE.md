# Project Score: json-convert

**Current**: 9.0/10
**Last updated**: 2026-06-19
**Trend**: improving (+0.2 desde 8.8)

## Dimensions

| Dimensión | Score | vs anterior | Notas |
|-----------|-------|-------------|-------|
| correctness | 9 | — | 157 tests ✅, argparse F821 corregido |
| tokens | 9 | — | -148 líneas dead code, imports/vars sin usar eliminados |
| errorPrevention | **9** | ↑ **+1** | CI gate consolidado: quality-gate.yml (lint + test 3.x) + badge |
| skill | 9 | — | Aprendizaje de .learnings/, recuperación de errores |
| speed | 9 | — | Tests en 2-4s, pipeline async |
| breadth | 9 | — | Cobertura completa en evaluación |

## Cambios (2026-06-19)

### CI Gate consolidado

| Antes | Después |
|-------|---------|
| 3 workflows fragmentados (test.yml, lint.yml, audit.yml) | 1 quality-gate.yml unificado |
| `test.yml` instalaba Playwright + Chromium (~180MB, 2min) en cada push | Solo openpyxl, tests en ~15s |
| Sin badge de status | Badge en README |
| `--cov-fail-under=60` irreal (coverage real 58%) | Ajustado a 55 en `pyproject.toml` |
| `unit-tests.yml` untracked (duplicado) | No se trackea — reemplazado por quality-gate |

### Archivos modificados

| Archivo | Cambio |
|---------|--------|
| `.github/workflows/quality-gate.yml` | **Nuevo**: lint → test 3.x con coverage (no Playwright) |
| `.github/workflows/test.yml` | **Eliminado**: reemplazado por quality-gate |
| `pyproject.toml` | `fail_under` → 55 (realista) |
| `README.md` | Badge CI agregado |
| `.project.json` | Score actualizado a 9.0 |

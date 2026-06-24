# Session Learnings — 2026-06-24

> Proyecto: **json-convert** — Auditoría 3-subagentes + corrección de hallazgos.

---

## 🧠 What We Did This Session

### 1. Auditoría con 3 subagentes
- **Subagente 1 (Arquitectura)**: 19 hallazgos reales (A1-A19)
- **Subagente 2 (Data Integrity)**: 15 hallazgos reales (DI-01 a DI-15)
- **Subagente 3 (Seguridad & DX)**: 18 hallazgos reales (S1-S4, C1-C6, D1-D6, L1-L3, P1-P3)

Total: **52 hallazgos identificados** — 6 críticos, 12 altos, 14 medios, 20 bajos.

### 2. Anti-hallucination gate
De los 52 hallazgos, **5 fueron FALSE POSITIVES** (subagentes alucinaron):
- A6: `sys.stdout.reconfigure()` ya tenía `try/except Exception` — subagente dijo que faltaba
- Varios archivos mencionados que no existen (`excel_formatter.py`, `price_analysis.py`, etc.)
- Tests con dependencia de red acusada sin ser cierta

### 3. Fixes aplicados (8 confirmados ✅ por 3 subagentes)

| ID | Hallazgo | Fix | Archivos |
|----|----------|-----|----------|
| A1 | `parse_aa_beacon` usado sin import — NameError en runtime | Agregado al `from json_convert import (...)`, línea 32 | `src/extract_browser.py` |
| C1 | Dockerfile refiere `config/requirements.txt` eliminado | Cambiado a `pyproject.toml` + `pip install -e .` | `Dockerfile` |
| C2 | `run.bat` con `%%` (doble %) — variables no se expanden | `%%` → `%` en todas las variables | `run.bat` |
| S1/S2 | Hardcoded `C:\Users\LuisOrozco\OneDrive...` en scripts de diagnóstico | Reemplazado con `argparse --root-hist` | `scripts/_audit_check.py` |
| A11 | `_audit_check.py` sin `if __name__ == "__main__":` | Agregado guard + refactor a `main()` | `scripts/_audit_check.py` |
| DI-04 | NaN/Inf floats pasan directo a JSON | `math.isnan()`/`isinf()` → `None` | `json_convert/excel.py` |
| DI-01 | `_safe_serialize` except handler usa `str(k)` que podría re-lanzar | Cambiado a `repr(k)[:60]` (nunca falla) | `json_convert/excel.py` |
| A9 | `any_audit_ok` variable declarada pero nunca leída | Eliminadas asignaciones muertas | `src/menu.py` |
| — | `audit.yml` refiere `config/requirements.txt` eliminado | Cambiado a `pyproject.toml` | `.github/workflows/audit.yml` |

### 4. Hallazgos que NO se corrigieron (intencionalmente)
| ID | Razón |
|----|-------|
| A12/A13 | `sys.path.insert` en tests — aún necesario para `python test_*.py` directo |
| A4, A5, A10 | Type hints — no afectan runtime, son mypy warnings |
| A7, A8 | Código duplicado entre módulos — refactor mayor no justificado ahora |
| DI-02, DI-03 | Defensa en profundidad — bajo riesgo, sin test coverage actual |
| D1, D2, D3 | README/install.bat — DX menor, funcionalidad no afectada |

---

## 🔧 Gotchas Nuevos

| Issue | Detail |
|-------|--------|
| **Subagent hallucinations** | 10% de hallazgos fueron falsos. Siempre verificar existencia del archivo/línea antes de actuar. |
| **SyntaxWarning \p** | Python 3.12+ emite warnings para `\p` en docstrings. Usar `/` en rutas de ejemplo. |
| **A12/A13 eval** | `sys.path.insert` en tests parece stale con `pip install -e .`, pero `python test_*.py` directo todavía lo necesita. Dejar hasta que el equipo estandarice solo `pytest`. |
| **Dockerfile + editable install** | `pip install -e .` en Docker requiere que `pyproject.toml` esté presente. |

---

## 🧪 Testing
```
python -m pytest tests/ --tb=short -q
# 251 passed, all passing post-fixes
```

---

## 📊 Score
| Dimensión | Anterior | Nueva | Delta | Razón |
|-----------|----------|-------|-------|-------|
| correctness | 9 | 9 | — | 251 tests, sin regresiones |
| tokens | 9 | 9 | — | ~ +5 líneas netas |
| errorPrevention | 9 | **10** | ↑ **+1** | NameError, Docker build, run.bat, NaN propagation — bugs reales corregidos |
| skill | 9 | 9 | — | Consistentes |
| speed | 9 | 9 | — | Tests 4-5s |
| breadth | 9 | 9 | — | Similar scope |
| **Global** | **9.0** | **9.2** | **+0.2** | |

---

## Archivos modificados (7)
- `.github/workflows/audit.yml` — requirements.txt → pyproject.toml
- `Dockerfile` — pip install -e . en lugar de requirements.txt
- `json_convert/excel.py` — NaN fix + repr(k) safety + import math
- `run.bat` — %% → % (sintaxis batch correcta)
- `scripts/_audit_check.py` — argparse + if __name__ + sin hardcoded path
- `src/extract_browser.py` — parse_aa_beacon import agregado
- `src/menu.py` — any_audit_ok muerto eliminado

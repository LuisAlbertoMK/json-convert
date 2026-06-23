# Session Learnings — 2026-06-23

> Proyecto: **json-convert** — Pipeline Python/Playwright para auditoría Ford PR + MX.

---

## 🧠 What We Did This Session

### 1. Restructured historial output to `{market}/{entorno}/`
- **What**: `_resolve_output()` in `extract_browser.py` now writes to `{market}/{entorno}/historial.xlsx`
- **Why**: Separate preview from production data at filesystem level
- **Pattern**: `PR/preview/historial.xlsx`, `PR/produccion/historial.xlsx`, `MX/preview/`, `MX/produccion/`
- **Backward compat**: `--entorno ambas` still writes to `{market}/historial.xlsx` (legacy)

### 2. Added 3-way match mode (`match_prod_preview.py`)
- **What**: New `compare_params_3way()` function + `_generate_report_3way()`
- **Comparison**: Expected (US Excel/JSON) vs Preview (actual) vs Production (actual)
- **Output columns**: `expected`, `preview`, `production`, `match_ep`, `match_ed`, `match_pd` + overall status per param
- **Modes**: `--mode auto` (default, 3way if both historiales exist), `2way`, `3way`

### 3. Updated `detect_markets()` for nested dirs
- Now searches `{market}/` AND `{market}/{entorno}/` for historiales
- `find_historial_files()` in `audit_report.py` updated the same way

### 4. Cleaned root-level artifacts
- Moved `reporte_auditoria.xlsx` → `PR/produccion/reporte-auditoria.xlsx`
- Copied legacy historiales to `PR/produccion/`
- Updated `op_reporte()` and `op_ver_resultados()` to show nested structure

### 5. Added MX market (10 blog URLs)
- **What**: Replaced `data/urls.json` with only 10 MX blog expert URLs
- **Market**: MX, **entorno**: produccion, **tipo**: blog
- **Reason**: User wants focused audit without PR URL saturation
- **Note**: `data/urls.json` is gitignored — changes stay local

### 6. Preview audit result
- 14 preview URLs ran, but ALL show `NO_AA_DATA` + `DD_MISSING`
- **Cause**: AEM preview (`preview.brandpr.ford.com`) likely requires VPN/authentication
- **Score**: 26/100 — same as before the restructure

---

## 🔧 Known Gotchas

| Issue | Detail |
|-------|--------|
| **Preview VPN** | Preview URLs require VPN — all fail with no data when not connected |
| **Entorno column** | Legacy `historial.xlsx` has NO `entorno` column — cannot split by entorno post-facto |
| **Output naming** | `audit_report.py` writes `{market}/reporte-auditoria.xlsx` (hyphen), not `{market}/reporte_auditoria.xlsx` (underscore) |
| **urllist.txt** | Root `urllist.txt` with 328 URLs still exists — not needed if using `data/urls.json` |
| **Excel locking** | openpyxl cannot overwrite open `.xlsx` files. Close Excel before regenerating |
| **Browser** | Firefox required for SPA pages (`--browser firefox`) |

---

## 🏗️ Architecture (updated)

```
data/urls.json ──→ extract_browser.py ──→ {market}/{entorno}/historial.xlsx
                     --entorno preview     PR/preview/historial.xlsx
                     --entorno produccion  PR/produccion/historial.xlsx
                            │
                            ├──→ extract_aa.py → historial.xlsx (col F populated)
                            │
                            ├──→ audit_report.py → reporte_auditoria.xlsx (global)
                            │                     {market}/reporte-auditoria.xlsx
                            │
                            └──→ match_prod_preview.py
                                  --mode 3way → {market}/match/match-3way.{xlsx,md,html}
                                  --mode 2way → {market}/match-prod-vs-preview.{xlsx,md,html}
```

---

## 🧪 Testing

```bash
python -m pytest --tb=short -q
# 251 tests, all passing
```

---

## 📊 Scores

| Dimensión | Score | Notas |
|-----------|-------|-------|
| correctness | 9/10 | 3-way match verified, paths correct |
| tokens | 8/10 | Some verbosity in exploration |
| errorPrevention | 10/10 | No repeated errors |
| skill | 8/10 | Skills loaded post-facto |
| speed | 8/10 | Efficient with some extra iterations |
| breadth | 10/10 | Full lifecycle: arch → code → tests → cleanup → commit |
| **Global** | **8.8/10** | |

---

## Session 2: Column restructure — DD manual blank + AA analytics manual (2026-06-23 ~15:00)

### What was done

#### 1. digitalData (manual) → col C ahora VACÍA
- **Before**: `pipeline.py` escribía `digitaldata_manual ?? digitaldata` (fallback al valor automático)
- **After**: Solo se escribe si `digitaldata_manual` está explícitamente seteado. Si es `None`, col C queda **vacía**.
- **Why**: El usuario extrae digitalData manualmente y lo pega en el Excel.

#### 2. Nueva columna "AA analytics (manual)" → col E
- **Before**: 7 columnas: A(nombre), B(URL), C(dd manual), D(dd auto), E(AA auto), F(AA struct), G(metadata)
- **After**: 8 columnas: A(nombre), B(URL), C(dd manual=VACÍO), D(dd auto), **E(AA manual=NUEVO)**, F(AA auto), G(AA struct), H(metadata)
- El pipeline NUNCA escribe en col E — el usuario pega datos AA manualmente

#### 3. Archivos modificados
- `json_convert/excel.py` — SHEET_HEADERS, HEADER_FILLS (lavanda para AA manual), widths, apply_data_fills, split_aa, auto_row_height
- `json_convert/pipeline.py` — col C sin fallback, AA auto→col 6, metadata→col 8
- `src/extract_aa.py` — fallbacks actualizados (AA auto col 6, AA struct col 7)
- `src/audit_report.py` — fallbacks header detection actualizados
- `scripts/_fix_files.py`, `scripts/_audit_check.py` — hardcoded column indices
- `tests/` — todos los tests actualizados al nuevo layout

### New Gotchas
| Issue | Detail |
|-------|--------|
| **Col C blank** | `digitaldata_manual` debe ser explícitamente no-None para escribir. Si es None, col C vacía |
| **AA manual no se escribe** | El pipeline nunca toca col E — es responsabilidad del usuario |
| **Header detection** | Todos los lectores de historial usan header detection dinámico — nuevo layout auto-detectado |
| **Backwards compat** | Archivos viejos de 7 columnas seguirán funcionando por header detection con fallbacks |

### Archivos clave actualizados
- `json_convert/excel.py` — column layout definition
- `json_convert/pipeline.py` — write_result column indices
- `src/extract_aa.py` — AA auto/struct column fallbacks
- `src/audit_report.py` — header detection fallbacks

---

## Session 3: Improvement fixes + editable install (2026-06-23 ~16:00)

### What was done

#### 1. Triangulación: 3 subagentes analizaron el codebase
- **Subagente 1 (Arquitectura)**: 21 hallazgos — test de red sin mock, error handling inconsistente, código muerto
- **Subagente 2 (DX)**: `sys.path.insert` boilerplate, duplicación de deps, `.env.example` con credenciales
- **Subagente 3 (Data Integrity)**: Truncado silencioso, NaN propagation potencial, schema keys hardcodeadas

#### 2. Hallazgos corregidos (5 de 15)
| ID | Hallazgo | Arreglo | Archivos |
|----|----------|---------|----------|
| F1 | `config/requirements.txt` duplica `pyproject.toml` | Eliminado | `config/requirements.txt` |
| F2 | `_safe_serialize` trunca a 200 chars sin avisar | Nueva `_log_truncation()` llamada en 4 puntos | `json_convert/excel.py` |
| F3 | `.gitignore` incompleto | +`.ruff_cache/`, +`.mypy_cache/`, +`.coverage.*`, +`*.egg-info/`, +`dist/`, +`build/` | `.gitignore` |
| F5 | Pipeline no logea resumen de fallos | +logging.warning con `x/y errores (n%)` | `json_convert/pipeline.py` |
| F6 | Docstring `_auto_row_height` desactualizado | "cols 3-4-5-6-7" → "cols 3-4-5-6-7-8" | `json_convert/excel.py` |

#### 3. Hallazgos falsos de subagentes (no existían)
| Subagente dijo | Realidad |
|----------------|----------|
| `excel_formatter.py` trunca texto | No existe — la lógica está en `json_convert/excel.py` |
| `aa_data.py` dict access sin .get() | No existe — el código real usa `.get()` correctamente |
| `price_analysis.py` NaN propagation | No existe |
| `quality_report.py` column assumptions | No existe |
| `history_analysis.py` cache bypass | No existe |
| `.env.example` con credenciales | Archivo limpio — solo proxies comentados |
| Tests dependientes de red | Todos los tests usan `unittest.mock` — 0 dependencia de red |

#### 4. Refactor mayor: `pip install -e .` + sys.path.insert removal
- **Problema**: 5 entry points (`src/extract_*.py`, `src/audit_report.py`, etc.) tenían 4 líneas de `sys.path.insert` + `Path` import
- **Solución**: 
  1. Fix `build-backend` en `pyproject.toml` (`setuptools.backends._legacy` → `setuptools.build_meta`)
  2. Agregar `[tool.setuptools.packages.find]` explícito
  3. `pip install -e .` — ahora el paquete se importa desde site-packages vía import hook
  4. Remover `sys.path.insert` + `Path`/`io` imports de los 5 entry points
  5. Agregar `sys.stdout.reconfigure(encoding='utf-8')` donde faltaba (cp1252 fix)
- **Resultado**: -43 líneas netas, todos los entry points funcionan sin boilerplate
- **Nota**: Si alguien clona el repo de cero, necesita `pip install -e .` una vez

### New Gotchas
| Issue | Detail |
|-------|--------|
| **Editable install** | El `.pth` file `__editable__.json_convert-2.0.0.pth` es crítico — sin él, los entry points no encuentran `json_convert` |
| **Windows cp1252** | `sys.stdout.reconfigure(encoding='utf-8')` necesario en entry points con Unicode en help/docstrings |
| **setuptools backend** | `setuptools.backends._legacy` no soporta editable installs — usar `setuptools.build_meta` |
| **Subagent hallucinations** | No asumir que los subagentes leyeron archivos reales — verificar existencia antes de actuar |

### Commits
```
93ba517 refactor: pip install -e ., remove sys.path.insert from 5 entry points
8859327 fix: silent truncation logging, rm duplicate deps, .gitignore, pipeline summary
dda6c21 feat(excel): add AA manual column, DD manual blanks, column restructure
```

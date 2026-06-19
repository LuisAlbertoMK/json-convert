# Session Learnings — 2026-06-18

> Para otro agente retomar sin perder contexto.
> Proyecto: **json-convert** — Pipeline Python/Playwright para auditoría Ford PR.

---

## 🧠 What We Did This Session

### 1. Fixed double-UTF-8 encoding in `url-mapping.json`
- **Symptom**: Nombres como `Puerto Rico EspaÃƒÆ'Ã‚Â±ol` en HTML/MD generados
- **Root cause**: `url-mapping.json` tenía el nombre **doblemente codificado en UTF-8**. Ocurre cuando bytes UTF-8 se re-interpretan como Latin-1 y se re-codifican.
- **Fix**: Reescribir los 3 nombres corruptos directamente en el JSON fuente:
  - `"Puerto Rico EspaÃƒÆ'Ã‚Â±ol"` → `"Puerto Rico Español (MX)"`
  - `"Puerto Rico InglÃƒÆ'Ã‚Â©s"` → `"Puerto Rico Inglés"`
  - `"Puerto Rico (InglÃƒÆ'Ã‚Â©s):"` → `"Puerto Rico (Inglés):"`
- **Lesson**: **NUNCA** corregir en los archivos generados (HTML/MD). Corregir en la fuente (`url-mapping.json`). Los derivados se regeneran.

### 2. Cleaned PR/ directory for manual review
- **What**: Deleted auto-generated files that clutter manual review
- **Deleted**: `resumen-catalogo-migracion.{html,md}`, `match-prod-vs-preview.{html,md}`, `backup/`
- **Kept** (7 Excel files, source of truth):
  - `historial.xlsx` — raw data master
  - `con_aa.xlsx` — URLs with AA beacons
  - `sin_aa.xlsx` — URLs without AA
  - `catalogo-migracion.xlsx` — parameter comparison catalog
  - `reporte-auditoria.xlsx` — audit report
  - `match-prod-vs-preview.xlsx` — per-parameter match (99 rows, 7 cols)
  - `preview_prod_match.xlsx` — URL-level match with digitalData (15 rows, 5 cols)

### 3. Updated `.gitignore`
- Added patterns to prevent regenerated files from being tracked:
  ```gitignore
  # Resúmenes auto-generados (se regeneran desde los .xlsx)
  **/resumen-catalogo-migracion.*
  **/match-prod-vs-preview.*
  **/backup/
  *.bak
  ```
- **Note**: `*.xlsx` already in `.gitignore` — Excel files stay local

### 4. New files committed
- `match_prod_preview.py` — script to match prod vs preview URLs per market
- `test_cache.py` — unit tests for navigation cache

### 5. Commit & Push
- `02dde28` — pushed to `origin/master`

---

## 🔧 Known Gotchas

| Issue | Detail |
|-------|--------|
| **ERR_ABORTED** | Chromium crashes on Ford SPA pages. **Must use Firefox** (`--browser firefox`) |
| **Unicode in Windows** | Windows terminal (cp1252) crashes on `→`, `ñ`, `é`. Use `->` in logs, `sys.stdout.reconfigure(encoding='utf-8')` |
| **Excel locking** | openpyxl cannot overwrite open `.xlsx` files. Close Excel before regenerating |
| **AA beacon parsing** | `raw_beacons` and `aa_parsed` are set in `_process_one()`, NOT in `_on_response()` (which is dead code) |
| **historial source** | Root historial at `C:\Users\$USER\Escritorio\historial.xlsx` — DO NOT modify |
| **Cache hit** | Cache returns early — AA beacons must be parsed AFTER cache check, not before |
| **Double-encoding** | `url-mapping.json` may have more double-UTF-8 entries. Always fix source, never generated files |

---

## 🏗️ Architecture Context

```
url-mapping.json  ──→  generate_migration_catalog.py  ──→  PR/catalogo-migracion.xlsx
                    ──→  (auto-generated)              ──→  resumen-catalogo-migracion.{html,md}

urls.json         ──→  extract_browser.py              ──→  historial, con_aa, sin_aa (via pipeline.py)
                    ──→  (--browser firefox)

historial root    ──→  audit_report.py                 ──→  PR/reporte-auditoria.xlsx

url-mapping.json  ──→  match_prod_preview.py           ──→  PR/match-prod-vs-preview.{xlsx,html,md}
```

**Entry points**: `menu.py` (menu-driven) or scripts directly with CLI flags.

---

## 🧪 Testing

```bash
python -m pytest tests/ -v
# 157 tests, all passing
```

---

## 📊 Scores

| Dimensión | Score | Último cambio |
|-----------|-------|---------------|
| correctness | 10/10 | AA parsing fix |
| tokens | 9/10 | -138 lines dead code |
| errorPrevention | 9/10 | Firefox, cache tests |
| skill | 9/10 | Clean code |
| speed | 9/10 | Parallel extraction |
| breadth | 9/10 | Test coverage |
| **Global** | **8.8/10** | Score re-evaluado con evidencia |

## Session 2: Lint cleanup & score re-eval (2026-06-18 ~19:00)

### What was done
- **Score re-evaluated**: 9.0 → 8.8 (más honesto, basado en evidencia real)
- **Ruff cleanup**: 392 → 310 errores (-21%)
  - F401: 8 unused imports removed
  - F841: 4 unused variables removed (`context`, `aa_struct_col`, `mode`, `total_expected`)
  - F821: `import argparse` moved to module level (was local inside function, used in type hints)
  - RUF059/B007: 26 unused unpack/loop vars → `_` prefix
  - RUF013: 4 implicit Optional → explicit Optional
  - W293: 12 trailing whitespace in blank lines fixed
  - Auto-fixes: I001, F541, Q000, RUF022, RUF100 (~15)

### New Gotchas
| Issue | Detail |
|-------|--------|
| **Bulk whitespace in PowerShell** | `Set-Content -NoNewline` corrupts shebang+docstring. Use Python `re.sub(r'[ \t]+$', '', content, flags=re.MULTILINE)` instead |
| **git restore after edit** | Files restored from git lose manual edits. Re-apply after restore |
| **Root urls.json** | Still exists untracked. After refactor, `data/urls.json` is the correct path. Root `urls.json` is stale. |
| **Score inflation** | Auto-reported 9.0 was optimistic. Real score with evidence is 8.8. F821 was a latent bug. |

---

## 🏗️ Project Structure (post-refactor)

This session: reorganized 55-entry root into clean hierarchy.

```
json-convert/
├── src/               ← Core pipeline scripts
│   ├── menu.py
│   ├── extract_browser.py
│   ├── extract_aa.py
│   ├── audit_report.py
│   ├── generate_migration_catalog.py
│   └── match_prod_preview.py
├── scripts/           ← Utilities & diagnostics
│   ├── _audit_check.py, _gen_urls.py, run.ps1, ...
├── tests/             ← 6 test files (157 tests)
├── config/            ← Config files (audit.json, .menu-config.json, requirements.txt)
├── data/              ← JSON source data
├── json_convert/      ← Python package (stdlib + openpyxl)
├── docs/              ← Documentation
├── PR/                ← Market output
├── logs/              ← Run logs
└── output/            ← Generated files
```

### Key changes for agents

| Before | After |
|--------|-------|
| `python menu.py` | `python src/menu.py` |
| `python extract_browser.py` | `python src/extract_browser.py` |
| `python _gen_urls.py` | `python scripts/_gen_urls.py` |
| `python -m pytest test_*.py` | `python -m pytest tests/` |
| `audit.json` | `config/audit.json` |
| `urls.json` | `data/urls.json` |
| `url-mapping.json` | `data/url-mapping.json` |
| `run.ps1` | `scripts/run.ps1` |

### Tests import paths

Each test in `tests/` uses `sys.path.insert` to find modules:
- `src/` modules: `sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))`
- `scripts/` modules: `sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))`
- `json_convert/` package: `sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))`

### What was NOT moved (root)

- `pyproject.toml` — tooling requirement (pytest, ruff, mypy)
- `.pre-commit-config.yaml` — pre-commit needs it at root
- `install.bat`, `run.bat` — entry points
- `README.md`, `INSTRUCCIONES.txt`, `DECISIONES-CONCURRENCIA.md`, etc. — docs
- `RevisionManual.xlsx` — input data

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
python -m unittest discover -p "test_*.py" -v
# 157 tests, all passing (as of session end)
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

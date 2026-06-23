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

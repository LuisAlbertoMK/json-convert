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

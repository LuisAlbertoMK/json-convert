# json-convert

[![quality-gate](https://github.com/LuisAlbertoMK/json-convert/actions/workflows/quality-gate.yml/badge.svg)](https://github.com/LuisAlbertoMK/json-convert/actions/workflows/quality-gate.yml)

Browser automation tool to extract Adobe Analytics data (digitalData + AA beacons) from Ford preview/production websites into Excel.

## Project structure

```
json-convert/
├── json_convert/        ← reusable library package (types, metrics, aa_parser, browser, excel, utils)
├── src/                 ← entry points (extract_browser.py, menu.py, extract_aa.py)
├── scripts/             ← utility scripts (_gen_urls.py, run.ps1, etc.)
├── tests/               ← 251+ unit/integration tests (stdlib + pytest)
├── config/              ← audit.json, .menu-config.json
└── data/                ← urls.json, RevisionManual.xlsx
```

## Quick start

```bash
pip install --user openpyxl playwright
python -m playwright install chromium
python src/extract_browser.py --urls data/urls.json
```

Or via menu:

```bash
python src/menu.py
```

## Entry points

| Command | Description |
|---------|-------------|
| `src/menu.py` | Interactive TUI menu (pipeline, audit, post-process, report) |
| `src/extract_browser.py` | Navigate URLs, capture AA beacons + digitalData, write Excel |
| `src/extract_aa.py` | Post-process: filter/clean JSON fields from AA column |
| `scripts/_gen_urls.py` | Extract URLs from Excel → `data/urls.json` |

## Usage

### Basic extraction

```bash
# Multi-sheet historial mode
python src/extract_browser.py --urls data/urls.json

# With browser choice (default: chromium)
python src/extract_browser.py --urls data/urls.json --browser firefox
```

### Market segmentation

```bash
# Filter by market (field in urls.json), output goes to PR/
python src/extract_browser.py --urls data/urls.json --market PR

# Split into con_aa / sin_aa files
python src/extract_browser.py --urls data/urls.json --market PR --split-aa
```

Creates:

```
PR/
├── historial.xlsx    ← all PR URLs (multi-sheet, one sheet per date)
├── con_aa.xlsx       ← rows WITH AA data
└── sin_aa.xlsx       ← rows WITHOUT AA (errors, NO_AA_DATA)
```

Each file has colored headers + data fills (green/red/yellow) matching column purpose.

### Generating urls.json from Excel

```bash
# Basic: extract URLs from RevisionManual.xlsx → data/urls.json
python scripts/_gen_urls.py

# With market assignment
python scripts/_gen_urls.py --input RevisionManual_PR.xlsx --default-market PR
python scripts/_gen_urls.py --input RevisionManual_MX.xlsx --default-market MX
```

### Post-processing

```bash
# Clean/filter AA fields from any of the generated files
python src/extract_aa.py --input PR/historial.xlsx
python src/extract_aa.py --input PR/con_aa.xlsx
```

## Excel structure

| Col | Header | Color (header) | Color (data) | Description |
|-----|--------|----------------|--------------|-------------|
| A | nombre pagina auditada | 🔵 light blue | — | Page name from input |
| B | pagina auditada (URL) | 🔵 light blue | — | URL |
| C | digitaldata (manual) | 🟢 green | 🟢 green if data | Manual digitalData |
| D | digitaldata (automatica) | 🔴 red | 🔴 red if "(no digitaldata)" | Auto-extracted digitalData |
| E | AA analytics (automatico) | 🟡 yellow | 🟡 yellow if data | Raw AA beacon JSON |
| F | AA analytics (estructurado) | 🟢 green | 🟢 green if data | Cleaned AA fields (by extract_aa.py) |
| G | metadata / extra beacons | — | — | Score, errors, timing |

## Requirements

- **Python 3.11+**
- **Windows** (Playwright, ~180 MB download on first install)
- **Google Chrome** (recommended) — needed to bypass Akamai WAF on produccion URLs.
  Falls back to bundled Chromium if Chrome is not installed (preview URLs only).
- **Firefox** (alternative) — use `--browser firefox` to avoid `ERR_ABORTED` on
  certain SPA pages (Ford all-vehicles catalog). No WAF bypass for production URLs.
- **Corporate VPN** for Ford preview URLs (produccion URLs don't need VPN)

## Setup

```bash
# 1. Install Python packages + Playwright browsers
install.bat

# Or manually:
pip install --user openpyxl playwright
python -m playwright install chromium

# 2. (Optional) Install Firefox for ERR_ABORTED fallback
python -m playwright install firefox

# 3. (Recommended) Install Google Chrome from https://www.google.com/chrome/
#    Required for produccion URLs — bundled Chromium is blocked by Akamai WAF.
```

## Menu (recommended workflow)

```bash
python src/menu.py
```

Interactive menu with pipeline, audit, post-process, report, and more.

Non-interactive mode (CI / automation):
```bash
python src/menu.py --run auto
```

## Pipeline (step‑by‑step)

The main pipeline (`python src/menu.py`, option 1):

1. Run tests
2. Generate `urls.json` (if missing, from `RevisionManual.xlsx`)
3. Choose environment: **Preview** / **Produccion** / **Ambas**
4. Audit URLs (via Playwright — Chromium or Firefox)
5. Post-process AA data
6. Generate report
7. Clean up

### Environment selection

| Entorno | URLs | VPN needed | Akamai block |
|---------|------|------------|--------------|
| Preview | 14 URLs | Sí | No |
| Produccion | 314 URLs | No | Sí (bypassed with real Chrome) |
| Ambas | all 328 | Partial | Mixed |

- `channel="chrome"` bypasses Akamai WAF detection (real Chrome TLS fingerprint)
- Firefox (`--browser firefox`) avoids `ERR_ABORTED` on SPA pages (all-vehicles catalog)
- See `docs/fallback-err-aborted.md` for detailed browser comparison

## Flags

### `src/extract_browser.py`

| Flag | Description |
|------|-------------|
| `--urls <file>` | Input JSON → multi-sheet historial |
| `--market <name>` | Filter by market (e.g. PR, MX). Output in `<NAME>/` dir |
| `--entorno <env>` | Environment: `preview` (default), `produccion`, or `ambas` |
| `--split-aa` | Create con_aa.xlsx + sin_aa.xlsx alongside historial |
| `--browser <name>` | Browser engine: `chromium` (default) or `firefox` |
| `--wait-after <sec>` | Seconds to wait after page load (default: 2) |
| `--workers <n>` | Concurrent URLs (default: 3) |
| `--timeout <sec>` | Page load timeout in seconds (default: 35) |
| `--max-retry <n>` | Retries per URL (default: 1) |
| `--headless` | Force headless mode |
| `--headed` | Show browser window |
| `--proxy <url>` | HTTP proxy |
| `--progress` | Show progress bar |
| `--verbose` | Debug logging |
| `--resume` | Skip rows with existing data |

### `src/extract_aa.py`

| Flag | Description |
|------|-------------|
| `--urls <file>` | Input JSON for market/entorno resolution |
| `--market <name>` | Filter by market |
| `--input <file>` | Input Excel file |
| `--output <file>` | Output Excel file |
| `--keep <fields>` | Filter AA fields to extract |
| `--score` | Per-row AA extraction metrics |
| `--sheet` | Select specific sheet |

## urls.json format (in `data/urls.json`)

```json
[
  {"url": "https://...", "nombre": "Page Name", "market": "PR", "entorno": "preview"},
  {"url": "https://...", "nombre": "Another Page", "market": "MX", "entorno": "produccion"}
]
```

- `url` (required): full URL to audit
- `nombre` (optional): display name, falls back to URL
- `market` (required for segmentation): market code (PR, MX, etc.)
- `entorno` (optional): `"preview"` (default) or `"produccion"`

## Project status

Internal tool. Active development.
- **Tests**: 251 ✓ — pytest + stdlib unittest
- **Coverage**: excel.py 83%, aa_parser.py 99%, utils.py 100%, browser.py ~65%
- **Mypy**: 4 modules annotated, 0 errors
- **Pipeline**: Chromium + Firefox, cache, smart wait, parallel extraction

# json-convert

Browser automation tool to extract Adobe Analytics data (digitalData + AA beacons) from Ford preview websites into Excel.

## Quick start

```bash
pip install --user openpyxl playwright
python -m playwright install chromium
python extract_browser.py --urls urls.json --output historial.xlsx
```

## Commands

| Command | Description |
|---------|-------------|
| `extract_browser.py` | Navigate URLs, capture AA beacons + digitalData, write Excel |
| `extract_aa.py` | Post-process: filter JSON fields from col E → col F |
| `_gen_urls.py` | Extract URLs from Excel → `urls.json` |
| `test_parse.py` | 72 unit tests (stdlib only, no browser needed) |
| `test_integration.py` | 21 integration tests (Excel pipeline, config, I/O) |

## Usage

### Basic

```bash
# Classic mode: read Excel, write results back
python extract_browser.py

# Multi-sheet historial mode (recommended)
python extract_browser.py --urls urls.json
```

### Market segmentation

```bash
# Filter by market (field in urls.json), output goes to PR/
python extract_browser.py --urls urls.json --market PR

# Split into con_aa / sin_aa files
python extract_browser.py --urls urls.json --market PR --split-aa
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
# Basic: extract URLs from RevisionManual.xlsx → urls.json
python _gen_urls.py

# With market assignment (all URLs get "market": "PR")
python _gen_urls.py --input RevisionManual_PR.xlsx --default-market PR
python _gen_urls.py --input RevisionManual_MX.xlsx --default-market MX
```

### Post-processing

```bash
# Clean/filter AA fields from any of the generated files
python extract_aa.py --input PR/historial.xlsx
python extract_aa.py --input PR/con_aa.xlsx
```

### All flags

| Flag | Script | Description |
|------|--------|-------------|
| `--urls <file>` | extract_browser | Input JSON → multi-sheet historial |
| `--market <name>` | extract_browser | Filter by market field (e.g. PR, MX). Output in `<NAME>/` dir |
| `--split-aa` | extract_browser | Create con_aa.xlsx + sin_aa.xlsx alongside historial |
| `--wait-after <sec>` | extract_browser | Seconds to wait after page load (default: 4) |
| `--workers <n>` | extract_browser | Concurrent URLs (default: 1) |
| `--retry <n>` | extract_browser | Retries per URL (default: 1) |
| `--timeout <ms>` | extract_browser | Page load timeout (default: 35000) |
| `--headed` | extract_browser | Show browser window |
| `--proxy <url>` | extract_browser | HTTP proxy |
| `--resume` | extract_browser | Skip rows with existing data |
| `--retry-failed` | extract_browser | Only process previously failed URLs |
| `--verbose` | both | Debug logging |
| `--backup` | extract_browser | Backup Excel before overwriting |
| `--progress` | extract_browser | Show progress bar |
| `--diff` | extract_browser | Compare last 2 audits |
| `--diagnostic` | extract_browser | Check environment without browsing |
| `--config <file>` | extract_browser | JSON config file (default: audit.json) |
| `--keep <fields>` | extract_aa | Filter AA fields to extract |
| `--score` | extract_aa | Per-row AA extraction metrics |
| `--sheet` | extract_aa | Select specific sheet |

### Configuration via `audit.json`

Flags can be persisted in `audit.json` (auto-loaded):

```json
{
  "workers": 3,
  "retry": 2,
  "timeout": 35000,
  "wait_after": 4,
  "progress": true,
  "headed": false,
  "verbose": false
}
```

CLI flags override config values when set.

## urls.json format

```json
[
  {"url": "https://...", "nombre": "Page Name", "market": "PR"},
  {"url": "https://...", "nombre": "Another Page", "market": "MX"}
]
```

- `url` (required): full URL to audit
- `nombre` (optional): display name, falls back to URL
- `market` (optional): used with `--market` flag for segmentation

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

- Python 3.9+
- Windows (Playwright Chromium, ~180 MB download on first install)
- Corporate VPN for Ford preview URLs

## Project status

Internal tool. Active development.

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
| `test_parse.py` | 31 unit tests (stdlib only, no browser needed) |

## Documentation

- [`INSTRUCCIONES.txt`](INSTRUCCIONES.txt) — Setup guide (Spanish, non-technical users)
- [`DECISIONES-CONCURRENCIA.md`](DECISIONES-CONCURRENCIA.md) — Async concurrency ADR
- [`audit.json`](audit.json) — Default configuration

## Requirements

- Python 3.9+
- Windows (Playwright Chromium, ~180 MB download on first install)
- Corporate VPN for Ford preview URLs

## Project status

Internal tool. Active development.

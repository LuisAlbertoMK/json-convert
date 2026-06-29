#!/usr/bin/env python3
"""
update-urls.py -- Actualiza urls.json desde un .txt y dispara extraccion.

Uso:
    python scripts/update-urls.py              -> lee urls.txt (raiz del proyecto)
    python scripts/update-urls.py mis-urls.txt -> lee archivo especifico

Formato del .txt: una URL por linea (las que no empiecen con http se saltan).

Ejemplo urls.txt:
    https://www.ford.mx/
    https://www.ford.mx/distribuidores/
    https://www.ford.mx/mustang/

Flujo:
    1. Lee URLs del .txt
    2. Actualiza data/urls.json (mercado auto-detectado)
    3. Dispara extraccion con extract_browser.py para cada mercado
"""

import json
import os
import subprocess
import sys
from urllib.parse import urlparse

# -- Mapeo de dominio a mercado --
DOMAIN_MAP = {
    "www.ford.com.pr": "PR",
    "www.ford.mx": "MX",
    "ford.com.pr": "PR",
    "ford.mx": "MX",
    "wwwac.preview.es.brandpr.ford.com": "PR",
    "wwwac.preview.brandpr.ford.com": "PR",
}

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URLS_PATH = os.path.join(BASE_DIR, "data", "urls.json")
DEFAULT_TXT = os.path.join(BASE_DIR, "urls.txt")
EXTRACT_SCRIPT = os.path.join(BASE_DIR, "src", "extract_browser.py")


def detect_market(url: str) -> str:
    """Detecta mercado desde el dominio de la URL."""
    try:
        domain = urlparse(url).netloc.lower()
        for pattern, market in DOMAIN_MAP.items():
            if pattern in domain:
                return market
    except Exception:
        pass
    return "MX"  # default


def read_urls_from_file(path: str) -> list[str]:
    """Lee URLs desde un archivo de texto (una por linea)."""
    urls = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.startswith(("http://", "https://")):
                urls.append(line)
    return urls


def build_entries(urls: list[str]) -> list[dict]:
    """Construye lista de entries para urls.json."""
    seen = set()
    entries = []
    for url in urls:
        norm = url.rstrip("/")
        if norm in seen:
            continue
        seen.add(norm)
        market = detect_market(url)
        entries.append({
            "url": url,
            "market": market,
            "entorno": "produccion",
            "tipo": "produccion",
        })
    return entries


def save_urls(entries: list[dict]) -> None:
    """Guarda entries en urls.json."""
    os.makedirs(os.path.dirname(URLS_PATH), exist_ok=True)
    with open(URLS_PATH, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=4, ensure_ascii=False)
        f.write("\n")
    print(f"\n  [OK] {URLS_PATH} actualizado ({len(entries)} URLs)")


def run_extraction(market: str, entries: list[dict]) -> bool:
    """Ejecuta extract_browser.py para un mercado."""
    output_dir = os.path.join(BASE_DIR, market, "produccion")
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, "historial.xlsx")

    market_urls = [e for e in entries if e["market"] == market]

    print(f"\n  {'='*50}")
    print(f"  Iniciando extraccion para {market} ({len(market_urls)} URLs)...")
    print(f"  {'='*50}")

    cmd = [
        sys.executable,
        EXTRACT_SCRIPT,
        "--urls", URLS_PATH,
        "--market", market,
        "--entorno", "produccion",
        "--output", output_file,
        "--split-aa",
        "--workers", "4",
        "--browser", "firefox",
    ]

    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode == 0:
        print(f"\n  [OK] Extraccion {market} completada")
        return True
    else:
        print(f"\n  [ERR] Extraccion {market} fallo (codigo {result.returncode})")
        return False


def main():
    # -- 1. Leer URLs --
    if len(sys.argv) > 1:
        txt_path = sys.argv[1]
        if not os.path.exists(txt_path):
            print(f"[ERR] Archivo no encontrado: {txt_path}")
            sys.exit(1)
    else:
        txt_path = DEFAULT_TXT
        if not os.path.exists(txt_path):
            print(f"[ERR] Creame primero {txt_path} con una URL por linea")
            sys.exit(1)

    urls = read_urls_from_file(txt_path)
    if not urls:
        print(f"\n  [!] No se encontraron URLs validas en {txt_path}")
        print("     Asegurate que cada linea empiece con http:// o https://")
        sys.exit(1)

    print(f"\n  [{os.path.basename(txt_path)}] {len(urls)} URLs leidas")

    # -- 2. Construir y guardar urls.json --
    entries = build_entries(urls)

    markets = {}
    for e in entries:
        m = e["market"]
        markets[m] = markets.get(m, 0) + 1

    print(f"\n  {'='*45}")
    print(f"  URLs detectadas: {len(entries)}")
    for m, c in sorted(markets.items()):
        print(f"    {m}: {c} URLs")
    print(f"  {'='*45}")

    save_urls(entries)

    # -- 3. Preguntar si extraer --
    resp = input("\n  Ejecutar extraccion ahora? [S/n]: ").strip().lower()
    if resp not in ("", "s", "si", "y", "yes"):
        print("  Ok. Podes correr el menu manualmente despues.")
        return

    # -- 4. Extraer por cada mercado --
    ok = True
    for market in sorted(markets):
        if not run_extraction(market, entries):
            ok = False

    if ok:
        print(f"\n  {'='*45}")
        print(f"  [OK] TODO COMPLETADO -- {len(entries)} URLs procesadas")
        print(f"  {'='*45}")
    else:
        print(f"\n  [!] Algunas extracciones fallaron -- revisa los logs")


if __name__ == "__main__":
    main()

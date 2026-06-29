#!/usr/bin/env python3
"""
update-urls.py — Actualiza data/urls.json pegando URLs.

Uso:
    python scripts/update-urls.py
        → pega TODAS las URLs de una (bloque), Ctrl+Z + Enter para terminar
        → auto-detecta mercado y genera urls.json

    python scripts/update-urls.py < urls.txt
        → desde archivo via redirección stdin

    python scripts/update-urls.py urls.txt
        → lee URLs de un archivo de texto directamente

Ejemplo:
    python scripts/update-urls.py
    Pegá TODAS las URLs (una por línea) y presioná Ctrl+Z + Enter para terminar:
    https://www.ford.mx/
    https://www.ford.mx/distribuidores/
    ^Z
    Detectado: MX - 2 URLs
    ¿Guardar? [S/n]: s
    [OK] urls.json actualizado (2 URLs)
"""

import json
import os
import sys
from urllib.parse import urlparse

# ── Mapeo de dominio → mercado ──
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


def read_urls_from_stdin() -> list[str]:
    """Lee URLs desde stdin — pega todo el bloque de una, Ctrl+Z + Enter para terminar."""
    print("\n  Pegá TODAS las URLs (una por línea) y presioná Ctrl+Z + Enter para terminar:\n")
    urls = []
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        if line.startswith(("http://", "https://")):
            urls.append(line)
        else:
            print(f"  [!] Se saltea (no es URL válida): {line}")
    return urls


def read_urls_from_file(path: str) -> list[str]:
    """Lee URLs desde un archivo de texto."""
    urls = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and line.startswith(("http://", "https://")):
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


def main():
    # Fuente de URLs
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if not os.path.exists(file_path):
            print(f"[ERR] Archivo no encontrado: {file_path}")
            sys.exit(1)
        urls = read_urls_from_file(file_path)
        print(f"\n  Leídas {len(urls)} URLs desde {file_path}")
    else:
        urls = read_urls_from_stdin()

    if not urls:
        print("\n  [!] No se ingresaron URLs. Saliendo.")
        return

    entries = build_entries(urls)

    # Resumen
    markets = {}
    for e in entries:
        m = e["market"]
        markets[m] = markets.get(m, 0) + 1

    print(f"\n  {'='*45}")
    print(f"  URLs detectadas: {len(entries)}")
    for m, c in sorted(markets.items()):
        print(f"    {m}: {c} URLs")
    print(f"  {'='*45}")

    # Confirmar (solo si stdin es interactivo)
    if sys.stdin.isatty():
        resp = input("\n  Guardar urls.json? [S/n]: ").strip().lower()
        if resp not in ("", "s", "si", "y", "yes"):
            print("  Cancelado.")
            return
    save_urls(entries)


if __name__ == "__main__":
    main()

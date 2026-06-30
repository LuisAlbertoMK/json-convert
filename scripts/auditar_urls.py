"""
auditar_urls.py — Pipeline reutilizable de auditoría individual por URL.
Cada URL se audita por separado con su propia con_aa, sin_aa, historial y matriz.

Uso:
  python scripts/auditar_urls.py                          # Corre las URLs por defecto (Bronco PR)
  python scripts/auditar_urls.py --clean                  # Limpia PR/ antes de empezar
  python scripts/auditar_urls.py --urls mis-urls.json     # Usa lista personalizada
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import argparse
import json
import os
import shutil
import subprocess

# ── Emoji-safe para Windows terminal ──
_CLEAN = str.maketrans({c: '?' for c in '✅⚠️❌🔴🟢🟡🗑️📊📁'})
def _safe(text: str) -> str:
    return text.translate(_CLEAN)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PR_DIR = os.path.join(ROOT, "PR")
PROD_DIR = os.path.join(PR_DIR, "produccion")
TICKETS_DIR = os.path.join(PR_DIR, "tickets")
PYTHON = sys.executable

# ── URLs por defecto: Bronco PR ──
DEFAULT_URLS = [
    {
        "page_key": "bronco",
        "nombre": "Bronco SUV",
        "url": "https://www.ford.com.pr/esp/crossovers-suvs/bronco/",
        "slug": "bronco",
    },
    {
        "page_key": "bronco-models-base",
        "nombre": "Bronco Base",
        "url": "https://www.ford.com.pr/esp/crossovers-suvs/bronco/models/base/",
        "slug": "base",
    },
    {
        "page_key": "bronco-models-big-bend",
        "nombre": "Bronco Big Bend",
        "url": "https://www.ford.com.pr/esp/crossovers-suvs/bronco/models/big-bend/",
        "slug": "big-bend",
    },
    {
        "page_key": "bronco-models-outer-banks",
        "nombre": "Bronco Outer Banks",
        "url": "https://www.ford.com.pr/esp/crossovers-suvs/bronco/models/outer-banks/",
        "slug": "outer-banks",
    },
    {
        "page_key": "bronco-models-badlands",
        "nombre": "Bronco Badlands",
        "url": "https://www.ford.com.pr/esp/crossovers-suvs/bronco/models/badlands/",
        "slug": "badlands",
    },
]


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  [JSON] {os.path.relpath(path, ROOT)}")


def run(cmd, label):
    desc = " ".join(cmd[-3:]) if len(cmd) > 3 else " ".join(cmd[-2:])
    print(f"\n  [{label}] {desc}")
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace", cwd=ROOT, env=env
    )
    for stream, prefix in [(result.stdout, ""), (result.stderr, "[STDERR] ")]:
        if stream:
            for line in stream.split("\n"):
                stripped = line.strip()
                if stripped and any(kw in stripped for kw in (
                    "[OK]", "[ERR]", "[WARN]", "[INFO]", "Pipeline",
                    "Auditando", "Guardado", "Matriz", "Score",
                    "paginas", "error", "Error", "OK", "Sin", "fila",
                )):
                    print(f"    {_safe(stripped)}")
    return result.returncode


def auditar_url(url_entry: dict, clean_first: bool = False):
    page_key = url_entry["page_key"]
    nombre = url_entry["nombre"]
    url = url_entry["url"]
    slug = url_entry["slug"]
    out_dir = os.path.join(TICKETS_DIR, slug)

    print(f"\n{'='*60}")
    print(f"  AUDITANDO: {nombre}")
    print(f"  URL: {url}")
    print(f"  OUTPUT: {out_dir}")
    print(f"{'='*60}")

    # ── 1. Preparar archivos de config ──
    write_json(os.path.join(ROOT, "data/url-mapping.json"), [{
        "page_key": page_key,
        "production_url": url,
        "preview_url": url,
        "nombre": nombre,
    }])
    write_json(os.path.join(ROOT, "data/urls.json"), [{
        "url": url,
        "market": "PR",
        "entorno": "produccion",
        "tipo": "produccion",
    }])

    # ── 2. Limpiar PR/produccion/ ──
    if clean_first and os.path.exists(PROD_DIR):
        shutil.rmtree(PROD_DIR)
    os.makedirs(PROD_DIR, exist_ok=True)

    # ── 3. Auditoria con browser ──
    ret = run([
        PYTHON, "src/extract_browser.py",
        "--urls", "data/urls.json",
        "--market", "PR",
        "--entorno", "produccion",
        "--split-aa", "--progress",
        "--headless", "--browser", "firefox",
        "--timeout", "30",
    ], nombre)
    if ret != 0:
        print(f"  [SKIP] Error en auditoria ({ret})")
        return False

    # ── 4. Post-procesar AA ──
    run([
        PYTHON, "src/extract_aa.py",
        "--input", "PR/produccion/historial.xlsx",
        "--urls", "data/urls.json",
    ], f"{nombre} (AA)")

    # ── 5. Generar matriz ──
    run([
        PYTHON, "src/generate_validation_matrix.py",
        "--market", "PR",
        "--entorno", "produccion",
    ], f"{nombre} (matriz)")

    # ── 6. Copiar outputs ──
    os.makedirs(out_dir, exist_ok=True)
    for fname in ["con_aa.xlsx", "sin_aa.xlsx", "historial.xlsx"]:
        src = os.path.join(PROD_DIR, fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(out_dir, fname))
    src_matriz = os.path.join(PR_DIR, "matriz-validacion-produccion.xlsx")
    if os.path.exists(src_matriz):
        shutil.copy2(src_matriz, os.path.join(out_dir, "matriz-validacion-produccion.xlsx"))

    print(f"\n  [OK] {nombre} COMPLETADO -> {out_dir}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Pipeline de auditoria individual por URL"
    )
    parser.add_argument("--clean", action="store_true",
                        help="Limpia PR/ antes de empezar")
    parser.add_argument("--urls", type=str, default=None,
                        help="Archivo JSON con lista de URLs a auditar")
    args = parser.parse_args()

    # Cargar URLs
    if args.urls:
        with open(args.urls, encoding="utf-8") as f:
            urls_list = json.load(f)
    else:
        urls_list = DEFAULT_URLS

    print(f"\n{'='*60}")
    print(f"  PIPELINE DE AUDITORIA INDIVIDUAL")
    print(f"  {len(urls_list)} URLs a procesar")
    print(f"{'='*60}")

    # Limpiar si se pide
    if args.clean:
        for d in [PROD_DIR, TICKETS_DIR]:
            if os.path.exists(d):
                shutil.rmtree(d)
        os.makedirs(PROD_DIR, exist_ok=True)
        os.makedirs(TICKETS_DIR, exist_ok=True)

    # Auditar cada URL
    ok = 0
    fail = 0
    for entry in urls_list:
        if auditar_url(entry, clean_first=True):
            ok += 1
        else:
            fail += 1

    print(f"\n{'='*60}")
    print(f"  PIPELINE COMPLETADO: {ok} OK, {fail} errores")
    print(f"  Output: {TICKETS_DIR}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()

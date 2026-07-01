"""
auditar_urls.py — Pipeline reutilizable de auditoría individual por URL.
Cada URL se audita por separado con su propia con_aa, sin_aa, historial y matriz.

Uso:
  python scripts/auditar_urls.py                          # Corre las URLs por defecto (Bronco PR)
  python scripts/auditar_urls.py --clean                  # Limpia PR/ antes de empezar
  python scripts/auditar_urls.py --urls mis-urls.json     # Usa lista personalizada
  python scripts/auditar_urls.py --ticket 42399           # Ticket GTBEMEAPUB para nombrar archivos
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import argparse
from datetime import datetime
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
TEMP_DIR = os.path.join(ROOT, "data", ".tmp")  # archivos temporales unicos por URL
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
    # Verificar que se escribio correctamente
    if not os.path.exists(path):
        print(f"  [ERR] No se pudo escribir {path}")
        return False
    with open(path, "r", encoding="utf-8-sig") as f:
        try:
            loaded = json.load(f)
            if loaded != data:
                print(f"  [WARN] {os.path.relpath(path, ROOT)} escrito pero contenido no coincide")
        except json.JSONDecodeError:
            print(f"  [ERR] {os.path.relpath(path, ROOT)} escrito pero JSON invalido")
            return False
    print(f"  [JSON] {os.path.relpath(path, ROOT)}")
    return True


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


def _to_pascal(slug: str) -> str:
    """Convierte slug a PascalCase. Ej: 'big-bend' -> 'BigBend'"""
    return "".join(word.capitalize() for word in slug.replace("-", " ").replace("_", " ").split())


def _commit_staging(staging_dir: str) -> None:
    """Mueve archivos de staging → PR/ de forma atómica.

    Cada URL produce staging/{slug}/produccion/, staging/{slug}/tickets/
    y staging/{slug}/matriz-validacion-*.xlsx.
    Al committear: tickets → PR/tickets/{slug}/,
    produccion y matrices → PR/ (último slug sobrescribe).
    """
    print(f"\n{'='*60}")
    print(f"  COMMIT: Moviendo staging → PR/")
    print(f"{'='*60}")

    os.makedirs(PROD_DIR, exist_ok=True)
    os.makedirs(TICKETS_DIR, exist_ok=True)

    slugs = sorted(os.listdir(staging_dir))
    for slug in slugs:
        slug_path = os.path.join(staging_dir, slug)
        if not os.path.isdir(slug_path):
            continue

        # Tickets individuales por URL
        src_tickets = os.path.join(slug_path, "tickets")
        if os.path.isdir(src_tickets):
            dst = os.path.join(TICKETS_DIR, slug)
            os.makedirs(dst, exist_ok=True)
            for fname in os.listdir(src_tickets):
                shutil.move(os.path.join(src_tickets, fname), os.path.join(dst, fname))
            print(f"  [TICKETS] {slug}/ → PR/tickets/{slug}/")

        # Producción (shared — último slug sobrescribe)
        src_prod = os.path.join(slug_path, "produccion")
        if os.path.isdir(src_prod):
            for fname in os.listdir(src_prod):
                shutil.move(os.path.join(src_prod, fname), os.path.join(PROD_DIR, fname))

        # Matrices (shared — último slug sobrescribe)
        for fname in os.listdir(slug_path):
            if fname.startswith("matriz-validacion"):
                shutil.move(os.path.join(slug_path, fname), os.path.join(PR_DIR, fname))

    print(f"  [OK] Commit completado — {len(slugs)} slugs a PR/")


def auditar_url(url_entry: dict, clean_first: bool = False, ticket: str = "",
                staging_dir: str = ""):
    page_key = url_entry["page_key"]
    nombre = url_entry["nombre"]
    url = url_entry["url"]
    slug = url_entry["slug"]

    # ── Determinar directorios de salida (staging o final) ──
    if staging_dir:
        staging_slug = os.path.join(staging_dir, slug)
        prod_dir = os.path.join(staging_slug, "produccion")
        out_dir = os.path.join(staging_slug, "tickets")
    else:
        prod_dir = PROD_DIR
        out_dir = os.path.join(TICKETS_DIR, slug)

    print(f"\n{'='*60}")
    print(f"  AUDITANDO: {nombre}")
    print(f"  URL: {url}")
    print(f"  OUTPUT: {out_dir}")
    if staging_dir:
        print(f"  STAGING: {staging_dir}")
    print(f"{'='*60}")

    # ── 1. Preparar archivos de config (archivos UNICOS por URL para evitar race conditions) ──
    os.makedirs(TEMP_DIR, exist_ok=True)
    url_mapping_path = os.path.join(TEMP_DIR, f"{slug}_url-mapping.json")
    urls_path = os.path.join(TEMP_DIR, f"{slug}_urls.json")
    write_json(url_mapping_path, [{
        "page_key": page_key,
        "production_url": url,
        "preview_url": url,
        "nombre": nombre,
    }])
    write_json(urls_path, [{
        "url": url,
        "market": "PR",
        "entorno": "produccion",
        "tipo": "produccion",
    }])

    # ── 2. Limpiar directorio de salida ──
    if clean_first and os.path.exists(prod_dir):
        shutil.rmtree(prod_dir)
    os.makedirs(prod_dir, exist_ok=True)

    # ── 3. Auditoria con browser ──
    browser_args = [
        PYTHON, "src/extract_browser.py",
        "--urls", urls_path,
        "--market", "PR",
        "--entorno", "produccion",
        "--split-aa", "--progress",
        "--headless", "--browser", "firefox",
        "--timeout", "30",
    ]
    if staging_dir:
        browser_args.extend(["--output", os.path.join(prod_dir, "historial.xlsx")])
    ret = run(browser_args, nombre)
    if ret != 0:
        print(f"  [SKIP] Error en auditoria ({ret})")
        return False

    # ── 4. Post-procesar AA ──
    run([
        PYTHON, "src/extract_aa.py",
        "--input", os.path.join(prod_dir, "historial.xlsx"),
        "--urls", urls_path,
    ], f"{nombre} (AA)")

    # ── 5. Generar matriz (usando mapping TEMPORAL, no el global) ──
    matrix_args = [
        PYTHON, "src/generate_validation_matrix.py",
        "--market", "PR",
        "--entorno", "produccion",
        "--mapping", url_mapping_path,
    ]
    if staging_dir:
        matrix_args.extend([
            "--output", os.path.join(staging_slug, "matriz-validacion-produccion.xlsx"),
            "--historial-produccion", os.path.join(prod_dir, "historial.xlsx"),
        ])
    run(matrix_args, f"{nombre} (matriz)")

    # ── 6. Copiar outputs (con verificacion) ──
    os.makedirs(out_dir, exist_ok=True)
    if not os.path.exists(out_dir):
        print(f"  [ERR] No se pudo crear {out_dir}")
        return False
    for fname in ["con_aa.xlsx", "sin_aa.xlsx", "historial.xlsx"]:
        src = os.path.join(prod_dir, fname)
        if os.path.exists(src):
            dst = os.path.join(out_dir, fname)
            try:
                shutil.copy2(src, dst)
                if not os.path.exists(dst):
                    print(f"  [WARN] No se pudo copiar {fname} a {out_dir}")
            except Exception as e:
                print(f"  [WARN] Error copiando {fname}: {e}")
    if staging_dir:
        src_matriz = os.path.join(staging_slug, "matriz-validacion-produccion.xlsx")
    else:
        src_matriz = os.path.join(PR_DIR, "matriz-validacion-produccion.xlsx")
    if os.path.exists(src_matriz):
        dst_matriz = os.path.join(out_dir, "matriz-validacion-produccion.xlsx")
        try:
            shutil.copy2(src_matriz, dst_matriz)
        except Exception as e:
            print(f"  [WARN] Error copiando matriz: {e}")

    # ── 7. Generar matriz individual (formato simple — como ejemplos GTBEMEAPUB) ──
    if ticket:
        model_name = _to_pascal(slug)
        simple_name = f"GTBEMEAPUB-{ticket}-PR-ESP-{model_name}.xlsx"
        simple_path = os.path.join(out_dir, simple_name)
        simple_args = [
            PYTHON, "src/generate_validation_matrix.py",
            "--market", "PR",
            "--entorno", "produccion",
            "--mapping", url_mapping_path,
            "--simple",
            "--output", simple_path,
        ]
        if staging_dir:
            simple_args.extend([
                "--historial-produccion", os.path.join(prod_dir, "historial.xlsx"),
            ])
        run(simple_args, f"{nombre} (simple)")
    else:
        print("  [--ticket no especificado, omitiendo archivo individual]")

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
    parser.add_argument("--ticket", type=str, default="",
                        help="Ticket GTBEMEAPUB para nombrar archivos individuales (ej: 42399)")
    args = parser.parse_args()
    ticket = args.ticket.strip()

    # Cargar URLs (utf-8-sig para tolerar BOM en Windows)
    if args.urls:
        with open(args.urls, encoding="utf-8-sig") as f:
            urls_list = json.load(f)
    else:
        urls_list = DEFAULT_URLS

    print(f"\n{'='*60}")
    print(f"  PIPELINE DE AUDITORIA INDIVIDUAL")
    print(f"  {len(urls_list)} URLs a procesar")
    print(f"{'='*60}")

    # Limpiar si se pide (tolerante a archivos bloqueados en Windows)
    def _on_rm_error(func, path, exc_info):
        print(f"  [WARN] No se pudo eliminar (en uso): {os.path.relpath(path, ROOT)}")
    if args.clean:
        for d in [PROD_DIR, TICKETS_DIR, TEMP_DIR]:
            if os.path.exists(d):
                shutil.rmtree(d, onexc=_on_rm_error)
        os.makedirs(PROD_DIR, exist_ok=True)
        os.makedirs(TICKETS_DIR, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)

    # ── Staging area: aislamiento por corrida ──
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    staging_dir = os.path.join(ROOT, "data", ".tmp", f"staging_{run_id}")
    os.makedirs(staging_dir, exist_ok=True)
    print(f"  [STAGING] {staging_dir}")

    # Auditar cada URL
    ok = 0
    fail = 0
    try:
        for entry in urls_list:
            if auditar_url(entry, clean_first=True, ticket=ticket, staging_dir=staging_dir):
                ok += 1
            else:
                fail += 1

        # Commit atómico solo si TODO OK
        if fail == 0 and ok > 0:
            _commit_staging(staging_dir)
        else:
            print(f"\n  [STAGING] {fail} errores — PR/ NO modificado")
    finally:
        # Limpiar staging siempre
        if os.path.exists(staging_dir):
            shutil.rmtree(staging_dir, ignore_errors=True)

    # Limpiar archivos temporales
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)

    print(f"\n{'='*60}")
    print(f"  PIPELINE COMPLETADO: {ok} OK, {fail} errores")
    print(f"  Output: {TICKETS_DIR}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()

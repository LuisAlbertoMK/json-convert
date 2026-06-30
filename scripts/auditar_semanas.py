"""
auditar_semanas.py — Pipeline de auditoría por semanas.
Cada semana recibe un grupo de URLs, se auditan juntas por mercado+entorno,
y se guardan en {mercado}/semanas/{nombre_semana}/.

Uso:
  python scripts/auditar_semanas.py data/semanas.json

Formato entrada:
  [
    {"nombre_semana": "semana1", "urls": ["https://...", ...]},
    {"nombre_semana": "semana2", "urls": ["https://...", ...]}
  ]
  o
  {
    "semana1": ["https://...", ...],
    "semana2": ["https://...", ...]
  }

Output por (mercado, semana):
  {mercado}/semanas/{nombre_semana}/
    con_aa.xlsx              ← páginas CON beacon AA
    sin_aa.xlsx              ← páginas SIN beacon AA
    matriz-validacion.xlsx   ← matriz comparativa
  (historial.xlsx se elimina — no se requiere)
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import argparse
import json
import os
import subprocess
import time
from pathlib import Path
from urllib.parse import urlparse

# ── Emoji-safe ──
_CLEAN = str.maketrans({c: '?' for c in '✅⚠️❌🔴🟢🟡🗑️📊📁'})
def _safe(text: str) -> str:
    return text.translate(_CLEAN)

ROOT = Path(__file__).resolve().parent.parent
PYTHON = sys.executable


# ── Helpers ──

def _mercado(url: str) -> str:
    """Infiere mercado del dominio: .mx → MX, .pr → PR, default US."""
    d = urlparse(url).netloc.lower()
    for tld in ('.pr', '.mx'):
        if tld in d:
            return tld[1:].upper()
    return 'US'


def _entorno(url: str) -> str:
    """preview si la URL contiene 'preview', sino produccion."""
    return 'preview' if 'preview' in url.lower() else 'produccion'


def _run(cmd: list[str], label: str = "") -> int:
    """Ejecuta comando, muestra líneas relevantes del stdout."""
    desc = " ".join(cmd[-3:]) if len(cmd) > 3 else " ".join(cmd[-2:])
    print(f"\n  [{label}] {desc}")
    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace", cwd=str(ROOT)
    )
    for line in result.stdout.split("\n"):
        stripped = line.strip()
        if stripped and any(kw in stripped for kw in (
            "[OK]", "[ERR]", "[WARN]", "[INFO]", "Pipeline",
            "Auditando", "Guardado", "Matriz", "Score",
            "paginas", "error", "Error", "OK", "Sin", "fila",
            "FAILED", "completed", "Total",
        )):
            print(f"    {_safe(stripped)}")
    for line in result.stderr.split("\n"):
        stripped = line.strip()
        if stripped and "error" in stripped.lower():
            print(f"    [STDERR] {_safe(stripped)}")
    return result.returncode


def _cargar_semanas(path: str) -> list[dict]:
    """Carga entrada: objeto {semana: [urls]} o array [{nombre_semana, urls}]."""
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(raw, dict):
        return [{"nombre_semana": k, "urls": v} for k, v in raw.items()]
    return raw


# ── Core ──

def procesar_semana(semana: dict) -> tuple[int, int]:
    """Procesa una semana. Retorna (grupos_ok, grupos_err)."""
    nombre = semana["nombre_semana"]
    urls = semana["urls"]
    if not urls:
        print(f"\n  [WARN] {nombre}: sin URLs, saltando")
        return 0, 1

    # Agrupar por (mercado, entorno)
    grupos: dict[tuple[str, str], list[str]] = {}
    for url in urls:
        m = _mercado(url)
        e = _entorno(url)
        grupos.setdefault((m, e), []).append(url)

    ok_total = 0
    err_total = 0

    for (mercado, entorno), grupo in grupos.items():
        semana_dir = ROOT / mercado / "semanas" / nombre
        semana_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n{'='*55}")
        print(f"  {nombre} | {mercado}/{entorno} | {len(grupo)} URLs")
        print(f"  Output: {mercado}/semanas/{nombre}/")
        print(f"{'='*55}")

        # ── Preparar data files ──
        mapping = [
            {
                "page_key": f"p{i}",
                "production_url": url,
                "preview_url": url,
                "nombre": f"URL {i+1}",
            }
            for i, url in enumerate(grupo)
        ]
        urls_data = [
            {"url": url, "market": mercado, "entorno": entorno, "tipo": entorno}
            for url in grupo
        ]

        json.dump(mapping, open(ROOT / "data" / "url-mapping.json", "w", encoding="utf-8"),
                  indent=2, ensure_ascii=False)
        json.dump(urls_data, open(ROOT / "data" / "urls.json", "w", encoding="utf-8"),
                  indent=2, ensure_ascii=False)

        # ── 1. extract_browser ──
        historial_path = semana_dir / "historial.xlsx"
        ret = _run([
            PYTHON, "src/extract_browser.py",
            "--urls", str(ROOT / "data" / "urls.json"),
            "--market", mercado,
            "--entorno", entorno,
            "--output", str(historial_path),
            "--split-aa", "--progress",
            "--headless", "--browser", "firefox",
            "--timeout", "30", "--workers", "3",
        ], f"Browser {nombre}/{mercado}")

        if ret != 0:
            print(f"\n  [ERR] extract_browser falló (código {ret}) en {nombre}/{mercado}")
            err_total += 1
            # Limpiar historial aunque haya fallado
            if historial_path.exists():
                historial_path.unlink()
            continue

        # ── 2. generate_validation_matrix ──
        matriz_path = semana_dir / "matriz-validacion.xlsx"
        _run([
            PYTHON, "src/generate_validation_matrix.py",
            "--market", mercado,
            "--entorno", entorno,
            "--historial-produccion", str(historial_path),
            "--mapping", str(ROOT / "data" / "url-mapping.json"),
            "--output", str(matriz_path),
        ], f"Matriz {nombre}/{mercado}")

        # ── 3. Eliminar historial (no requerido) ──
        if historial_path.exists():
            historial_path.unlink()
            print(f"    [LIMPIO] historial.xlsx eliminado (no requerido)")

        # ── 4. Verificar outputs ──
        con_aa = semana_dir / "con_aa.xlsx"
        sin_aa = semana_dir / "sin_aa.xlsx"

        faltantes = [f.name for f in [con_aa, sin_aa, matriz_path] if not f.exists()]
        if faltantes:
            print(f"\n  [WARN] Archivos no generados en {nombre}/{mercado}: {', '.join(faltantes)}")
            err_total += 1
        else:
            print(f"\n  [OK] {nombre}/{mercado} COMPLETADO")
            ok_total += 1

    return ok_total, err_total


# ── Main ──

def main():
    parser = argparse.ArgumentParser(
        description="Pipeline de auditoría individual por semana"
    )
    parser.add_argument("input", help="Archivo JSON con semanas y URLs")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"[ERR] Archivo no encontrado: {args.input}")
        sys.exit(1)

    semanas = _cargar_semanas(args.input)

    print(f"\n{'='*55}")
    print(f"  PIPELINE DE AUDITORÍA POR SEMANAS")
    print(f"  {len(semanas)} semanas, {sum(len(s['urls']) for s in semanas)} URLs totales")
    print(f"{'='*55}")

    t0 = time.time()
    total_ok = 0
    total_err = 0

    for semana in semanas:
        ok, err = procesar_semana(semana)
        total_ok += ok
        total_err += err

    elapsed = int(time.time() - t0)
    print(f"\n{'='*55}")
    print(f"  PIPELINE COMPLETADO en {elapsed}s")
    print(f"  {total_ok} grupos OK, {total_err} errores")
    print(f"  Revisa: */semanas/*/")
    print(f"{'='*55}")


if __name__ == "__main__":
    main()

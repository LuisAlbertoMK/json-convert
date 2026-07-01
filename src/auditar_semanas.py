"""
auditar_semanas.py — Pipeline de auditoría por semanas.
Cada semana recibe un grupo de URLs, se auditan juntas por mercado+entorno,
y se guardan en {mercado}/semanas/{nombre_semana}/{entorno}/.

Uso:
  python scripts/auditar_semanas.py <archivo>
  python scripts/auditar_semanas.py <archivo> --weeks 1,3,5
  python scripts/auditar_semanas.py <archivo> --weeks 3-7 --no-matrix

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
  {mercado}/semanas/{nombre_semana}/{entorno}/
    con_aa.xlsx              ← páginas CON beacon AA
    sin_aa.xlsx              ← páginas SIN beacon AA
    historial.xlsx           ← raw audit data (para matriz)
    matriz-validacion-{entorno}.xlsx   ← matriz comparativa
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


def _infer_page_key(url: str) -> str:
    """Infiere page_key de la URL — mismo algoritmo que generate_validation_matrix.py."""
    if not url:
        return "unknown"
    path = urlparse(url).path.rstrip("/")
    if not path or path == "/":
        return "home"
    segments = [s for s in path.split("/")
                if s and s not in ("esp", "en", "content", "na",
                                    "es_pr", "en_pr", "es_mx", "en_mx")]
    if not segments:
        return "home"
    return segments[-1]


def _nombre_desde_url(url: str) -> str:
    """Convierte el último segmento de la URL a un nombre legible."""
    pk = _infer_page_key(url)
    return pk.replace("-", " ").title()


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

def procesar_semana(semana: dict, no_matrix: bool = False) -> tuple[int, int]:
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
        semana_dir = ROOT / mercado / "semanas" / nombre / entorno
        semana_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n{'='*55}")
        print(f"  {nombre} | {mercado}/{entorno} | {len(grupo)} URLs")
        print(f"  Output: {mercado}/semanas/{nombre}/{entorno}/")
        print(f"{'='*55}")

        # ── Preparar urls.json TEMPORAL (no tocar data/urls.json global) ──
        urls_data = [
            {"url": url, "market": mercado, "entorno": entorno, "tipo": entorno}
            for url in grupo
        ]
        urls_temp = semana_dir / ".urls.json"
        json.dump(urls_data, open(urls_temp, "w", encoding="utf-8"),
                  indent=2, ensure_ascii=False)

        # ── Preparar url-mapping.json TEMPORAL (no tocar data/url-mapping.json global) ──
        mapping = [
            {
                "page_key": _infer_page_key(url),
                "production_url": url,
                "preview_url": url,
                "nombre": _nombre_desde_url(url),
            }
            for url in grupo
        ]
        mapping_temp = semana_dir / ".url-mapping.json"
        # Solo escribir JUSTO antes de la matriz, no antes del browser

        # ── 1. extract_browser ──
        historial_path = semana_dir / "historial.xlsx"
        ret = _run([
            PYTHON, "src/extract_browser.py",
            "--urls", str(urls_temp),
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

        # ── 2. generate_validation_matrix (opcional) ──
        if no_matrix:
            print(f"    [SKIP] Matriz saltada (--no-matrix)")
            matriz_path = None
        else:
            # Escribir mapping temporal justo antes de la matriz
            json.dump(mapping, open(mapping_temp, "w", encoding="utf-8"),
                      indent=2, ensure_ascii=False)
            matriz_path = semana_dir / f"matriz-validacion-{entorno}.xlsx"
            _run([
                PYTHON, "src/generate_validation_matrix.py",
                "--market", mercado,
                "--entorno", entorno,
                "--historial-produccion", str(historial_path),
                "--mapping", str(mapping_temp),
                "--output", str(matriz_path),
            ], f"Matriz {nombre}/{mercado}")

        # ── 3. Verificar outputs ──
        con_aa = semana_dir / "con_aa.xlsx"
        sin_aa = semana_dir / "sin_aa.xlsx"
        historial_out = semana_dir / "historial.xlsx"

        esperados = [con_aa, sin_aa, historial_out]
        if matriz_path:
            esperados.append(matriz_path)
        faltantes = [f.name for f in esperados if not f.exists()]
        if faltantes:
            print(f"\n  [WARN] Archivos no generados en {nombre}/{mercado}: {', '.join(faltantes)}")
            err_total += 1
        else:
            print(f"\n  [OK] {nombre}/{mercado} COMPLETADO")
            ok_total += 1

    return ok_total, err_total


# ── Main ──

def _filtrar_semanas(semanas: list[dict], weeks_arg: str | None) -> list[dict]:
    """Filtra semanas segun --weeks.

    Formatos aceptados:
      "1,3,5"  → semanas 1, 3, 5 (1-based index en el array)
      "3-7"    → rango
      "" / None → todas
    """
    if not weeks_arg:
        return semanas

    indices: set[int] = set()
    for part in weeks_arg.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            indices.update(range(int(a.strip()) - 1, int(b.strip())))
        else:
            indices.add(int(part) - 1)

    result = [s for i, s in enumerate(semanas) if i in indices]
    omitidas = len(semanas) - len(result)
    if omitidas:
        print(f"  (filtro --weeks: {omitidas} semanas omitidas)")
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Pipeline de auditoría individual por semana"
    )
    parser.add_argument("input", help="Archivo JSON con semanas y URLs")
    parser.add_argument("--weeks", help="Semanas a procesar: 1,3,5 / 3-7 / 1,3-5")
    parser.add_argument("--no-matrix", action="store_true",
                        help="Saltar generación de matriz (solo browser audit)")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"[ERR] Archivo no encontrado: {args.input}")
        sys.exit(1)

    semanas = _filtrar_semanas(_cargar_semanas(args.input), args.weeks)

    print(f"\n{'='*55}")
    print(f"  PIPELINE DE AUDITORÍA POR SEMANAS")
    print(f"  {len(semanas)} semanas, {sum(len(s['urls']) for s in semanas)} URLs totales")
    if args.no_matrix:
        print(f"  Modo: solo auditoria (sin matriz)")
    print(f"{'='*55}")

    t0 = time.time()
    total_ok = 0
    total_err = 0

    for semana in semanas:
        ok, err = procesar_semana(semana, no_matrix=args.no_matrix)
        total_ok += ok
        total_err += err

    elapsed = int(time.time() - t0)
    print(f"\n{'='*55}")
    print(f"  PIPELINE COMPLETADO en {elapsed}s")
    print(f"  {total_ok} grupos OK, {total_err} errores")
    print(f"  Revisa: */semanas/*/")
    if args.weeks and total_err > 0:
        print(f"  Reintenta fallidos: --weeks ...")
    print(f"{'='*55}")


if __name__ == "__main__":
    main()

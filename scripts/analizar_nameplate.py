"""
analizar_nameplate.py — Analiza Nameplate Excel y genera CSV con estructura original.

Workflow:
  1. Lee el Excel FPR_MYCO_Nameplate_Bronco.xlsx (todas las hojas visibles)
  2. Extrae URLs desde las filas donde el tipo = "URL"
  3. Preserva la estructura original del Excel (Visualización, Módulo, SEO/Formato, etc.)
  4. Detecta anomalías de calidad automáticamente (reporte aparte)
  5. Genera un CSV unificado con naming: {ticket}-{mercado}-{pagina}.csv

Uso:
  python scripts/analizar_nameplate.py
  python scripts/analizar_nameplate.py --ticket GTBEMEAPUB-42479
  python scripts/analizar_nameplate.py --input docs/FPR_MYCO_Nameplate_Bronco.xlsx --output-dir ./

Columnas del CSV (reflejan el Excel original):
  Visualización | Módulo / Sección | SEO/Formato | Copies | Contadores
  | Otros/Notas/Comentarios | URL | Origen

Dependencias: openpyxl
"""

import argparse
import csv
import json
import logging
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("[ERR] openpyxl no instalado. Ejecutá: pip install openpyxl")

logger = logging.getLogger("analizar_nameplate")

# ── Constantes ──
DEFAULT_TICKET = "GTBEMEAPUB-42479"
DEFAULT_MERCADO = "PR"
DEFAULT_PAGINA = "Bronco"
EXCEL_PATH = "docs/FPR_MYCO_Nameplate_Bronco.xlsx"
SHEETS_A_INCLUIR = [
    "SEO Recommendations | ESP",
    "FPR_MYCO_Formato_Nameplate_v5(R",
    "Bronco Version Base",
    "Bronco Version Big Bend",
    "Bronco Version Outer Banks",
    "Bronco Version Badlands",
]
SHEET_VERSION_MAP = {
    "Bronco Version Base": "models/base",
    "Bronco Version Big Bend": "models/big-bend",
    "Bronco Version Outer Banks": "models/outer-banks",
    "Bronco Version Badlands": "models/badlands",
}

# Columnas originales del Excel (0-indexed)
COL_VIS = 0    # A — Visualización
COL_MOD = 1    # B — Módulo / Sección
COL_TIPO = 2   # C — SEO/Formato
COL_CONT = 3   # D — Copies
COL_CTD = 4    # E — Contadores
COL_NOTAS = 5  # F — Otros/Notas/Comentarios
COL_G = 6      # G — (notas ocasionales)

# CSV fields que reflejan el Excel original
CSV_FIELDS = [
    "Visualización",
    "Módulo / Sección",
    "SEO/Formato",
    "Copies",
    "Contadores",
    "Otros/Notas/Comentarios",
    "URL",
    "Origen",
]


def _normalize(val) -> str:
    return "" if val is None else str(val).strip()


def _is_url_row(tipo: str, modulo: str, contenido: str) -> bool:
    """True si esta fila representa una URL (col C o col B)."""
    return contenido.startswith("http") and (tipo.upper() == "URL" or modulo.upper() == "URL")


def _extract_urls(ws) -> list[str]:
    """Extrae todas las URLs de la hoja en orden de aparición."""
    urls: list[str] = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        tipo = _normalize(row[COL_TIPO])
        modulo = _normalize(row[COL_MOD])
        contenido = _normalize(row[COL_CONT])
        if _is_url_row(tipo, modulo, contenido):
            if contenido not in urls:
                urls.append(contenido)
    return urls


def _page_from_url(url: str) -> str:
    """Deriva nombre de página legible desde la URL."""
    m = re.search(r"/bronco/(.*)", url)
    if not m:
        return "Bronco"
    path = m.group(1).strip("/")
    if not path:
        return "Bronco"
    parts = [p for p in path.split("/") if p and p not in ("models",)]
    if parts:
        name = " ".join(p.capitalize() for p in parts)
        return f"Bronco {name}"
    return "Bronco"


def _sheet_label(sheet_name: str) -> str:
    """Etiqueta descriptiva para la columna Origen."""
    labels = {
        "SEO Recommendations | ESP": "SEO Metadatos",
        "FPR_MYCO_Formato_Nameplate_v5(R": "Maestra (Landing)",
        "Bronco Version Base": "Version Base",
        "Bronco Version Big Bend": "Version Big Bend",
        "Bronco Version Outer Banks": "Version Outer Banks",
        "Bronco Version Badlands": "Version Badlands",
    }
    return labels.get(sheet_name, sheet_name)


# ── Validaciones de calidad ──

def _check_label(sheet_name: str, modulo: str) -> str | None:
    """Labels 'Version Base' en sheets que no son Base."""
    if sheet_name in SHEET_VERSION_MAP and sheet_name != "Bronco Version Base":
        if "Version Base" in modulo:
            target = sheet_name.replace("Bronco Version ", "")
            return f"Label '{modulo}' debería ser 'Version {target}'"
    return None


def _check_title(sheet_name: str, tipo: str, contenido: str) -> str | None:
    """Título 'Bronco Base' en hoja Badlands."""
    if sheet_name == "Bronco Version Badlands" and tipo in ("Title", "Sub Title", "Copy"):
        if "Bronco" in contenido and "Base" in contenido and "Badlands" not in contenido:
            return f"Texto copiado de Base: '{contenido[:60]}'"
    return None


def _check_contador(val: str) -> str | None:
    """Valor no numérico en Contadores."""
    if val and not val.isdigit() and val not in ("", "Caracteres", "Palabras", "0"):
        return f"Contador mixto: '{val}'"
    return None


def _check_color(modulo: str, seen: set) -> tuple[str | None, set]:
    """Número de color duplicado."""
    m = re.search(r"Color\s+(\d+)", modulo)
    if m:
        num = m.group(1)
        tag = f"Color {num}"
        if tag in seen:
            return f"{tag} duplicado", seen
        seen.add(tag)
    return None, seen


# ── Extracción de filas ──

def _extract_sheet(wb, sheet_name: str) -> list[dict]:
    """Extrae filas manteniendo la estructura original del Excel."""
    ws = wb[sheet_name]
    rows: list[dict] = []
    page_url = ""
    seccion_actual = ""
    seen_colors: set = set()
    origen = _sheet_label(sheet_name)
    page_url_captured = False

    for row in ws.iter_rows(min_row=1, values_only=True):
        vis = _normalize(row[COL_VIS])
        mod = _normalize(row[COL_MOD])
        tipo = _normalize(row[COL_TIPO])
        cont = _normalize(row[COL_CONT])
        ctd = _normalize(row[COL_CTD])
        notas = _normalize(row[COL_NOTAS])
        notas_g = _normalize(row[COL_G])

        # Detectar sección
        if vis:
            seccion_actual = vis

        # Saltar filas header de columna
        if mod in ("Módulo / Sección",) or tipo == "SEO/Formato":
            if cont in ("Copies", "Contadores", "Otros/Notas/Comentarios") or \
               "SEO RECOMMENDATIONS" in cont:
                continue

        # Capturar la URL de página (la PRIMERA URL del sheet)
        es_url = _is_url_row(tipo, mod, cont)
        if es_url and not page_url_captured:
            page_url = cont
            page_url_captured = True

        # Si no hay contenido ni tipo, es fila estructural
        if not tipo and not cont:
            continue

        notas_final = notas or notas_g

        # --- Quality checks (solo para stats) ---
        flag = _check_label(sheet_name, mod)
        if not flag:
            flag = _check_title(sheet_name, tipo, cont)
        if not flag:
            flag = _check_contador(ctd)
        if not flag and sheet_name in SHEET_VERSION_MAP:
            flag, seen_colors = _check_color(mod, seen_colors)

        # Filas SECTION no tienen URL (son estructurales)
        url_fila = "" if mod.upper() in ("SECTION", "SECCIÓN") else page_url

        rows.append({
            "Visualización": seccion_actual if vis else "",
            "Módulo / Sección": mod,
            "SEO/Formato": tipo,
            "Copies": cont,
            "Contadores": ctd,
            "Otros/Notas/Comentarios": notas_final,
            "URL": url_fila,
            "Origen": origen,
        })

    return rows, page_url


# ── Validaciones globales ──

def _global_validations(all_rows: list[dict], sheet_names: list[str]) -> list[dict]:
    """Valida cobertura y consistencia global."""
    issues: list[dict] = []
    all_text = " ".join(r.get("Copies", "") + " " + r.get("Módulo / Sección", "") for r in all_rows).lower()

    # 1. Raptor sin hoja
    if "raptor" in all_text:
        has_sheet = any("Raptor" in s for s in sheet_names)
        if not has_sheet:
            issues.append({
                "tipo": "cobertura",
                "severidad": "media",
                "descripcion": "Raptor existe en hoja maestra pero no tiene hoja dedicada",
            })

    # 2. Comentario SEO desactualizado
    for r in all_rows:
        c = r.get("Otros/Notas/Comentarios", "") + r.get("Copies", "")
        if "falta actualizar" in c.lower():
            issues.append({
                "tipo": "obsolescencia",
                "severidad": "media",
                "descripcion": "Comentario: 'falta actualizar estas SEO Recommendations'",
            })
            break

    return issues


# ── Quality stats ──

def _quality_stats(all_rows: list[dict]) -> dict:
    """Analiza calidad de todos los datos y devuelve estadísticas."""
    checks = {"labels_incorrectos": 0, "titulo_incorrecto": 0, "contador_mixto": 0, "color_duplicado": 0}

    # Re-aplicar checks contra el contenido
    sheet_cache = {}
    for r in all_rows:
        sheet = r.get("Origen", "")
        sheet_name = None
        for k, v in {
            "SEO Recommendations | ESP": "SEO",
            "FPR_MYCO_Formato_Nameplate_v5(R": "Master",
            "Bronco Version Base": "Version Base",
            "Bronco Version Big Bend": "Version Big Bend",
            "Bronco Version Outer Banks": "Version Outer Banks",
            "Bronco Version Badlands": "Version Badlands",
        }.items():
            if v == sheet:
                sheet_name = k
                break

        if not sheet_name:
            continue

        # Label check
        if _check_label(sheet_name, r.get("Módulo / Sección", "")):
            checks["labels_incorrectos"] += 1

        # Title check
        if _check_title(sheet_name, r.get("SEO/Formato", ""), r.get("Copies", "")):
            checks["titulo_incorrecto"] += 1

        # Contador check
        if _check_contador(r.get("Contadores", "")):
            checks["contador_mixto"] += 1

    # Color duplicates (by sheet version)
    for sheet_name in SHEET_VERSION_MAP:
        seen = set()
        sheet_label = _sheet_label(sheet_name)
        for r in all_rows:
            if r.get("Origen") != sheet_label:
                continue
            flag, seen = _check_color(r.get("Módulo / Sección", ""), seen)
            if flag:
                checks["color_duplicado"] += 1

    return checks


# ── Output ──

def _write_csv(output_path: Path, all_rows: list[dict]):
    """Escribe CSV con estructura del Excel original."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS, extrasaction="ignore")
        w.writeheader()
        for row in all_rows:
            w.writerow(row)
    print(f"  CSV: {output_path.resolve()}")
    print(f"  Filas: {len(all_rows)}")


def _write_summary(output_dir: Path, all_rows: list[dict], issues: list[dict],
                   quality: dict, ticket: str, mercado: str, pagina: str,
                   sheets_ok: list[str]):
    """Escribe resumen JSON con estadísticas y calidad."""
    by_sheet: dict[str, int] = {}
    by_url: dict[str, int] = {}
    for r in all_rows:
        s = r.get("Origen", "(sin origen)")
        u = r.get("URL", "") or "(sin URL)"
        by_sheet[s] = by_sheet.get(s, 0) + 1
        by_url[u] = by_url.get(u, 0) + 1

    summary = {
        "archivo": f"{ticket}-{mercado}-{pagina}.csv",
        "ticket": ticket,
        "mercado": mercado,
        "pagina": pagina,
        "total_filas": len(all_rows),
        "sheets_procesadas": len(sheets_ok),
        "sheets": sheets_ok,
        "urls_encontradas": list(by_url.keys()),
        "filas_por_sheet": by_sheet,
        "filas_por_url": by_url,
        "calidad": {
            "total_flags": sum(quality.values()),
            "detalle": quality,
        },
        "incidencias": [
            {"tipo": i["tipo"], "severidad": i["severidad"], "descripcion": i["descripcion"]}
            for i in issues
        ],
    }

    json_path = output_dir / f"{ticket}-{mercado}-{pagina}-summary.json"
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print(f"  JSON: {json_path.resolve()}")

    # Console summary
    print(f"\n{'='*50}")
    print(f"  RESUMEN — {ticket}-{mercado}-{pagina}")
    print(f"{'='*50}")
    print(f"  Sheets:             {len(sheets_ok)}")
    print(f"  Filas totales:      {len(all_rows)}")
    print(f"  URLs encontradas:   {len(by_url)}")
    print(f"  Quality flags:      {sum(quality.values())}")
    for k, v in quality.items():
        if v:
            print(f"    {k}: {v}")
    for i in issues:
        print(f"  [{i['severidad'].upper()}] {i['descripcion']}")
    print(f"{'='*50}\n")


# ── Main ──

def main():
    parser = argparse.ArgumentParser(
        description="Analiza Nameplate Excel y genera CSV con estructura original.",
    )
    parser.add_argument("--input", default=EXCEL_PATH)
    parser.add_argument("--ticket", default=DEFAULT_TICKET)
    parser.add_argument("--mercado", default=DEFAULT_MERCADO)
    parser.add_argument("--pagina", default=DEFAULT_PAGINA)
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERR] No se encuentra: {input_path.resolve()}")
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else Path(args.ticket)
    output_filename = f"{args.ticket}-{args.mercado}-{args.pagina}.csv"
    output_path = output_dir / output_filename

    # ── Cargar Excel ──
    print(f"\n[1/4] Leyendo Excel: {input_path.resolve()}")
    wb = openpyxl.load_workbook(str(input_path), data_only=True, read_only=True)

    # ── Extraer ──
    print("[2/4] Extrayendo sheets...")
    all_rows: list[dict] = []
    sheets_ok: list[str] = []
    url_encontrada = ""

    for sheet_name in SHEETS_A_INCLUIR:
        if sheet_name not in wb.sheetnames:
            logger.warning("Sheet no encontrada: %s", sheet_name)
            continue
        ws = wb[sheet_name]
        if ws.sheet_state == "hidden":
            logger.info("Sheet oculta: %s", sheet_name)
            continue

        logger.info("Procesando: %s", sheet_name)
        rows, url = _extract_sheet(wb, sheet_name)
        all_rows.extend(rows)
        sheets_ok.append(sheet_name)
        if url and not url_encontrada:
            url_encontrada = url
        logger.debug("  -> %d filas", len(rows))

    if not all_rows:
        print("[ERR] No se extrajeron datos")
        sys.exit(1)

    # ── Validar ──
    print("[3/4] Analizando calidad...")
    issues = _global_validations(all_rows, SHEETS_A_INCLUIR)
    quality = _quality_stats(all_rows)

    # ── Generar ──
    print("[4/4] Generando reportes...")
    _write_csv(output_path, all_rows)
    _write_summary(output_dir, all_rows, issues, quality, args.ticket, args.mercado, args.pagina, sheets_ok)

    print(f"[OK] Analisis completado -> {output_dir}/")


if __name__ == "__main__":
    main()

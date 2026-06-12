"""
extract_aa.py — Extrae campos seleccionables del JSON de Adobe Analytics en Excel.

Lee col E (AA analytics automatico), extrae campos elegidos, escribe en col F (estructurado).

Uso:
  python extract_aa.py                                          # valores por defecto
  python extract_aa.py --keep page,request,props,evars          # solo estos 4
  python extract_aa.py --keep events,visitor,products           # campos adicionales
  python extract_aa.py --keep all                               # TODO el JSON (solo pretty-print)
  python extract_aa.py --input historial.xlsx                   # otro archivo
  python extract_aa.py --score                                  # métricas detalladas
  python extract_aa.py --verbose                               # logging debug

Maneja 2 formatos detectados en col E:
  - Grupo 1 (filas 2-10): keys "eVars" + "props" con prop1/eVar1
  - Grupo 2 (filas 11-15): keys "evars" + "props" con c1/v1
"""

import json
import logging
import os
import sys
import openpyxl
from openpyxl.styles import Alignment, PatternFill

DEFAULT_KEEP = ["page", "request", "props", "evars"]
ALL_FIELDS = [
    "solution", "page", "request", "visitor", "hit", "page_data",
    "browser", "events", "eVars", "evars", "props",
    "products", "frame", "adobe", "pageName", "channel", "page_attributes", "environment",
]


def _save_workbook(wb, path):
    try:
        wb.save(path)
        return path
    except PermissionError:
        name, ext = os.path.splitext(path)
        fallback = f"{name}_limpio{ext}"
        wb.save(fallback)
        logging.warning("Archivo bloqueado, guardado como %s", fallback)
        return fallback


def extract_fields(data: dict, keep: list[str]) -> dict:
    """
    Extrae del JSON original solo los campos en `keep`.
    Maneja eVars/evars indistintamente.
    """
    result = {}
    for field in keep:
        if field == "evars":
            # Unifica eVars (Grupo 1) y evars (Grupo 2)
            val = data.get("eVars") or data.get("evars")
            if val is not None:
                result["evars"] = val
        elif field == "eVars":
            val = data.get("eVars") or data.get("evars")
            if val is not None:
                result["eVars"] = val
        else:
            val = data.get(field)
            if val is not None:
                result[field] = val
    return result


def count_values(obj) -> dict:
    """Cuenta elementos de cada campo en el JSON extraído (para métricas)."""
    stats = {}
    for key, val in obj.items():
        if isinstance(val, dict):
            stats[key] = f"{len(val)} keys"
        elif isinstance(val, list):
            stats[key] = f"{len(val)} items"
        elif isinstance(val, str):
            stats[key] = f"{len(val)} chars" if len(val) > 50 else val[:50]
        else:
            stats[key] = str(val)[:50]
    return stats


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Extrae campos del JSON de Adobe Analytics en Excel")
    parser.add_argument("--input", default="RevisionManual.xlsx", help="Archivo Excel")
    parser.add_argument("--sheet", help="Nombre del sheet (default: auto-detecta el mas reciente)")
    parser.add_argument("--keep", default=",".join(DEFAULT_KEEP),
                        help=f"Campos a conservar (separados por coma). Opciones: {','.join(ALL_FIELDS)}. Usar 'all' para todo.")
    parser.add_argument("--score", action="store_true", help="Mostrar metricas detalladas por fila")
    parser.add_argument("--verbose", action="store_true", help="Logging detallado (debug)")
    args = parser.parse_args()

    # ── Logging ──
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format="%(asctime)s | %(levelname)s | %(message)s",
                        handlers=[logging.StreamHandler(sys.stderr)])

    # Parsear keep
    if args.keep == "all":
        keep = ALL_FIELDS
    else:
        keep = [k.strip() for k in args.keep.split(",") if k.strip()]
        # Validar
        unknown = [k for k in keep if k not in ALL_FIELDS and k != "all"]
        if unknown:
            logging.error("Campos desconocidos: %s. Opciones validas: %s o 'all'",
                          unknown, ','.join(ALL_FIELDS))
            sys.exit(1)

    wb = openpyxl.load_workbook(args.input)

    # Seleccionar sheet
    if args.sheet:
        if args.sheet not in wb.sheetnames:
            logging.error("Sheet '%s' no encontrado", args.sheet)
            sys.exit(1)
        ws = wb[args.sheet]
    else:
        # Auto-detect: el sheet de datos mas reciente (excluye _control)
        candidates = [s for s in wb.sheetnames if s != "_control"]
        candidates.sort(key=lambda s: s if "-" in s else "0", reverse=True)
        if not candidates:
            logging.error("No hay sheets de datos en el archivo")
            sys.exit(1)
        ws = wb[candidates[0]]
        logging.info("Sheet auto-detectado: '%s'", candidates[0])

    # Validar col E
    e_header = str(ws.cell(1, 5).value or "").strip().lower()
    if e_header and "analytics" not in e_header:
        logging.warning("Col E header no esperado: '%s'", ws.cell(1, 5).value)

    total = 0
    errores = []
    stats_rows = []

    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row, 5).value  # col E
        if not raw:
            errores.append((row, "COL_E_EMPTY", "col E vacia"))
            continue

        raw_str = str(raw).strip()
        if not raw_str or raw_str.startswith("(no") or raw_str.startswith("(error"):
            code = "NO_AA_DATA"
            errores.append((row, code, f"col E sin datos validos: {raw_str[:60]}"))
            continue

        try:
            data = json.loads(raw_str)
        except json.JSONDecodeError as e:
            errores.append((row, "JSON_INVALID", f"JSON invalido: {e}"))
            continue

        extracted = extract_fields(data, keep)

        # Si no se extrajo nada, avisar
        if not extracted:
            errores.append((row, "NO_FIELDS_MATCHED", "ningun campo coincidio en el JSON origen"))
            continue

        pretty = json.dumps(extracted, indent=2, ensure_ascii=False)
        cell = ws.cell(row, 6)
        cell.value = pretty
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        total += 1

        if args.score:
            stats_rows.append({"row": row, "fields": list(extracted.keys()), **count_values(extracted)})

    # Ancho col F
    ws.column_dimensions["F"].width = 80

    # Guardar
    out = _save_workbook(wb, args.input)
    wb.close()

    logging.info("Guardado: %s", out)
    logging.info("OK Procesadas: %d filas", total)
    logging.info("Campos extraidos: %s", keep)

    if args.score and stats_rows:
        print(f"\n{'='*55}", file=sys.stderr)
        print(f"  MÉTRICAS POR FILA (--score)", file=sys.stderr)
        print(f"{'='*55}", file=sys.stderr)
        for s in stats_rows:
            print(f"  Fila {s['row']}: campos={s['fields']}", file=sys.stderr)
            for k, v in s.items():
                if k not in ("row", "fields"):
                    print(f"    {k}: {v}", file=sys.stderr)
            print(file=sys.stderr)

    if errores:
        logging.warning("Errores: %d", len(errores))
        for r, code, msg in errores:
            logging.debug("  Fila %d [%s]: %s", r, code, msg)
    else:
        logging.info("Sin errores")


if __name__ == "__main__":
    main()

"""Extrae URLs de Excel → urls.json con soporte de mercado.

Uso:
  python _gen_urls.py                                         # urls.json desde RevisionManual.xlsx
  python _gen_urls.py --input otro.xlsx                       # desde otro archivo
  python _gen_urls.py --output mis_urls.json                  # nombre custom de salida
  python _gen_urls.py --default-market PR                     # asigna "market": "PR" a todas
  python _gen_urls.py --input PR.xlsx --default-market PR     # combo tipico
"""
import openpyxl, json, argparse

DEFAULT_INPUT = "RevisionManual.xlsx"
DEFAULT_OUTPUT = "urls.json"

def extract(input_path: str, default_market: str = None) -> list:
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    entries = []
    for row in range(2, ws.max_row + 1):
        url = ws.cell(row, 2).value
        if not url:
            continue
        url = str(url).strip()
        nombre = ws.cell(row, 1).value

        if nombre and str(nombre).strip():
            entry = {"url": url, "nombre": str(nombre).strip()}
        else:
            entry = {"url": url}

        if default_market:
            entry["market"] = default_market.upper()

        entries.append(entry)
    return entries

def main():
    parser = argparse.ArgumentParser(description="Extrae URLs de Excel a JSON")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Excel de entrada")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="JSON de salida")
    parser.add_argument("--default-market", help="Asigna market a todas las URLs (ej: PR, MX, ES)")
    args = parser.parse_args()

    entries = extract(args.input, args.default_market)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)
    print(f"Generado {args.output} con {len(entries)} URLs desde {args.input}")
    if args.default_market:
        print(f"  Market asignado: {args.default_market.upper()}")

if __name__ == "__main__":
    main()

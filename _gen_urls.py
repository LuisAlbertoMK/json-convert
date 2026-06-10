"""Extrae URLs de RevisionManual.xlsx → urls.json

Uso:
  python _gen_urls.py                              # urls.json desde RevisionManual.xlsx
  python _gen_urls.py --input otro.xlsx             # desde otro archivo
  python _gen_urls.py --output mis_urls.json        # nombre custom de salida
"""
import openpyxl, json, argparse

DEFAULT_INPUT = "RevisionManual.xlsx"
DEFAULT_OUTPUT = "urls.json"

def extract(input_path: str) -> list:
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
            entries.append({"url": url, "nombre": str(nombre).strip()})
        else:
            entries.append(url)
    return entries

def main():
    parser = argparse.ArgumentParser(description="Extrae URLs de Excel a JSON")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Excel de entrada")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="JSON de salida")
    args = parser.parse_args()

    entries = extract(args.input)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)
    print(f"Generado {args.output} con {len(entries)} URLs desde {args.input}")

if __name__ == "__main__":
    main()

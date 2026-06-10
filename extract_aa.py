"""
extract_aa.py — Extrae campos clave de Adobe Analytics JSON en RevisionManual.xlsx

Lee col D (adobe analytics original), extrae:
  page, title, url, request, props, evars

Maneja 2 formatos detectados en el archivo:
  - Grupo 1 (filas 2-10): keys "eVars" + "props" con prop1/eVar1
  - Grupo 2 (filas 11-15): keys "evars" + "props" con c1/v1

Escribe resultado pretty-print en col E del mismo archivo.
"""

import json
import openpyxl
from openpyxl.styles import Alignment
from copy import copy

INPUT_FILE = "RevisionManual.xlsx"
OUTPUT_FILE = None  # se determina en main()


def extract_fields(data: dict) -> dict:
    """
    Extrae page, title, url, request, props, evars del JSON original.
    Detecta automáticamente el grupo por las keys presentes.
    """
    result = {}

    # page (contiene title + url adentro)
    result["page"] = data.get("page", {})

    # request
    result["request"] = data.get("request", {})

    # props — está en "props" en AMBOS grupos, con keys distintas
    result["props"] = data.get("props", {})

    # evars — puede ser "eVars" (Grupo 1) o "evars" (Grupo 2)
    result["evars"] = data.get("eVars") or data.get("evars") or {}

    return result


def main():
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active

    total = 0
    errores = []

    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row, 4).value  # col D
        if not raw:
            errores.append((row, "col D vacía"))
            continue

        try:
            data = json.loads(str(raw))
            extracted = extract_fields(data)
            pretty = json.dumps(extracted, indent=2, ensure_ascii=False)
            cell = ws.cell(row, 5)
            cell.value = pretty  # col E
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            total += 1
        except Exception as e:
            errores.append((row, str(e)))

    # Ajustar ancho de col E para legibilidad
    ws.column_dimensions["E"].width = 80

    # Guardar — intenta mismo archivo, fallback a _limpio si está bloqueado
    out = INPUT_FILE
    try:
        wb.save(out)
    except PermissionError:
        import os
        name, ext = os.path.splitext(INPUT_FILE)
        out = f"{name}_limpio{ext}"
        wb.save(out)
    wb.close()
    print(f"Guardado: {out}")

    print(f"OK Procesadas: {total} filas")
    if errores:
        print(f"ERROR Errores: {len(errores)}")
        for r, e in errores:
            print(f"   Fila {r}: {e}")
    else:
        print("OK Sin errores")


if __name__ == "__main__":
    main()

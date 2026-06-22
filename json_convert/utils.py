"""
utils.py — Utilidades compartidas del paquete json_convert.
"""

import json


def load_json(path: str) -> dict | list:
    """Lee un archivo JSON y retorna su contenido."""
    with open(path, encoding="utf-8") as f:
        return json.load(f)

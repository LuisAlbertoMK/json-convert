"""
test_utils.py — Tests para json_convert/utils.py.

Requiere: unittest (stdlib, 0 deps externas).
"""

import json
import os
import tempfile
import unittest

from json_convert.utils import load_json


class TestLoadJson(unittest.TestCase):
    """Tests para load_json()."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def _write(self, name: str, content: str) -> str:
        path = os.path.join(self.tmpdir, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return path

    def test_load_dict(self):
        """JSON object → retorna dict."""
        path = self._write("dict.json", '{"page": "home", "id": 42}')
        result = load_json(path)
        self.assertEqual(result, {"page": "home", "id": 42})

    def test_load_list(self):
        """JSON array → retorna list."""
        path = self._write("list.json", '["a", "b", "c"]')
        result = load_json(path)
        self.assertEqual(result, ["a", "b", "c"])

    def test_load_nested(self):
        """JSON anidado."""
        path = self._write("nested.json", '{"data": {"items": [1, 2]}}')
        result = load_json(path)
        self.assertEqual(result["data"]["items"], [1, 2])

    def test_file_not_found(self):
        """Archivo inexistente → FileNotFoundError."""
        with self.assertRaises(FileNotFoundError):
            load_json(os.path.join(self.tmpdir, "nope.json"))

    def test_invalid_json(self):
        """Archivo con JSON inválido → JSONDecodeError."""
        path = self._write("bad.json", "{not: json}")
        with self.assertRaises(json.JSONDecodeError):
            load_json(path)

    def test_empty_file(self):
        """Archivo vacío → JSONDecodeError."""
        path = self._write("empty.json", "")
        with self.assertRaises(json.JSONDecodeError):
            load_json(path)

    def test_utf8_bom(self):
        """BOM UTF-8 → JSONDecodeError (encoding=utf-8 no tolera BOM)."""
        path = self._write("bom.json", "\ufeff{}")
        with self.assertRaises(json.JSONDecodeError):
            load_json(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)

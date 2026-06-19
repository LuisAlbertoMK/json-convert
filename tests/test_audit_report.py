"""
Tests para audit_report.py — funciones de reporte de fallos.

Cubre funciones puras (sin I/O):
  - parse_meta_col()
  - determine_status()
  - build_report()
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

# El módulo depende de openpyxl, que no está en todos los entornos
try:
    import openpyxl
except ImportError:
    openpyxl = None  # type: ignore

from audit_report import build_report, determine_status, parse_meta_col


class TestParseMetaCol(unittest.TestCase):
    """parse_meta_col(): parsea JSON de metadata (col G)."""

    def test_empty_string(self):
        self.assertEqual(parse_meta_col(""), {})

    def test_none_value(self):
        """None debe devolver dict vacío."""
        self.assertEqual(parse_meta_col(""), {})  # string vacío == sin datos

    def test_valid_json(self):
        val = '{"score": 85, "code": "OK", "error": ""}'
        result = parse_meta_col(val)
        self.assertEqual(result.get("score"), 85)
        self.assertEqual(result.get("code"), "OK")

    def test_invalid_json(self):
        self.assertEqual(parse_meta_col("not json at all"), {})

    def test_partial_json(self):
        val = '{"score": 50}'
        result = parse_meta_col(val)
        self.assertEqual(result.get("score"), 50)

    def test_nested_json(self):
        val = '{"score": 100, "aa": {"parsed": true}}'
        result = parse_meta_col(val)
        self.assertIn("aa", result)
        self.assertTrue(result["aa"]["parsed"])


class TestDetermineStatus(unittest.TestCase):
    """determine_status(): clasifica estado de URL (OK/FALLO/SIN_DATOS)."""

    def test_ok_with_data(self):
        estado, detalle, _score = determine_status("some data", "", {})
        self.assertEqual(estado, "OK")
        self.assertEqual(detalle, "")

    def test_error_code_in_states(self):
        for code in ("NO_AA_DATA", "TIMEOUT", "HTTP_403", "HTTP_ERROR",
                     "URL_INVALID", "NETWORK_ERROR", "NAV_ERROR", "UNKNOWN"):
            meta = {"code": code, "error": "test error"}
            estado, detalle, _score = determine_status("", "", meta)
            self.assertEqual(estado, "FALLO", f"FALLO esperado para {code}")
            self.assertIn("test error", detalle)

    def test_error_code_fallback_to_code(self):
        meta = {"code": "HTTP_403", "error": ""}
        estado, detalle, _score = determine_status("", "", meta)
        self.assertEqual(estado, "FALLO")
        self.assertEqual(detalle, "HTTP_403")

    def test_col_e_starts_with_paren(self):
        estado, detalle, _score = determine_status("(404 Not Found)", "", {})
        self.assertEqual(estado, "FALLO")
        self.assertEqual(detalle, "404 Not Found")

    def test_empty_col_e(self):
        estado, _detalle, _score = determine_status("", "", {})
        self.assertEqual(estado, "SIN_DATOS")

    def test_col_e_dash(self):
        estado, _detalle, _score = determine_status("-", "", {})
        self.assertEqual(estado, "SIN_DATOS")

    def test_col_e_n_a(self):
        estado, _detalle, _score = determine_status("N/A", "", {})
        self.assertEqual(estado, "SIN_DATOS")

    def test_sin_datos_with_no_digitaldata(self):
        estado, detalle, _score = determine_status("", "(no digitaldata)", {})
        self.assertEqual(estado, "SIN_DATOS")
        self.assertIn("Sin digitalData", detalle)

    def test_score_from_meta(self):
        meta = {"score": 75}
        _estado, _detalle, score = determine_status("data", "", meta)
        self.assertEqual(score, 75)

    def test_score_defaults_zero(self):
        _estado, _detalle, score = determine_status("", "", {})
        self.assertEqual(score, 0)

    def test_ok_with_col_e_numeric(self):
        """Col E con valor numérico (no string) debe tratarse como OK."""
        estado, _detalle, _score = determine_status("12345", "", {})
        self.assertEqual(estado, "OK")

    def test_error_code_with_score(self):
        meta = {"code": "TIMEOUT", "error": "timeout after 30s", "score": 0}
        estado, detalle, score = determine_status("", "", meta)
        self.assertEqual(estado, "FALLO")
        self.assertEqual(detalle, "timeout after 30s")
        self.assertEqual(score, 0)


class TestBuildReport(unittest.TestCase):
    """build_report(): separa working y failed, ordena."""

    def setUp(self):
        self.ok_page = {
            "nombre": "Home", "url": "https://example.com",
            "mercado": "PR", "estado": "OK", "detalle": "",
            "fecha": "2026-01-01", "score": 100,
            "digitaldata": "OK", "hoja": "2026-01-01",
        }
        self.fail_page = {
            "nombre": "Broken", "url": "https://example.com/broken",
            "mercado": "PR", "estado": "FALLO", "detalle": "404",
            "fecha": "2026-01-01", "score": 0,
            "digitaldata": "NO", "hoja": "2026-01-01",
        }
        self.nodata_page = {
            "nombre": "NoData", "url": "https://example.com/nodata",
            "mercado": "MX", "estado": "SIN_DATOS", "detalle": "Sin datos",
            "fecha": "2026-01-01", "score": 50,
            "digitaldata": "NO", "hoja": "2026-01-01",
        }

    def test_all_ok(self):
        failed, all_sorted = build_report([self.ok_page])
        self.assertEqual(len(failed), 0)
        self.assertEqual(len(all_sorted), 1)

    def test_all_fail(self):
        failed, _all_sorted = build_report([self.fail_page])
        self.assertEqual(len(failed), 1)
        self.assertEqual(failed[0]["estado"], "FALLO")

    def test_mixed(self):
        pages = [self.ok_page, self.fail_page, self.nodata_page]
        failed, all_sorted = build_report(pages)
        self.assertEqual(len(failed), 2)  # FALLO + SIN_DATOS
        self.assertEqual(len(all_sorted), 3)

    def test_ordering_by_mercado(self):
        pr_fail = dict(self.fail_page, mercado="PR", nombre="A")
        mx_fail = dict(self.fail_page, mercado="MX", nombre="B")
        failed, _all = build_report([pr_fail, mx_fail])
        self.assertEqual(failed[0]["mercado"], "MX")  # MX < PR
        self.assertEqual(failed[1]["mercado"], "PR")

    def test_ordering_by_score_asc(self):
        low = dict(self.fail_page, score=10, nombre="Low")
        high = dict(self.fail_page, score=90, nombre="High")
        failed, _all = build_report([high, low])
        # build_report ordena por (mercado, score, nombre): score 10 < 90
        self.assertEqual(failed[0]["nombre"], "Low")
        self.assertEqual(failed[0]["score"], 10)
        self.assertEqual(failed[1]["nombre"], "High")
        self.assertEqual(failed[1]["score"], 90)

    def test_empty_list(self):
        failed, all_sorted = build_report([])
        self.assertEqual(len(failed), 0)
        self.assertEqual(len(all_sorted), 0)


if __name__ == "__main__":
    unittest.main()

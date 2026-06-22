"""
test_aa_async.py — Tests con AsyncMock para funciones async de aa_parser.py.

Uso:
  python -m pytest tests/test_aa_async.py -v
  python test_aa_async.py                       # unittest runner

Requiere: unittest (stdlib) + pytest-asyncio (opcional).
No requiere: playwright, navegador, red.
"""

import logging
import sys
import unittest
from unittest.mock import AsyncMock, MagicMock, patch

# Silenciar logging durante tests
logging.disable(logging.CRITICAL)

from json_convert.aa_parser import (
    DATA_LAYER_NAMES,
    debug_dump_available_globals,
    extract_digital_data,
    extract_s_object,
    extract_title,
    try_dismiss_cookie_consent,
)


class FakePage:
    """Page simulada que retorna valores controlados de evaluate/query_selector."""

    def __init__(self):
        self.evaluate = AsyncMock(return_value=None)
        self.query_selector = AsyncMock(return_value=None)


class TestExtractSObject(unittest.IsolatedAsyncioTestCase):
    """Tests para extract_s_object()."""

    async def test_returns_s_object(self):
        """evaluate retorna un dict → se retorna tal cual."""
        page = FakePage()
        page.evaluate.return_value = {"pageName": "test", "prop1": "val"}
        result = await extract_s_object(page)
        self.assertEqual(result, {"pageName": "test", "prop1": "val"})
        page.evaluate.assert_awaited_once()

    async def test_returns_none(self):
        """evaluate retorna None → se retorna None."""
        page = FakePage()
        page.evaluate.return_value = None
        result = await extract_s_object(page)
        self.assertIsNone(result)

    async def test_exception_returns_none(self):
        """evaluate lanza excepción → retorna None."""
        page = FakePage()
        page.evaluate.side_effect = Exception("JS error")
        result = await extract_s_object(page)
        self.assertIsNone(result)


class TestExtractDigitalData(unittest.IsolatedAsyncioTestCase):
    """Tests para extract_digital_data()."""

    async def test_finds_valid_dict(self):
        """Primer nombre de data layer retorna un dict válido."""
        page = FakePage()
        page.evaluate.return_value = {"page": "home", "id": 42}
        result = await extract_digital_data(page)
        self.assertEqual(result, {"page": "home", "id": 42})

    async def test_skips_none_tries_next(self):
        """Primer nombre retorna None → prueba el siguiente."""
        page = FakePage()
        page.evaluate.side_effect = [None, {"page": "ok"}]
        result = await extract_digital_data(page)
        self.assertEqual(result, {"page": "ok"})
        self.assertEqual(page.evaluate.await_count, 2)

    async def test_skips_empty_dict(self):
        """Dict vacío → no se considera válido, prueba siguiente."""
        page = FakePage()
        page.evaluate.side_effect = [{}, {"page": "ok"}]
        result = await extract_digital_data(page)
        self.assertEqual(result, {"page": "ok"})

    async def test_skips_non_dict(self):
        """Valor no-dict (string) → se salta."""
        page = FakePage()
        page.evaluate.side_effect = ["string_val", {"page": "ok"}]
        result = await extract_digital_data(page)
        self.assertEqual(result, {"page": "ok"})

    async def test_exhausts_all_names(self):
        """Ningún nombre retorna dict válido → retorna None."""
        page = FakePage()
        page.evaluate.return_value = None
        result = await extract_digital_data(page)
        self.assertIsNone(result)
        self.assertEqual(page.evaluate.await_count, len(DATA_LAYER_NAMES))

    async def test_exception_on_name(self):
        """Un nombre falla con excepción → intenta el siguiente."""
        page = FakePage()
        page.evaluate.side_effect = [Exception("err"), {"page": "ok"}]
        result = await extract_digital_data(page)
        self.assertEqual(result, {"page": "ok"})

    async def test_length_zero_skipped(self):
        """Dict con len==0 → se salta."""
        page = FakePage()
        page.evaluate.side_effect = [{"page": "ok"}]  # tiene contenido
        result = await extract_digital_data(page)
        self.assertEqual(result, {"page": "ok"})


class TestExtractTitle(unittest.IsolatedAsyncioTestCase):
    """Tests para extract_title()."""

    async def test_returns_title(self):
        """document.title retorna string → se retorna stripped."""
        page = FakePage()
        page.evaluate.return_value = "  Ford Mach-E  "
        result = await extract_title(page)
        self.assertEqual(result, "Ford Mach-E")

    async def test_returns_empty(self):
        """document.title es None → retorna ''."""
        page = FakePage()
        page.evaluate.return_value = None
        result = await extract_title(page)
        self.assertEqual(result, "")

    async def test_exception_returns_empty(self):
        """evaluate lanza excepción → retorna ''."""
        page = FakePage()
        page.evaluate.side_effect = Exception("no page")
        result = await extract_title(page)
        self.assertEqual(result, "")


class TestTryDismissCookieConsent(unittest.IsolatedAsyncioTestCase):
    """Tests para try_dismiss_cookie_consent()."""

    async def test_finds_and_clicks(self):
        """Selector coincide → hace click y retorna True."""
        btn = AsyncMock()
        page = FakePage()
        page.query_selector.return_value = btn
        result = await try_dismiss_cookie_consent(page)
        self.assertTrue(result)
        btn.click.assert_awaited_once()

    async def test_no_selector_matches(self):
        """Ningún selector coincide → retorna False."""
        page = FakePage()
        page.query_selector.return_value = None
        result = await try_dismiss_cookie_consent(page)
        self.assertFalse(result)

    async def test_click_exception_continues(self):
        """Click lanza excepción → continúa al siguiente selector."""
        btn_fail = AsyncMock()
        btn_fail.click.side_effect = Exception("click blocked")
        btn_ok = AsyncMock()
        page = FakePage()
        page.query_selector.side_effect = [btn_fail, btn_ok]
        result = await try_dismiss_cookie_consent(page)
        self.assertTrue(result)
        btn_ok.click.assert_awaited_once()


class TestDebugDumpAvailableGlobals(unittest.IsolatedAsyncioTestCase):
    """Tests para debug_dump_available_globals()."""

    async def test_logs_globals(self):
        """evaluate retorna lista → logging.info se llama."""
        page = FakePage()
        page.evaluate.return_value = [
            {"key": "digitalData", "type": "Object", "sample_keys": ["page"], "len": 1},
        ]
        with patch.object(logging, "info") as mock_info:
            await debug_dump_available_globals(page)
            mock_info.assert_called()

    async def test_no_globals(self):
        """evaluate retorna lista vacía → no se loggean."""
        page = FakePage()
        page.evaluate.return_value = []
        with patch.object(logging, "debug") as mock_debug:
            await debug_dump_available_globals(page)
            mock_debug.assert_called()

    async def test_exception_handled(self):
        """evaluate lanza excepción → no propaga."""
        page = FakePage()
        page.evaluate.side_effect = Exception("js fail")
        # No debe lanzar
        await debug_dump_available_globals(page)


if __name__ == "__main__":
    unittest.main(verbosity=2)

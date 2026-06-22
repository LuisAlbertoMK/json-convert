"""
test_browser.py — Tests para browser.py (process_url, _backoff_delay, _page_name_from_url).

Uso:
  python -m pytest tests/test_browser.py -v
  python test_browser.py                       # unittest runner

No requiere: playwright, navegador, red.
"""

import asyncio
import logging
import unittest
from unittest.mock import AsyncMock, MagicMock, patch

# Silenciar logging durante tests
logging.disable(logging.CRITICAL)

from playwright.async_api import TimeoutError as PwTimeout

from json_convert.browser import (
    _backoff_delay,
    _page_name_from_url,
    _shutdown_flag,
    process_url,
)


class FakePage:
    """Page simulada con métodos async configurables vía AsyncMock."""

    def __init__(self):
        self.goto = AsyncMock()
        self.on = MagicMock()
        self.remove_listener = MagicMock()
        self.wait_for_function = AsyncMock()
        self.wait_for_timeout = AsyncMock()
        self.close = AsyncMock()
        self.setContent = AsyncMock()
        self.query_selector = AsyncMock(return_value=None)
        self.evaluate = AsyncMock()


# ═══════════════════════════════════════════════════════════════════
# _page_name_from_url — función pura
# ═══════════════════════════════════════════════════════════════════

class TestPageNameFromUrl(unittest.TestCase):
    """Tests para _page_name_from_url(). Función pura, sin mocks."""

    def test_standard_html_page(self):
        """URL con .html → nombre del archivo sin extensión, title case."""
        name = _page_name_from_url(
            "https://www.ford.com/pr/driving-precautions.html"
        )
        self.assertEqual(name, "Driving Precautions")

    def test_no_html_extension(self):
        """URL sin .html → último segmento del path."""
        name = _page_name_from_url(
            "https://www.ford.com/pr/all-vehicles"
        )
        self.assertEqual(name, "All Vehicles")

    def test_multi_segment_name(self):
        """Guiones múltiples → espacios."""
        name = _page_name_from_url(
            "https://www.ford.com/pr/fuel-saving-tips.html"
        )
        self.assertEqual(name, "Fuel Saving Tips")

    def test_url_with_trailing_slash(self):
        """URL con trailing slash → último segmento después del /."""
        name = _page_name_from_url("https://www.ford.com/pr/")
        self.assertEqual(name, "Pr")

    def test_root_domain(self):
        """Solo dominio → el hostname title-cased."""
        name = _page_name_from_url("https://www.ford.com")
        self.assertEqual(name, "Www.Ford.Com")

    def test_no_slashes_returns_title_cased(self):
        """String sin '/' → title-case."""
        name = _page_name_from_url("some-string-without-path")
        self.assertEqual(name, "Some String Without Path")

    def test_non_ford_domain(self):
        """Dominio no-ford → igual toma el último segmento."""
        name = _page_name_from_url(
            "https://other.com/path/legal-page.html"
        )
        self.assertEqual(name, "Legal Page")


# ═══════════════════════════════════════════════════════════════════
# _backoff_delay — función pura con random
# ═══════════════════════════════════════════════════════════════════

class TestBackoffDelay(unittest.TestCase):
    """Tests para _backoff_delay(). Jitter controlado con patch."""

    def test_attempt_0_base_2(self):
        """attempt=0 → base*2^0=2 + jitter(0,1) = 2-3."""
        with patch("random.uniform", return_value=0.5):
            delay = _backoff_delay(0, base=2.0)
        self.assertEqual(delay, 2.5)

    def test_attempt_1_doubles(self):
        """attempt=1 → base*2^1=4 + jitter(0,2) ≈ 4-6."""
        with patch("random.uniform", return_value=1.0):
            delay = _backoff_delay(1, base=2.0)
        self.assertEqual(delay, 5.0)

    def test_exponential_increase(self):
        """Cada intento incrementa el delay (jitter=0 para comparar)."""
        with patch("random.uniform", return_value=0):
            d0 = _backoff_delay(0)
            d1 = _backoff_delay(1)
            d2 = _backoff_delay(2)
        self.assertLess(d0, d1)
        self.assertLess(d1, d2)

    def test_capped_at_max_delay(self):
        """Intentos altos → cap en max_delay."""
        with patch("random.uniform", return_value=0):
            delay = _backoff_delay(10, base=2.0, max_delay=30.0)
        self.assertEqual(delay, 30.0)

    def test_jitter_never_exceeds_max(self):
        """Jitter incluido no supera max_delay."""
        with patch("random.uniform", return_value=1.0):  # jitter = delay * 0.5
            delay = _backoff_delay(10, base=2.0, max_delay=30.0)
        self.assertLessEqual(delay, 30.0)

    def test_custom_base(self):
        """Base personalizada (ej: 1s)."""
        with patch("random.uniform", return_value=0):
            delay = _backoff_delay(0, base=1.0)
        self.assertEqual(delay, 1.0)


# ═══════════════════════════════════════════════════════════════════
# process_url — tests con Page simulado
# ═══════════════════════════════════════════════════════════════════

class TestProcessUrlShutdown(unittest.IsolatedAsyncioTestCase):
    """Tests de graceful shutdown — no requieren mocks de extracción."""

    async def asyncSetUp(self):
        self.page = FakePage()
        # Asegurar flag limpio antes de cada test
        import json_convert.browser as B
        B._shutdown_flag = False

    async def test_shutdown_flag_returns_immediately(self):
        """_shutdown_flag=True → error sin navegar."""
        import json_convert.browser as B
        B._shutdown_flag = True
        try:
            result = await process_url(self.page, 1, "https://ford.com")
            self.assertEqual(result["error"], "Shutdown requested")
            self.page.goto.assert_not_called()
        finally:
            B._shutdown_flag = False

    async def test_invalid_url_returns_error(self):
        """URL inválida → error de validación sin navegar."""
        result = await process_url(
            self.page, 1, "ftp://not-allowed.com"
        )
        self.assertIn("error", result)
        self.page.goto.assert_not_called()

    async def test_localhost_blocked(self):
        """localhost → bloqueado por validate_url."""
        result = await process_url(
            self.page, 1, "http://localhost:3000/test"
        )
        self.assertIn("error", result)
        self.page.goto.assert_not_called()


class TestProcessUrlSuccess(unittest.IsolatedAsyncioTestCase):
    """Tests del happy path — extracción exitosa."""

    async def asyncSetUp(self):
        self.page = FakePage()
        self.page.goto = AsyncMock(return_value=None)
        self.page.wait_for_function = AsyncMock(return_value=MagicMock())
        import json_convert.browser as B
        B._shutdown_flag = False

    @patch("json_convert.browser.extract_digital_data", return_value={"page": "home"})
    @patch("json_convert.browser.extract_title", return_value="Test Page Title")
    @patch("json_convert.browser.extract_s_object", return_value={"prop1": "val1"})
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_full_success(self, mock_cookie, mock_s, mock_title, mock_dd):
        """Happy path: goto + smart wait + gather extraen datos."""
        result = await process_url(
            self.page, 2,
            "https://www.ford.com/pr/driving-precautions.html",
        )

        self.assertEqual(result["url"], "https://www.ford.com/pr/driving-precautions.html")
        self.assertEqual(result["row"], 2)
        self.assertEqual(result["title"], "Test Page Title")
        self.assertEqual(result["digitaldata"], {"page": "home"})
        self.assertEqual(result["metadata"]["s_object"], {"prop1": "val1"})
        self.assertEqual(result["page_name"], "Driving Precautions")
        self.assertNotIn("error", result)
        self.assertEqual(result["retries_used"], 0)
        self.page.goto.assert_awaited_once()

    @patch("json_convert.browser.extract_digital_data", return_value={"page": "home"})
    @patch("json_convert.browser.extract_title", return_value="Test Page")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_s_object_none(self, mock_cookie, mock_s, mock_title, mock_dd):
        """s_object=None → metadata['s_object'] es None (no crash)."""
        result = await process_url(self.page, 3, "https://www.ford.com/pr/page.html")
        self.assertIsNone(result["metadata"]["s_object"])
        self.assertNotIn("error", result)

    @patch("json_convert.browser.extract_digital_data", return_value=None)
    @patch("json_convert.browser.extract_title", return_value="")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_no_digital_data(self, mock_cookie, mock_s, mock_title, mock_dd):
        """Sin digitalData → todo None, sin error de navegación."""
        result = await process_url(self.page, 4, "https://www.ford.com/pr/page.html")
        self.assertIsNone(result["digitaldata"])
        self.assertEqual(result["title"], "")
        # No hay error de navegación porque goto no falló
        self.assertNotIn("error", result)

    @patch("json_convert.browser.extract_digital_data", return_value=None)
    @patch("json_convert.browser.extract_title", return_value="")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_smart_wait_fallback(self, mock_cookie, mock_s, mock_title, mock_dd):
        """Smart wait timeout → fallback a wait_for_timeout(1000)."""
        self.page.wait_for_function.side_effect = Exception("timeout")
        result = await process_url(self.page, 5, "https://www.ford.com/pr/page.html")
        self.page.wait_for_timeout.assert_awaited_with(1000)
        self.assertNotIn("error", result)

    @patch("json_convert.browser.extract_digital_data", return_value={"page": "home"})
    @patch("json_convert.browser.extract_title", return_value="Popup Test")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_popup_detected_and_closed(self, mock_cookie, mock_s, mock_title, mock_dd):
        """Popup abierto → se cierra automáticamente."""
        popup_mock = AsyncMock()
        # Simular que page.on("popup", handler) llama al handler con popup_mock
        def _on_popup_side_effect(event, handler):
            if event == "popup":
                handler(popup_mock)
        self.page.on.side_effect = _on_popup_side_effect

        result = await process_url(self.page, 6, "https://www.ford.com/pr/page.html")
        popup_mock.close.assert_awaited_once()
        self.assertNotIn("error", result)


class TestProcessUrlErrors(unittest.IsolatedAsyncioTestCase):
    """Tests de errores de navegación y recuperación."""

    async def asyncSetUp(self):
        self.page = FakePage()
        import json_convert.browser as B
        B._shutdown_flag = False

    def _goto_err_aborted(self, url: str) -> AsyncMock:
        """Crea un mock goto que lanza ERR_ABORTED para URLs reales
        pero permite about:blank (usado en fetch+setContent fallback)."""
        async def _side_effect(url, **kwargs):
            if url == "about:blank":
                return None
            raise Exception("net::ERR_ABORTED at " + url)
        return AsyncMock(side_effect=_side_effect)

    @patch("json_convert.browser.extract_digital_data", return_value=None)
    @patch("json_convert.browser.extract_title", return_value="")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_timeout_with_retry(self, mock_cookie, mock_s, mock_title, mock_dd):
        """PwTimeout → retry, segunda vez OK."""
        self.page.goto.side_effect = [
            PwTimeout("Timeout"),
            None,  # segundo intento exitoso
        ]
        self.page.wait_for_function.return_value = MagicMock()
        # mock wait_for_timeout para el backoff entre retries
        self.page.wait_for_timeout = AsyncMock()

        result = await process_url(
            self.page, 1, "https://www.ford.com/pr/page.html",
            max_retry=1,
        )
        self.assertEqual(self.page.goto.await_count, 2)
        self.assertEqual(result["retries_used"], 1)
        self.assertNotIn("error", result)

    @patch("json_convert.browser.extract_digital_data", return_value=None)
    @patch("json_convert.browser.extract_title", return_value="")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_timeout_exhausted(self, mock_cookie, mock_s, mock_title, mock_dd):
        """PwTimeout en todos los intentos → error final."""
        self.page.goto.side_effect = PwTimeout("Timeout")
        self.page.wait_for_timeout = AsyncMock()

        result = await process_url(
            self.page, 1, "https://www.ford.com/pr/page.html",
            max_retry=1,
        )
        self.assertEqual(result["error"], "Timeout al navegar")
        self.assertEqual(self.page.goto.await_count, 2)

    async def test_cancelled_error_propagates(self):
        """CancelledError → se propaga (no se captura)."""
        self.page.goto.side_effect = asyncio.CancelledError()

        with self.assertRaises(asyncio.CancelledError):
            await process_url(
                self.page, 1, "https://www.ford.com/pr/page.html",
            )

    @patch("json_convert.browser.extract_digital_data", return_value={"page": "recovered"})
    @patch("json_convert.browser.extract_title", return_value="Recovered Page")
    async def test_err_aborted_partial_dom(self, mock_title, mock_dd):
        """ERR_ABORTED con digitalData en DOM parcial → recuperado."""
        self.page.goto.side_effect = Exception(
            "net::ERR_ABORTED at https://ford.com/all-vehicles"
        )

        result = await process_url(
            self.page, 1, "https://www.ford.com/pr/all-vehicles.html",
        )
        self.assertEqual(result["digitaldata"], {"page": "recovered"})
        self.assertEqual(result["title"], "Recovered Page")
        # Error limpiado porque se recuperó digitalData
        self.assertNotIn("error", result)

    @patch("json_convert.browser._fetch_html_via_http", return_value="<html><body>data</body></html>")
    async def test_err_aborted_fetch_fallback(self, mock_fetch):
        """ERR_ABORTED sin DD en DOM → fallback fetch+setContent recupera."""
        self.page.goto = self._goto_err_aborted("https://ford.com/all-vehicles")
        self.page.wait_for_timeout = AsyncMock()
        self.page.setContent = AsyncMock()

        import json_convert.browser as B
        # side_effect como lista: [call1, call2, call3, ...]
        # calls 1-2 (ERR handler en 2 intentos) → None
        # call 3 (post-fetch) → datos
        mock_dd = AsyncMock(side_effect=[None, None, {"page": "fetch_recovered"}])
        mock_title = AsyncMock(return_value="Fetch Recovered")

        orig_dd = B.extract_digital_data
        orig_title = B.extract_title
        B.extract_digital_data = mock_dd
        B.extract_title = mock_title

        try:
            result = await process_url(
                self.page, 1, "https://www.ford.com/pr/all-vehicles.html",
                max_retry=1,
            )
            self.page.setContent.assert_awaited_once()
            self.assertEqual(result["digitaldata"], {"page": "fetch_recovered"})
            self.assertEqual(result["title"], "Fetch Recovered")
            self.assertNotIn("error", result)
        finally:
            B.extract_digital_data = orig_dd
            B.extract_title = orig_title

    @patch("json_convert.browser._fetch_html_via_http", return_value="<html/>")
    async def test_err_aborted_fetch_fails(self, mock_fetch):
        """ERR_ABORTED + fetch falla → error persistente."""
        self.page.goto = self._goto_err_aborted("https://ford.com/all-vehicles")
        self.page.wait_for_timeout = AsyncMock()
        self.page.setContent.side_effect = Exception("setContent failed")

        import json_convert.browser as B
        orig_dd = B.extract_digital_data
        orig_title = B.extract_title
        B.extract_digital_data = AsyncMock(return_value=None)
        B.extract_title = AsyncMock(return_value="")

        try:
            result = await process_url(
                self.page, 1, "https://www.ford.com/pr/all-vehicles.html",
                max_retry=1,
            )
            self.assertIn("error", result)
            self.assertIn("ERR_ABORTED", result["error"])
        finally:
            B.extract_digital_data = orig_dd
            B.extract_title = orig_title

    @patch("json_convert.browser.extract_digital_data", return_value=None)
    @patch("json_convert.browser.extract_title", return_value="")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_generic_error_with_retry(self, mock_cookie, mock_s, mock_title, mock_dd):
        """Error genérico → retry, segundo intento OK."""
        self.page.goto.side_effect = [
            Exception("Some generic error"),
            None,  # segundo intento exitoso
        ]
        self.page.wait_for_function.return_value = MagicMock()
        self.page.wait_for_timeout = AsyncMock()

        result = await process_url(
            self.page, 1, "https://www.ford.com/pr/page.html",
            max_retry=1,
        )
        self.assertEqual(self.page.goto.await_count, 2)
        self.assertEqual(result["retries_used"], 1)
        self.assertNotIn("error", result)

    @patch("json_convert.browser.extract_digital_data", return_value=None)
    @patch("json_convert.browser.extract_title", return_value="")
    @patch("json_convert.browser.extract_s_object", return_value=None)
    @patch("json_convert.browser.try_dismiss_cookie_consent", return_value=False)
    async def test_page_closed_during_retry_wait(self, mock_cookie, mock_s, mock_title, mock_dd):
        """Page cerrada durante backoff → break sin crash."""
        self.page.goto.side_effect = PwTimeout("Timeout")
        self.page.wait_for_timeout.side_effect = Exception("page closed")

        result = await process_url(
            self.page, 1, "https://www.ford.com/pr/page.html",
            max_retry=1,
        )
        # El wait_for_timeout falló → break, retry no completo
        self.assertEqual(result["retries_used"], 0)  # no llegó a retry exitoso
        self.assertIn("error", result)


if __name__ == "__main__":
    unittest.main(verbosity=2)

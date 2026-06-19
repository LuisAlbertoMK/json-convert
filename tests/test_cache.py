"""
test_cache.py — Tests unitarios para json_convert/cache.py.

Requiere: solo stdlib (no requiere playwright ni navegador).
"""

import json
import os
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from json_convert.cache import UrlCache


class TestUrlCache(unittest.TestCase):
    """Tests para UrlCache — caché de resultados de navegación."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.cache = UrlCache(cache_dir=self.tmpdir, ttl=3600)

    def tearDown(self):
        self.cache.clear()
        if os.path.isdir(self.tmpdir):
            os.rmdir(self.tmpdir)

    # ── Set / Get básico ──

    def test_set_and_get(self):
        result = {"url": "https://ford.com", "digitaldata": {"page": {"pageName": "test"}}}
        self.cache.set("https://ford.com", result)
        cached = self.cache.get("https://ford.com")
        self.assertIsNotNone(cached)
        self.assertEqual(cached["url"], "https://ford.com")
        self.assertEqual(cached["digitaldata"]["page"]["pageName"], "test")

    def test_get_nonexistent(self):
        cached = self.cache.get("https://no-existe.com")
        self.assertIsNone(cached)

    def test_set_and_get_multiple_urls(self):
        urls = {
            "https://ford.com/a": {"url": "a", "data": 1},
            "https://ford.com/b": {"url": "b", "data": 2},
            "https://ford.com/c": {"url": "c", "data": 3},
        }
        for url, result in urls.items():
            self.cache.set(url, result)
        for url, expected in urls.items():
            cached = self.cache.get(url)
            self.assertIsNotNone(cached)
            self.assertEqual(cached["data"], expected["data"])

    def test_get_returns_copy_not_reference(self):
        result = {"url": "https://ford.com", "counter": 1}
        self.cache.set("https://ford.com", result)
        cached = self.cache.get("https://ford.com")
        self.assertIsNotNone(cached)
        # Modificar el resultado no debería afectar la caché
        cached["counter"] = 999
        cached2 = self.cache.get("https://ford.com")
        self.assertIsNotNone(cached2)
        self.assertEqual(cached2["counter"], 1)

    # ── TTL y expiración ──

    def test_expired_cache_returns_none(self):
        cache_short = UrlCache(cache_dir=self.tmpdir, ttl=0)  # TTL cero = expira al instante
        cache_short.set("https://ford.com", {"url": "ford"})
        # Con TTL=0, el get debería fallar por expirado
        cached = cache_short.get("https://ford.com")
        self.assertIsNone(cached)

    def test_cache_hits_within_ttl(self):
        cache_short = UrlCache(cache_dir=self.tmpdir, ttl=5)
        cache_short.set("https://ford.com", {"url": "ford"})
        # Dentro del TTL, debería devolver el resultado
        cached = cache_short.get("https://ford.com")
        self.assertIsNotNone(cached)
        self.assertEqual(cached["url"], "ford")

    def test_cache_ttl_default(self):
        """Verifica que el TTL por defecto sea 24h."""
        default_cache = UrlCache(cache_dir=self.tmpdir)
        self.assertEqual(default_cache.ttl, 86400)

    # ── Clear ──

    def test_clear_all(self):
        for i in range(5):
            self.cache.set(f"https://ford.com/{i}", {"idx": i})
        cleared = self.cache.clear()
        self.assertEqual(cleared, 5)
        # Verificar que todos expiraron
        for i in range(5):
            self.assertIsNone(self.cache.get(f"https://ford.com/{i}"))

    def test_clear_empty_returns_zero(self):
        self.cache.clear()  # limpia primero
        cleared = self.cache.clear()
        self.assertEqual(cleared, 0)

    def test_clear_single_url(self):
        self.cache.set("https://ford.com/keep", {"keep": True})
        self.cache.set("https://ford.com/remove", {"remove": True})
        removed = self.cache.clear_url("https://ford.com/remove")
        self.assertTrue(removed)
        # La removida no debe estar
        self.assertIsNone(self.cache.get("https://ford.com/remove"))
        # La otra debe seguir
        self.assertIsNotNone(self.cache.get("https://ford.com/keep"))

    def test_clear_nonexistent_url(self):
        result = self.cache.clear_url("https://no-existe.com")
        self.assertFalse(result)

    # ── Persistencia en disco ──

    def test_cache_persists_to_disk(self):
        self.cache.set("https://ford.com", {"msg": "hello"})
        cache_file = os.path.join(self.tmpdir, os.listdir(self.tmpdir)[0])
        self.assertTrue(os.path.exists(cache_file))
        with open(cache_file, encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data["_url"], "https://ford.com")
        self.assertIn("_cached_at", data)
        self.assertEqual(data["result"]["msg"], "hello")

    def test_cache_key_is_md5(self):
        """La clave de caché debe ser un hash MD5 de 32 caracteres."""
        from json_convert.cache import UrlCache
        key = UrlCache._key("https://ford.com")
        self.assertEqual(len(key), 32)
        self.assertTrue(all(c in "0123456789abcdef" for c in key))

    # ── Tolerancia a errores ──

    def test_corrupted_cache_returns_none(self):
        self.cache.set("https://ford.com", {"ok": True})
        # Corromper el archivo
        cache_file = os.path.join(self.tmpdir, os.listdir(self.tmpdir)[0])
        with open(cache_file, "w", encoding="utf-8") as f:
            f.write("not valid json{{{")
        cached = self.cache.get("https://ford.com")
        self.assertIsNone(cached)

    def test_empty_cache_dir_clear_returns_zero(self):
        """clear() en directorio vacío debe retornar 0."""
        empty_dir = tempfile.mkdtemp()
        cache = UrlCache(cache_dir=empty_dir)
        cleared = cache.clear()
        self.assertEqual(cleared, 0)
        os.rmdir(empty_dir)

    def test_no_cache_dir_yet(self):
        """get() en directorio que no existe aún debe retornar None."""
        fake_dir = os.path.join(self.tmpdir, "noexiste")
        cache = UrlCache(cache_dir=fake_dir)
        self.assertIsNone(cache.get("https://ford.com"))


if __name__ == "__main__":
    unittest.main()

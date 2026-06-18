"""
json_convert/cache.py — Caché de resultados de navegación por URL.

Reduce el tiempo en corridas consecutivas al reutilizar resultados
previos de URLs que ya fueron auditadas. Útil durante debugging
o cuando se itera sobre el mismo set de URLs.

Uso:
    cache = UrlCache()
    cached = cache.get("https://...")
    if cached:
        return cached
    result = await process_url(...)
    cache.set("https://...", result)
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import time

CACHE_DIR = ".url_cache"
DEFAULT_TTL = 86400  # 24 horas


class UrlCache:
    """Caché de resultados URL→dict, persistida en archivos JSON.

    Cada URL se guarda como .url_cache/{md5}.json con timestamp.
    """

    def __init__(self, cache_dir: str = CACHE_DIR, ttl: int = DEFAULT_TTL) -> None:
        self.cache_dir = cache_dir
        self.ttl = ttl

    @staticmethod
    def _key(url: str) -> str:
        """Retorna hash MD5 de la URL (nombre de archivo)."""
        return hashlib.md5(url.encode("utf-8")).hexdigest()

    def _path(self, url: str) -> str:
        return os.path.join(self.cache_dir, f"{self._key(url)}.json")

    def get(self, url: str) -> dict | None:
        """Retorna resultado cacheado o None si no existe/expirado."""
        path = self._path(url)
        if not os.path.exists(path):
            return None
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logging.debug("Cache corrupto para %s: %s", url, e)
            return None

        cached_at = data.get("_cached_at", 0)
        if time.time() - cached_at > self.ttl:
            logging.debug("Cache expirado para %s", url)
            return None

        return data.get("result")

    def set(self, url: str, result: dict) -> None:
        """Guarda resultado en caché."""
        os.makedirs(self.cache_dir, exist_ok=True)
        path = self._path(url)
        to_save = {
            "_cached_at": time.time(),
            "_url": url,
            "result": result,
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(to_save, f, ensure_ascii=False)
        except OSError as e:
            logging.warning("No se pudo cachear %s: %s", url, e)

    def clear(self) -> int:
        """Elimina todos los archivos de caché. Retorna cantidad eliminada."""
        if not os.path.isdir(self.cache_dir):
            return 0
        count = 0
        for fname in os.listdir(self.cache_dir):
            if fname.endswith(".json"):
                try:
                    os.remove(os.path.join(self.cache_dir, fname))
                    count += 1
                except OSError:
                    pass
        return count

    def clear_url(self, url: str) -> bool:
        """Elimina caché de una URL específica."""
        path = self._path(url)
        if os.path.exists(path):
            try:
                os.remove(path)
                return True
            except OSError:
                pass
        return False

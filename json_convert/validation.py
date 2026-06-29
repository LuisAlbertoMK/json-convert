"""
json_convert/validation.py — Validación y sanitización de URLs.

Constantes y funciones para validar URLs antes de navegación,
con chequeo SSRF básico y sanitización para logging.
"""

from urllib.parse import parse_qs, urlparse

# Schemes permitidos para navegación
VALID_URL_SCHEMES = ("http", "https")

# Dominios conocidos del proyecto Ford preview + Adobe
ALLOWED_HOSTNAME_SUFFIXES = (
    ".ford.com",
    ".brandpr.ford.com",
    ".ford.mx",
    ".ford.com.pr",
    ".lincoln.mx",
    ".lincoln.com",
    ".omtrdc.net",
    ".adobedc.net",
    "2o7.net",
)


def validate_url(url: str) -> str | None:
    """Valida URL antes de navegar. Retorna mensaje de error o None si OK."""
    if not url or not isinstance(url, str):
        return "URL vacía o inválida"
    url = url.strip()
    if not url:
        return "URL vacía después de trim"
    try:
        parsed = urlparse(url)
    except Exception as e:
        return f"URL no parseable: {e}"
    if parsed.scheme not in VALID_URL_SCHEMES:
        return f"Scheme '{parsed.scheme}' no permitido (solo http/https)"
    if not parsed.netloc:
        return "URL sin hostname"
    # Chequeo de SSRF básico: solo dominios conocidos
    hostname = parsed.netloc.lower()
    # Remover user:password@ si existe
    if "@" in hostname:
        hostname = hostname.split("@")[-1]
    # Remover :puerto si existe
    if ":" in hostname:
        hostname = hostname.split(":")[0]
    if hostname in ("localhost", "127.0.0.1", "0.0.0.0", "::1"):
        return "URL apunta a localhost (posible SSRF)"
    if not hostname.endswith(ALLOWED_HOSTNAME_SUFFIXES):
        return f"Dominio '{hostname}' no está en la whitelist de proyectos"
    return None


def sanitize_url_for_log(url: str, max_len: int = 80) -> str:
    """Limpia URL para logging: trunca y redacta query params sensibles.

    Redacta valores de query params que podrían contener PII (email, token, etc).
    """
    if not url:
        return ""
    try:
        parsed = urlparse(url)
        if parsed.query:
            qs = parse_qs(parsed.query)
            sensitive_keys = {"email", "token", "key", "secret", "password", "pass", "auth"}
            cleaned = {k: ("[REDACTED]" if k.lower() in sensitive_keys else v[0][:60])
                       for k, v in qs.items()}
            base = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"[:max_len]
            return f"{base}?{cleaned}" if cleaned else base
        return url[:max_len]
    except Exception:
        return url[:max_len]

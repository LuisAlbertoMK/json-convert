"""
json_convert/metrics.py — Códigos de error, clasificación y scoring.

Funciones para calcular score de auditoría y clasificar errores
de Adobe Analytics de forma estandarizada.
"""

# Códigos de error estándar para toda la aplicación
ERROR_CODES = {
    "TIMEOUT":       "Tiempo de espera agotado al navegar",
    "HTTP_403":      "Acceso denegado (HTTP 403)",
    "HTTP_ERROR":    "Error HTTP al navegar",
    "NO_AA_DATA":    "No se capturó data de Adobe Analytics",
    "URL_INVALID":   "URL no válida o no permitida",
    "NETWORK_ERROR": "Error de red o conexión fallida",
    "NAV_ERROR":     "Error durante la navegación",
    "UNKNOWN":       "Error desconocido",
}


def _error_code_from_detail(err: str) -> str:
    """Asigna código de error estándar según el texto del error."""
    if not err:
        return "UNKNOWN"
    err_lower = err.lower()
    if "timeout" in err_lower:
        return "TIMEOUT"
    if "403" in err:
        return "HTTP_403"
    if "no aa" in err_lower or "no AA" in err:
        return "NO_AA_DATA"
    if "url_invalid" in err_lower or "url inv" in err_lower:
        return "URL_INVALID"
    if "network" in err_lower or "connection" in err_lower or "dns" in err_lower:
        return "NETWORK_ERROR"
    if "naveg" in err_lower or "nav" in err_lower:
        return "NAV_ERROR"
    return "NETWORK_ERROR"


def classify_errors(errors_detail: list[dict]) -> dict[str, list[int]]:
    """Agrupa errores por categoría para output legible."""
    categories: dict[str, list[int]] = {
        "HTTP 403 (acceso denegado)": [],
        "Timeout": [],
        "Sin dato AA (no beacon)": [],
        "Error de red/conexión": [],
    }
    for e in errors_detail:
        code = e.get("code", _error_code_from_detail(e.get("error", "")))
        err = e.get("error", "")
        if code == "TIMEOUT":
            categories["Timeout"].append(e["row"])
        elif code == "HTTP_403" or "403" in err:
            categories["HTTP 403 (acceso denegado)"].append(e["row"])
        elif code in ("NO_AA_DATA",):
            categories["Sin dato AA (no beacon)"].append(e["row"])
        elif code == "URL_INVALID":
            pass
        else:
            categories["Error de red/conexión"].append(e["row"])
    return {k: v for k, v in categories.items() if v}


# ═══════════════════════════════════════════════════════════════════════════
# SCORE
# ═══════════════════════════════════════════════════════════════════════════

def compute_score(metrics: dict) -> int:
    """Score global 0-100 de la corrida."""
    success_rate = (metrics["ok_aa"] / max(metrics["total"], 1)) * 100
    dd_rate = (metrics["ok_dd"] / max(metrics["total"], 1)) * 100
    avg_time = sum(metrics["times"]) / max(len(metrics["times"]), 1)
    beacons_per_url = metrics["total_beacons"] / max(metrics["total"], 1)
    retry_efficiency = max(0, 1 - metrics["retries"] / max(metrics["total"], 1)) * 100
    time_score = (1 - min(avg_time, 60) / 60) * 100
    return int(
        success_rate * 0.40 +
        dd_rate * 0.25 +
        max(time_score, 0) * 0.15 +
        min(beacons_per_url, 3) / 3 * 100 * 0.10 +
        retry_efficiency * 0.10
    )


def compute_url_score(result: dict) -> int:
    """Score 0-100 por URL individual."""
    s = 0
    if result.get("digitaldata"):
        s += 30
    if result.get("aa_parsed"):
        s += 30
    if result.get("extra_beacons"):
        s += 10
    if not result.get("error") and result.get("status", 0) not in (-1, -2, 403):
        s += 20
    elapsed = result.get("elapsed_s", 99)
    if elapsed < 5:
        s += 10
    elif elapsed < 15:
        s += 5
    return s

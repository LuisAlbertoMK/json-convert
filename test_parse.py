"""
test_parse.py — Tests unitarios para parse_aa_beacon(), build_aa_from_s()
y extract_fields().

Uso:
  python test_parse.py                       # verbose (default)
  python test_parse.py -v                    # aún más detalle
  python test_parse.py TestParseBeacon       # solo una clase
  python test_parse.py test_beacon_basic     # solo un test
  python test_parse.py TestExtractFields     # tests de extract_aa.py

Requiere: unittest (stdlib, 0 deps externas)
No requiere: playwright, openpyxl, red, navegador.
"""

import json
import os
import sys
import unittest

# Importar funciones del script principal
sys.path.insert(0, os.path.dirname(__file__))
from extract_aa import extract_fields
from extract_browser import (
    _error_code_from_detail,
    build_aa_from_s,
    classify_errors,
    compute_score,
    compute_url_score,
    parse_aa_beacon,
    sanitize_url_for_log,
    validate_url,
)

# ═══════════════════════════════════════════════════════════════════════════
# BEACONS DE EJEMPLO (formato real de Adobe Analytics)
# ═══════════════════════════════════════════════════════════════════════════

# Beacon estándar s.t() — page view con props, eVars, events
BEACON_PAGEVIEW = (
    "https://smetrics.ford.com/b/ss/fordglobal/1/JS-2.25.0/s431234567890?"
    "g=https%3A%2F%2Fpreview.ford.com%2Fes%2Fvehiculos%2Fmach-e"
    "&pageName=ford%3Amach-e%3Apreview"
    "&c1=home"
    "&c2=vehiculos"
    "&v1=%7B%22id%22%3A%22mach-e%22%7D"
    "&v5=preview"
    "&events=event1%2Cevent2"
    "&ch=automotriz"
    "&res=1920x1080"
    "&bw=1920&bh=1080"
    "&cd=24-bit&ce=UTF-8"
    "&mid=1234567890"
    "&aamlh=12345"
    "&products=cars%3Bmach-e"
    "&t=1718000000000"
)

# Beacon s.tl() — custom link (sin products, sin page URL en g)
BEACON_CUSTOM_LINK = (
    "https://smetrics.ford.com/b/ss/fordglobal/1/JS-2.25.0/s987654321?"
    "pe=lnk_o&pev2=ford%3Aclick%3Acta%3Aver-mas"
    "&pageName=ford%3Amach-e%3Apreview"
    "&c1=interaccion"
    "&v1=cta-ver-mas"
    "&events=event3"
    "&ch=automotriz"
    "&res=1920x1080"
    "&mid=9876543210"
)

# Beacon mínimo — solo pageName + evento
BEACON_MINIMAL = (
    "https://sc.omtrdc.net/b/ss/fordglobal/1/JS-2.25.0/s0?"
    "pageName=test%3Aminimal&events=event1"
)

# Beacon con eVars sin props
BEACON_EVARS_ONLY = (
    "https://smetrics.omtrdc.net/b/ss/fordglobal/1/JS-2.25.0/s111?"
    "g=https%3A%2F%2Fford.com"
    "&pageName=ford%3Ahome"
    "&v1=home"
    "&v10=usuario-nuevo"
    "&v15=canal-organic"
    "&events=event1"
)

# Beacon con props sin eVars
BEACON_PROPS_ONLY = (
    "https://smetrics.ford.com/b/ss/fordglobal/1/JS-2.25.0/s222?"
    "g=https%3A%2F%2Fford.com%2Fes"
    "&pageName=ford%3Aes%3Ahome"
    "&c1=home-es"
    "&c5=navegacion"
    "&events=event1"
)

# URL sin query string (borde)
BEACON_NO_QUERY = (
    "https://smetrics.ford.com/b/ss/fordglobal/1/JS-2.25.0/s333?"
)

# URL con hostname atípico pero válido
BEACON_ALT_DOMAIN = (
    "https://data.adobedc.net/b/ss/fordglobal/1/JS-2.25.0/s444?"
    "pageName=test%3Aalt-domain&events=event4"
)


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: parse_aa_beacon
# ═══════════════════════════════════════════════════════════════════════════

class TestParseBeacon(unittest.TestCase):
    """Tests para parse_aa_beacon(beacon_url, page_title)."""

    def test_beacon_basic(self):
        """Beacon pageview completo: verifica estructura y valores clave."""
        result = parse_aa_beacon(BEACON_PAGEVIEW, "Ford Mach-E Preview")
        self.assertEqual(result["solution"], "analytics")
        self.assertEqual(result["page"]["title"], "Ford Mach-E Preview")
        self.assertIn("ford.com", result["page"]["url"])
        self.assertEqual(result["request"]["method"], "GET")
        self.assertEqual(result["request"]["hostname"], "smetrics.ford.com")
        self.assertIn("hit", result)
        self.assertEqual(result["hit"]["reportSuiteId"], "fordglobal")
        self.assertEqual(result["hit"]["type"], "pageView")

    def test_beacon_props(self):
        """Props se parsean como prop1, prop2, etc."""
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertIn("prop1", result["props"])
        self.assertEqual(result["props"]["prop1"], "home")
        self.assertIn("prop2", result["props"])
        self.assertEqual(result["props"]["prop2"], "vehiculos")

    def test_beacon_evars(self):
        """eVars se parsean como eVar1, eVar2, etc."""
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertIn("eVar1", result["eVars"])
        self.assertEqual(result["eVars"]["eVar1"], '{"id":"mach-e"}')
        self.assertIn("eVar5", result["eVars"])
        self.assertEqual(result["eVars"]["eVar5"], "preview")

    def test_beacon_events(self):
        """Events se parsean como lista."""
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertEqual(result["events"], ["event1", "event2"])

    def test_beacon_visitor(self):
        """Visitor fields: mid y aamlh."""
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertEqual(result["visitor"]["experienceCloudId"], "1234567890")
        self.assertEqual(result["visitor"]["audienceManagerHint"], "12345")

    def test_beacon_browser(self):
        """Browser fields: resolution, width, height, colorDepth, charset."""
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertEqual(result["browser"]["resolution"], "1920x1080")
        self.assertEqual(result["browser"]["browserWidth"], 1920)
        self.assertEqual(result["browser"]["browserHeight"], 1080)
        self.assertEqual(result["browser"]["colorDepth"], "24-bit")
        self.assertEqual(result["browser"]["charset"], "UTF-8")

    def test_beacon_channel(self):
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertEqual(result["channel"], "automotriz")

    def test_beacon_products(self):
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertEqual(result["products"], "cars;mach-e")

    def test_beacon_timestamp(self):
        result = parse_aa_beacon(BEACON_PAGEVIEW)
        self.assertEqual(result["request"]["collectedTimestamp"], "1718000000000")

    def test_beacon_custom_link(self):
        """s.tl() beacon: sin products, sin pageURL."""
        result = parse_aa_beacon(BEACON_CUSTOM_LINK)
        self.assertEqual(result["hit"]["reportSuiteId"], "fordglobal")
        self.assertEqual(result["events"], ["event3"])
        # s.tl. no incluye g= normalmente
        self.assertEqual(result["page"]["url"], "")
        self.assertEqual(result["channel"], "automotriz")

    def test_beacon_minimal(self):
        """Beacon mínimo: solo pageName + evento, sin props/evars/visitor."""
        result = parse_aa_beacon(BEACON_MINIMAL)
        self.assertEqual(result["events"], ["event1"])
        self.assertEqual(result["pageName"], "test:minimal")
        self.assertEqual(result["props"], {})
        self.assertEqual(result["eVars"], {})
        self.assertEqual(result["visitor"], {})
        self.assertEqual(result["browser"], {})

    def test_beacon_evars_only(self):
        """Solo eVars, sin props."""
        result = parse_aa_beacon(BEACON_EVARS_ONLY)
        self.assertIn("eVar1", result["eVars"])
        self.assertIn("eVar15", result["eVars"])
        self.assertEqual(result["props"], {})

    def test_beacon_props_only(self):
        """Solo props, sin eVars."""
        result = parse_aa_beacon(BEACON_PROPS_ONLY)
        self.assertIn("prop1", result["props"])
        self.assertEqual(result["eVars"], {})

    def test_beacon_no_query(self):
        """URL sin query string: no debe fallar."""
        result = parse_aa_beacon(BEACON_NO_QUERY)
        self.assertEqual(result["hit"]["reportSuiteId"], "fordglobal")
        self.assertEqual(result["events"], [])
        self.assertEqual(result["props"], {})
        self.assertEqual(result["eVars"], {})

    def test_alt_domain(self):
        """Dominio Adobe alternativo (data.adobedc.net)."""
        result = parse_aa_beacon(BEACON_ALT_DOMAIN)
        self.assertEqual(result["request"]["hostname"], "data.adobedc.net")
        self.assertEqual(result["events"], ["event4"])

    def test_page_title_empty(self):
        """page_title vacío no debe romper."""
        result = parse_aa_beacon(BEACON_MINIMAL, "")
        self.assertEqual(result["page"]["title"], "")

    def test_page_title_none(self):
        """page_title None no debe romper."""
        result = parse_aa_beacon(BEACON_MINIMAL, None)
        self.assertEqual(result["page"]["title"], "")

    def test_output_json_serializable(self):
        """El resultado debe ser serializable a JSON sin errores."""
        result = parse_aa_beacon(BEACON_PAGEVIEW, "Test")
        # json.dumps no debe lanzar
        dumped = json.dumps(result, ensure_ascii=False)
        self.assertGreater(len(dumped), 50)  # debe tener contenido


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: build_aa_from_s
# ═══════════════════════════════════════════════════════════════════════════

class TestBuildFromS(unittest.TestCase):
    """Tests para build_aa_from_s(s_obj, page_title)."""

    def test_basic(self):
        s = {
            "pageName": "ford:mach-e:preview",
            "pageURL": "https://preview.ford.com/es/mach-e",
            "channel": "automotriz",
            "prop1": "home",
            "prop2": "vehiculos",
            "eVar1": '{"id":"mach-e"}',
            "events": "event1,event2",
            "products": "cars;mach-e",
        }
        result = build_aa_from_s(s, "Ford Mach-E")
        self.assertEqual(result["solution"], "analytics")
        self.assertEqual(result["pageName"], "ford:mach-e:preview")
        self.assertEqual(result["page"]["title"], "Ford Mach-E")
        self.assertIn("prop1", result["props"])
        self.assertIn("eVar1", result["eVars"])
        self.assertEqual(result["events"], ["event1", "event2"])
        self.assertEqual(result["channel"], "automotriz")
        self.assertEqual(result["request"]["source"], "window.s")

    def test_empty_s(self):
        """s object vacío: no debe romper."""
        result = build_aa_from_s({}, "")
        self.assertEqual(result["pageName"], "")
        self.assertEqual(result["events"], [])
        self.assertEqual(result["props"], {})
        self.assertEqual(result["eVars"], {})

    def test_case_insensitive_props(self):
        """prop1, PROP1, Prop1 deben ser capturados (regex IGNORECASE)."""
        s = {"prop1": "value1", "eVar2": "value2"}
        result = build_aa_from_s(s)
        self.assertEqual(result["props"]["prop1"], "value1")
        self.assertEqual(result["eVars"]["eVar2"], "value2")

    def test_extra_fields_ignored(self):
        """Campos no-AA en window.s no deben contaminar props/evars."""
        s = {
            "pageName": "test",
            "customVar": "should_not_appear",
            "somePlugin": "data",
        }
        result = build_aa_from_s(s)
        self.assertEqual(result["pageName"], "test")
        self.assertNotIn("customVar", result)
        self.assertNotIn("somePlugin", result)


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: extract_fields (de extract_aa.py)
# ═══════════════════════════════════════════════════════════════════════════

SAMPLE_BEACON_JSON = {
    "solution": "analytics",
    "page": {"title": "Ford Mach-E", "url": "https://ford.com/es/mach-e"},
    "request": {"method": "GET", "hostname": "smetrics.ford.com"},
    "visitor": {"experienceCloudId": "abc123"},
    "hit": {"id": "s123", "type": "pageView", "reportSuiteId": "fordglobal"},
    "events": ["event1", "event2"],
    "eVars": {"eVar1": '{"id":"mach-e"}', "eVar5": "preview"},
    "props": {"prop1": "home", "prop2": "vehiculos"},
    "pageName": "ford:mach-e:preview",
    "channel": "automotriz",
}

SAMPLE_BEACON_JSON_GROUP2 = {
    "solution": "analytics",
    "evars": {"v1": "home", "v10": "nuevo"},
    "props": {"c1": "home-es", "c5": "navegacion"},
    "pageName": "ford:es:home",
}

SAMPLE_BEACON_JSON_MINIMAL = {
    "solution": "analytics",
    "pageName": "test:minimal",
    "events": ["event1"],
}


class TestExtractFields(unittest.TestCase):
    """Tests para extract_fields(data, keep) de extract_aa.py."""

    def test_default_fields(self):
        """Campos default: page, request, props, evars."""
        result = extract_fields(SAMPLE_BEACON_JSON, ["page", "request", "props", "evars"])
        self.assertIn("page", result)
        self.assertIn("request", result)
        self.assertIn("props", result)
        self.assertIn("evars", result)
        self.assertEqual(result["props"], {"prop1": "home", "prop2": "vehiculos"})
        self.assertEqual(result["evars"], {"eVar1": '{"id":"mach-e"}', "eVar5": "preview"})

    def test_evars_unification(self):
        """'evars' unifica eVars (Grupo 1) y evars (Grupo 2)."""
        r1 = extract_fields(SAMPLE_BEACON_JSON, ["evars"])
        self.assertEqual(r1["evars"]["eVar1"], '{"id":"mach-e"}')
        r2 = extract_fields(SAMPLE_BEACON_JSON_GROUP2, ["evars"])
        self.assertEqual(r2["evars"]["v1"], "home")

    def test_all_fields(self):
        """'all' extrae todo lo disponible."""
        result = extract_fields(SAMPLE_BEACON_JSON, [
            "solution", "page", "request", "visitor", "hit",
            "events", "eVars", "props", "pageName", "channel",
        ])
        self.assertEqual(len(result), 10)
        self.assertEqual(result["solution"], "analytics")
        self.assertEqual(result["events"], ["event1", "event2"])
        self.assertEqual(result["channel"], "automotriz")

    def test_extrae_solo_lo_pedido(self):
        """No debe incluir campos no solicitados."""
        result = extract_fields(SAMPLE_BEACON_JSON, ["pageName"])
        self.assertEqual(result, {"pageName": "ford:mach-e:preview"})
        self.assertNotIn("page", result)
        self.assertNotIn("eVars", result)

    def test_campo_inexistente(self):
        """Campo pedido que no existe en data → no aparece en resultado."""
        result = extract_fields(SAMPLE_BEACON_JSON, ["products"])
        self.assertEqual(result, {})

    def test_data_vacio(self):
        """data={} → resultado vacío."""
        result = extract_fields({}, ["page", "props"])
        self.assertEqual(result, {})

    def test_keep_vacio(self):
        """keep=[] → resultado vacío."""
        result = extract_fields(SAMPLE_BEACON_JSON, [])
        self.assertEqual(result, {})

    def test_grupo2_evars(self):
        """Grupo 2 con keys 'evars' (minúscula)."""
        result = extract_fields(SAMPLE_BEACON_JSON_GROUP2, ["evars", "props"])
        self.assertEqual(result["evars"]["v1"], "home")
        self.assertEqual(result["props"]["c1"], "home-es")

    def test_datos_minimos(self):
        """JSON mínimo sin props/evars."""
        result = extract_fields(SAMPLE_BEACON_JSON_MINIMAL, ["pageName", "events"])
        self.assertEqual(result["pageName"], "test:minimal")
        self.assertEqual(result["events"], ["event1"])
        self.assertNotIn("props", result)
        self.assertNotIn("eVars", result)


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: validate_url
# ═══════════════════════════════════════════════════════════════════════════

class TestValidateUrl(unittest.TestCase):
    """Tests para validate_url()."""

    def test_valid_https_ford(self):
        """URL HTTPS a ford.com es válida."""
        self.assertIsNone(validate_url("https://www.ford.com/es/mustang"))

    def test_valid_https_preview(self):
        """URL HTTPS a preview.ford.com es válida."""
        self.assertIsNone(validate_url("https://brandpr.ford.com/preview/mach-e"))

    def test_empty_url(self):
        """URL vacía retorna error."""
        self.assertIsNotNone(validate_url(""))
        self.assertIsNotNone(validate_url(None))  # type: ignore

    def test_invalid_scheme_ftp(self):
        """Scheme FTP no permitido."""
        err = validate_url("ftp://ford.com/file.txt")
        self.assertIsNotNone(err)
        self.assertIn("Scheme", err or "")

    def test_no_hostname(self):
        """URL sin hostname retorna error."""
        err = validate_url("http:///path")
        self.assertIsNotNone(err)
        self.assertIn("hostname", err or "")

    def test_localhost_blocked(self):
        """localhost bloqueado por SSRF."""
        err = validate_url("http://localhost:8080/audit")
        self.assertIsNotNone(err)
        self.assertIn("SSRF", err or "")

    def test_loopback_blocked(self):
        """127.0.0.1 bloqueado."""
        err = validate_url("http://127.0.0.1/secret")
        self.assertIsNotNone(err)
        self.assertIn("SSRF", err or "")

    def test_domain_not_in_whitelist(self):
        """Dominio externo no permitido."""
        err = validate_url("https://evil.com/phish")
        self.assertIsNotNone(err)
        self.assertIn("whitelist", err or "")

    def test_url_with_credentials(self):
        """URL con user:password se parsea ok."""
        err = validate_url("https://user:pass@brandpr.ford.com/preview")
        self.assertIsNone(err)

    def test_url_with_port(self):
        """URL con puerto se parsea ok."""
        err = validate_url("https://brandpr.ford.com:8443/preview")
        self.assertIsNone(err)


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: sanitize_url_for_log
# ═══════════════════════════════════════════════════════════════════════════

class TestSanitizeUrl(unittest.TestCase):
    """Tests para sanitize_url_for_log()."""

    def test_empty_url(self):
        """URL vacía retorna string vacío."""
        self.assertEqual(sanitize_url_for_log(""), "")

    def test_no_sensitive_params(self):
        """URL sin params sensibles pasa limpia (truncada)."""
        url = "https://brandpr.ford.com/preview/mach-e"
        result = sanitize_url_for_log(url, max_len=100)
        self.assertIn("ford.com", result)
        self.assertNotIn("[REDACTED]", result)

    def test_email_redacted(self):
        """Query param 'email' se redacta."""
        url = "https://brandpr.ford.com/preview?email=john.doe@ford.com&token=abc123&page=1"
        result = sanitize_url_for_log(url)
        self.assertIn("[REDACTED]", result)
        self.assertNotIn("john.doe", result)
        self.assertNotIn("abc123", result)

    def test_password_redacted(self):
        """Query param 'password' se redacta."""
        url = "https://brandpr.ford.com/login?password=supersecret&user=admin"
        result = sanitize_url_for_log(url, max_len=120)
        self.assertIn("[REDACTED]", result)

    def test_max_len_truncation(self):
        """URL se trunca a max_len."""
        url = "https://brandpr.ford.com/" + "a" * 200
        result = sanitize_url_for_log(url, max_len=50)
        self.assertLessEqual(len(result), 50)

    def test_none_url(self):
        """None retorna string vacío."""
        self.assertEqual(sanitize_url_for_log(None), "")  # type: ignore

    def test_invalid_url_no_crash(self):
        """URL malformada no causa excepción."""
        result = sanitize_url_for_log("not a url at all!!!")
        self.assertIsInstance(result, str)
        self.assertTrue(len(result) > 0)


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: _error_code_from_detail
# ═══════════════════════════════════════════════════════════════════════════

class TestErrorCodeFromDetail(unittest.TestCase):
    """Tests para _error_code_from_detail()."""

    def test_timeout(self):
        """Texto con timeout → TIMEOUT."""
        self.assertEqual(_error_code_from_detail("Timeout after 30000ms"), "TIMEOUT")
        self.assertEqual(_error_code_from_detail("timeout navegando"), "TIMEOUT")

    def test_http_403(self):
        """Texto con 403 → HTTP_403."""
        self.assertEqual(_error_code_from_detail("HTTP 403 Forbidden"), "HTTP_403")
        self.assertEqual(_error_code_from_detail("Error 403"), "HTTP_403")

    def test_no_aa(self):
        """Texto con 'no AA' → NO_AA_DATA."""
        self.assertEqual(_error_code_from_detail("no AA data found"), "NO_AA_DATA")
        self.assertEqual(_error_code_from_detail("No AA beacon detected"), "NO_AA_DATA")

    def test_url_invalid(self):
        """Texto con URL invalid → URL_INVALID."""
        self.assertEqual(_error_code_from_detail("url_invalid: bad host"), "URL_INVALID")
        self.assertEqual(_error_code_from_detail("URL inválida"), "URL_INVALID")

    def test_network_error(self):
        """Texto con network/connection → NETWORK_ERROR."""
        self.assertEqual(_error_code_from_detail("Network error"), "NETWORK_ERROR")
        self.assertEqual(_error_code_from_detail("connection refused"), "NETWORK_ERROR")
        self.assertEqual(_error_code_from_detail("DNS resolution failed"), "NETWORK_ERROR")

    def test_nav_error(self):
        """Texto con naveg/nav → NAV_ERROR."""
        self.assertEqual(_error_code_from_detail("navegación fallida"), "NAV_ERROR")
        self.assertEqual(_error_code_from_detail("Navigation error"), "NAV_ERROR")

    def test_empty_string(self):
        """String vacío → UNKNOWN."""
        self.assertEqual(_error_code_from_detail(""), "UNKNOWN")

    def test_none_string(self):
        """None → UNKNOWN."""
        self.assertEqual(_error_code_from_detail(None), "UNKNOWN")  # type: ignore

    def test_unknown_error(self):
        """Texto genérico → NETWORK_ERROR (fallback)."""
        self.assertEqual(_error_code_from_detail("something weird happened"), "NETWORK_ERROR")


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: classify_errors
# ═══════════════════════════════════════════════════════════════════════════

class TestClassifyErrors(unittest.TestCase):
    """Tests para classify_errors()."""

    def test_empty_list(self):
        """Lista vacía → dict vacío."""
        self.assertEqual(classify_errors([]), {})

    def test_timeout_classified(self):
        """Error TIMEOUT va a categoría Timeout."""
        result = classify_errors([{"row": 3, "code": "TIMEOUT"}])
        self.assertIn("Timeout", result)
        self.assertEqual(result["Timeout"], [3])

    def test_http403_classified(self):
        """Error HTTP_403 va a categoría HTTP 403."""
        result = classify_errors([{"row": 5, "code": "HTTP_403", "error": "403"}])
        self.assertIn("HTTP 403 (acceso denegado)", result)
        self.assertEqual(result["HTTP 403 (acceso denegado)"], [5])

    def test_no_aa_classified(self):
        """Error NO_AA_DATA va a categoría Sin dato AA."""
        result = classify_errors([{"row": 7, "code": "NO_AA_DATA"}])
        self.assertIn("Sin dato AA (no beacon)", result)
        self.assertEqual(result["Sin dato AA (no beacon)"], [7])

    def test_url_invalid_skipped(self):
        """URL_INVALID se omite (ya reportado en validación)."""
        result = classify_errors([{"row": 9, "code": "URL_INVALID"}])
        self.assertNotIn("URL_INVALID", result)

    def test_mixed_errors(self):
        """Múltiples errores se agrupan correctamente."""
        errors = [
            {"row": 2, "code": "TIMEOUT"},
            {"row": 3, "code": "HTTP_403", "error": "403"},
            {"row": 5, "code": "NO_AA_DATA"},
            {"row": 7, "code": "NETWORK_ERROR"},
        ]
        result = classify_errors(errors)
        self.assertEqual(result.get("Timeout"), [2])
        self.assertEqual(result.get("HTTP 403 (acceso denegado)"), [3])
        self.assertEqual(result.get("Sin dato AA (no beacon)"), [5])
        self.assertIn("Error de red/conexión", result)
        self.assertEqual(result["Error de red/conexión"], [7])


# ═══════════════════════════════════════════════════════════════════════════
# TESTS: compute_score
# ═══════════════════════════════════════════════════════════════════════════

class TestComputeScore(unittest.TestCase):
    """Tests para compute_score() y compute_url_score()."""

    def test_perfect_score(self):
        """100% éxito → score ~98 (tiempo 5s resta ~2 pts)."""
        metrics = {
            "total": 10, "ok_aa": 10, "ok_dd": 10,
            "times": [5]*10, "total_beacons": 30, "retries": 0, "errors": 0,
            "errores_detalle": [],
        }
        score = compute_score(metrics)
        self.assertGreater(score, 90)
        self.assertLessEqual(score, 100)

    def test_zero_all(self):
        """Sin datos → score bajo pero no crash."""
        metrics = {
            "total": 0, "ok_aa": 0, "ok_dd": 0,
            "times": [], "total_beacons": 0, "retries": 0, "errors": 0,
            "errores_detalle": [],
        }
        score = compute_score(metrics)
        self.assertIsInstance(score, int)
        self.assertGreaterEqual(score, 0)

    def test_half_success(self):
        """50% captura AA → score ~50."""
        metrics = {
            "total": 10, "ok_aa": 5, "ok_dd": 8,
            "times": [10]*10, "total_beacons": 15, "retries": 2, "errors": 2,
            "errores_detalle": [],
        }
        score = compute_score(metrics)
        self.assertIsInstance(score, int)
        self.assertGreater(score, 0)
        self.assertLess(score, 100)

    def test_fast_scores_higher(self):
        """Misma captura, tiempo menor → score mayor."""
        fast = {
            "total": 5, "ok_aa": 5, "ok_dd": 5,
            "times": [3]*5, "total_beacons": 15, "retries": 0, "errors": 0,
            "errores_detalle": [],
        }
        slow = {
            "total": 5, "ok_aa": 5, "ok_dd": 5,
            "times": [55]*5, "total_beacons": 15, "retries": 0, "errors": 0,
            "errores_detalle": [],
        }
        fast_score = compute_score(fast)
        slow_score = compute_score(slow)
        self.assertGreater(fast_score, slow_score,
                           f"fast={fast_score} should be > slow={slow_score}")

    def test_retries_penalize(self):
        """Más reintentos → score menor."""
        low_retry = {
            "total": 10, "ok_aa": 8, "ok_dd": 8,
            "times": [10]*10, "total_beacons": 20, "retries": 0, "errors": 2,
            "errores_detalle": [],
        }
        high_retry = {
            "total": 10, "ok_aa": 8, "ok_dd": 8,
            "times": [10]*10, "total_beacons": 20, "retries": 5, "errors": 2,
            "errores_detalle": [],
        }
        self.assertGreater(compute_score(low_retry), compute_score(high_retry))


class TestComputeUrlScore(unittest.TestCase):
    """Tests para compute_url_score()."""

    def test_perfect_url(self):
        """URL con todo OK → 100."""
        result = {"digitaldata": {"page": "ok"}, "aa_parsed": {"events": []},
                   "extra_beacons": [{}], "error": None, "status": 200, "elapsed_s": 3}
        self.assertEqual(compute_url_score(result), 100)

    def test_no_data(self):
        """URL sin datos → 0."""
        result = {"digitaldata": None, "aa_parsed": None,
                   "extra_beacons": None, "error": "timeout", "status": -1, "elapsed_s": 99}
        self.assertEqual(compute_url_score(result), 0)

    def test_only_dd(self):
        """Solo digitaldata → 30."""
        result = {"digitaldata": {"page": "ok"}, "aa_parsed": None,
                   "extra_beacons": None, "error": None, "status": 200, "elapsed_s": 3}
        self.assertEqual(compute_url_score(result), 60)  # 30 dd + 20 no error + 10 fast

    def test_slow_url_penalty(self):
        """URL lenta (30s) → sin bonus rapidez."""
        result = {"digitaldata": {"page": "ok"}, "aa_parsed": {"events": []},
                   "extra_beacons": None, "error": None, "status": 200, "elapsed_s": 30}
        self.assertEqual(compute_url_score(result), 80)  # 30 dd + 30 aa + 20 no error (sin extra, slow)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    unittest.main(verbosity=2)

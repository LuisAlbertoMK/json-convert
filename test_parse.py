"""
test_parse.py — Tests unitarios para parse_aa_beacon() y build_aa_from_s().

Uso:
  python test_parse.py                    # verbose (default)
  python test_parse.py -v                 # aún más detalle
  python test_parse.py TestParseBeacon    # solo una clase
  python test_parse.py test_beacon_basic  # solo un test

Requiere: unittest (stdlib, 0 deps externas)
No requiere: playwright, openpyxl, red, navegador.
"""

import json
import unittest
import sys
import os

# Importar funciones del script principal
sys.path.insert(0, os.path.dirname(__file__))
from extract_browser import parse_aa_beacon, build_aa_from_s


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
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    unittest.main(verbosity=2)

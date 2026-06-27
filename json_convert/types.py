"""
types.py — TypedDicts y type aliases para el pipeline de json_convert.

Provee tipos reutilizables entre módulos sin importaciones circulares.
"""

from __future__ import annotations

from typing import TypedDict


# ═══════════════════════════════════════════════════════════════════════════
# aa_parser.py
# ═══════════════════════════════════════════════════════════════════════════


class BeaconPage(TypedDict, total=False):
    title: str
    url: str


class BeaconRequest(TypedDict, total=False):
    method: str
    hostname: str | None
    pathname: str
    collectedTimestamp: str
    source: str


class BeaconVisitor(TypedDict, total=False):
    experienceCloudId: str
    audienceManagerHint: str


class BeaconHit(TypedDict, total=False):
    id: str
    type: str
    reportSuiteId: str


class BeaconBrowser(TypedDict, total=False):
    resolution: str
    browserWidth: int
    browserHeight: int
    colorDepth: str
    charset: str


class BeaconResult(TypedDict, total=False):
    """Estructura retornada por parse_aa_beacon() y build_aa_from_s()."""
    solution: str
    page: BeaconPage
    request: BeaconRequest
    visitor: BeaconVisitor
    hit: BeaconHit
    browser: BeaconBrowser
    events: list[str]
    eVars: dict[str, str]
    props: dict[str, str]
    pageName: str
    channel: str
    products: str


# ═══════════════════════════════════════════════════════════════════════════
# browser.py
# ═══════════════════════════════════════════════════════════════════════════


class UrlMetadata(TypedDict):
    title: str
    s_object: dict | None
    elapsed_s: float
    beacon_count: int


class UrlResult(TypedDict, total=False):
    """Estructura retornada por process_url().

    total=False porque 'error' es condicional.
    """
    url: str
    row: int
    page_name: str
    digitaldata: dict | None
    digitaldata_auto: dict | None
    digitaldata_manual: None
    raw_beacons: list[str]
    aa_parsed: None
    extra_beacons: list[str]
    metadata: UrlMetadata
    elapsed_s: float
    status: int
    aa_source: None
    title: str
    code: None
    retries_used: int
    error: str


# ═══════════════════════════════════════════════════════════════════════════
# pipeline.py
# ═══════════════════════════════════════════════════════════════════════════


class PipelineMetrics(TypedDict):
    total: int
    ok_aa: int
    ok_dd: int
    errors: int
    retries: int
    total_beacons: int
    times: list[float]
    errores_detalle: list[dict]
    total_time: float


# ═══════════════════════════════════════════════════════════════════════════
# metrics.py
# ═══════════════════════════════════════════════════════════════════════════


class ErrorDetail(TypedDict):
    row: int
    error: str


# ═══════════════════════════════════════════════════════════════════════════
# excel.py
# ═══════════════════════════════════════════════════════════════════════════

# setup_multisheet retorna una tupla simple (Workbook, Worksheet, str, bool).
# No necesita TypedDict — NamedTuple en el estándar es suficiente.
# Dejamos constancia del tipo para documentación:
#   Tuple[openpyxl.Workbook, openpyxl.Worksheet, str, bool]

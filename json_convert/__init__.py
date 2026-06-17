# json_convert - Paquete de módulos para extracción de datos Adobe Analytics
from json_convert.validation import (
    VALID_URL_SCHEMES,
    ALLOWED_HOSTNAME_SUFFIXES,
    validate_url,
    sanitize_url_for_log,
)

from json_convert.aa_parser import (
    AA_DOMAINS,
    DATA_LAYER_NAMES,
    parse_aa_beacon,
    build_aa_from_s,
    extract_s_object,
    extract_digital_data,
    extract_title,
    try_dismiss_cookie_consent,
)

from json_convert.excel import (
    INPUT_FILE,
    SAVE_EVERY_N,
    SHEET_HEADERS,
    CONTROL_HEADERS,
    HEADER_FILLS,
    DATA_FILLS,
    _pretty_json,
    _set_col_widths,
    _auto_row_height,
    _write_cell,
    validate_sheet,
    save_workbook,
    _is_json_error,
    _has_json_data,
    apply_data_fills,
    split_aa_workbooks,
    setup_multisheet,
    update_control,
    update_vars_sheet,
    print_progress,
)

from json_convert.metrics import (
    ERROR_CODES,
    _error_code_from_detail,
    classify_errors,
    compute_score,
    compute_url_score,
)

__all__ = [
    "VALID_URL_SCHEMES", "ALLOWED_HOSTNAME_SUFFIXES",
    "validate_url", "sanitize_url_for_log",
    "AA_DOMAINS", "DATA_LAYER_NAMES",
    "parse_aa_beacon", "build_aa_from_s",
    "extract_s_object", "extract_digital_data", "extract_title",
    "try_dismiss_cookie_consent",
    "INPUT_FILE", "SAVE_EVERY_N", "SHEET_HEADERS", "CONTROL_HEADERS",
    "HEADER_FILLS", "DATA_FILLS",
    "_pretty_json", "_set_col_widths", "_auto_row_height", "_write_cell",
    "validate_sheet", "save_workbook",
    "_is_json_error", "_has_json_data", "apply_data_fills",
    "split_aa_workbooks", "setup_multisheet", "update_control",
    "update_vars_sheet", "print_progress",
    "ERROR_CODES", "_error_code_from_detail", "classify_errors",
    "compute_score", "compute_url_score",
]

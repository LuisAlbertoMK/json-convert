"""
menu.py - Panel de control para json-convert (zero-dependency TUI).

Uso:
  python menu.py                        # modo interactivo
  python menu.py --run 1                # ejecuta opcion 1 y sale
  python menu.py --run auto             # ejecuta pipeline completo (non-interactive)

Sin dependencias externas - solo stdlib.
"""

import argparse

# Force UTF-8 stdout for ANSI codes to work on Windows
import io
import json
import logging
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

if sys.stdout.encoding and sys.stdout.encoding.upper() != "UTF-8":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

# ── Browser engine configuration ──
# Priority: EXTRACT_BROWSER env var > .menu-config.json > default "chromium"
_MENU_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config/.menu-config.json"
_EXTRACT_BROWSER = os.environ.get("EXTRACT_BROWSER", "").strip().lower()

if not _EXTRACT_BROWSER:
    try:
        if _MENU_CONFIG_PATH.exists():
            _cfg = json.loads(_MENU_CONFIG_PATH.read_text(encoding="utf-8"))
            _EXTRACT_BROWSER = (_cfg.get("browser") or "").strip().lower()
    except Exception:
        logging.debug("Failed to load .menu-config.json")


def _browser_args() -> list[str]:
    """Devuelve ['--browser', 'firefox'] si está configurado Firefox, sino [].

    Configurable via:
      1. Variable de entorno EXTRACT_BROWSER=firefox
      2. Archivo .menu-config.json: {"browser": "firefox"}
    """
    if _EXTRACT_BROWSER == "firefox":
        return ["--browser", "firefox"]
    return []


# -- Detect terminal color support --
_HAS_ANSI = (hasattr(sys.stdout, "isatty") and sys.stdout.isatty()) or bool(os.environ.get("TERM"))
if sys.platform == "win32":
    _HAS_ANSI = True  # Windows 10+ supports ANSI natively

# -- Non-interactive mode flag (set by --run) --
_NON_INTERACTIVE = False


def _c(code: str, text: str) -> str:
    """Wrap text in ANSI code."""
    if not _HAS_ANSI:
        return text
    codes = {
        "bold": "\033[1m", "dim": "\033[2m",
        "green": "\033[92m", "red": "\033[91m", "yellow": "\033[93m",
        "cyan": "\033[96m", "blue": "\033[94m", "magenta": "\033[95m",
        "reset": "\033[0m",
    }
    cstart = codes.get(code, "")
    return cstart + text + codes["reset"]


def c_print(code: str, text: str, end: str = "\n") -> None:
    print(_c(code, text), end=end)


def separator(char: str = "=", width: int = 55) -> None:
    print(_c("dim", char * width))


def header(title: str) -> None:
    print()
    separator()
    c_print("bold", "  " + title)
    separator()


def ask_int(prompt: str, min_v: int, max_v: int, default: int | None = None) -> int:
    """Pide un numero, con default opcional."""
    # Modo no-interactivo: retornar default o min_v
    if _NON_INTERACTIVE:
        return default if default is not None else min_v
    while True:
        raw = input(prompt).strip()
        if not raw and default is not None:
            return default
        try:
            v = int(raw)
            if min_v <= v <= max_v:
                return v
        except ValueError:
            pass
        msg = "  Ingresa un numero entre " + str(min_v) + " y " + str(max_v)
        if default is not None:
            msg += " (Enter = " + str(default) + ")"
        print(_c("yellow", msg))


def confirm(prompt: str, default: bool = True) -> bool:
    """Confirmacion SI/NO."""
    # Modo no-interactivo: retornar default
    if _NON_INTERACTIVE:
        return default
    hint = " [Y/n]" if default else " [y/N]"
    while True:
        raw = input(prompt + hint + ": ").strip().lower()
        if not raw:
            return default
        if raw in ("s", "si", "y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print(_c("yellow", '  Responde con "s" o "n"'))


def run_step(cmd: list[str], label: str, cwd: str | None = None,
             timeout: int | None = None) -> int:
    """Ejecuta un comando mostrando output en tiempo real.

    Returns:
        exit code del comando
    """
    print("\n  -> " + _c("cyan", label))
    separator("-", 40)
    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        proc = subprocess.Popen(
            cmd,
            cwd=cwd or BASE_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )
        for line in iter(proc.stdout.readline, ""):
            print("    " + line, end="")
        proc.wait(timeout=timeout)
        return proc.returncode
    except FileNotFoundError:
        print(_c("red", "    [ERROR] No se encuentra: " + cmd[0]))
        return -1
    except subprocess.TimeoutExpired:
        print(_c("red", "    [ERROR] Tiempo de espera agotado (" + str(timeout) + "s)"))
        return -2
    except Exception as e:
        print(_c("red", "    [ERROR] " + str(e)))
        return -3


def run_ps1(script: str, args: str = "", timeout: int | None = None) -> int:
    """Ejecuta un script PowerShell."""
    cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-NoProfile",
           "-File", os.path.join(BASE_DIR, script)]
    if args:
        cmd += args.split()
    return run_step(cmd, "PowerShell: " + script + " " + args, timeout=timeout)


def get_markets_from_urls() -> list[str]:
    """Lee los mercados únicos desde urls.json."""
    urls_path = os.path.join(BASE_DIR, "data/urls.json")
    if not os.path.exists(urls_path):
        return []
    try:
        with open(urls_path, encoding="utf-8") as f:
            entries = json.load(f)
        markets = sorted(set(e.get("market") for e in entries if e.get("market")))
        return markets
    except Exception:
        return []


def _count_urls(market: str | None = None, env: str = "preview") -> int:
    """Cuenta URLs en urls.json (por mercado y entorno)."""
    urls_path = os.path.join(BASE_DIR, "data/urls.json")
    if not os.path.exists(urls_path):
        return 0
    try:
        with open(urls_path, encoding="utf-8") as f:
            entries = json.load(f)
        if env == "ambas":
            if market:
                return sum(1 for e in entries
                           if e.get("market", "").upper() == market.upper())
            return len(entries)
        if market:
            return sum(1 for e in entries
                       if e.get("market", "").upper() == market.upper()
                       and e.get("entorno", "preview") == env)
        return sum(1 for e in entries if e.get("entorno", "preview") == env)
    except Exception:
        return 0


def _choose_entorno(non_interactive: bool = False) -> str:
    """Muestra submenu de entorno y retorna 'preview', 'produccion' o 'ambas'.

    Args:
        non_interactive: si True, retorna 'ambas' sin preguntar.
    """
    if non_interactive:
        return "ambas"
    n_preview = _count_urls(env="preview")
    n_prod = _count_urls(env="produccion")
    n_ambas = _count_urls(env="ambas")
    print(_c("cyan", "\n  Entorno a auditar:"))
    print("    " + _c("bold", "1") + f". Preview  ({n_preview} URLs)")
    print("    " + _c("bold", "2") + f". Produccion  ({n_prod} URLs)")
    print("    " + _c("bold", "3") + f". Ambas  ({n_ambas} URLs)")
    idx = ask_int("  Selecciona entorno [1-3]: ", 1, 3)
    return {1: "preview", 2: "produccion", 3: "ambas"}[idx]


def detect_markets() -> list[tuple[str, str]]:
    """Detecta directorios de mercado con archivos de auditoria.

    Busca en subdirectorios (PR/, MX/, etc.), subdirectorios de entorno
    ({market}/{entorno}/) y tambien en la raiz del proyecto.
    Retorna UNA entrada por combinacion (mercado, entorno).
    Excluye directorios de codigo/documentacion.
    """
    _EXCLUDED_DIRS = frozenset({
        "src", "scripts", "docs", "tests", "logs", "config",
        "json_convert", ".pytest_cache", "__pycache__", "node_modules",
        "match",  # match/ no es un directorio de mercado
    })
    results: list[tuple[str, str]] = []
    base = Path(BASE_DIR)

    # Evitar duplicados: set de (label, dir_path)
    seen: set[tuple[str, str]] = set()

    def _add(label: str, filepath: str, dirpath: str) -> None:
        key = (label, dirpath)
        if key not in seen:
            seen.add(key)
            results.append((label, filepath))

    for d in base.iterdir():
        dname = d.name
        if not d.is_dir() or dname.startswith(".") or dname.lower() in _EXCLUDED_DIRS:
            continue

        market_label = dname.upper()
        dir_str = str(d)

        # 1) Buscar en subdirectorios de entorno: {market}/preview/, {market}/produccion/
        entornos = sorted(
            e for e in d.iterdir()
            if e.is_dir() and not e.name.startswith(".")
            and e.name.lower() in ("preview", "produccion")
        )
        for ent in entornos:
            entorno_name = ent.name.capitalize()
            label = f"{market_label} - {entorno_name}"
            for candidate in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
                fp = ent / candidate
                if fp.exists():
                    _add(label, str(fp), dir_str)
                    break  # solo un entry por (market, entorno)

        # 2) Legacy: {market}/historial.xlsx (solo si no hay subdirectorios de entorno)
        if not entornos:
            for candidate in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
                fp = d / candidate
                if fp.exists():
                    _add(market_label, str(fp), dir_str)
                    break

    # 3) Legacy: raiz del proyecto
    for candidate in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
        fp = base / candidate
        if fp.exists():
            _add("RAIZ", str(fp), str(base))
            break

    return results


def get_project_status() -> dict:
    """Retorna estado resumido del proyecto."""
    base = Path(BASE_DIR)
    py_files = list(base.glob("*.py"))
    markets = detect_markets()

    # Ultima auditoria: buscar sheets en historial.xlsx
    last_audit = "-"
    try:
        import openpyxl
        for _m, hpath in markets:
            wb = openpyxl.load_workbook(hpath, data_only=True)
            sheets = [s for s in wb.sheetnames if s not in ("_control", "_vars") and not s.startswith("_")]
            wb.close()
            if sheets:
                last_sheet = max(sheets)
                if last_sheet > last_audit or last_audit == "-":
                    last_audit = last_sheet
    except Exception:
        pass

    return {
        "py_files": len(py_files),
        "markets": [m for m, _ in markets],
        "last_audit": last_audit,
        "has_data": len(markets) > 0,
    }


# ============================================
#  HELPERS DE MERCADO
# ============================================

ALL_MARKETS = "__ALL__"  # sentinel para "todos los mercados"


def choose_market(source: str = "detect") -> tuple[str | None, str | None]:
    """Muestra submenu de mercados y retorna (nombre|ALL_MARKETS, archivo|None).

    Args:
        source: "detect" → detect_markets() (busca archivos Excel)
                "urls" → get_markets_from_urls() (desde urls.json)

    Returns:
        (market_name, filepath) o (ALL_MARKETS, None) para "todos"
    """
    if source == "urls":
        markets = get_markets_from_urls()
        items = [(m, None) for m in markets]
    else:
        items = detect_markets()  # [(name, path), ...]

    if not items:
        print(_c("yellow", "\n  [!] No hay mercados disponibles."))
        return None, None

    if len(items) == 1:
        name = items[0][0]
        print(_c("dim", f"\n  Mercado único: {name} (seleccionado automáticamente)"))
        return items[0]

    print(_c("cyan", "\n  Mercados disponibles:"))
    print("    " + _c("bold", "0") + ". Todos los mercados")
    for i, item in enumerate(items, 1):
        name = item[0]
        extra = f" ({os.path.relpath(item[1], BASE_DIR)})" if len(item) > 1 and item[1] else ""
        print(f"    {i}. {name}{extra}")

    idx = ask_int("  Selecciona mercado [0-" + str(len(items)) + "]: ", 0, len(items))
    if idx == 0:
        return (ALL_MARKETS, None)
    return items[idx - 1]


# ============================================
#  OPCIONES DEL MENU
# ============================================

def op_auditar(non_interactive: bool = False) -> None:
    """Opcion 2: Solo auditoria (extract_browser) por mercado."""
    header("AUDITORIA - extract_browser.py")

    urls_path = os.path.join(BASE_DIR, "data/urls.json")
    if not os.path.exists(urls_path):
        print(_c("yellow", "  No se encuentra urls.json"))
        if confirm("  Generarlo desde RevisionManual.xlsx?"):
            run_step([sys.executable, "scripts/_gen_urls.py"], "Generando urls.json")
        else:
            print(_c("yellow", "  Cancelado. Necesitas urls.json para auditar."))
            return

    entorno = _choose_entorno(non_interactive)

    market, _ = choose_market(source="urls")
    if market is None:
        print(_c("yellow", "  No hay mercados en urls.json. Generalo primero."))
        return

    if market == ALL_MARKETS:
        markets = get_markets_from_urls()
        for m in markets:
            n = _count_urls(m, entorno)
            cmd = ([sys.executable, "src/extract_browser.py", "--urls", "data/urls.json",
                    "--market", m, "--entorno", entorno, "--split-aa", "--progress"]
                   + _browser_args())
            run_step(cmd, f"Auditando {m} ({n} URLs [{entorno}])...", timeout=600)
    else:
        n = _count_urls(market, entorno)
        cmd = ([sys.executable, "src/extract_browser.py", "--urls", "data/urls.json",
                "--market", market, "--entorno", entorno, "--split-aa", "--progress"]
               + _browser_args())
        run_step(cmd, f"Auditando {market} ({n} URLs [{entorno}])...", timeout=600)

    print(_c("green", "\n  [OK] Auditoria finalizada."))


def op_postprocesar(non_interactive: bool = False) -> None:
    """Opcion 3: Solo post-procesar (extract_aa)."""
    header("POST-PROCESO - extract_aa.py")

    market, hpath = choose_market(source="detect")
    if market is None:
        # Sin archivos → auto-bootstrap desde urls.json
        print(_c("yellow", "  No se encontraron archivos de auditoria."))
        urls_path = os.path.join(BASE_DIR, "data/urls.json")
        if os.path.exists(urls_path) and confirm("  Generar desde urls.json automaticamente?"):
            run_step([sys.executable, "src/extract_aa.py", "--input", "historial.xlsx",
                      "--urls", urls_path],
                     "Generando y procesando desde urls.json...", timeout=600)
        else:
            print(_c("yellow", "  Cancelado. Pasa --urls o ejecuta primero una auditoria."))
        return

    if market == ALL_MARKETS:
        markets = detect_markets()
        for _m, hpath in markets:
            _run_extract_aa(hpath)
            _run_extract_aa_companions(hpath)
    else:
        _run_extract_aa(hpath)
        _run_extract_aa_companions(hpath)

    print(_c("green", "\n  [OK] Post-proceso finalizado."))


def _run_extract_aa(hpath: str) -> None:
    """Ejecuta extract_aa.py sobre un archivo historial (auto-bootstrap si falta)."""
    run_step([sys.executable, "src/extract_aa.py", "--input", hpath, "--urls", "data/urls.json"],
             "Procesando: " + hpath, timeout=600)


def _run_extract_aa_companions(hpath: str) -> None:
    """Ejecuta extract_aa.py sobre con_aa.xlsx y sin_aa.xlsx del mismo dir."""
    base_dir = os.path.dirname(hpath)
    for fname in ["con_aa.xlsx", "sin_aa.xlsx"]:
        fp = os.path.join(base_dir, fname)
        if os.path.exists(fp):
            run_step([sys.executable, "src/extract_aa.py", "--input", fp, "--urls", "data/urls.json"],
                     "Procesando: " + fp, timeout=600)


def op_reporte(non_interactive: bool = False) -> None:
    """Opcion 4: Generar reporte de fallos."""
    header("REPORTE DE FALLOS - audit_report.py")

    entorno = _choose_entorno(non_interactive)
    urls_path = os.path.join(BASE_DIR, "data/urls.json")
    urls_arg = ["--urls", urls_path] if os.path.exists(urls_path) else []

    code = run_step(
        [sys.executable, "src/audit_report.py"] + urls_arg + ["--entorno", entorno],
        f"Generando reporte (entorno: {entorno}, puede tardar 15-30 min en auto-bootstrap)...",
        timeout=3600)

    if code == 0:
        print(_c("green", "\n  [OK] Reporte global generado: reporte_auditoria.xlsx"))
        markets = detect_markets()
        for m, _ in markets:
            m_code = m.split(" - ")[0]
            if m_code != "RAIZ":
                for ent in ("produccion", "preview"):
                    ep = os.path.join(BASE_DIR, m_code, ent, "historial.xlsx")
                    if os.path.exists(ep):
                        print(_c("green", f"       {m_code}/{ent}/historial.xlsx"))
                print(_c("green", f"       {m_code}/reporte-auditoria.xlsx"))
        if confirm("  Abrir el global?"):
            open_file("reporte_auditoria.xlsx")
    else:
        print(_c("red", "\n  [ERROR] Fallo la generacion del reporte."))


def op_catalogo(non_interactive: bool = False) -> None:
    """Opcion 7: Generar catalog de migracion."""
    header("CATALOGO DE MIGRACION - generate_migration_catalog.py")

    # Verificar que existe historial
    markets = detect_markets()
    if not markets:
        print(_c("yellow", "  No se encontraron archivos de auditoria."))
        print("  Ejecuta primero una auditoria (opcion 1 o 2).")
        return

    # Elegir mercado
    print(_c("cyan", "  Mercados detectados:"))
    for i, (m, hpath) in enumerate(markets, 1):
        print("    " + str(i) + ". " + m + " (" + os.path.relpath(hpath, BASE_DIR) + ")")

    if len(markets) == 1:
        m_idx = 1
    else:
        m_idx = ask_int("  Selecciona mercado [1-" + str(len(markets)) + "]: ", 1, len(markets))

    market_name, historial_path = markets[m_idx - 1]

    # Verificar url-mapping.json
    mapping_path = os.path.join(BASE_DIR, "data/url-mapping.json")
    if not os.path.exists(mapping_path):
        print(_c("yellow", "\n  No se encuentra url-mapping.json"))
        if confirm("  Generar template desde RevisionManual.xlsx?", default=True):
            input_path = os.path.join(BASE_DIR, "RevisionManual.xlsx")
            if os.path.exists(input_path):
                run_step(
                    [sys.executable, "src/generate_migration_catalog.py",
                     "--gen-template", "--input", input_path,
                     "--mapping", "data/url-mapping.json"],
                    "Generando template url-mapping.json...")
                print(_c("yellow", "\n  [!] EDITAR url-mapping.json con production_url, aem_path y page_key"))
                print("      Despues ejecutar esta opcion nuevamente.")
            else:
                print(_c("red", "  No se encuentra RevisionManual.xlsx"))
        return

    # Verificar expected.json
    expected_path = os.path.join(BASE_DIR, "data/expected.json")
    if not os.path.exists(expected_path):
        print(_c("red", "  No se encuentra expected.json"))
        print("  Debe existir con las reglas del estandar US para cada mercado.")
        return

    # Generar catalogo en la misma carpeta del historial (PR/preview/ o PR/produccion/)
    output_name = "catalogo-migracion.xlsx"
    historial_dir = os.path.dirname(historial_path)
    output_path = os.path.join(historial_dir, output_name)

    # Mostrar resumen antes de ejecutar
    print()
    print(_c("cyan", "  Resumen:"))
    print("    Mercado:      " + market_name)
    print("    Historial:    " + os.path.relpath(historial_path, BASE_DIR))
    print("    Mapping:      url-mapping.json")
    print("    Expected:     expected.json")
    print("    Output:       " + os.path.relpath(output_path, BASE_DIR))
    print()

    if confirm("  Generar catalogo?", default=True):
        run_step(
            [sys.executable, "src/generate_migration_catalog.py",
             "--historial", historial_path,
             "--mapping", "data/url-mapping.json",
             "--expected", "data/expected.json",
             "--market", market_name.split(" - ")[0],
             "--output", output_path],
            "Generando catalogo de migracion...", timeout=60)

        if os.path.exists(output_path):
            print(_c("green", "\n  [OK] Catalogo generado: " + os.path.relpath(output_path, BASE_DIR)))
            if confirm("  Abrirlo ahora?", default=True):
                open_file(output_path)


def op_limpieza() -> None:
    """Opcion 5: Limpieza (run.ps1 -SkipTests)."""
    header("LIMPIEZA - run.ps1")
    run_ps1("scripts/run.ps1", "-SkipTests", timeout=60)
    print(_c("green", "\n  [OK] Limpieza finalizada."))





def op_tests() -> None:
    """Opcion 9: Ejecutar tests."""
    header("TESTS - pytest")
    rc = run_step([sys.executable, "-m", "pytest", "--tb=long", "-q"],
                  "pytest completo", timeout=120)
    if rc == 0:
        print(_c("green", "\n  [OK] Todos los tests pasaron."))
    else:
        print(_c("red", "\n  [ERROR] Fallaron tests (codigo: " + str(rc) + ")"))


def op_match() -> None:
    """Opcion 10: Match prod vs preview."""
    header("MATCH PROD VS PREVIEW - match_prod_preview.py")

    # Detectar mercados
    markets = detect_markets()
    if not markets:
        print(_c("yellow", "  No se encontraron archivos de auditoria."))
        print("  Ejecuta primero una auditoria (opcion 1 o 2).")
        return

    print(_c("cyan", "  Mercados detectados:"))
    for i, (m, hpath) in enumerate(markets, 1):
        m_code = m.split(" - ")[0]
        extra = ""
        prev_path = os.path.join(BASE_DIR, m_code, "historial_preview.xlsx")
        if os.path.exists(prev_path):
            extra = _c("green", " [preview OK]")
        else:
            extra = _c("yellow", " [sin preview — usando expected.json]")
        print(f"    {i}. {m} ({os.path.relpath(hpath, BASE_DIR)}){extra}")

    if len(markets) == 1:
        m_idx = 1
    else:
        m_idx = ask_int("  Selecciona mercado [1-" + str(len(markets)) + "]: ", 1, len(markets))

    market_label, historial_path = markets[m_idx - 1]
    market_name = market_label.split(" - ")[0]  # "PR - Preview" → "PR"
    mapping_path = os.path.join(BASE_DIR, "data/url-mapping.json")
    expected_path = os.path.join(BASE_DIR, "data/expected.json")

    if not os.path.exists(mapping_path):
        print(_c("red", "  No se encuentra url-mapping.json"))
        return
    if not os.path.exists(expected_path):
        print(_c("red", "  No se encuentra expected.json"))
        return

    # Chequear si hay preview
    preview_path = os.path.join(BASE_DIR, market_name, "historial_preview.xlsx")
    preview_arg = []
    if os.path.exists(preview_path):
        preview_arg = ["--preview", preview_path]
        print(_c("green", f"\n  Preview detectado: {market_name}/historial_preview.xlsx"))
    else:
        print(_c("yellow", "\n  Sin preview — usando expected.json como prometido"))
        print("  Para extraer preview: corre opcion 2 con --entorno preview y VPN activa.")
        print(f"  Luego renombrar a: {market_name}/historial_preview.xlsx")

    if confirm("  Generar match report?", default=True):
        run_step(
            [sys.executable, "src/match_prod_preview.py",
             "--production", historial_path,
             "--mapping", mapping_path,
             "--expected", expected_path,
             "--market", market_name] + preview_arg,
            f"Match {market_label}...", timeout=60)

        output_path = os.path.join(BASE_DIR, market_name, "match-prod-vs-preview.xlsx")
        if os.path.exists(output_path):
            print(_c("green", f"\n  [OK] Match generado: {market_name}/match-prod-vs-preview.xlsx"))
            print(_c("green", f"       HTML: {market_name}/match-prod-vs-preview.html"))
            if confirm("  Abrir el HTML?", default=True):
                open_file(os.path.join(BASE_DIR, market_name, "match-prod-vs-preview.html"))


def op_validacion() -> None:
    """Opcion 11: Generar matriz de validación."""
    header("MATRIZ DE VALIDACION - generate_validation_matrix.py")

    markets = detect_markets()
    if not markets:
        print(_c("yellow", "  No se encontraron archivos de auditoria."))
        print("  Ejecuta primero una auditoria (opcion 1 o 2).")
        return

    # Detectar mercados únicos (sin duplicar preview/produccion)
    seen = {}
    for label, hpath in markets:
        m_code = label.split(" - ")[0]
        if m_code not in seen:
            seen[m_code] = hpath

    market_list = sorted(seen.items())
    print(_c("cyan", "  Mercados detectados:"))
    for i, (m, hpath) in enumerate(market_list, 1):
        extra = ""
        preview_dir = os.path.join(BASE_DIR, m, "preview")
        prod_dir = os.path.join(BASE_DIR, m, "produccion")
        preview_ok = os.path.exists(os.path.join(preview_dir, "historial.xlsx"))
        prod_ok = os.path.exists(os.path.join(prod_dir, "historial.xlsx"))
        if preview_ok or prod_ok:
            extra = _c("green", "  [preview+produccion OK]" if preview_ok and prod_ok
                       else _c("yellow", "  [parcial]"))
        else:
            extra = _c("dim", "  [sin historial]")
        print(f"    {i}. {m}{extra}")

    idx = ask_int("  Selecciona mercado [1-" + str(len(market_list)) + "]: ", 1, len(market_list))
    market_name = market_list[idx - 1][0]

    print()
    print("  " + _c("bold", "Entorno:"))
    print("    " + _c("bold", "1") + ") Preview")
    print("    " + _c("bold", "2") + ") Produccion")
    print("    " + _c("bold", "3") + ") Ambas (Preview vs Produccion lado a lado)")
    e_idx = ask_int("  Elige [1-3]: ", 1, 3)
    entorno_map = {1: "preview", 2: "produccion", 3: "ambas"}
    entorno = entorno_map[e_idx]

    mapping_path = os.path.join(BASE_DIR, "data/url-mapping.json")
    expected_path = os.path.join(BASE_DIR, "data/expected.json")
    if not os.path.exists(mapping_path):
        print(_c("red", "  No se encuentra data/url-mapping.json"))
        return
    if not os.path.exists(expected_path):
        print(_c("red", "  No se encuentra data/expected.json"))
        return

    print()
    print("  " + _c("bold", "Formato de salida:"))
    print("    " + _c("bold", "1") + ") Todo en un mismo xlsx (varias sheets — una por URL)")
    print("    " + _c("bold", "2") + ") Todo en una misma sheet")
    fmt_idx = ask_int("  Elige [1-2]: ", 1, 2)
    split_files = (fmt_idx == 1)

    if confirm(f"  Generar matriz de validación para {market_name} ({entorno})?", default=True):
        cmd = [sys.executable, "src/generate_validation_matrix.py",
               "--market", market_name,
               "--entorno", entorno,
               "--mapping", mapping_path,
               "--expected", expected_path]
        if split_files:
            cmd.append("--split")

        # Add explicit historial paths for non-standard layout
        preview_path = os.path.join(BASE_DIR, market_name, "preview", "historial.xlsx")
        prod_path = os.path.join(BASE_DIR, market_name, "produccion", "historial.xlsx")
        if entorno in ("preview", "ambas") and os.path.exists(preview_path):
            cmd += ["--historial-preview", preview_path]
        if entorno in ("produccion", "ambas") and os.path.exists(prod_path):
            cmd += ["--historial-produccion", prod_path]

        run_step(cmd, f"Matriz validación {market_name}...", timeout=120)

        if split_files:
            split_dir = os.path.join(BASE_DIR, market_name, "split")
            if os.path.exists(split_dir):
                count = len([f for f in os.listdir(split_dir) if f.endswith(".xlsx")])
                print(_c("green", f"\n  [OK] {count} matrices generadas en: {market_name}/split/"))
                if confirm("  Abrir carpeta?", default=True):
                    open_file(split_dir)
        else:
            output_name = f"matriz-validacion-{entorno}.xlsx"
            output_path = os.path.join(BASE_DIR, market_name, output_name)
            if os.path.exists(output_path):
                print(_c("green", f"\n  [OK] Matriz: {market_name}/{output_name}"))
                if confirm("  Abrir el Excel?", default=True):
                    open_file(output_path)


def op_ver_resumen() -> None:
    """Opcion 12: Abrir resumen del catálogo .html."""
    header("RESUMEN CATALOGO MIGRACION")

    # Buscar resumenes .html en directorios de mercado (incluyendo preview/produccion/)
    resumenes = []
    base = Path(BASE_DIR)
    for d in base.iterdir():
        if not d.is_dir() or d.name.startswith("."):
            continue
        # Buscar en {market}/preview/ y {market}/produccion/
        for sub in ("preview", "produccion"):
            html_path = d / sub / "resumen-catalogo-migracion.html"
            if html_path.exists():
                label = f"{d.name.upper()} - {sub.capitalize()}"
                resumenes.append((label, str(html_path)))
        # Legacy: {market}/resumen-catalogo-migracion.html
        html_path = d / "resumen-catalogo-migracion.html"
        if html_path.exists():
            resumenes.append((d.name.upper(), str(html_path)))

    # Tambien en la raiz
    root_html = base / "resumen-catalogo-migracion.html"
    if root_html.exists():
        resumenes.append(("RAIZ", str(root_html)))

    if not resumenes:
        print(_c("yellow", "  No se encontraron resumenes de catálogo."))
        print("  Genera primero el catálogo (opcion 7) que auto-genera el resumen.")
        return

    print(_c("cyan", "  Resúmenes disponibles:"))
    for i, (m, path) in enumerate(resumenes, 1):
        size = os.path.getsize(path)
        print(f"    {i}. {m} ({size // 1024} KB)")

    idx = ask_int("  Cuál querés ver [1-" + str(len(resumenes)) + "]: ", 1, len(resumenes))
    open_file(resumenes[idx - 1][1])


def op_ver_resultados() -> None:
    """Opcion 6: Abrir Excel."""
    header("RESULTADOS")

    targets = []

    # Global report in root (generated on next run) or in market/entorno
    global_report = os.path.join(BASE_DIR, "reporte_auditoria.xlsx")
    if os.path.exists(global_report):
        targets.append(global_report)

    # Per-market files: historiales + reports
    for _m, hpath in detect_markets():
        base = os.path.dirname(hpath)
        for fname in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
            fp = os.path.join(base, fname)
            if os.path.exists(fp) and fp not in targets:
                targets.append(fp)

    # Legacy root historiales
    for fname in ["historial.xlsx", "reporte_auditoria.xlsx"]:
        fp = os.path.join(BASE_DIR, fname)
        if os.path.exists(fp) and fp not in targets:
            targets.append(fp)

    if not targets:
        print(_c("yellow", "  No hay archivos Excel para mostrar."))
        print("  Ejecuta primero una auditoria.")
        return

    print(_c("cyan", "  Archivos disponibles:"))
    for i, t in enumerate(targets, 1):
        size = os.path.getsize(t)
        label = os.path.relpath(t, BASE_DIR)
        print("    " + str(i) + ". " + label + " (" + str(size // 1024) + " KB)")

    idx = ask_int("\n  Cual queres abrir [1-" + str(len(targets)) + "]: ", 1, len(targets))
    open_file(targets[idx - 1])


def _run_audit_for_env(markets: list[str], env: str, results: list) -> bool:
    """Corre extract_browser para un entorno particular. Retorna True si todo OK."""
    ok = True
    for m in markets:
        n = _count_urls(m, env)
        rc = run_step(
            [sys.executable, "src/extract_browser.py", "--urls", "data/urls.json",
             "--market", m, "--entorno", env, "--split-aa", "--progress"]
            + _browser_args(),
            f"Auditando {m} ({n} URLs [{env}])...", timeout=600)
        results.append((f"Auditoria {m} ({env})", rc))
        if rc != 0:
            ok = False
    return ok


def op_todo_en_uno(target_market=None, non_interactive=False):
    """Opcion 1: Pipeline completo — multi-entorno con match 3‑vías.

    Args:
        target_market: None → preguntar (interactivo) o ALL (--run)
                       "PR" → solo ese mercado
                       ALL_MARKETS → todos
        non_interactive: si True, no hace preguntas al usuario.
    """
    separator("=", 55)
    c_print("bold", "  [!] PIPELINE COMPLETO")
    separator("=", 55)
    print()
    print("  Inicio: " + datetime.now().strftime("%H:%M:%S"))
    print()

    # Elegir mercado
    if target_market is None:
        market_choice, _ = choose_market(source="urls")
        if market_choice is None:
            print(_c("yellow", "  No hay mercados disponibles."))
            return
        target_market = market_choice

    markets_to_run = []
    if target_market == ALL_MARKETS:
        markets_to_run = get_markets_from_urls()
    else:
        markets_to_run = [target_market]

    # Elegir entorno
    entorno = _choose_entorno(non_interactive)
    entornos_to_run = ["preview", "produccion"] if entorno == "ambas" else [entorno]

    market_label = "todos los mercados" if target_market == ALL_MARKETS else target_market
    print(_c("dim", f"  Mercado(s): {market_label}  |  Entorno(s): {', '.join(entornos_to_run)}"))
    print()

    results = []

    # Paso 1: Tests
    print()
    print(_c("cyan", "  [1] Ejecutando tests..."))
    rc = run_step([sys.executable, "-m", "pytest", "--tb=short", "-q"],
                  "pytest --tb=short -q", timeout=120)
    if rc != 0:
        print(_c("yellow", "    ⚠ Tests fallaron (codigo: " + str(rc) + ")"))
        if not confirm("    Continuar pipeline igual?", default=False):
            print(_c("red", "    Pipeline abortado por fallo en tests."))
            results.append(("Tests", rc))
            print()
            _pipeline_summary(results)
            return
    results.append(("Tests", rc))

    # Paso 2: Verificar urls.json
    print()
    print(_c("cyan", "  [2] Verificando entorno..."))
    urls_path = os.path.join(BASE_DIR, "data/urls.json")
    has_urls = os.path.exists(urls_path)
    if not has_urls:
        print(_c("yellow", "    No se encuentra urls.json"))
        rp = os.path.join(BASE_DIR, "RevisionManual.xlsx")
        if os.path.exists(rp):
            if confirm("    Generarlo desde RevisionManual.xlsx?", default=True):
                rc = run_step([sys.executable, "scripts/_gen_urls.py"], "Generando urls.json")
                results.append(("Generar urls.json", rc))
                has_urls = (rc == 0)
            else:
                print(_c("yellow", "    Saltando generacion de urls.json"))
        else:
            print(_c("yellow", "    No hay RevisionManual.xlsx para generar urls.json"))
    else:
        print(_c("green", "    urls.json OK"))
        results.append(("Verificar entorno", 0))

    # Pasos 3‑N: Auditoria + Post-proceso por cada entorno
    post_markets = []

    for env_idx, env in enumerate(entornos_to_run):
        step_n = 3 + env_idx * 2

        # Auditoría — SIEMPRE se ejecuta (el usuario exige datos frescos)
        print()
        print(_c("cyan", f"  [{step_n}] Auditoría ({env})..."))
        audit_ok = True

        if has_urls:
            audit_ok = _run_audit_for_env(markets_to_run, env, results)
        else:
            rc = run_step([sys.executable, "src/extract_browser.py"] + _browser_args(),
                          "Ejecutando (modo Excel plano)...", timeout=600)
            results.append(("Auditoria", rc))
            audit_ok = (rc == 0)

        if not audit_ok:
            print(_c("red", f"  ⚠ Auditoría ({env}) falló."))
            continue

        # Post-procesar
        print()
        print(_c("cyan", f"  [{step_n + 1}] Post-procesando ({env})..."))
        all_m = detect_markets()
        for m, hpath in all_m:
            m_code = m.split(" - ")[0]  # "PR - Preview" → "PR"
            if m_code in markets_to_run or target_market == ALL_MARKETS:
                _run_extract_aa(hpath)
                _run_extract_aa_companions(hpath)
                post_markets.append((m_code, hpath))  # guardar solo el codigo de mercado
                results.append((f"Post-proceso {m} ({env})", 0))

    # Reportes de fallos por entorno
    for env in entornos_to_run:
        print()
        print(_c("cyan", f"  Reporte de fallos ({env})..."))
        rc = run_step(
            [sys.executable, "src/audit_report.py", "--urls", "data/urls.json",
             "--entorno", env, "--output-dir", "."],
            f"audit_report ({env})", timeout=600)
        results.append((f"Reporte fallos ({env})", rc))

    # Match 3‑vías si tenemos ambos entornos
    if len(entornos_to_run) > 1:
        print()
        print(_c("cyan", "  Match 3‑vías (US Expected vs Preview vs Producción)..."))
        for m in markets_to_run:
            rc = run_step(
                [sys.executable, "src/match_prod_preview.py",
                 "--market", m, "--mode", "3way"],
                f"Match 3‑vías {m}...", timeout=60)
            results.append((f"Match 3‑vías {m}", rc))

    # (Catálogo de migración removido del pipeline automático)
    #

    # Resultados
    print()
    print(_c("cyan", "  Resultados..."))
    match_html = os.path.join(BASE_DIR, "PR", "match", "match-3way.html")
    if os.path.exists(match_html):
        if confirm("  Abrir match 3‑vías?", default=True):
            open_file(match_html)
    elif post_markets:
        for m, hpath in post_markets[:1]:
            if confirm(f"  Abrir {os.path.dirname(hpath)}/historial.xlsx?", default=True):
                open_file(hpath)
                break

    _pipeline_summary(results)


def _pipeline_summary(results: list) -> None:
    """Resumen final del pipeline."""
    separator("=", 55)
    c_print("bold", "  [R] RESUMEN DEL PIPELINE")
    separator("-", 55)
    ok_count = 0
    for label, rc in results:
        icon = _c("green", "+") if rc == 0 else _c("red", "x")
        print("  " + icon + " " + label)
        if rc == 0:
            ok_count += 1
    separator("-", 55)
    print("  " + str(ok_count) + "/" + str(len(results)) + " pasos completados")
    print("  Fin: " + datetime.now().strftime("%H:%M:%S"))
    if ok_count == len(results):
        print(_c("green", "  Pipeline completado exitosamente."))
    else:
        print(_c("yellow", "  Algunos pasos fallaron - revisa los logs arriba."))
    separator("=", 55)
    print()


def open_file(path: str) -> None:
    """Abre un archivo con la aplicacion asociada."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as e:
        print(_c("red", "  Error al abrir: " + str(e)))


def show_menu() -> int:
    """Muestra el menu principal y retorna la opcion elegida."""
    status = get_project_status()
    markets_str = ", ".join(status["markets"]) if status["markets"] else "-"
    last_audit_str = status["last_audit"] if status["last_audit"] != "-" else "-"
    data_indicator = _c("green", "con datos") if status["has_data"] else _c("yellow", "sin datos")

    print()
    separator("=", 55)
    c_print("bold", "     JSON-CONVERT  -  Panel de Control")
    c_print("dim", "     Extraccion Adobe Analytics | Ford Preview")
    separator("=", 55)
    print("  Proyecto: " + _c("cyan", str(status["py_files"])) + " scripts  |  "
          "Mercados: " + _c("cyan", markets_str))
    print("  Ultima auditoria: " + _c("cyan", last_audit_str)
          + "  |  " + data_indicator)
    separator("-", 55)
    print()
    print("  " + _c("bold", "1") + ") " + _c("cyan", "[!] TODO EN UNO") + "        pipeline completo")
    print("  " + _c("bold", "2") + ") " + _c("green", ">") + "  Solo auditoria       extract_browser.py")
    print("  " + _c("bold", "3") + ") " + _c("magenta", "[R]") + "  Solo reporte fallos  audit_report.py")
    print("  " + _c("bold", "4") + ") " + _c("green", "[T]") + "  Ejecutar tests       pytest")
    print("  " + _c("bold", "5") + ") " + _c("magenta", "[D]") + "  Match prod vs preview  comparar prometido vs entregado")
    print("  " + _c("bold", "6") + ") " + _c("green", "[V]") + "  Matriz validación     generate_validation_matrix.py")
    print("  " + _c("bold", "7") + ") " + _c("cyan", "[N]") + "  Auditoría individual  auditar_urls.py → PR/tickets/")
    print("  " + _c("bold", "8") + ") " + _c("magenta", "[S]") + "  Auditoría por semanas auditar_semanas.py")
    print()
    separator("-", 55)
    print("  " + _c("dim", "0") + ") " + _c("dim", "x  Salir"))
    separator("=", 55)

    return ask_int("  Opcion [0-8]: ", 0, 8)


def op_auditar_urls() -> None:
    """Opcion 7: Pipeline de auditoria individual por URL -> PR/tickets/."""
    header("AUDITORIA INDIVIDUAL - auditar_urls.py")
    # El script ya limpia con --clean
    rc = run_step([sys.executable, "scripts/auditar_urls.py", "--clean"],
                  "auditar_urls", cwd=str(BASE_DIR))
    if rc == 0:
        c_print("green", "\n  [OK] Pipeline completado. Revisa PR/tickets/")
    else:
        c_print("red", f"\n  [ERR] Pipeline fallo (exit code {rc})")


def op_auditar_semanas() -> None:
    """Opcion 8: Pipeline de auditoria por semanas."""
    # Detectar archivos disponibles
    posibles = ["SEMANAS_MX.json", "SEMANAS_PR.json"]
    existentes = [f for f in posibles if os.path.exists(os.path.join(BASE_DIR, f))]
    default = os.path.join(BASE_DIR, existentes[0]) if existentes else ""

    print()
    if len(existentes) > 1:
        print("  Archivos disponibles:")
        for i, f in enumerate(existentes, 1):
            print(f"    {i}. {f}")
        idx = ask_int("  Selecciona [1-" + str(len(existentes)) + "]: ", 1, len(existentes), default=1)
        json_path = os.path.join(BASE_DIR, existentes[idx - 1])
    elif existentes:
        raw = input(f"  Archivo JSON [{existentes[0]}]: ").strip()
        json_path = os.path.join(BASE_DIR, raw or existentes[0])
    else:
        raw = input("  Ruta del archivo JSON con semanas: ").strip()
        json_path = os.path.join(BASE_DIR, raw) if raw else ""
        if not json_path or not os.path.exists(json_path):
            c_print("red", "  [ERR] Archivo no encontrado")
            return

    header("AUDITORIA POR SEMANAS")
    rc = run_step([sys.executable, "scripts/auditar_semanas.py", json_path],
                  "auditar_semanas", cwd=str(BASE_DIR))
    if rc == 0:
        c_print("green", "\n  [OK] Pipeline completado. Revisa */semanas/*/")
    else:
        c_print("red", f"\n  [ERR] Pipeline fallo (exit code {rc})")


def run_option(opt: int, non_interactive: bool = False) -> bool:
    """Ejecuta la opcion. Retorna False si hay que salir."""
    start = time.time()

    if opt == 1:
        op_todo_en_uno(target_market=ALL_MARKETS if non_interactive else None,
                       non_interactive=non_interactive)
    elif opt == 2:
        op_auditar(non_interactive=non_interactive)
    elif opt == 3:
        op_reporte(non_interactive=non_interactive)
    elif opt == 4:
        op_tests()
    elif opt == 5:
        op_match()
    elif opt == 6:
        op_validacion()
    elif opt == 7:
        op_auditar_urls()
    elif opt == 8:
        op_auditar_semanas()
    elif opt == 0:
        print()
        c_print("green", "  Hasta luego!")
        return False

    elapsed = time.time() - start
    if elapsed > 2:
        print("  (" + _c("dim", str(int(elapsed)) + "s") + ")")
    print()
    if not _NON_INTERACTIVE:
        input(_c("dim", "  Presiona Enter para volver al menu..."))
    return True


# ============================================
#  MAIN
# ============================================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Panel de control json-convert")
    parser.add_argument("--run", type=str,
                        help="Ejecutar opcion directa: numero 1-8, 0, o 'auto'")
    args = parser.parse_args()

    BASE_DIR = Path(__file__).resolve().parent.parent

    # Verificar proyecto
    is_project = any(
        os.path.exists(os.path.join(BASE_DIR, f))
        for f in ["src/extract_browser.py", "src/extract_aa.py", "scripts/run.ps1"]
    )
    if not is_project:
        print(_c("red", "[ERROR] Este script debe ejecutarse desde la raiz del proyecto json-convert"))
        print("        No se encontraron extract_browser.py, extract_aa.py o run.ps1")
        sys.exit(1)

    if sys.version_info < (3, 10):
        print(_c("red", "[ERROR] Python 3.10+ requerido"))
        sys.exit(1)

    # Modo directo --run
    if args.run:
        opt_map = {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5,
                   "6": 6, "7": 7, "8": 8, "0": 0, "auto": 1}
        opt = opt_map.get(args.run, -1)
        if opt < 0:
            print(_c("red", "[ERROR] Opcion invalida: " + args.run))
            print("  Valores validos: 1-8, 0, 'auto'")
            sys.exit(1)

        # Modo no-interactivo: respuestas automaticas
        _NON_INTERACTIVE = True

        run_option(opt, non_interactive=True)
        sys.exit(0)

    # Modo interactivo
    print(_c("cyan", "  Iniciando menu..."))
    while True:
        try:
            opt = show_menu()
            if not run_option(opt):
                break
        except KeyboardInterrupt:
            print()
            c_print("green", "\n  Hasta luego!")
            break
        except EOFError:
            print()
            c_print("green", "\n  Hasta luego!")
            break

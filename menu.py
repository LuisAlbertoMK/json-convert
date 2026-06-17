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


# -- Detect terminal color support --
_HAS_ANSI = (hasattr(sys.stdout, "isatty") and sys.stdout.isatty()) or bool(os.environ.get("TERM"))
if sys.platform == "win32":
    _HAS_ANSI = True  # Windows 10+ supports ANSI natively


def _c(code, text):
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


def c_print(code, text, end="\n"):
    print(_c(code, text), end=end)


def separator(char="=", width=55):
    print(_c("dim", char * width))


def header(title):
    print()
    separator()
    c_print("bold", "  " + title)
    separator()


def ask_int(prompt, min_v, max_v, default=None):
    """Pide un numero, con default opcional."""
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


def confirm(prompt, default=True):
    """Confirmacion SI/NO."""
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


def run_step(cmd, label, cwd=None, timeout=None):
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


def run_ps1(script, args="", timeout=None):
    """Ejecuta un script PowerShell."""
    cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-NoProfile",
           "-File", os.path.join(BASE_DIR, script)]
    if args:
        cmd += args.split()
    return run_step(cmd, "PowerShell: " + script + " " + args, timeout=timeout)


def get_markets_from_urls() -> list[str]:
    """Lee los mercados únicos desde urls.json."""
    urls_path = os.path.join(BASE_DIR, "urls.json")
    if not os.path.exists(urls_path):
        return []
    try:
        with open(urls_path, encoding="utf-8") as f:
            entries = json.load(f)
        markets = sorted(set(e.get("market") for e in entries if e.get("market")))
        return markets
    except Exception:
        return []


def detect_markets():
    """Detecta directorios de mercado con archivos de auditoria.

    Busca en subdirectorios (PR/, MX/, etc.) y tambien en la raiz del proyecto.
    """
    results = []
    base = Path(BASE_DIR)

    # Buscar en subdirectorios de mercado
    for d in base.iterdir():
        if not d.is_dir() or d.name.startswith("."):
            continue
        for candidate in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
            if (d / candidate).exists():
                results.append((d.name.upper(), str(d / candidate)))
                break

    # Buscar tambien en la raiz del proyecto
    for candidate in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
        fp = base / candidate
        if fp.exists() and not any(r[1] == str(fp) for r in results):
            results.append(("RAIZ", str(fp)))

    return results


def get_project_status():
    """Retorna estado resumido del proyecto."""
    base = Path(BASE_DIR)
    py_files = list(base.glob("*.py"))
    markets = detect_markets()

    # Ultima auditoria: buscar sheets en historial.xlsx
    last_audit = "-"
    try:
        import openpyxl
        for m, hpath in markets:
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


def choose_market(source="detect"):
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

def op_auditar():
    """Opcion 2: Solo auditoria (extract_browser) por mercado."""
    header("AUDITORIA - extract_browser.py")

    urls_path = os.path.join(BASE_DIR, "urls.json")
    if not os.path.exists(urls_path):
        print(_c("yellow", "  No se encuentra urls.json"))
        if confirm("  Generarlo desde RevisionManual.xlsx?"):
            run_step([sys.executable, "_gen_urls.py"], "Generando urls.json")
        else:
            print(_c("yellow", "  Cancelado. Necesitas urls.json para auditar."))
            return

    market, _ = choose_market(source="urls")
    if market is None:
        print(_c("yellow", "  No hay mercados en urls.json. Generalo primero."))
        return

    if market == ALL_MARKETS:
        markets = get_markets_from_urls()
        for m in markets:
            cmd = [sys.executable, "extract_browser.py", "--urls", "urls.json",
                   "--market", m, "--split-aa", "--progress"]
            run_step(cmd, "Auditando " + m + "...", timeout=600)
    else:
        cmd = [sys.executable, "extract_browser.py", "--urls", "urls.json",
               "--market", market, "--split-aa", "--progress"]
        run_step(cmd, "Auditando " + market + "...", timeout=600)

    print(_c("green", "\n  [OK] Auditoria finalizada."))


def op_postprocesar():
    """Opcion 3: Solo post-procesar (extract_aa)."""
    header("POST-PROCESO - extract_aa.py")

    market, hpath = choose_market(source="detect")
    if market is None:
        print(_c("yellow", "  No se encontraron archivos de auditoria para procesar."))
        print("  Ejecuta primero una auditoria (opcion 1 o 2).")
        return

    if market == ALL_MARKETS:
        markets = detect_markets()
        for m, hpath in markets:
            _run_extract_aa(hpath)
            _run_extract_aa_companions(hpath)
    else:
        _run_extract_aa(hpath)
        _run_extract_aa_companions(hpath)

    print(_c("green", "\n  [OK] Post-proceso finalizado."))


def _run_extract_aa(hpath):
    """Ejecuta extract_aa.py sobre un archivo historial."""
    run_step([sys.executable, "extract_aa.py", "--input", hpath],
             "Procesando: " + hpath, timeout=120)


def _run_extract_aa_companions(hpath):
    """Ejecuta extract_aa.py sobre con_aa.xlsx y sin_aa.xlsx del mismo dir."""
    base_dir = os.path.dirname(hpath)
    for fname in ["con_aa.xlsx", "sin_aa.xlsx"]:
        fp = os.path.join(base_dir, fname)
        if os.path.exists(fp):
            run_step([sys.executable, "extract_aa.py", "--input", fp],
                     "Procesando: " + fp, timeout=120)


def op_reporte():
    """Opcion 4: Generar reporte de fallos."""
    header("REPORTE DE FALLOS - audit_report.py")

    markets = detect_markets()
    if not markets:
        print(_c("yellow", "  No se encontraron archivos de auditoria."))
        print("  Ejecuta primero una auditoria (opcion 1 o 2).")
        return

    code = run_step([sys.executable, "audit_report.py"],
                    "Generando reporte (global + por mercado)...", timeout=60)

    if code == 0:
        print(_c("green", "\n  [OK] Reporte global generado: reporte_auditoria.xlsx"))
        markets = detect_markets()
        for m, _ in markets:
            if m != "RAIZ":
                print(_c("green", f"       {m}/reporte-auditoria.xlsx"))
        if confirm("  Abrir el global?"):
            open_file("reporte_auditoria.xlsx")
    else:
        print(_c("red", "\n  [ERROR] Fallo la generacion del reporte."))


def op_catalogo():
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
    mapping_path = os.path.join(BASE_DIR, "url-mapping.json")
    if not os.path.exists(mapping_path):
        print(_c("yellow", "\n  No se encuentra url-mapping.json"))
        if confirm("  Generar template desde RevisionManual.xlsx?", default=True):
            input_path = os.path.join(BASE_DIR, "RevisionManual.xlsx")
            if os.path.exists(input_path):
                run_step(
                    [sys.executable, "generate_migration_catalog.py",
                     "--gen-template", "--input", input_path,
                     "--mapping", "url-mapping.json"],
                    "Generando template url-mapping.json...")
                print(_c("yellow", "\n  [!] EDITAR url-mapping.json con production_url, aem_path y page_key"))
                print("      Despues ejecutar esta opcion nuevamente.")
            else:
                print(_c("red", "  No se encuentra RevisionManual.xlsx"))
        return

    # Verificar expected.json
    expected_path = os.path.join(BASE_DIR, "expected.json")
    if not os.path.exists(expected_path):
        print(_c("red", "  No se encuentra expected.json"))
        print("  Debe existir con las reglas del estandar US para cada mercado.")
        return

    # Generar catalogo
    output_name = "catalogo-migracion.xlsx"
    output_path = os.path.join(BASE_DIR, market_name, output_name)

    # Mostrar resumen antes de ejecutar
    print()
    print(_c("cyan", "  Resumen:"))
    print("    Mercado:      " + market_name)
    print("    Historial:    " + os.path.relpath(historial_path, BASE_DIR))
    print("    Mapping:      url-mapping.json")
    print("    Expected:     expected.json")
    print("    Output:       " + market_name + "/" + output_name)
    print()

    if confirm("  Generar catalogo?", default=True):
        run_step(
            [sys.executable, "generate_migration_catalog.py",
             "--historial", historial_path,
             "--mapping", "url-mapping.json",
             "--expected", "expected.json",
             "--market", market_name],
            "Generando catalogo de migracion...", timeout=60)

        if os.path.exists(output_path):
            print(_c("green", "\n  [OK] Catalogo generado: " + market_name + "/" + output_name))
            if confirm("  Abrirlo ahora?", default=True):
                open_file(output_path)


def op_limpieza():
    """Opcion 5: Limpieza (run.ps1 -SkipTests)."""
    header("LIMPIEZA - run.ps1")
    run_ps1("run.ps1", "-SkipTests", timeout=60)
    print(_c("green", "\n  [OK] Limpieza finalizada."))


def op_prune():
    """Opcion 8: Limpiar columnas muertas de Excel."""
    header("LIMPIAR COLUMNAS - prune_excel_columns.py")

    market, _ = choose_market(source="detect")
    if market is None:
        print(_c("yellow", "  No se encontraron archivos Excel."))
        return

    if market == ALL_MARKETS:
        cmd = [sys.executable, "prune_excel_columns.py"]
        label = "Todos los mercados"
    else:
        cmd = [sys.executable, "prune_excel_columns.py", "--dir", market]
        label = "Mercado: " + market

    code = run_step(cmd, label, timeout=30)
    if code == 0:
        print(_c("green", "\n  [OK] Columnas limpiadas."))
    else:
        print(_c("red", "\n  [ERROR] Fallo la limpieza."))


def op_tests():
    """Opcion 9: Ejecutar tests."""
    header("TESTS - pytest")
    rc = run_step([sys.executable, "-m", "pytest", "--tb=long", "-q"],
                  "pytest completo", timeout=120)
    if rc == 0:
        print(_c("green", "\n  [OK] Todos los tests pasaron."))
    else:
        print(_c("red", "\n  [ERROR] Fallaron tests (codigo: " + str(rc) + ")"))


def op_ver_resultados():
    """Opcion 6: Abrir Excel."""
    header("RESULTADOS")

    targets = []
    report = os.path.join(BASE_DIR, "reporte_auditoria.xlsx")
    if os.path.exists(report):
        targets.append(report)

    for m, hpath in detect_markets():
        base = os.path.dirname(hpath)
        for fname in ["historial.xlsx", "con_aa.xlsx", "sin_aa.xlsx"]:
            fp = os.path.join(base, fname)
            if os.path.exists(fp) and fp not in targets:
                targets.append(fp)

    for fname in ["historial.xlsx"]:
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


def op_todo_en_uno(target_market=None):
    """Opcion 1: Pipeline completo.
    
    Args:
        target_market: None → preguntar (interactivo) o ALL (--run)
                       "PR" → solo ese mercado
                       ALL_MARKETS → todos
    """
    separator("=", 55)
    c_print("bold", "  [!] PIPELINE COMPLETO")
    separator("=", 55)
    print()
    print("  Inicio: " + datetime.now().strftime("%H:%M:%S"))
    print()

    # Elegir mercado si no se especificó
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

    market_label = "todos los mercados" if target_market == ALL_MARKETS else target_market
    print(_c("dim", f"  Mercado(s): {market_label}"))
    print()

    results = []

    # Paso 1: Correr tests
    print()
    print(_c("cyan", "  [1/8] Ejecutando tests..."))
    rc = run_step([sys.executable, "-m", "pytest", "--tb=short", "-q"],
                  "pytest --tb=short -q", timeout=120)
    if rc != 0:
        print(_c("yellow", "    ⚠ Tests fallaron (codigo: " + str(rc) + ")"))
        if not confirm("    Continuar pipeline igual?", default=False):
            print(_c("red", "    Pipeline abortado por fallo en tests."))
            results.append(("Tests", rc))
            # Saltar al resumen
            print()
            _pipeline_summary(results)
            return
    results.append(("Tests", rc))

    # Paso 2: Verificar urls.json
    print()
    print(_c("cyan", "  [2/8] Verificando entorno..."))
    urls_path = os.path.join(BASE_DIR, "urls.json")
    has_urls = os.path.exists(urls_path)
    if not has_urls:
        print(_c("yellow", "    No se encuentra urls.json"))
        rp = os.path.join(BASE_DIR, "RevisionManual.xlsx")
        if os.path.exists(rp):
            if confirm("    Generarlo desde RevisionManual.xlsx?", default=True):
                rc = run_step([sys.executable, "_gen_urls.py"], "Generando urls.json")
                results.append(("Generar urls.json", rc))
                has_urls = (rc == 0)
            else:
                print(_c("yellow", "    Saltando generacion de urls.json"))
        else:
            print(_c("yellow", "    No hay RevisionManual.xlsx para generar urls.json"))
    else:
        print(_c("green", "    urls.json OK"))
        results.append(("Verificar entorno", 0))

    # Paso 3: Auditoria (por mercado)
    print()
    print(_c("cyan", "  [3/8] Auditoria (extract_browser)..."))
    if has_urls:
        audit_ok = True
        for m in markets_to_run:
            rc = run_step(
                [sys.executable, "extract_browser.py", "--urls", "urls.json",
                 "--market", m, "--split-aa", "--progress"],
                "Auditando " + m + "...", timeout=600)
            results.append(("Auditoria " + m, rc))
            if rc != 0:
                audit_ok = False
    else:
        rc = run_step([sys.executable, "extract_browser.py"],
                      "Ejecutando (modo Excel plano)...", timeout=600)
        results.append(("Auditoria", rc))

    # Paso 4: Post-procesar
    print()
    print(_c("cyan", "  [4/8] Post-procesando (extract_aa)..."))
    all_markets = detect_markets()
    # Filtrar solo los markets objetivo
    post_markets = [(m, p) for m, p in all_markets if m in markets_to_run or target_market == ALL_MARKETS]
    processed_any = False
    for m, hpath in post_markets:
        _run_extract_aa(hpath)
        results.append(("Post-proceso " + m, 0))
        processed_any = True
        _run_extract_aa_companions(hpath)

    if not processed_any:
        print(_c("yellow", "    No hay archivos de auditoria para post-procesar."))
        results.append(("Post-proceso", -1))

    # Paso 5: Limpiar columnas muertas
    print()
    print(_c("cyan", "  [5/8] Limpiando columnas inutiles..."))
    rc = run_step([sys.executable, "prune_excel_columns.py"],
                  "prune_excel_columns.py", timeout=30)
    results.append(("Prune columnas", rc))

    # Paso 6: Reporte de fallos (global + por mercado)
    print()
    print(_c("cyan", "  [6/8] Generando reporte de fallos..."))
    rc = run_step([sys.executable, "audit_report.py"],
                  "audit_report.py", timeout=60)
    results.append(("Reporte de fallos", rc))

    # Paso 7: Catalogo de migracion
    print()
    print(_c("cyan", "  [7/8] Catalogo de migracion..."))
    mapping_path = os.path.join(BASE_DIR, "url-mapping.json")
    if not os.path.exists(mapping_path):
        rp = os.path.join(BASE_DIR, "RevisionManual.xlsx")
        if os.path.exists(rp):
            print(_c("yellow", "    No se encuentra url-mapping.json"))
            if confirm("    Generar template desde RevisionManual.xlsx?", default=True):
                rc = run_step(
                    [sys.executable, "generate_migration_catalog.py",
                     "--gen-template", "--input", rp, "--mapping", "url-mapping.json"],
                    "Generando url-mapping.json...")
                if rc == 0:
                    print(_c("yellow", "    [!] EDITAR url-mapping.json con production_url, aem_path y page_key"))
                results.append(("Template url-mapping", rc))
            else:
                results.append(("Template url-mapping", -1))
        else:
            print(_c("yellow", "    No hay RevisionManual.xlsx ni url-mapping.json - saltando catalogo"))
            results.append(("Catalogo migracion", -1))

    if os.path.exists(mapping_path) and os.path.exists(os.path.join(BASE_DIR, "expected.json")):
        for m, hpath in post_markets:
            rc = run_step(
                [sys.executable, "generate_migration_catalog.py",
                 "--historial", hpath,
                 "--mapping", "url-mapping.json",
                 "--expected", "expected.json",
                 "--market", m],
                "Catalogo " + m + "...", timeout=60)
            results.append(("Catalogo " + m, rc))

    # Paso 8: Resultados
    print()
    print(_c("cyan", "  [8/8] Resultados..."))
    report_path = os.path.join(BASE_DIR, "reporte_auditoria.xlsx")
    if os.path.exists(report_path):
        if confirm("  Abrir el reporte?", default=True):
            open_file(report_path)
    elif post_markets:
        for m, hpath in post_markets:
            if confirm("  Abrir " + m + "/historial.xlsx?", default=True):
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


def open_file(path):
    """Abre un archivo con la aplicacion asociada."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as e:
        print(_c("red", "  Error al abrir: " + str(e)))


def show_menu():
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
    print("  " + _c("bold", "3") + ") " + _c("yellow", "*") + "  Solo post-procesar   extract_aa.py")
    print("  " + _c("bold", "4") + ") " + _c("magenta", "[R]") + "  Solo reporte fallos  audit_report.py")
    print("  " + _c("bold", "5") + ") " + _c("yellow", "[C]") + "  Solo limpieza        run.ps1")
    print("  " + _c("bold", "6") + ") " + _c("blue", "[V]") + "  Ver resultados       abrir Excel")
    print("  " + _c("bold", "7") + ") " + _c("magenta", "[M]") + "  Catalogo migracion   generate_migration_catalog.py")
    print("  " + _c("bold", "8") + ") " + _c("yellow", "[P]") + "  Limpiar columnas     prune_excel_columns.py")
    print("  " + _c("bold", "9") + ") " + _c("green", "[T]") + "  Ejecutar tests       pytest")
    print()
    separator("-", 55)
    print("  " + _c("dim", "0") + ") " + _c("dim", "x  Salir"))
    separator("=", 55)

    return ask_int("  Opcion [0-9]: ", 0, 9)


def run_option(opt, non_interactive=False):
    """Ejecuta la opcion. Retorna False si hay que salir."""
    start = time.time()

    if opt == 1:
        op_todo_en_uno(target_market=ALL_MARKETS if non_interactive else None)
    elif opt == 2:
        op_auditar()
    elif opt == 3:
        op_postprocesar()
    elif opt == 4:
        op_reporte()
    elif opt == 5:
        op_limpieza()
    elif opt == 6:
        op_ver_resultados()
    elif opt == 7:
        op_catalogo()
    elif opt == 8:
        op_prune()
    elif opt == 9:
        op_tests()
    elif opt == 0:
        print()
        c_print("green", "  Hasta luego!")
        return False

    elapsed = time.time() - start
    if elapsed > 2:
        print("  (" + _c("dim", str(int(elapsed)) + "s") + ")")
    print()
    input(_c("dim", "  Presiona Enter para volver al menu..."))
    return True


# ============================================
#  MAIN
# ============================================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Panel de control json-convert")
    parser.add_argument("--run", type=str,
                        help="Ejecutar opcion directa: numero 1-9, 0, o 'auto'")
    args = parser.parse_args()

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # Verificar proyecto
    is_project = any(
        os.path.exists(os.path.join(BASE_DIR, f))
        for f in ["extract_browser.py", "extract_aa.py", "run.ps1"]
    )
    if not is_project:
        print(_c("red", "[ERROR] Este script debe ejecutarse desde la raiz del proyecto json-convert"))
        print("        No se encontraron extract_browser.py, extract_aa.py o run.ps1")
        sys.exit(1)

    if sys.version_info < (3, 9):
        print(_c("red", "[ERROR] Python 3.9+ requerido"))
        sys.exit(1)

    # Modo directo --run
    if args.run:
        opt_map = {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8,
                   "0": 0, "auto": 1}
        opt = opt_map.get(args.run, -1)
        if opt < 0:
            print(_c("red", "[ERROR] Opcion invalida: " + args.run))
            print("  Valores validos: 1-8, 0, 'auto'")
            sys.exit(1)

        # Modo no-interactivo: auto-confirm sin preguntar mercado
        _old_input = input
        input = lambda prompt="": "y"

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

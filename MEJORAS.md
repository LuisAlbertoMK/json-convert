# Mejoras del Proyecto json-convert

Registro de todas las mejoras solicitadas, implementadas, y su solución aplicada.

## Convención del proyecto: JSON-first

Todo dato estructurado en el proyecto se representa en **JSON**. No se usan formatos propietarios, XML, INI, YAML, ni schemas ad-hoc. Las razones:

- **`urls.json`** — lista de URLs con metadatos (nombre, mercado)
- **`audit.json`** — configuración del extractor (workers, timeouts, proxy)
- **Adobe Analytics beacons** — los datos extraídos de las páginas son JSON
- **Data layer** (`window.digitaldata`) — es JSON nativo del navegador
- **Columnas del Excel** — almacenan JSON pretty-printed como texto plano
- **Intercambio entre scripts** — el pipeline se comunica via archivos JSON

Ejemplo de `urls.json`:
```json
[
  {
    "url": "https://wwwac.preview.es.brandpr.ford.com/.../donativos-ambientales.html",
    "nombre": "Puerto Rico Español",
    "market": "PR"
  },
  {
    "url": "https://wwwac.preview.brandpr.ford.com/.../fuel-saving.html",
    "nombre": "Fuel Saving",
    "market": "PR"
  }
]
```

Ejemplo de `audit.json`:
```json
{
  "workers": 3,
  "retry": 2,
  "timeout": 35000,
  "wait_after": 4,
  "headed": false
}
```

Ejemplo del data layer (`window.digitaldata`) extraído:
```json
{
  "page": {
    "pageName": "fpr:corporate:ford environmental grants",
    "siteSection": "home",
    "hierarchy": "home"
  }
}
```
---

## 1. Script de análisis y limpieza one-click

**Solicitud**: Poder analizar y limpiar el proyecto con un solo comando, sin ejecutar comandos manuales.

**Solución**:
- **`run.ps1`**: Script PowerShell con 3 fases:
  1. Análisis (versión Python, dependencias, tests, git status, tamaño del proyecto)
  2. Limpieza (`__pycache__`, `.pyc`/`.pyo`, `.log`, `.history/`)
  3. Reporte final con scores y espacio liberado
- **`run.bat`**: Wrapper batch para doble clic, compatible con Windows

**Bug encontrado y corregido**: `Get-ChildItem -Include "*.pyc" -LiteralPath $ROOT -Recurse` en PowerShell 5.1 **ignora `-Include`** cuando el path no tiene wildcard, devolviendo TODOS los archivos. Se reemplazó por `Where-Object { $_.Extension -match '\.py[co]$' }`.

**Archivos**: `run.ps1`, `run.bat`

---

## 2. Panel de control interactivo (menú TUI)

**Solicitud**: Tener un menú centralizado para ejecutar cualquier parte del pipeline sin recordar comandos.

**Solución**:
- **`menu.py`**: Menú zero-dependencias (solo stdlib) con 7 opciones:
  - **1** `[!] TODO EN UNO` — Pipeline completo
  - **2** `> Solo auditoría` — Solo `extract_browser.py`
  - **3** `* Solo post-procesar` — Solo `extract_aa.py`
  - **4** `[R] Solo reporte` — Solo `audit_report.py`
  - **5** `[C] Solo limpieza` — `run.ps1`
  - **6** `[V] Ver resultados` — Abrir Excel
- Modo headless: `python menu.py --run auto` para CI/no-interactive
- Pipeline completo: verifica entorno → audita → post-procesa → reporta → muestra resumen
- Manejo de errores: cada paso falla independientemente, no rompe el flujo

**Formato de datos en el Excel**: cada columna almacena datos en JSON pretty-printed:

```json
// Columna D (digitaldata extraído)
{ "page": { "pageName": "fpr:corporate:ford environmental grants", "siteSection": "home" } }

// Columna E (AA beacon original parseado)
{ "solution": "analytics", "page": { "title": "Donativos Ambientales", "url": "https://..." } }

// Columna F (AA post-procesado - extract_aa.py)
{ "pageName": "fpr:corporate:ford environmental grants", "channel": "home" }

// Columna G (metadata del error)
{ "error": "no AA data captured", "code": "NO_AA_DATA", "url": "https://..." }
```

**Bug de encoding**: Windows cp1252 no soporta Unicode box-drawing. Solución: solo ASCII-safe (`[!]`, `[R]`, `[C]`, `[V]`, `=`, `-`, `+`, `|`, `x`).

**Archivos**: `menu.py`

---

## 3. Bugfix: import openpyxl redundante en extract_browser.py

**Problema**: `extract_browser.py` tenía `import openpyxl` global (línea 123) y otro `import openpyxl` local dentro de una función (línea 892). Python lo trataba como variable local sin asignación completa, causando `UnboundLocalError: cannot access local variable 'openpyxl' where it is not associated with a value`.

**Solución**: Eliminar el `import openpyxl` local redundante.

**Archivos**: `extract_browser.py`

---

## 4. Bugfix: f-strings partidos en audit_report.py

**Problema**: 3 f-strings partidos en dos líneas causaban `SyntaxError: unterminated f-string literal`:
```python
# Incorrecto (antes)
text = (f"...{valor}"
        f"...{otro}")
```

**Solución**: Unificar cada f-string en una sola línea con `\n` explícito donde sea necesario.

**Archivos**: `audit_report.py`

---

## 5. Bugfix: detect_markets() no encontraba archivos en la raíz

**Problema**: `detect_markets()` en `menu.py` solo buscaba `historial.xlsx` dentro de subdirectorios (PR/, MX/, etc.). Cuando la auditoría se ejecutaba sin `--market`, el archivo quedaba en la raíz y el paso de post-procesado nunca lo encontraba.

**Solución**: Agregar búsqueda también en la raíz del proyecto como fallback, etiquetando como mercado "RAIZ".

**Archivos**: `menu.py` — función `detect_markets()`

---

## 6. URLs truncadas corregidas (causa raíz de 9/14 errores)

**Problema**: 3 URLs en `urls.json` estaban truncadas con caracter `…` (Unicode ellipsis) en lugar de `.html` completo:

```json
// Antes (truncado)
{ "url": "https://.../duenos/tips/driving-precautio…" }

// Después (corregido)
{ "url": "https://.../duenos/tips/driving-precautions.html" }
```

| Truncado | Corregido |
|----------|-----------|
| `driving-precautio…` | `driving-precautions.html` |
| `collision-repair.…` | `collision-repair.html` |
| `driving-precautions.…` | `driving-precautions.html` |

Esto causaba que Playwright no pudiera navegar, generando **9 errores de 14 URLs** por corrida. Además, en `RevisionManual.xlsx` las mismas 3 URLs estaban truncadas, más 9 trailing non-breaking spaces (`\xa0`) adicionales.

**Solución**:
- Reemplazar las 3 URLs truncadas en `urls.json`
- Reemplazar las 3 URLs truncadas + limpiar 9 `\xa0` en `RevisionManual.xlsx`
- Eliminar `historial.xlsx` y `reporte_auditoria.xlsx` (datos obsoletos con errores)

**Archivos**: `urls.json`, `RevisionManual.xlsx`

---

## 7. Auto-generar con_aa.xlsx y sin_aa.xlsx (--split-aa)

**Problema**: El flag `--split-aa` de `extract_browser.py` nunca se activaba desde el menú. `con_aa.xlsx` (URLs con Adobe Analytics) y `sin_aa.xlsx` (URLs sin AA) nunca se generaban.

**Solución**: Agregar `--split-aa` en ambos lugares de `menu.py`:
- `op_todo_en_uno()` — Pipeline completo
- `op_auditar()` — Solo auditoría

**Cómo funciona**: `extract_browser.py` clasifica cada URL según si se detectó un beacon AA en la página:

```json
// URL con AA → va a con_aa.xlsx
{ "url": ".../donativos-ambientales.html", "aa_parsed": { "pageName": "fpr:corporate:ford environmental grants", "reportSuite": "fordglobal" } }

// URL sin AA → va a sin_aa.xlsx
{ "url": ".../fuel-saving.html", "aa_parsed": null, "error": "no AA data captured" }
```

**Archivos**: `menu.py`

---

## 8. Detección automática de mercado desde la URL

**Solicitud**: El campo `market` debe inferirse de la URL (ej: `es_pr`, `en_pr`), no asignarse manualmente.

**Solución**:
- **`_gen_urls.py`**: Nueva función `detect_market_from_url()` que usa regex `/([a-z]{2})_([a-z]{2})/` en el path de la URL

```json
// Antes (--default-market manual)
{ "url": ".../es_pr/index/legal.html", "market": "PR" }

// Automático (detectado desde URL)
{ "url": ".../es_pr/index/legal.html", "market": "PR" }
{ "url": ".../en_pr/index/legal.html", "market": "PR" }
{ "url": ".../es_mx/index/legal.html", "market": "MX" }
```

  - `/es_pr/...` → `"market": "PR"`
  - `/en_pr/...` → `"market": "PR"`
  - `/es_mx/...` → `"market": "MX"`
- Se eliminó la dependencia de `--default-market` (aunque queda como override manual)
- `urls.json` regenerado con los markets detectados

**Archivos**: `_gen_urls.py`, `urls.json`

---

## 9. Estructura de directorios por mercado

**Solicitud**: Los archivos deben organizarse en carpetas por mercado (`PR/`, `MX/`, etc.), no todo en la raíz.

**Problema**: Aunque `extract_browser.py` ya soportaba `--market` (crea directorio y guarda dentro), el menú nunca lo usaba. Todo se guardaba en la raíz.

**Solución**:
- `menu.py` ahora detecta los markets únicos desde `urls.json` con `get_markets_from_urls()`
- Ejecuta `extract_browser.py --market <X>` por cada mercado detectado
- El output se guarda en `PR/historial.xlsx`, `PR/con_aa.xlsx`, `PR/sin_aa.xlsx`

**Estructura resultante**:
```
PR/
  historial.xlsx
  con_aa.xlsx
  sin_aa.xlsx
```

**Archivos**: `menu.py`

---

## 10. Auto-generar nombre de página desde la URL

**Problema**: 8 de 14 URLs no tenían `nombre` en `RevisionManual.xlsx`. `extract_browser.py` usaba la URL completa como fallback (`entry.get("nombre", entry.get("url", ""))`), mostrando la URL cruda como título.

```json
// Antes (sin nombre → URL completa)
{ "url": "https://.../driving-precautions.html" }
// → en Excel columna A: "https://wwwac.preview.../driving-precautions.html"

// Después (nombre generado desde URL)
{ "url": "https://.../driving-precautions.html", "nombre": "Driving Precautions" }
// → en Excel columna A: "Driving Precautions"
```

**Solución**:
- **`_gen_urls.py`**: Nueva función `make_name_from_url()` que extrae el filename del path, remueve `.html`, reemplaza guiones por espacios y capitaliza

```json
{ "url": ".../driving-precautions.html" }  → "nombre": "Driving Precautions"
{ "url": ".../fuel-saving.html" }          → "nombre": "Fuel Saving"
{ "url": ".../privacy-policy.html" }       → "nombre": "Privacy Policy"
{ "url": ".../collision-repair.html" }     → "nombre": "Collision Repair"
{ "url": ".../donativos-ambientales.html" }→ "nombre": "Donativos Ambientales"
```

- Las URLs sin path (dominio desnudo) se dejan sin `nombre`

**Archivos**: `_gen_urls.py`, `urls.json`

---

## 11. Análisis de velocidad de recolección — ✅ IMPLEMENTADO

**Solicitud**: Evaluar si se puede acelerar la auditoría de URLs.

**Análisis original**: 3 cuellos de botella identificados en `extract_browser.py`:

| # | Problema | Solución aplicada | Dónde |
|---|----------|-------------------|-------|
| 1 | `wait_until="load"` | Estrategia progresiva: `["domcontentloaded", "commit", "load"]` | `browser.py:process_url()` |
| 2 | `asyncio.sleep(4)` fijo | Smart wait: `wait_for_function("window.digitalData !== undefined")` con fallback 1s | `browser.py:process_url()` |
| 3 | Extracciones secuenciales | `asyncio.gather(title, digitaldata, s_object)` paralelizado | `browser.py:process_url()` |

**Ajustes adicionales (2026-06-22)**:
- `--wait-after` default reducido de **4 → 2** (el smart wait timeout real era `min(4000, 3000)=3000ms`, ahora `min(2000, 2000)=2000ms`)
- Ahorro estimado: ~1-2s por URL en páginas sin digitalData (314 URLs → ~5-10 min menos)

**Estado actual del pipeline**:
- Workers: 3 (default, configurable vía `--workers`)
- `wait_until`: `domcontentloaded` en primer intento
- Smart wait: hasta 2s por `digitalData`, +1s fallback
- Extracciones: paralelizadas con `asyncio.gather`
- Caché de navegación: persiste 24h (evita re-navegar URLs ya auditadas)

**Próximo paso opcional**: Evaluar Firefox como default browser (Gecko no sufre ERR_ABORTED, documentado en `docs/fallback-err-aborted.md`).

---

## Resumen de archivos modificados

| Archivo | Mejoras aplicadas |
|---------|-------------------|
| `run.ps1` | Script análisis + limpieza one-click, fix PS 5.1 `-Include` bug |
| `run.bat` | Wrapper doble clic |
| `menu.py` | Menú TUI, pipeline completo, `--split-aa`, `--market` por mercado, fix `detect_markets()` |
| `extract_browser.py` | Fix `UnboundLocalError` (import redundante). `--wait-after` default 4→2 |
| `audit_report.py` | Fix 3 f-strings partidos |
| `_gen_urls.py` | Auto-detect mercado desde URL + auto-generar nombre desde filename |
| `urls.json` | Regenerado: markets detectados, nombres auto-generados, URLs truncadas corregidas |
| `RevisionManual.xlsx` | URLs truncadas corregidas, trailing spaces limpiados |
| `test_gen_urls.py` | Test actualizado para nuevo comportamiento de auto-nombre |
| `json_convert/browser.py` | Smart wait (digitalData), progressive `wait_until`, `asyncio.gather`, `wait_after` default 4→2, `_backoff_delay` con jitter |
| `src/extract_browser.py` | `--wait-after` default 4→2, `--browser` flag (chromium/firefox), `UrlCache` integrado |
| `docs/fallback-err-aborted.md` | Documentación de estrategia ERR_ABORTED + Firefox testing |

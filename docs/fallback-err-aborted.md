# Fallback para ERR_ABORTED en Playwright

## Objetivo

El script `extract_browser.py` navega URLs de Ford (México, Puerto Rico, etc.) usando Playwright + Chrome para extraer:

- **Adobe Analytics beacons** (requests a `*.sc.omtrdc.net`, `*.2o7.net`, etc.)
- **digitalData** (`window.digitalData`, `window.s`. etc.)
- **Título y metadatos** de cada página

El objetivo de este documento es documentar por qué ciertas URLs fallan con `net::ERR_ABORTED` y qué estrategias se implementaron para recuperar los datos igualmente.

---

## El Problema: `net::ERR_ABORTED`

### Síntoma

En URLs del catálogo Ford (`all-vehicles`), Playwright lanza:

```
net::ERR_ABORTED at https://.../all-vehicles
```

Esto **no** es un bloqueo de Akamai WAF ni un timeout de red. Es la **propia página** la que aborta la navegación — típicamente por:

1. **SPA client-side routing**: la página inicia una navegación, detecta que es el mismo "espacio" y la cancela para hacer routing interno.
2. **`window.stop()` o `document.open()`**: scripts analytics o de terceros que llaman a estas APIs durante la carga inicial.
3. **Service Worker interference**: un SW intercepta el fetch y aborta la solicitud original.
4. **Ciclo de redirección/resource**: un recurso (CSS/JS) falla y la página decide abortar todo.

### Consecuencia

`page.goto()` lanza excepción → el flujo normal de extracción se salta → no se obtiene `digitalData`, no se capturan beacons → la URL queda marcada como error.

---

## Solución Implementada: Fallback en 2 Capas

### Capa 1 — Extracción del DOM Parcial (intra-retry)

Cuando `page.goto()` falla con `ERR_ABORTED`, la página **puede tener el DOM parcialmente cargado**. El navegador abortó el evento de navegación, pero el JavaScript que setea `window.digitalData` ya se ejecutó.

```python
if "ERR_ABORTED" in str(e):
    last_dd = await extract_digital_data(page)
    if last_dd:
        navigation_error = ""  # recuperado!
        break
```

**Costo**: ~0 (ya estamos en el handler de error, solo agregamos un intento de extracción)

### Capa 2 — Fetch + setContent (post-retry loop)

Si todos los reintentos fallaron y no hay `digitalData`, traemos el HTML vía `urllib.request` (stdlib, **cero dependencias nuevas**) y lo inyectamos con `page.setContent()`:

```python
html = await _fetch_html_via_http(url, timeout_ms)
html = html.replace("<head>", '<head><base href="{url}">', 1)
await page.goto("about:blank")
await page.setContent(html, wait_until="domcontentloaded")
last_dd = await extract_digital_data(page)
```

**Por qué funciona**:
- `ERR_ABORTED` es un abort de **navegación** — `setContent()` no navega, inyecta DOM directamente
- `urllib` tiene un fingerprint de red completamente distinto a Playwright/CDP
- Las URLs relativas se resuelven correctamente gracias al `<base href="...">`
- Sin dependencias extra: `urllib.request` viene con Python

**Costo**: 1 request HTTP extra + 1 `setContent()` por URL fallida (~2-5s adicionales)

---

## Tradeoffs y Alternativas

| Enfoque | Esfuerzo | Confianza | Por qué NO se eligió (primario) |
|---------|----------|-----------|----------------------------------|
| **✅ Fetch + setContent** | 1 día | 🟢 Alta | Ganador: bypassea el abort por completo |
| Firefox (`pw.firefox`) | 30 min | 🟢 Alta | Pendiente de probar — diferente engine, no sufre ERR_ABORTED de la misma forma |
| `wait_until="networkidle"` | 5 min | 🔶 Media | Si la página hace polling infinito, timeout igual |
| `page.route()` bloquear script | 2-4 hrs | 🔶 Media | Requiere identificar QUÉ script causa el abort |
| `connectOverCDP` + Chrome real | 5 min | 🟢 Alta | Solo manual, no sirve para CI/automatizado |
| Stealth plugins | 15 min | 🔶 Media | En 2026 ya los detectan |
| Selenium | 2-3 días | 🔶 Media | Cambiar de herramienta es costo muy alto |

---

## Resultado: Firefox TEST ✅

**Firefox NO sufre ERR_ABORTED en Ford all-vehicles.**

```
>>> https://www.ford.com.pr/en/all-vehicles/
   [OK]  | keys: version, _il, _in, _c, F...
   title: All Vehicles | Showroom | Ford Puerto Rico
```

Firefox (Gecko engine) no tiene el mismo bug de abort de navegación que Chromium. La URL problemática se cargó sin errores y digitalData se extrajo correctamente.

### Implementado: `--browser {chromium|firefox}`

Flag CLI agregado a `extract_browser.py`:

```bash
# Default: Chromium (mismo comportamiento que antes)
python extract_browser.py --urls urls.json

# Firefox: evita ERR_ABORTED en URLs problemáticas
python extract_browser.py --urls urls.json --browser firefox
```

Firefox se lanza con:
- User-Agent Firefox 150.0 (`Mozilla/5.0 ... Firefox/150.0`)
- Anti-detección minimalista (solo `navigator.webdriver = undefined`)
- Sin `--disable-blink-features=AutomationControlled` (no existe en Gecko)

Chromium mantiene su lógica original: Chrome real → fallback a Chromium bundled.

## Próximos Pasos

1. ✅ ~~Probar Firefox~~ — **Funciona. No sufre ERR_ABORTED.**
2. ✅ ~~Flag `--browser`~~ — **Implementado.**
3. Evaluar si Firefox como default es viable (performance, estabilidad, memoria)

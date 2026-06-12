# Resumen de Errores — Extracción AA

ADVERTENCIA: Documento vivo. Los fixes aplicados estan listos, pero el score real requiere re-ejecucion.

---

## Linea base (historica)

**Ejecucion**: `extract_browser.py` sobre `audit_final.xlsx` (16 URLs)
**Fecha**: 2026-06-12
**Score global**: 46/100

| Metric | Valor |
|---|---|
| URLs procesadas | 16 |
| AA capturados | 8 (50%) |
| Errores | 8 |

Todos los errores: **Sin dato AA (no beacon)**.

### URLs con error (linea base)

| # | Dominio | URLs | Posible causa |
|---|---------|------|---------------|
| 6 | `brandpr.preview` (legal/privacy) | ES + EN | Preview sin AA, o carga condicional solo en prod |
| 2 | `ford.mx` | autos, blog | GA4 en vez de Adobe, o bloqueo a headless |

### URLs OK (linea base)

Tips / Driving Precautions / Fuel Saving / Collision Repair - ES (filas 4-7) y EN (filas 8-11).

---

## Correcciones Aplicadas

### 1. `UnboundLocalError` en cleanup
- Fix: mover `import subprocess, shutil` a imports globales
- Resultado: resuelto

### 2. Cookie consent ahora siempre se intenta cerrar
- Antes: opt-in con `--discard-cookies`
- Ahora: se ejecuta `try_dismiss_cookie_consent()` en TODAS las URLs
- Resultado: resuelto

### 3. `--wait-after` configurable (default: 4s)
- Antes: `sleep(2)` hardcodeado
- Ahora: `sleep(wait_after)`, default 4s, configurable por flag o `audit.json`
- Resultado: resuelto

### 4. `--market` + `--split-aa`
- `urls.json` ahora tiene campo `"market"` en cada entrada
- `--market PR` filtra y dirige output a `PR/historial.xlsx`
- `--split-aa` genera `con_aa.xlsx` y `sin_aa.xlsx` por mercado
- Resultado: resuelto

### 5. Colores en celdas de datos
- `apply_data_fills()` pinta: C verde, D rojo (si fallo), E amarillo, F verde
- Resultado: resuelto

### 6. Formato texto en celdas JSON
- `number_format = "@"` en columnas D, E, F, G - evita que Excel interprete JSON como numero/fecha
- Resultado: resuelto

---

## Score: Pendiente de verificacion

Los fixes aplicados **deberian** mejorar el score, pero no se ha re-ejecutado para medirlo.

| Fix | Impacto esperado |
|-----|-----------------|
| `wait_after: 4` (vs 2s) | Mas paginas alcanzan a emitir beacon AA |
| Cookie consent siempre | Menos falsos NO_AA_DATA por banner bloqueando |
| `--market PR` (excluye MX) | Score mas limpio por mercado |

**Pendiente**: ejecutar `extract_browser.py --urls urls.json --market PR --split-aa` y comparar score vs 46/100.

---

## A debate

Estos puntos requieren decision, no tienen fix tecnico:

1. **6 URLs legal/privacy sin AA en preview** - Es esperado? Debemos ignorarlas o desactivar la alerta?
2. **2 URLs ford.mx** - No estan en `urls.json` actual. Usan GA4? Agregarlas con `"market": "MX"`?
3. **Score general vs por mercado** - Medimos score global o por mercado? El score mezcla markets con comportamientos muy distintos.
4. **Linea de aceptacion** - Que score consideramos "suficiente"? 70? 80? Definir target por mercado.

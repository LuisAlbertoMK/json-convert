# Decisiones de Concurrencia — extract_browser.py

## Contexto

El script original procesaba URLs **secuencialmente**: 1 URL → esperar (carga + 2s) → extraer → escribir → siguiente. Con ~15 URLs a ~10s cada una, el tiempo total era ~150s (2.5 min). Con 100 URLs, ~17 min.

Objetivo: reducir tiempo manteniendo **corrección** (sin datos duplicados/perdidos) y **seguridad** (sin corrupción del Excel).

---

## Opciones Evaluadas

### Opción A ✅ — asyncio + Playwright async (IMPLEMENTADA)

```
1 browser       → N contextos/páginas
1 asyncio.Semaphore(N)  → controla cuántos vuelan simultáneamente
1 asyncio.Lock          → protege escritura Excel
```

| Pro | Contra |
|-----|--------|
| Un solo proceso, una sola memoria | Requiere migrar a async API |
| Playwright async API está diseñada para esto | openpyxl es sincrónico → lock |
| Semáforo = control fino de concurrencia | |
| Misma RAM ~200MB sin importar workers | |
| Workers comparten browser (cache HTTP, SSL) | |

**Veredicto**: Elegida. Es la forma correcta de escalar Playwright.

---

### Opción B ❌ — multiprocessing (DESCARTADA)

```
N procesos          → N navegadores independientes
N * ~200MB RAM      → 1GB con solo 5 workers
Cola de resultados  → pipe/queue entre procesos
```

| Pro | Contra |
|-----|--------|
| Aislamiento total entre procesos | **RAM: 200MB × N workers** — inviable |
| CPU-bound real? | I/O bound, no CPU. No gana nada |
| Multiprocesamiento en Windows es tortuoso | `if __name__ == "__main__"` + pickling |
| | Contexto nuevo = sin cache de SSL/TLS compartido |

**Razón de descarte**: El cuello de botella es **I/O de red** (esperar que cargue una página), no CPU. Multiprocessing agrega overhead de memoria sin beneficio real. En Windows, además, el multiprocessing es frágil (spawn, pickling, manejadores).

---

### Opción C ❌ — threading (DESCARTADA)

```
N threads           → 1 browser, N pages
GIL                 → todo corre en 1 core
openpyxl            → NO thread-safe
```

| Pro | Contra |
|-----|--------|
| Fácil de escribir | **Playwright sync API NO es thread-safe** |
| Baja memoria extra | openpyxl requiere Lock en cada celda |
| | GIL anula cualquier ganancia de CPU |
| | Debugging pesadillesco |

**Razón de descarte**: Playwright documenta explícitamente que la sync API no debe compartirse entre threads. Además, openpyxl no es thread-safe para escritura concurrente. Habría que poner locks hasta en la sopa. asyncio es la solución correcta para I/O-bound en Python.

---

### Opción D ❌ — ProcessPoolExecutor (DESCARTADA)

Variante de multiprocessing con `concurrent.futures`. Mismos problemas de RAM + Windows + pickling. Se descarta por las mismas razones que la Opción B.

---

## Riesgos de la Opción A (asyncio) y Mitigaciones

| Riesgo | Mitigación |
|--------|-----------|
| **Excel corrupto** si 2 workers escriben simultáneo | `asyncio.Lock` exclusivo en `write_result()`. Solo 1 worker escribe a la vez |
| **openpyxl.save() bloquea el event loop** | Se ejecuta directo (es I/O bound pero rápido). Si fuera lento, se migra a `loop.run_in_executor()` |
| **Handler de beacons race condition** | Cada worker tiene su propia page + su propia lista `beacon_urls` + su propio handler. **No hay estado compartido** |
| **Timeout sumado** (todos los workers timeout a la vez) | Cada worker tiene su propio timeout individual. Un timeout no afecta a los demás |
| **Demasiados workers saturan la red** | `--workers` default = 1. El usuario decide el grado de concurrencia. Semáforo límite superior |

---

## Benchmark Esperado (estimado)

| Workers | 15 URLs | 100 URLs | RAM |
|---------|---------|----------|-----|
| 1 (secuencial) | ~150s | ~17 min | ~200MB |
| 3 | ~50-60s | ~5-6 min | ~200MB |
| 5 | ~30-40s | ~3-4 min | ~200MB |
| 10 | ~20-30s | ~2-3 min | ~200MB |

El límite real es el **servidor destino** (Ford preview). Si el servidor rate-limitea, más workers = más 429/503. `--workers=3` es un buen punto de partida.

---

## --output: Por qué no es trivial

`--output` parece simple ("copia el archivo y escribe allá"), pero openpyxl no ofrece un método `copy_workbook()` desde siempre. La implementación actual:
1. Carga el workbook fuente con `openpyxl.load_workbook(input)`
2. Si hay `--output`, crea un workbook nuevo y copia **celda por celda** (incluyendo estilos)
3. Trabaja sobre la copia
4. Guarda en `--output`

Esto asegura que **el archivo original nunca se modifica**, incluso si el script crashea.

**Alternativa descartada**: `shutil.copy(input, output)` + trabajar sobre el output. Descartada porque si el script crashea antes de guardar, el output queda a medio escribir (posible corrupto). Con la copia en memoria, el output solo se escribe al final.

---

## Conclusión

asyncio + Playwright async + Semaphore + Lock es el patrón correcto para este caso:
- I/O bound (red) — asyncio brilla
- Múltiples contexts aislados — sin race conditions en beacons
- Excel compartido — Lock bien acotado solo para escritura
- RAM constante — sin importar workers
- Compatible Windows — sin multiprocessing

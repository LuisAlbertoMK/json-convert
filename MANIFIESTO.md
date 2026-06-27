# Manifiesto — Migración Adobe Experience Cloud Ford PR

> **Proyecto**: `json-convert` · Pipeline automatizado de validación y migración
> **Rol**: Especialista en migración de sitios Adobe Experience Cloud (Ford & Lincoln)
> **Mercado objetivo**: Puerto Rico (PR) — extensible a México (MX) y El Caribe
> **Fecha**: 2026-06-26

---

## 1. Misión

Automatizar la validación de Page Variables (digitalData) durante la migración al
**Estándar US Global de Adobe Experience Cloud** para Ford Puerto Rico,
eliminando la revisión manual y garantizando datos correctos de Adobe Analytics,
Target y Launch en cada URL.

**El problema**: Sin este pipeline, cada URL debe auditarse manualmente:
abrir navegador, inspeccionar digitalData, comparar contra expected.json,
anotar diferencias, generar Excel. Con 25+ URLs y múltiples entornos (preview +
producción + ambos), es imposible escalar.

**La solución**: Un pipeline completo que:
1. Navega cada URL con Playwright (browser real)
2. Extrae digitalData automático + beacons de AA
3. Compara contra el estándar esperado (expected.json)
4. Genera matrices de validación en Excel con estados ✅⚠️❌
5. Produce splits multi-sheet (1 archivo, 1 hoja por URL) con notas explicativas

---

## 2. Rol: Especialista en Migración Adobe Experience Cloud

### Responsabilidades

| Área | Qué hacemos |
|------|-------------|
| **Validación de Page Variables** | Comparamos valores REALES de digitalData.page (pageName, siteSection, client, site, variantName, pageType) contra el estándar US Global definido en expected.json |
| **Generación de documentos de implementación** | Producimos matrices de validación Excel para equipos de Authoring AEM, indicando exactamente qué cambiar en cada parámetro |
| **Aseguramiento de configuración** | Verificamos Adobe Analytics (AA beacons), Adobe Target, Adobe Launch y AEM durante el proceso de migración |
| **Multi-mercado** | Arquitectura diseñada para PR, MX y El Caribe — cada mercado con su propio expected.json, prefix (`fpr:`, `fmx:`) y config |

### Mercados actuales

| Mercado | Prefix | Estado |
|---------|--------|--------|
| **PR** (Puerto Rico) | `fpr:` | Pipeline completo, 25 URLs mapeadas, auditoría en progreso |
| **MX** (México) | `fmx:` | expected.json definido para blog/experto, pendiente de integrar |

---

## 3. Stack Tecnológico

| Componente | Tecnología | Propósito |
|------------|-----------|-----------|
| Browser automation | **Playwright** + Chromium/Firefox | Navegar URLs reales, extraer digitalData + AA beacons |
| Procesamiento Excel | **openpyxl** | Generar historiales, catálogos, matrices de validación |
| Parseo AA | **Python** (stdlib) | Extraer y estructurar beacons de Adobe Analytics |
| Pipeline | **CLI + menú interactivo** (menu.py) | Orquestar pasos: test → audit → post-process → report |
| Tests | **pytest + stdlib unittest** | 251+ tests, coverage 65-100% por módulo |
| Almacenamiento | **JSON** (url-mapping, expected) + **Excel** (historiales, catálogos) |

### Flujo de datos

```
url-mapping.json ──→ extract_browser.py ──→ historial.xlsx
     │                                          │
     │                                          ▼
     │                              generate_migration_catalog.py
     │                                          │
     ├── expected.json                          ▼
     │                              catalogo-migracion.xlsx
     │                                          │
     └──→ match_prod_preview.py ──→ match-3way.xlsx
          ↓
     generate_validation_matrix.py ──→ matriz-validacion.xlsx
                                       matriz-split-{entorno}.xlsx
```

---

## 4. Pipeline de Migración (Flujo Completo)

### Paso 1: Definir mapping (`data/url-mapping.json`)

Cada URL que se migra tiene una entrada con:
- `production_url` — URL real en producción (ford.com.pr)
- `preview_url` — URL en AEM preview (wwwac.preview.*.brandpr.ford.com) — *opcional*
- `aem_path` — Ruta en AEM (/content/ford/pr/...) — *opcional*
- `page_key` — Identificador único que conecta con expected.json
- `nombre` — Nombre descriptivo para las hojas Excel

**25 entradas actuales**:
- 14 con preview_url + aem_path (sustainability, owners, legal, privacy — ES/EN)
- 6 con aem_path sin preview_url (home, home-redirect, brochures, dealerships, dashboard, super-duty)
- 5 sin preview_url ni aem_path (performance/ev/trucks/all-vehicles-ev/maverick)

> ⚠️ **Aprendizaje crítico**: preview_url NO se puede derivar de aem_path.
> La relación no es 1:1 — AEM usa paths traducidos/localizados
> (ej: `sustainability/environment` → `index/donativos-ambientales.html`).

### Paso 2: Definir estándar esperado (`data/expected.json`)

Cada `page_key` define valores esperados para 7 parámetros:

| Parámetro | Regla | Ejemplo PR |
|-----------|-------|------------|
| pageName | pattern | `fpr:home`, `fpr:shopping:brochures` |
| siteSection | mapping | `home`, `shopping`, `owners`, `vehicles` |
| pageNameNoVehicle | mirror | = pageName |
| client | fixed | `ford-puerto rico` |
| site | fixed | `ford-brand` |
| variantName | mapping | `home-load`, `gallery-load`, `vhp-load` |
| pageType | required | `home`, `shopping-tools`, `vhp` |

**12 page_keys actuales**: home, home-redirect, brochures, dealerships, owners,
super-duty-f250, sustainability, legal, privacy, **showroom**, **ev**, **trucks**

### Paso 3: Auditar URLs (Playwright)

```bash
python src/extract_browser.py --urls data/urls.json --market PR --entorno preview
python src/extract_browser.py --urls data/urls.json --market PR --entorno produccion
```

Cada URL se navega con un browser real:
- Preview: Chromium headless (VPN necesaria)
- Producción: Chrome real (bypass Akamai WAF)
- Firefox para SPA con ERR_ABORTED (all-vehicles catalog)

Extrae: digitalData.page, AA beacons, errores, metadata (score, timing).

### Paso 4: Generar catálogo de migración

```bash
python src/generate_migration_catalog.py --market PR --entorno produccion
```

Produce `PR/produccion/catalogo-migracion.xlsx` con:
- URL, AEM path, parámetro, valor actual, valor esperado, acción

### Paso 5: Comparar (match) preview vs producción vs expected

```bash
python src/match_prod_preview.py --market PR --mode auto
```

Modos:
- **3-way**: US Expected vs Preview real vs Producción real (cuando existen ambos historiales)
- **2-way**: Prometido vs Entregado (fallback a expected.json si no hay preview)

### Paso 6: Generar matriz de validación

```bash
python src/generate_validation_matrix.py --market PR --entorno ambas --split
```

**Split** (formato actual): 1 archivo, 1 hoja por URL — `PR/split/matriz-split-{entorno}.xlsx`
Cada hoja muestra URL solo en primera fila, con columna "Observación / Nota" al final.

---

## 5. Base de Conocimiento (Aprendizajes Esenciales)

### 5.1 DigitalData en Ford PR

**Valores reales verificados en producción** (contra el sitio real):

| URL | pageName real | Estado |
|-----|---------------|--------|
| `ford.com.pr/` | `home` | ✅ Verificado |
| `ford.com.pr/esp/` | `home` | ✅ Verificado |
| `ford.com.pr/esp/duenos/` | `owner support:dashboard` | ✅ Verificado |
| `ford.com.pr/esp/shopping/brochures/` | `direct sales:qr brochures` | ✅ Verificado |
| `ford.com.pr/esp/concesionarios/` | `dealer locater` | ✅ Verificado |
| `ford.com.pr/esp/trucks/super-duty-f250/` | `vehicle:home:ford super duty f-250` | ✅ Verificado |
| `ford.com.pr/esp/ev/` | `showroom:electrified` | ✅ Verificado (sin prefijo `fpr:`) |

**Patrones observados**:
- El prefijo `fpr:` NO siempre está presente en producción actual (solo en expected futuro)
- `siteSection: "showroom"` para EV (no "vehicles" como se esperaría)
- `client` puede variar: `ford-puerto rico` (con espacio) vs `ford-puertorico` (sin espacio, preview)
- Ninguna URL auditada tiene Adobe Analytics (AA) configurado — es site-wide, no bug del pipeline

### 5.2 Relación Preview ↔ Producción

| Aspecto | Preview | Producción |
|---------|---------|------------|
| Dominio | `wwwac.preview.*.brandpr.ford.com` | `www.ford.com.pr` |
| RSID | `fmcprpredevp` (dev) | `fmcprprod` (real) |
| client | `ford-puertorico` (sin espacio) | `ford-puerto rico` (con espacio) |
| VPN | Requerida | No requerida |
| Akamai WAF | No | Sí (bypass con Chrome real) |
| URLs disponibles | ~14 de 25 mapeadas | 25 de 25 mapeadas |

### 5.3 Bug Crítico: Substring URL Matching

**Problema**: `"" in str(url_cell)` = **siempre True en Python**. Causó que 11 entries
sin preview_url (y por lo tanto `url_fragment=""`) matchearan SIEMPRE el primer
row del historial, independientemente de la URL real.

```python
# ANTES (BUG): url_fragment="" → SIEMPRE True
if url_cell and url_fragment in str(url_cell):
    # "" in "anything" → True → matcheaba cualquier URL
    
# DESPUÉS (FIX): guard explícito
if url_cell and url_fragment and url_fragment in str(url_cell):
    # "" es falsy → short-circuit False → no falso positivo
```

**Mismo bug** existía en `generate_migration_catalog.py:116` y
`match_prod_preview.py:113`. Ambos corregidos con el mismo fix.

### 5.4 Cache vs Datos Reales

Decisión firme del usuario: **cero caché, siempre datos frescos del sitio**.
- Cache global eliminado de `extract_browser.py` (UrlCache, --no-cache, --cache-ttl, --clear-cache)
- Playwright navega URLs reales cada vez
- No alucinar valores — solo lo que el sitio devuelve realmente

### 5.5 Error Pages

Cuando una URL devuelve 404, el servidor de Ford genera:
- digitalData.page.pageName = `"error page"` (o "errorPage", "error-page")
- Esto es un FALLBACK del servidor, NO digitalData real

El pipeline detecta automáticamente "error page" y agrega nota explicativa
en la columna "Observación / Nota".

### 5.6 AA (Adobe Analytics) no Configurado

En todas las URLs auditadas (ambos entornos): **0% con AA**.
- digitalData → presente en ~80% de URLs
- AA beacons → ausente en 100%

Es una limitación site-wide del sitio Ford PR actual (pre-migración).
No es bug del pipeline. Se documenta en la nota cuando corresponde.

---

## 6. Decisiones Arquitectónicas

| Decisión | Opción descartada | Por qué |
|----------|------------------|---------|
| **Split multi-sheet** (1 archivo, 1 hoja/URL) | Múltiples archivos separados | Más simple, 1 archivo por entorno. URL solo en 1ª fila para evitar repetirla 8 veces |
| **Columna nota al final** | Sin nota | Explica "error page" y otros fallbacks directamente en el Excel que recibe Authoring |
| **Catalogo docs como fallback** | Solo pipeline catalogo | Los docs tienen valores reales verificados manualmente. El pipeline catalogo puede tener datos vacíos si no se corrió la auditoría |
| **Exact match para docs catalog** | Path-prefix flexible | `/esp/` en docs causaría falsos positivos con otras URLs |
| **3-way match** (Expected vs Preview vs Prod) | Solo 2-way | Revela discrepancias entre entornos (preview vs production) que 2-way no detecta |
| **Git Bash para scripting** | PowerShell 5.1 | PS5.1 rechaza `&&`, `||`, `@{var}`. Git Bash es el shell portable confiable en Windows |

---

## 7. Gotchas y Anti-Patrones

### 🔴 Anti-Patrones (NO hacer)

1. **Derivar preview_url de aem_path** — NO se puede. Los paths AEM usan
   traducciones/localizaciones que no siguen un patrón predecible.

2. **Cachear resultados de auditoría** — NO. El usuario exige datos frescos
   del sitio en cada ejecución. El sitio cambia (AEM authoring actualiza páginas).

3. **Substring URL matching** — `a in b` para URLs causa falsos positivos.
   `www.ford.com.pr/` es substring de TODAS las URLs del mismo dominio.
   Siempre usar `urlparse` + path component matching.

4. **Alucinar valores** — Cero invención. Si una URL no se audita, su valor
   es "—" (sin datos), no un valor inventado.

5. **Asumir que preview_url existe** — 11/25 entries no tienen preview_url.
   El pipeline debe manejar esto sin crash.

### ⚠️ Gotchas Técnicos

| Gotcha | Explicación |
|--------|-------------|
| **`"" in str` es True** | En Python, `"" in "cualquier_cosa"` → `True`. Siempre agregar `and url_fragment and` antes. |
| **Playwright + Akamai WAF** | Bundled Chromium es bloqueado por Akamai. Usar `channel="chrome"` (Chrome real instalado) para producción. |
| **ERR_ABORTED en SPA** | Ford all-vehicles catalog causa ERR_ABORTED en Chromium. Firefox lo evita. |
| **VPN para preview** | URLs de preview requieren VPN corporativa. URLs de producción no. |
| **client con/sin espacio** | Preview usa `ford-puertorico` (sin espacio). Producción usa `ford-puerto rico` (con espacio). expected.json debe reflejar producción. |
| **pageName sin prefijo `fpr:`** | EV devuelve `showroom:electrified` (sin `fpr:`). Es el valor actual real, no necesariamente el esperado. |

### 🧪 Patrón de Prevención de Errores

Cada vez que comparas URLs en este proyecto, usa:

```python
from urllib.parse import urlparse

def _paths_match(short: str, long: str) -> bool:
    """True si short es prefijo de long en frontera de path component."""
    if short == long:
        return True
    if short == "/":
        return False  # root solo matchea root
    if long.startswith(short) and len(long) > len(short) and long[len(short)] == "/":
        return True
    if short.startswith(long) and len(short) > len(long) and short[len(long)] == "/":
        return True
    return False
```

---

## 8. Estado Actual y Roadmap

### Estado: 🟡 En Progreso

| Componente | Estado | Notas |
|-----------|--------|-------|
| Pipeline base (extract_browser) | ✅ Completo | Playwright + multi-entorno + split-AA |
| Catálogo de migración | ✅ Completo | generate_migration_catalog.py con catalogo_docs |
| Match 3-way | ✅ Completo | match_prod_preview.py con 3 modos |
| Matriz validación + split | ✅ Completo | generate_validation_matrix.py con columna nota |
| expected.json PR | ✅ Completo | 12 page_keys con valores reales |
| url-mapping.json PR | ✅ 25 entries | 11 sin preview_url, pendientes de auditoría |
| Bug `"" in url` | ✅ Corregido | Ambos archivos (catalog + match) |
| Cache eliminado | ✅ Completo | UrlCache removido globalmente |
| Tests | ✅ 251/251 | Pasando |
| **Reporte auditoría** | 🟡 Parcial | 10 URLs auditadas, 2 fallos (sandbox) |
| **Auditoría completa** | 🔴 Pendiente | Faltan 15 URLs de producción + 14 preview |
| **Integrar reporte en nota** | 🔴 Pendiente | Usar score y fallos del reporte_auditoria.xlsx |
| **BITACORA/METRICAS/.learnings** | 🟡 Parcial | Actualizar con hallazgos de esta sesión |
| **MX market** | 🔴 Pendiente | expected.json definido, pipeline no probado |

### Roadmap Inmediato

```
Fase 1: Completar datos base
  ✅ expected.json con showroom/ev/trucks
  ⬜ url-mapping.json: agregar preview_url de auditorías reales
  ⬜ expected.json: agregar page_keys showroom/ev/trucks con siteSection correcto

Fase 2: Re-ejecutar pipeline
  ⬜ Auditar URLs faltantes (Playwright: producción + preview)
  ⬜ Regenerar catálogo con datos reales
  ⬜ Regenerar matriz validación + split

Fase 3: Integrar reporte
  ⬜ Usar scores y fallos de reporte_auditoria.xlsx en columna nota
  ⬜ Documentar hallazgos en BITACORA.md, METRICAS.md, .learnings

Fase 4: Calidad
  ⬜ Score target: 9.0+ en todas las dimensiones
  ⬜ Auto-metrics post-ejecución
  ⬜ Ciclo de mejora continua
```

---

## 9. Comandos Rápidos

```bash
# Pipeline completo
python src/menu.py --run auto

# Generar mapping desde Excel
python scripts/_gen_urls.py --input RevisionManual_PR.xlsx --default-market PR

# Auditar (Playwright)
python src/extract_browser.py --urls data/urls.json --market PR --entorno preview
python src/extract_browser.py --urls data/urls.json --market PR --entorno produccion

# Catálogo
python src/generate_migration_catalog.py --market PR --entorno produccion
python src/generate_migration_catalog.py --market PR --entorno ambas

# Match (comparar)
python src/match_prod_preview.py --market PR

# Matriz validación
python src/generate_validation_matrix.py --market PR --entorno produccion
python src/generate_validation_matrix.py --market PR --entorno preview
python src/generate_validation_matrix.py --market PR --entorno ambas --split

# Tests
python -m pytest tests/ -v
```

---

> **Nota final**: Este manifiesto debe actualizarse cuando:
> - Se complete la auditoría de las 25 URLs
> - Se verifiquen valores de producción para todos los page_keys
> - Se expanda a MX market
> - Se identifique un nuevo anti-patrón o gotcha
>
> *"Datos reales del sitio, cero invención, procesos automáticos."*

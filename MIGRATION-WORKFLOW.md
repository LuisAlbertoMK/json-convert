# Flujo de Trabajo: Migración de Mercados Ford al Estándar US Global

## Visión General

Pipeline completo para auditar y migrar mercados Ford (PR, MX, etc.) al estándar
de Adobe Analytics US Global.

```
urls.json → extract_browser.py → historial.xlsx
                                          ↓
url-mapping.json + expected.json → generate_migration_catalog.py
                                          ↓
                              catálogo-migracion.xlsx + resumen
                                          ↓
                              implementacion-{mercado}.md (especificación)
```

---

## 1. Preparación

### 1.1 urls.json — Definir URLs a auditar

```json
{
  "url": "https://www.ford.mx/blog/experto/.../",
  "nombre": "Nombre Descriptivo",
  "market": "MX",
  "tipo": "blog",
  "entorno": "produccion"
}
```

| Campo | Descripción |
|-------|-------------|
| `url` | URL completa (preview o producción) |
| `nombre` | Nombre legible para el reporte |
| `market` | Código ISO del mercado (PR, MX, …) |
| `tipo` | Categoría: `home`, `nameplate`, `blog`, `estaticas`, … |
| `entorno` | `preview` (VPN, staging) o `produccion` (público) |

### 1.2 data/url-mapping-{market}.json — Mapear page_key

Cada URL recibe un `page_key` que la agrupa con otras del mismo tipo:

```json
{
  "production_url": "https://...",
  "page_key": "blog",
  "nombre": "Artículo"
}
```

### 1.3 data/expected.json — Definir valores esperados

```json
{
  "markets": {
    "MX": {
      "prefix": "fmx:",
      "params": {
        "pageName": {
          "rule": "pattern",
          "patterns": { "blog": "blog:expert:" }
        },
        "siteSection": {
          "rule": "mapping",
          "mapping": { "blog": "blog" }
        },
        "variantName": {
          "rule": "mapping",
          "mapping": { "blog": "blog-load" }
        },
        "pageType": {
          "rule": "required",
          "mapping": { "blog": "blog" }
        }
      }
    }
  }
}
```

**Tipos de regla:**
| Regla | Descripción |
|-------|-------------|
| `pattern` | Valor esperado (comparación con `startswith` para pageName) |
| `mapping` | Mapeo por page_key (siteSection, variantName) |
| `fixed` | Valor fijo en todas las páginas (client, site) |
| `required` | Nuevo parámetro que debe crearse (pageType) |
| `mirror` | Copia de otro parámetro (pageNameNoVehicle → pageName) |
| `deprecated` | Parámetro que debe eliminarse (userLanguage) |

---

## 2. Extracción

### extract_browser.py

```bash
python src/extract_browser.py \
  --urls urls.json \
  --market MX \
  --entorno produccion \
  --workers 3
```

**Flags clave:**
| Flag | Descripción |
|------|-------------|
| `--market MX` | Filtra URLs del mercado MX |
| `--entorno produccion` | Solo URLs en producción |
| `--entorno ambas` | Preview + producción |
| `--workers 3` | Concurrencia (default 3) |
| `--timeout 35` | Timeout por URL en segundos |
| `--browser chrome` | Usa Chrome (bypass Akamai WAF en prod) |

**Output:** `{market}/{entorno}/historial.xlsx`

**Sheets del historial:**
- `{fecha}` — datos crudos por URL
- `_control` — métricas de la auditoría
- `_vars` — variables Adobe Analytics (si existen)

---

## 3. Generación de Catálogo

### generate_migration_catalog.py

```bash
python src/generate_migration_catalog.py \
  --historial MX/produccion/historial.xlsx \
  --mapping data/url-mapping-mx.json \
  --expected data/expected.json \
  --market MX \
  --output MX/catalogo-migracion.xlsx
```

**Outputs:**
| Archivo | Contenido |
|---------|-----------|
| `MX/catalogo-migracion.xlsx` | Catálogo completo por URL + parámetro |
| `MX/resumen-catalogo-migracion.md` | Resumen markdown |
| `MX/resumen-catalogo-migracion.html` | Resumen HTML portátil |

**Interpretación del catálogo:**

| Estado | Significado |
|--------|-------------|
| ✅ Alineado | Valor actual = valor esperado |
| ⚠️ Requiere cambios | Valor legacy → valor estándar |
| ❌ No existe | Parámetro nuevo que debe crearse |
| 🗑️ Deprecado | Parámetro que debe eliminarse |

---

## 4. Documentación de Implementación

Una vez generado el catálogo, crear la especificación de implementación:

```
MX/implementacion-mx-blog.md
```

Debe incluir:
- Diferencias exactas (`diff`) por parámetro
- Instrucciones para el equipo de Launch/DTM
- Lo que NO cambia (para evitar confusiones)
- Comando de verificación post-implementación

---

## 5. Verificación Post-Implementación

```bash
# 1. Re-extraer después del cambio
python src/extract_browser.py --urls urls.json --market MX --entorno produccion

# 2. Regenerar catálogo
python src/generate_migration_catalog.py \
  --historial MX/produccion/historial.xlsx \
  --mapping data/url-mapping-mx.json \
  --expected data/expected.json \
  --market MX \
  --output MX/catalogo-migracion.xlsx

# 3. Validar que todos los parámetros estén en ✅
```

---

## 6. Menú Interactivo

```bash
python src/menu.py
```

Opciones:
1. **Auditar URLs** — ejecuta extract_browser.py
2. **Extraer AA** — procesa beacons Adobe Analytics
3. **Match 3-way** — compara preview vs producción vs expected
4. **Generar catálogo** — genera catálogo de migración
5. **Reporte de auditoría** — reporte consolidado
6. **Pipeline completo** — todo en uno

---

## Anatomía de digitalData.page

```yaml
# Estándar US Global — campos relevantes para migración
pageName:         "blog:expert:<title>"        # pattern — sin prefijo de mercado
siteSection:      "blog"                       # mapping — eVar9 / hierarchy
pageNameNoVehicle: "(mirror de pageName)"      # mirror — eVar11
client:           "ford-mexico"                # fixed — eVar14
site:             "ford-brand"                 # fixed — eVar15
variantName:      "blog-load"                  # mapping — Direct Call
pageType:         "blog"                       # required — eVar129
userLanguage:     "esp"                        # deprecated — eVar4
hierarchy:        "blog:expert"                # derivado de siteSection
```

---

## Notas por Mercado

### MX — Blog/Experto
- pageName usa `blog:expert:<title>` SIN prefijo `fmx:`
- Los títulos están en inglés (ej: "how to jump start a car")
- 0% Adobe Analytics actualmente — solo digitalData + GA4
- siteSection legacy: `incentives` → target: `blog`
- variantName legacy: `gallery-load` → target: `blog-load`
- pageType no existe → debe crearse como `blog`

### PR — General
- pageName usa `fpr:<categoria>:<subcategoria>` CON prefijo `fpr:`
- 100% Adobe Analytics presente
- Catálogo completo generado en `PR/{entorno}/catalogo-migracion.xlsx` (ej: `PR/produccion/catalogo-migracion.xlsx`)

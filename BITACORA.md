# Bitácora de Sesiones — json-convert

## 2026-06-30 — Limpieza profunda + Auditoría Myco Performance GTBEMEAPUB-42399

### Limpieza de raíz (52→33 entradas)
- Eliminados 11 scripts debug (`_check_*.py`, `_scan*.py`, `verify_*.py`, etc.)
- Eliminados `logs/` (10+ .log), `config/` (326B), `.learnings/` (8 md), `.atl/` (2 files)
- Movidos SEMANAS_*.json → `data/`, `.menu-config.json` y `.env.example` → raíz
- Actualizada ruta en `menu.py` para SEMANAS y menu-config
- Actualizado `.gitignore`: logs/, .learnings/, .atl/, rutas nuevas
- Caches regenerables eliminados (.pytest_cache, .url_cache, egg-info)

### Auditoría Myco All Vehicles Performance (GTBEMEAPUB-42399)
- Extraídas 6 URLs desde `docs/FPR_MYCO_All_vehicles_performance_v2.xlsx`
- Creado `myco_urls.json` → ejecutado `auditar_urls.py --ticket 42399`
- 6/6 URLs auditadas OK: Performance, Raptor, Ranger Raptor, Mustang Dark Horse, Bronco Raptor, Home
- Todos los archivos de ticket consolidados en `PR/GTBEMEAPUB-42399/`
- Score digitalData: 51/80 en Performance (esperado), Home ≥80 ✅

## 2026-06-26 (Parte 3) — Corrección: revertir datos inventados, solo docs reales

### Corrección
- **Revertí** `expected.json` y `url-mapping.json` a `f9760d2` — los 8 page_keys que agregué (showroom, ev, trucks, all-vehicles, crossovers-suvs, maverick, sandbox-es, sandbox-en) NO existen en ningún doc
- Los únicos page_keys documentados son 9: home, home-redirect, brochures, dealerships, owners, super-duty-f250, sustainability, legal, privacy
- url-mapping.json de `f9760d2` ya tenía showroom/ev/trucks/maverick (URLs de producción reales) pero SIN valores esperados — correcto así
- Moraleja: NUNCA inferir valores que no están explícitos en docs

### Cambios que se mantienen
- `src/generate_validation_matrix.py`: mirror rule fix (línea 194-196)
- `src/generate_migration_catalog.py`: mirror rule fix (línea 147-149)
- `src/menu.py`: +opción 11 (matriz validación), soporte entornos preview/produccion
- `METRICAS.md`: session metrics format
- `MIGRATION-WORKFLOW.md`: path PR/{entorno}/
- `.gitignore`: excluye PR/, debug scripts, PDFs
- MX blog y RevisionManual eliminados (manual)

## 2026-06-26 (Parte 2) — pageNameNoVehicle fix + pipeline completo + !ship

### Objetivo
Corregir mirror rule de `pageNameNoVehicle` que resolvía literal `(ver pageName)` y completar pipeline.

### Bug corregido
- Ambos scripts (`generate_validation_matrix.py` y `generate_migration_catalog.py`) tenían el mismo bug: la regla mirror leía el placeholder literal del expected.json en lugar del valor real de pageName
- Fix: resolver pageName real desde datos resueltos, no desde el placeholder en expected

### Cambios
- `data/expected.json`: +8 page_keys (showroom, ev, trucks, all-vehicles, crossovers-suvs, maverick, sandbox-es, sandbox-en)
- `data/url-mapping.json`: maverick separado de trucks
- `src/generate_validation_matrix.py`: mirror rule fix (línea 194-196)
- `src/generate_migration_catalog.py`: mirror rule fix (línea 147-149)
- `src/menu.py`: +opción 11 (matriz validación), soporte entornos preview/produccion
- `METRICAS.md`: session metrics format

### Estado
- 251/251 tests pasando ✅
- 10 URLs × 7 parámetros en catálogo migración
- Score: 9.2/10 (+0.7 vs baseline)
- Pendiente PR/produccion/historial.xlsx para match prod vs preview

## 2026-06-26 — Integración docs catalog + URL matching fix

### Objetivo
Integrar el `docs/ford-pr-catalogo-valores-pre-preview.xlsx` como fuente de valores reales de producción en el pipeline de validación.

### Descubrimientos
- El docs catalog tiene digitalData real de 6 URLs de producción de Ford PR (home, home-redirect, brochures, dealerships, owners, super-duty-f250)
- El formato del docs catalog es IDÉNTICO al del pipeline catalogo (col1=URL, col3=param, col4=value) — `load_catalogo()` ya lo soporta
- La matriz de referencia (`ford-pr-matriz_migracion_datos_v2-.xlsx`) tiene el MISMO formato que nuestra salida actual — confirma dirección correcta

### Bug corregido
- `_catalog_lookup()` usaba `a in b` (substring) para URL matching — causaba que root `/` y `/esp/` matchearan TODAS las URLs
- Corregido con URL-aware path component matching usando `urlparse`
- Docs catalog usa matching exacto; pipeline catalogo usa path-prefix

### Cambios
- `data/url-mapping.json`: +6 entries con URLs reales de producción
- `src/generate_validation_matrix.py`: +`--catalogo-docs` param, +auto-detect docs catalog, fixed `_catalog_lookup()`

### Estado
- 25 páginas en mapping, 11 con datos (6 docs catalog + 5 pipeline catalogo)
- 14 páginas sin datos — requieren auditoría pipeline
- Commit: `f9760d2`

## 2026-06-26 (Parte 4) — Reporte auditoría + expected.json completo + MANIFIESTO

### Objetivo
Completar Fase 1 de la recomendación (expected.json con showroom/ev/trucks), crear manifiesto del proyecto, y verificar pipeline completo.

### Descubrimientos clave
- **preview_url NO se deriva de aem_path**: los paths en preview usan segmentos traducidos/localizados (ej: `sustainability/environment` → `index/donativos-ambientales.html`). No hay correlación directa.
- **Reporte auditoría (10 URLs)**: 80% tasa de éxito, 0% con AA (site-wide), todas las URLs de producción funcionan con digitalData
- **EV real**: `pageName: "showroom:electrified"` (SIN prefijo `fpr:`) — verificado contra producción
- **showroom real (preview)**: `pageName: "fpr:showroom:all vehicles"` — verificado contra preview
- **2 orphans en historial**: `all-vehicles.html` y `crossovers-suvs.html` existen en historial pero NO en url-mapping.json

### Bug encontrado
- `menu.py` match 3-way pasa el MISMO historial para preview y producción (detecta "PR - Preview" y "PR - Produccion" pero match_prod_preview.py recibe ambos como preview)

### Cambios
- `data/expected.json`: +3 page_keys (showroom, ev, trucks) con valores REALES verificados
- `MANIFIESTO.md`: nuevo — documento fuente con misión, rol, aprendizajes, anti-patrones, roadmap
- Engram: guardado patrón "leer MANIFIESTO.md cada sesión"

### Pipeline verificado
- ✅ 251/251 tests
- ✅ Catálogo migración: 25 URLs generadas
- ✅ Match 3-way: 25 entradas, 175 parámetros
- ✅ Matriz validación: 216 filas, ✅52 ⚠️41 ❌123
- ⬜ Audit Playwright pendiente (14 URLs sin datos)
- ⬜ Integrar reporte_auditoria.xlsx en nota (Fase 3)

# Session Learnings — 2026-06-30 (Nameplate Analysis Pipeline)

> Proyecto: **json-convert** — Pipeline de análisis para Nameplate Excel de Ford Bronco PR

---

## 🧠 What We Did This Session

### 1. Nuevo script: `scripts/analizar_nameplate.py`
Pipeline autónomo que analiza el Excel `FPR_MYCO_Nameplate_Bronco.xlsx` y genera:
- **CSV unificado** con estructura idéntica al Excel original (columnas A-F + URL + Origen)
- **JSON de resumen** con estadísticas y calidad
- Nomenclatura: `{ticket}-{mercado}-{pagina}.csv` dentro de carpeta `{ticket}/`

### 2. Estructura del Nameplate Excel
El archivo contiene 7 hojas (6 visibles, 1 oculta):

| Hoja | Propósito |
|------|-----------|
| SEO Recommendations \| ESP | Metadata SEO final (titles, descriptions, OG tags) |
| FPR_MYCO_Formato_Nameplate_v5(R | **Maestra** — especificación completa de la landing page Bronco (710 filas, 40 cols) |
| Bronco Version Base | Página del trim Base |
| Bronco Version Big Bend | Página del trim Big Bend |
| Bronco Version Outer Banks | Página del trim Outer Banks |
| Bronco Version Badlands | Página del trim Badlands |
| Expedition \| Eng (oculta) | Expedition 2026 en inglés (referencia futura) |

### 3. Columnas del Excel original
- **A**: Visualización (sección de página — 17 valores únicos)
- **B**: Módulo / Sección (componente UI — 72 valores únicos)
- **C**: SEO/Formato (tipo de contenido — 23 tipos)
- **D**: Copies (contenido real)
- **E**: Contadores (conteo de caracteres, semántica mixta)
- **F**: Otros/Notas/Comentarios (constraints, reglas)
- **G**: Notas ocasionales

### 4. Anomalías de calidad detectadas
- **9 labels incorrectos**: "Version Base - Color 2/3/4" en hojas Big Bend, Outer Banks y Badlands (deberían decir su propia versión)
- **2 títulos incorrectos**: "Bronco Base" aparece en contenido de Badlands
- **1 contador mixto**: valor '#REF!' en Contadores
- **8 colores duplicados**: Color 8 y 9 aparecen dos veces en cada hoja de versión
- **1 incidencia global**: Raptor existe en maestra pero no tiene hoja dedicada

---

## 🔧 Gotchas Nuevos

### URL detection varies by sheet layout
- En hojas Maestra y Versiones: las URLs están en **col C** (SEO/Formato = "URL")
- En hoja SEO: las URLs están en **col B** (Módulo/Sección = "URL")
- Solución: función `_is_url_row()` que revisa ambas columnas

### Page URL vs Navigation URLs
- La PRIMERA URL de cada hoja es la **URL de página** (identificador)
- Las URLs siguientes son **links de navegación** (CTA, anchor links, etc.)
- No propagar navegación URLs como página — usar solo la primera URL como page_url

### SECTION rows are structural, not content
- En la hoja SEO, las filas "SECTION" agrupan bloques de metadata por versión
- No deben tener URL (son estructurales) ni deben confundirse con content rows

### Contadores column has mixed semantics
- A veces es un entero (character count)
- A veces es un string ("Caracteres", "Palabras")
- A veces es una nota ("Este aplica para todas las páginas")
- No se puede asumir tipo de dato

---

## 📁 Archivos relevantes
- `scripts/analizar_nameplate.py` — Pipeline de análisis
- `GTBEMEAPUB-42479/GTBEMEAPUB-42479-PR-Bronco.csv` — Output CSV (887 filas)
- `GTBEMEAPUB-42479/GTBEMEAPUB-42479-PR-Bronco-summary.json` — Resumen con calidad

---

## ✅ Próximos pasos
- Integrar al menú principal del pipeline (`menu.py`)
- Agregar flag para incluir/ocultar hoja Expedition
- Agregar flag para output en Excel (multi-sheet como el original)

# Implementación: Migración MX Blog → US Global

## Objetivo
Alinear el data layer (`digitalData.page`) de las páginas blog/experto de ford.mx
con el estándar US Global de Ford.

## URLs afectadas
Todas las páginas bajo `https://www.ford.mx/blog/experto/*`

## Cambios requeridos en digitalData.page

### 1. siteSection — CAMBIAR
```diff
- "siteSection": "incentives"
+ "siteSection": "blog"
```
**Dónde**: DTM/Launch — variable eVar9 / `digitalData.page.siteSection`
**Por qué**: `incentives` es un valor legacy incorrecto para contenido blog.

### 2. variantName — CAMBIAR
```diff
- "variantName": "gallery-load"
+ "variantName": "blog-load"
```
**Dónde**: DTM/Launch — Direct Call rule / `digitalData.page.variantName`
**Por qué**: `gallery-load` corresponde a galerías de imágenes, no a contenido blog.

### 3. pageType — AGREGAR
```diff
+ "pageType": "blog"
```
**Dónde**: DTM/Launch — variable eVar129 / `digitalData.page.pageType`
**Por qué**: Es un nuevo parámetro requerido por el componente estándar de US Global.
Captura el tipo de página que el visitante está viendo.

### 4. userLanguage — ELIMINAR
```diff
- "userLanguage": "esp"
```
**Dónde**: DTM/Launch — variable eVar4 / `digitalData.page.userLanguage`
**Por qué**: Deprecado en el nuevo componente. El idioma se unifica desde Adobe Launch.

## Lo que NO cambia

| Parámetro | Valor actual | Estándar | Acción |
|-----------|-------------|----------|--------|
| pageName | `blog:expert:<title>` | ✅ blog:expert:<title> | Mantener |
| pageNameNoVehicle | mirror de pageName | ✅ igual | Mantener |
| client | `ford-mexico` | ✅ | Mantener |
| site | `ford-brand` | ✅ | Mantener |

## Notas técnicas
- El pageName **no lleva prefijo de mercado** (`fmx:`). Las páginas blog usan
  `blog:expert:<slug>` directamente, igual que ford.com US.
- Los valores de pageName se generan a partir del título del artículo en inglés
  (ej: `blog:expert:how to jump start a car`), no de la URL en español.
- No hay Adobe Analytics en las páginas blog MX actualmente (0% AA). Solo existe
  digitalData. La migración agrega el tracking AA vía el componente US Global.

## Verificación post-implementación
Correr el pipeline de auditoría:
```bash
python src/extract_browser.py --urls urls.json --market MX --entorno produccion
python src/generate_migration_catalog.py \
  --historial MX/produccion/historial.xlsx \
  --mapping data/url-mapping-mx.json \
  --expected data/expected.json \
  --market MX \
  --output MX/catalogo-migracion.xlsx
```
Validar que los 4 parámetros aparezcan como ✅ (alineados) en el catálogo.

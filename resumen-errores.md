# Resumen de Errores — Extracción AA

**Ejecución**: `extract_browser.py` sobre `audit_final.xlsx` (16 URLs)
**Fecha**: 2026-06-12
**Score global**: 46/100

---

## Totales

| Métrica | Valor |
|---|---|
| URLs procesadas | 16 |
| AA capturados | 8 (50%) |
| Errores | 8 |

Todos los errores son del mismo tipo: **Sin dato AA (no beacon)**. Las páginas cargaron pero no emitieron ningún beacon de Adobe Analytics.

---

## URLs con Error

### brandpr.preview — Sin Adobe Analytics

| Fila | Nombre | URL |
|------|--------|-----|
| 2 | Donativos Ambientales (ES) | `https://wwwac.preview.es.brandpr.ford.com/content/na/ford/es_pr/index/donativos-ambientales.html` |
| 3 | Environmental Grants (EN) | `https://wwwac.preview.brandpr.ford.com/content/na/ford/en_pr/index/environmental-grants.html` |
| 12 | Legal (ES) | `https://wwwac.preview.es.brandpr.ford.com/content/na/ford/es_pr/index/legal.html` |
| 13 | Privacy Policy (ES) | `https://wwwac.preview.es.brandpr.ford.com/content/na/ford/es_pr/index/legal/privacy-policy.html` |
| 14 | Legal (EN) | `https://wwwac.preview.brandpr.ford.com/content/na/ford/en_pr/index/legal.html` |
| 15 | Privacy Policy (EN) | `https://wwwac.preview.brandpr.ford.com/content/na/ford/en_pr/index/legal/privacy-policy.html` |

Posible causa: estas páginas en el entorno preview no tienen implementado Adobe Analytics, o el data layer se carga condicionalmente (solo en producción).

### ford.mx — Sin Adobe Analytics

| Fila | Nombre | URL |
|------|--------|-----|
| 16 | Ford MX Autos | `https://www.ford.mx/autos/` |
| 17 | Ford MX Blog - Pasar Corriente | `https://www.ford.mx/blog/experto/instrucciones-pasar-corriente-auto-ago2020/` |

Posible causa: el sitio ford.mx puede estar bloqueando el headless de Playwright, o usar un proveedor de analytics diferente a Adobe (GA4, etc.).

---

## URLs OK (con AA capturado)

| Filas | Grupo |
|-------|-------|
| 4 – 7 | Tips / Driving Precautions / Fuel Saving / Collision Repair (ES) |
| 8 – 11 | Tips / Driving Precautions / Collision Repair / Fuel Saving (EN) |

Todas las URLs con AA exitoso son del entorno **brandpr preview** en las secciones de Tips/Consejos para dueños.

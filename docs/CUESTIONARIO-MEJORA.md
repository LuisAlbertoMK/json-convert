# Cuestionario de Mejora — Proyecto json-convert

Basado en el pipeline actual (url-mapping → extract_browser → historial → catalogo-migracion → matriz-validacion) y los datos reales capturados en producción.

---

## 1. Valores Esperados (expected.json)

### 1.1 page_keys nuevos
Agregué 7 page_keys a `expected.json`. ¿Estos valores son correctos o necesitan ajustes?

| page_key | pageName (actual) | pageName (esperado) |
|---|---|---|
| `showroom` | — | `fpr:vehicles:showroom` |
| `ev` | — | `fpr:vehicles:ev` |
| `trucks` | — | `fpr:vehicles:trucks` |
| `all-vehicles` | — | `fpr:vehicle-shopping:all-vehicles` |
| `crossovers-suvs` | — | `fpr:vehicle-shopping:crossovers-suvs` |
| `sandbox-es` | — | `fpr:sandbox:es` (placeholder) |
| `sandbox-en` | — | `fpr:sandbox:en` (placeholder) |

**Preguntas:**
- ¿Los expected values de `showroom`, `ev`, `trucks`, `all-vehicles` y `crossovers-suvs` son los correctos según el estándar US Global final?
- ¿`sandbox-es` y `sandbox-en` son páginas temporales de prueba? ¿Deben tener valores esperados reales o solo ignorarse?

### 1.2 page_keys legacy
Los 9 page_keys originales (home, home-redirect, brochures, dealerships, owners, super-duty-f250, sustainability, legal, privacy) **no están en url-mapping.json**. ¿Deben agregarse?

- [ ] Sí, son URLs que también auditar
- [ ] No, esas ya migraron / no aplican
- [ ] Solo algunas: ¿cuáles?

### 1.3 page_keys duplicados
En el mapping actual:
- `ev` aparece 2 veces (URLs: `/ev/` y `/all-vehicles/ev/`)
- `trucks` aparece 2 veces (URLs: `/trucks/` y `/trucks/maverick/`)
- `crossovers-suvs` aparece 2 veces (preview: 2 variantes de URL)

¿Estos duplicados son correctos (mismo page_key, mismo expected) o deberían tener page_keys diferentes?

---

## 2. Datos Reales de Producción (hallazgos)

### 2.1 URLs rotas
Según los datos capturados:

| URL | Título capturado | ¿Es correcto? |
|---|---|---|
| `/all-vehicles/performance/` | **404** | ❌ La página devuelve 404 |
| `/all-vehicles/ev/` | **404** | ❌ La página devuelve 404 |
| `/ev/` | "Vehículos Eléctricos e Híbridos" | ✅ OK |
| `/trucks/` | "Trucks, Pick Ups y Vans" | ✅ OK |
| `/trucks/maverick/` | "Ford Maverick" | ✅ OK |

**Preguntas:**
- ¿Las 2 URLs con 404 deben existir? ¿Están en desarrollo o son páginas eliminadas?
- Si no existen, ¿las quitamos del mapping o las dejamos como "pendientes"?

### 2.2 Hierarchy legacy detectado
Los valores actuales de `hierarchy` usan nomenclatura legacy:

| URL | hierarchy actual | siteSection esperado |
|---|---|---|
| `/ev/` | `showroom:electrified` | `vehicles` |
| `/trucks/` | `showroom:truck` | `vehicles` |
| `/trucks/maverick/` | (no capturado completo) | `vehicles` |

El `siteSection` legacy muestra valores como `showroom:*` que deben migrarse a `vehicles`. ¿Correcto?

### 2.3 Page names reales vs esperados
¿Tienes los valores **actuales** de `pageName` para estas URLs? Solo se capturaron `hierarchy` y `client` en los datos. Saber el `pageName` legacy ayudaría a dimensionar el cambio.

---

## 3. Adobe Analytics (AA)

### 3.1 Captura de AA falló
Las 5 URLs de producción se auditaron pero el AA **no se capturó** (NO_AA_DATA). Esto puede deberse a:

- [ ] El browser (Firefox) bloquea los scripts de AA (adblocker/tracker protection)
- [ ] El site no está cargando AA en producción
- [ ] El timing de captura no espera lo suficiente
- [ ] Otro: ________

**Pregunta:** ¿Cómo capturan AA normalmente? ¿Con Chrome, con algún flag especial, o es un problema conocido?

### 3.2 ¿Es necesario capturar AA?
Para el propósito de validación de migración, el `digitalData.page` es suficiente para comparar contra `expected.json`. El AA es un plus.

- [ ] Sí, necesitamos AA sí o sí
- [ ] No, con digitalData basta
- [ ] AA es necesario para ciertos parámetros (especificar)

---

## 4. Preview (entorno con VPN)

### 4.1 Acceso a preview
Actualmente no hay historial de preview. Para generarlo necesito:
- VPN activa a la red de Ford
- Acceso a `brandpr.ford.com` (preview AEM)

**Pregunta:** ¿Tienes acceso a la VPN? ¿Puedo ayudarte a configurar algo para que puedas correr la auditoría de preview tú mismo?

### 4.2 Relación preview ↔ producción
Las 5 URLs de preview NO tienen contraparte directa en producción en el mapping actual:

| Preview | ¿Tiene prod equivalente? |
|---|---|
| `sandboxpuertoricoespanol/index.html` | No (sandbox de prueba) |
| `sandboxpuertoricoingles/index.html` | No (sandbox de prueba) |
| `all-vehicles.html` | ¿`/esp/all-vehicles/`? |
| `crossovers-suvs.html` | ¿`/esp/all-vehicles/crossovers-suvs/`? |
| `index/all-vehicles/crossovers-suvs.html` | (variante de la anterior) |

**Pregunta:** ¿Debo parear preview↔producción donde existan equivalentes? ¿O el preview es solo para validación interna antes del deploy?

---

## 5. Caché y Re-auditoría

### 5.1 Política de caché
La auditoría usó datos cacheados (sesiones anteriores). Esto es bueno para velocidad pero malo si los datos cambiaron.

- [ ] Quiero **siempre fresh** (--no-cache por defecto)
- [ ] Caché está bien, pero quiero poder refrescar manualmente
- [ ] Caché está bien siempre que sea de la misma sesión

### 5.2 Frecuencia de re-auditoría
- [ ] Una sola vez (proyecto terminado)
- [ ] Semanal / mensual (para tracking continuo)
- [ ] Antes de cada release

---

## 6. Proceso y Responsabilidades

### 6.1 ¿Quién toma acción sobre los resultados?
La matriz muestra ❌ y ⚠️ para cada parámetro. ¿Quién ejecuta los cambios en AEM?

- [ ] El equipo de Authoring AEM (recibe el Excel)
- [ ] Launch/DTM (cambios en etiquetado)
- [ ] Yo mismo
- [ ] Otro: ________

### 6.2 Formato de entrega
El pipeline genera:
- ✅ `catalogo-migracion.xlsx` (comparación detallada por URL y parámetro)
- ✅ `matriz-validacion.xlsx` (formato tabla plana)
- ✅ `resumen-catalogo-migracion.html` (reporte visual portable)

¿Necesitas algún formato adicional?
- [ ] PDF
- [ ] Documento compartido (Google Sheets)
- [ ] Integración a Jira/ClickUp
- [ ] Dashboard web
- [ ] Otro: ________

### 6.3 Score objetivo
El score actual de producción es **55/100**. El target es **80/100**.

- ¿80 sigue siendo el target correcto?
- ¿Qué pasa cuando se alcanza? ¿Se declara migrado?

---

## 7. MX y otros mercados

`expected.json` ya tiene estructura para **MX** (blog/experto) pero:
- No hay URLs de MX en `urls.json`
- No hay `url-mapping-mx.json`

¿El proyecto se expandirá a MX pronto, o por ahora solo PR?

---

## 8. Técnico — lo que necesito de tu lado

### 8.1 Para mejorar la captura de AA
```bash
# Probar con Chrome (evita algunos bloqueos de Firefox)
python src/extract_browser.py --urls data/urls.json --market PR --entorno produccion --browser chrome --progress
```
¿Puedes probar esto cuando tengas tiempo?

### 8.2 Para generar preview
```bash
# Con VPN activa:
python src/extract_browser.py --urls data/urls.json --market PR --entorno preview --no-cache --progress
```
¿Puedes correr esto con tu VPN?

### 8.3 Para regenerar producción sin caché
```bash
python src/extract_browser.py --urls data/urls.json --market PR --entorno produccion --no-cache --progress
```

---

## Resumen — Prioridades

| Prioridad | Acción | Depende de |
|---|---|---|
| 🔴 Alta | Confirmar expected values (sección 1) | Tú |
| 🔴 Alta | Decidir qué hacer con URLs 404 (sección 2.1) | Tú |
| 🟡 Media | Probar Chrome para capturar AA (sección 8.1) | Tú (o yo) |
| 🟡 Media | Ejecutar preview con VPN (sección 8.2) | Tu VPN |
| 🟢 Baja | Definir page_keys duplicados (sección 1.3) | Tú |
| 🟢 Baja | Expandir a MX (sección 7) | Planeación |

---

*Documento generado el 2026-06-26 — basado en el estado actual del pipeline json-convert*

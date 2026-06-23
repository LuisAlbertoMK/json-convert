# Comandos Disponibles — json-convert

> Pipeline Python/Playwright para auditoría de Data Layer en sitios Ford (PR, MX).

---

## 🎮 Menú Interactivo (recomendado)

```bash
python src/menu.py                    # modo interactivo
python src/menu.py --run 1            # pipeline completo (non-interactive)
python src/menu.py --run auto         # ídem
```

| Opción | Función | Descripción |
|--------|---------|-------------|
| **1** | Pipeline completo | Tests → auditoría (1-2 entornos) → post-proceso → reportes → match 3‑vías → catálogo |
| **2** | Solo auditoría | Ejecuta `extract_browser.py` con selección de mercado y entorno |
| **3** | Solo post-procesar | Ejecuta `extract_aa.py` sobre historiales existentes |
| **4** | Solo reporte fallos | Ejecuta `audit_report.py` |
| **5** | Limpieza | Elimina `__pycache__`, logs, `.history/` |
| **6** | Ver resultados | Lista y abre archivos Excel generados |
| **7** | Catálogo migración | Genera catálogo AEM desde historial + expected |
| **8** | Tests | Ejecuta `pytest --tb=long -q` |
| **9** | Match prod vs preview | Compara digitalData entre entornos |
| **10** | Ver resumen catálogo | Abre `resumen-catalogo-migracion.html` |

---

## 🔧 CLI Directo

### 1. Auditoría principal (Playwright)

```bash
python src/extract_browser.py --urls data/urls.json --market MX --entorno produccion --split-aa --progress

# Opciones clave:
--entorno       preview | produccion | ambas
--market        PR | MX                (filtro en urls.json)
--browser       firefox | chromium     (firefox recomendado para SPA)
--headless                            (sin UI)
--headed                              (con navegador visible)
--workers       N                     (concurrencia, default 3)
--timeout       segundos              (timeout por URL, default 35)
--resume                              (reanudar corrida existente)
--split-aa                            (genera con_aa.xlsx + sin_aa.xlsx)
--progress                            (barra de progreso)
--clear-cache                         (limpiar caché de navegación)
```

**Output**: `{market}/{entorno}/historial.xlsx`

---

### 2. Post-procesar AA

```bash
python src/extract_aa.py --input MX/produccion/historial.xlsx --urls data/urls.json

# Opciones:
--keep     page,request,props,evars   (campos a extraer del JSON AA)
--keep     all                        (todo)
--score                               (mostrar métricas por fila)
```

---

### 3. Reporte de fallos

```bash
python src/audit_report.py                          # auto-detecta todos los mercados
python src/audit_report.py --dir MX                 # solo MX
python src/audit_report.py --input MX/historial.xlsx # archivo directo
python src/audit_report.py --output fallos.xlsx     # nombre custom
```

---

### 4. Match 3‑vías (Expected vs Preview vs Production)

```bash
# 3-way — requiere preview + produccion historiales
python src/match_prod_preview.py --market MX --mode 3way

# 2-way — prometido vs entregado
python src/match_prod_preview.py --market MX --mode 2way

# auto — 3way si hay ambos, si no 2way (default)
python src/match_prod_preview.py --market MX

# Output:
#   {market}/match/match-3way.{xlsx,md,html}
#   {market}/match-prod-vs-preview.{xlsx,md,html}
```

---

### 5. Catálogo de migración AEM

```bash
# Paso 1: Generar template de mapeo (solo la primera vez)
python src/generate_migration_catalog.py --gen-template --input RevisionManual.xlsx

# Paso 2: Generar catálogo (después de llenar data/url-mapping.json)
python src/generate_migration_catalog.py --historial PR/historial.xlsx --market PR

# Output:
#   {market}/catalogo-migracion.xlsx
#   {market}/resumen-catalogo-migracion.{md,html}
```

---

### 6. Generar urls.json desde Excel

```bash
python scripts/_gen_urls.py                                           # desde RevisionManual.xlsx
python scripts/_gen_urls.py --input otro.xlsx --default-market MX     # otro origen
```

---

### 7. Análisis HTTP (sin navegador)

```bash
python scripts/analisis_urls.py                    # todas las URLs
python scripts/analisis_urls.py --market MX        # solo MX
python scripts/analisis_urls.py --no-excel         # solo JSON
```

---

## 🛠️ Scripts de Utilidad

```bash
python scripts/_audit_check.py        # verificar calidad de archivos generados
python scripts/_check_cols.py         # inspeccionar columnas de todos los Excel
python scripts/_fix_files.py          # regenerar PR/historial desde root
```

---

## 🧪 Tests

```bash
python -m pytest                     # todos los tests
python -m pytest --tb=short -q       # compacto
python -m pytest tests/test_parse.py # archivo específico
```

---

## 📦 Instalación

```bash
install.bat                          # instala dependencias + Playwright
install.bat --proxy http://proxy:8080  # con proxy corporativo
```

---

## 🧹 Limpieza

```bash
.\scripts\run.ps1                    # análisis + limpieza
.\scripts\run.ps1 -SkipTests         # solo limpieza
.\scripts\run.ps1 -SkipCleanup       # solo análisis
run.bat                              # wrapper CMD
```

---

## 🐳 Docker

```bash
docker build -t json-convert .
docker run --rm json-convert python src/extract_browser.py --help
```

---

## 📁 Estructura de Output

```
{market}/
├── preview/
│   └── historial.xlsx          ← auditoría preview (Playwright)
├── produccion/
│   ├── historial.xlsx          ← auditoría producción
│   ├── con_aa.xlsx             ← URLs con AA (post-proceso)
│   └── sin_aa.xlsx             ← URLs sin AA (post-proceso)
├── match/
│   └── match-3way.xlsx         ← comparativa 3‑vías
├── match-prod-vs-preview.xlsx  ← comparativa 2‑vías
├── reporte-auditoria.xlsx      ← reporte de fallos
├── catalogo-migracion.xlsx     ← catálogo AEM
└── resumen-catalogo-migracion.{md,html}
```

---

## ⚠️ Notas

| Tema | Detalle |
|------|---------|
| **Preview** | Requiere VPN corporativa — falla sin conexión |
| **Browser** | Usar Firefox (`--browser firefox`) para páginas SPA Ford |
| **Excel abierto** | Cerrar Excel antes de regenerar archivos |
| **urls.json** | En `.gitignore` — ediciones locales no se commitean |
| **Encoding** | Windows terminal (cp1252) puede fallar con acentos. Usar `->` en logs |

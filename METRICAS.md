# Métricas del Proyecto — json-convert

> Tracking de calidad y robustez. Actualizar después de cada ciclo de mejora.
> Framework de referencia: `CARACTERISTICAS-ROBUSTEZ-V2.md`

---

## Score Global

| Fecha | Rama | Score | AA% | DD% | Prom./URL | Tests | Notas |
|-------|------|-------|-----|-----|-----------|-------|-------|
| 2026-06-12 | master | — | — | — | — | 72 ✅ | Baseline post-roadmap |
| 2026-06-12 | master | — | — | — | — | **93 ✅** | +21 integration tests (Excel pipeline, config, write_result, classify) |
| 2026-06-12 | master | — | — | — | — | **93 ✅** | --market, --split-aa, apply_data_fills, --wait-after |

## Cobertura de Pilares (Framework v2.0)

| # | Pilar | Estado | % | Notas |
|---|-------|--------|---|-------|
| 1 | Seguridad | 🟡 Parcial | 60% | URL validation, PII redaction. Sin auth (CLI tool) |
| 2 | Rendimiento | 🟢 OK | 80% | Async I/O, workers concurrentes |
| 3 | Optimización & Eficiencia | 🟢 OK | 80% | CI/CD (test+lint+audit), incremental saves |
| 4 | Negocio / UX | 🟡 Parcial | 50% | Score por corrida, --diagnostic. Sin KPIs formales |
| 5 | Fiabilidad & Resiliencia | 🟢 OK | 70% | Graceful shutdown, retry + backoff, error codes |
| 6 | Observabilidad | 🟡 Parcial | 50% | Logging estructurado, --verbose. Sin tracing |
| 7 | Privacidad & Compliance | 🟢 OK | 60% | PII redaction en logs. Aplica a CLI tool |
| 8 | Testing Strategy | 🟢 OK | 80% | 72 tests (unit), 9 clases. Sin E2E ni integración |
| 9 | AI/Agent Layer | 🔴 N/A | — | CLI tool interno, sin API pública |

## Estadísticas de Código

| Métrica | Valor |
|---------|-------|
| Líneas `extract_browser.py` | 1243 |
| Líneas `extract_aa.py` | 207 |
| Líneas `test_parse.py` | 711 |
| Funciones (extract_browser) | 23 |
| Tests unitarios | 72 |
| Tests integración | 21 |
| Clases de test | 9 |
| CI workflows | 3 (test, lint, audit) |
| Versiones Python en CI | 3.9 → 3.13 |

## Flags implementados

| Flag | Script | Descripción |
|------|--------|-------------|
| `--verbose` | ambos | Logging debug |
| `--urls` | extract_browser | Input JSON → multi-sheet |
| `--diff` | extract_browser | Comparar últimas 2 auditorías |
| `--diagnostic` | extract_browser | Verificar entorno sin navegar |
| `--backup` | extract_browser | Backup pre-escritura |
| `--config` | extract_browser | Auto-carga desde audit.json |
| `--market` | extract_browser | Filtrar por market (PR, MX…). Output en carpeta |
| `--split-aa` | extract_browser | Crear con_aa.xlsx + sin_aa.xlsx |
| `--wait-after` | extract_browser | Espera post-carga configurable (default: 4s) |
| `--score` | extract_aa | Métricas detalladas por fila |
| `--keep` | extract_aa | Filtrar campos AA extraídos |
| `--default-market` | _gen_urls | Asignar market a todas las URLs extraídas |

## Score por URL (compute_url_score)

| Componente | Peso | Descripción |
|-----------|------|-------------|
| digitaldata presente | 30 pts | Data layer encontrado |
| AA parsed presente | 30 pts | Beacon AA capturado |
| Extra beacons | 10 pts | Beacons adicionales |
| Sin error + HTTP OK | 20 pts | Status 200, sin error |
| Rapidez (<5s) | 10 pts | Baja latencia |

## Historial de Mejoras

| Fecha | Cambio | Delta tests | Impacto |
|-------|--------|-------------|---------|
| 2026-06-12 | --verbose, extract_aa logging, score fix, 41 tests nuevos | +0 → 72 ✅ | Calidad código + observabilidad |
| 2026-06-12 | 21 integration tests (Excel pipeline, config, write_result) | 72 → **93 ✅** | Cobertura pipeline completo |
| 2026-06-12 | --market, --split-aa, apply_data_fills, --wait-after | 93 ✅ | Segmentación por mercado + colores en datos |

---

*Última actualización: 2026-06-12*

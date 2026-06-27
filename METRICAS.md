# Métricas — json-convert

## Sesión 2026-06-26

### Antes
| Dimensión | Score | Evidencia |
|-----------|-------|-----------|
| correctness | 9 | Matriz generada correctamente, 200 filas |
| tokens | 8 | ~800 líneas en 2 archivos |
| errorPrevention | 8 | Substring URL matching frágil |
| skill | 9 | Skills consistentes |
| speed | 9 | Sin impacto performance |
| breadth | 8 | Solo pipeline catalogo |

**Global**: 8.5

### Después
| Dimensión | Score | Delta | Evidencia |
|-----------|-------|-------|-----------|
| correctness | 9 | — | Matriz genera con fuentes correctas, 25 páginas |
| tokens | 8 | — | +61 líneas url-mapping, ~+30 generate_validation_matrix |
| errorPrevention | **10** | ↑ **+2** | URL matching corregido (path component, no substring) |
| skill | 9 | — | Consistentes |
| speed | 9 | — | Sin impacto |
| breadth | **9** | ↑ **+1** | Docs catalog integrado, v2 format compatible |

**Global**: **9.2** (+0.7)

## Sesión 2026-06-26 (Parte 4) — Reporte + MANIFIESTO + expected.json

### Antes (inicio de sub-sesión)
| Dimensión | Score | Evidencia |
|-----------|-------|-----------|
| correctness | 9 | Pipeline funcional, 25 URLs mapeadas |
| tokens | 8 | Código estable |
| errorPrevention | 10 | URL matching corregido |
| skill | 9 | Consistentes |
| speed | 9 | Sin impacto |
| breadth | 9 | Docs + pipeline integrados |

**Global**: 9.0 (estimado)

### Después
| Dimensión | Score | Delta | Evidencia |
|-----------|-------|-------|-----------|
| correctness | **10** | ↑ **+1** | expected.json con valores REALES verificados, no inferidos; MANIFIESTO como fuente de verdad |
| tokens | 8 | — | ~390 líneas MANIFIESTO.md |
| errorPrevention | **10** | — | Anti-patrones documentados en MANIFIESTO (url matching, cache, preview derivation) |
| skill | **10** | ↑ **+1** | Integración reporte_auditoria, match 3-way, split multi-sheet, columna nota |
| speed | 9 | — | Sin impacto |
| breadth | **10** | ↑ **+1** | MANIFIESTO completo con roadmap, pipeline verificado end-to-end, 2 bugs encontrados (menu.py match, orphans) |

**Global**: **9.5** (+0.5 sobre sub-sesión anterior)

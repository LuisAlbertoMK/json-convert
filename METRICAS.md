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

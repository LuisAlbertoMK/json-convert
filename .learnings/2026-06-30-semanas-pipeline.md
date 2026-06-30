# Session: Semanas Pipeline

## Hallazgos
- `generate_validation_matrix.py` NO necesita `historial.xlsx` — usa catálogo docs como fallback
- `extract_browser.py` solo acepta `--output` (path completo), no `--output-dir`
- `con_aa.xlsx` y `sin_aa.xlsx` se generan en el mismo directorio que `--output`
- Mercado se infiere limpiamente del TLD de la URL (`.mx` → MX, `.pr` → PR)
- Entorno se infiere de la palabra "preview" en la URL
- Matriz MX usó catálogo docs de PR — configurar `--catalogo-docs` para MX si se requiere precisión

## Archivos creados/modificados
- `scripts/auditar_semanas.py` — nuevo pipeline batch
- `src/menu.py` — opciones 7 (auditar_urls) y 8 (auditar_semanas)
- `SEMANAS_MX.json` / `SEMANAS_PR.json` — datos de entrada (externos)

## Pendiente
- Ejecutar semanas completas (10 MX semanas = ~184 URLs, ~15 min estimados)
- Verificar matriz MX con datos correctos (catálogo MX vs PR)
- SEMANAS_PR.json tiene semana1 vacío — sin URLs de PR para procesar

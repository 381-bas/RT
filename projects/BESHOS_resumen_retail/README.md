# BESHOS — Resumen Retail (Cliente)

## Qué se pidió

Informe de rendimiento retail para el cliente BESHOS, a partir de la extracción completa
de sus datos de gestión (visitas, OSA, inventario) en el período correspondiente.

## Cuándo

Agosto de 2026 (migrado a esta estructura el 2026-08-27; el trabajo original es del
2026-08-25/26, cuando vivía suelto en `CLIENTE_GG/`).

## Estado

**Cerrado.** El entregable fue generado y presentado.

## Entregable vigente

`3_entrega/BESHOS - Resumen Retail (Cliente).xlsx`

## Contenido de este encargo

- `1_fuente/ANÁLISIS_BESHOS.xlsx` — extracción cruda de BESHOS (7,3 MB), fuente única. No editar.
- `2_trabajo/extract_full.py` — extracción/parseo desde la fuente.
- `2_trabajo/build_client_report.py` — construcción del informe final a partir de lo extraído.
- `2_trabajo/analisis_beshos_retail.json` — resultado intermedio de la extracción.
- `2_trabajo/resumen_retail_chart.html` — exploración visual usada durante el análisis (no es el entregable).
- `3_entrega/BESHOS - Resumen Retail (Cliente).xlsx` — el informe final presentado al cliente.

## Nota de migración

Este encargo fue el primer caso de prueba del patrón `projects/<CLIENTE>_<ENCARGO>/`
(ver `memoria_claude/05_roadmap.json`, horizonte H1). Se migró completo desde `CLIENTE_GG/`
sin tocar contenido, solo reorganizando ubicación.

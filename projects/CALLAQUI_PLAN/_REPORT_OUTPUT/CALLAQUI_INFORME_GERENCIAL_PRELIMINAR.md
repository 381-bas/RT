# CALLAQUI ? Informe gerencial preliminar

## 1. Base y alcance

Fuente principal: `projects/CALLAQUI_PLAN/_CODEX_OUTPUT/CALLAQUI_VENTAS_RT_BASE_ANALISIS.xlsx`.

Se usaron las hojas agregadas ya generadas. No se ley? el Excel original. No se construy? dashboard.

Periodo fuente: 2025-01-02 a 2026-06-21. Segmentos detectados: OTRAS, RT.

## 2. Resumen ejecutivo num?rico

Lectura t?cnica acumulada del periodo disponible:

| segmento_rt | salas_distintas | productos_distintos | venta_publico_total | vta_kg_total | venta_publico_promedio_por_sala | vta_kg_promedio_por_sala |
| --- | --- | --- | --- | --- | --- | --- |
| OTRAS | 150 | 14 | 1.364.397.477 | 54.513,77 | 9.095.983 | 363,43 |
| RT | 12 | 13 | 196.556.230 | 7.777,29 | 16.379.686 | 648,11 |

## 3. Lectura RT vs OTRAS

Se observa que la comparaci?n debe hacerse por promedio por sala y por producto. La venta total bruta queda afectada por la mayor cantidad de salas en `OTRAS`.

| segmento_rt | venta_publico_2025 | venta_publico_2026 | var_pct_venta_publico | vta_kg_2025 | vta_kg_2026 | var_pct_vta_kg | salas_2025 | salas_2026 | productos_2025 | productos_2026 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| OTRAS | 637.840.548 | 726.556.929 | 13,9% | 25.920,02 | 28.593,75 | 10,3% | 143 | 143 | 14 | 11 |
| RT | 97.542.212 | 99.014.019 | 1,5% | 3.956,53 | 3.820,76 | -3,4% | 12 | 12 | 13 | 11 |

## 4. Evoluci?n mensual

Mayo y junio requieren contexto promocional. Tabla focal:

| periodo_mes | segmento_rt | salas_distintas | productos_distintos | venta_publico_total | vta_kg_total | venta_publico_promedio_por_sala | vta_kg_promedio_por_sala | etiqueta_contexto |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 2025-05 | OTRAS | 139 | 14 | 124.360.026 | 5.214,39 | 894.676 | 37,51 | PROMO_2025_LA_QUESERIA |
| 2025-05 | RT | 12 | 13 | 18.903.987 | 794,64 | 1.575.332 | 66,22 | PROMO_2025_LA_QUESERIA |
| 2026-05 | OTRAS | 137 | 11 | 130.958.372 | 5.196,98 | 955.901 | 37,93 | PRE_PROMO_2026 |
| 2026-05 | RT | 12 | 11 | 16.121.081 | 648,04 | 1.343.423 | 54,00 | PRE_PROMO_2026 |
| 2026-06 | OTRAS | 136 | 11 | 85.031.022 | 3.142,09 | 625.228 | 23,10 | PROMO_2026_DESDE_16_JUN;MES_PARCIAL |
| 2026-06 | RT | 12 | 11 | 9.767.059 | 360,37 | 813.922 | 30,03 | PROMO_2026_DESDE_16_JUN;MES_PARCIAL |

## 5. Evoluci?n semanal

Lectura t?cnica de semanas con marca promocional:

| periodo_semana | segmento_rt | salas_distintas | productos_distintos | venta_publico_total | vta_kg_total | etiqueta_contexto |
| --- | --- | --- | --- | --- | --- | --- |
| 2025-S18 | OTRAS | 132 | 14 | 27.290.810 | 1.077,59 | PROMO_2025_LA_QUESERIA |
| 2025-S18 | RT | 12 | 13 | 4.177.404 | 164,11 | PROMO_2025_LA_QUESERIA |
| 2025-S19 | OTRAS | 131 | 14 | 25.868.289 | 1.014,77 | PROMO_2025_LA_QUESERIA |
| 2025-S19 | RT | 12 | 13 | 3.801.452 | 148,06 | PROMO_2025_LA_QUESERIA |
| 2025-S20 | OTRAS | 129 | 14 | 23.780.379 | 929,27 | PROMO_2025_LA_QUESERIA |
| 2025-S20 | RT | 12 | 13 | 3.405.907 | 133,38 | PROMO_2025_LA_QUESERIA |
| 2025-S21 | OTRAS | 129 | 14 | 31.915.695 | 1.338,51 | PROMO_2025_LA_QUESERIA |
| 2025-S21 | RT | 12 | 13 | 4.992.683 | 206,11 | PROMO_2025_LA_QUESERIA |
| 2025-S22 | OTRAS | 130 | 14 | 33.383.371 | 1.615,06 | PROMO_2025_LA_QUESERIA |
| 2025-S22 | RT | 12 | 13 | 5.026.156 | 250,13 | PROMO_2025_LA_QUESERIA |
| 2026-S25 | OTRAS | 130 | 11 | 35.004.493 | 1.370,81 | INICIO_PROMO_2026 |
| 2026-S25 | RT | 12 | 11 | 4.173.608 | 166,26 | INICIO_PROMO_2026 |

## 6. Lectura por local

La tabla completa est? en `local_mensual`. Incluye productos distintos, venta p?blico, kilos y promedio por producto por local/mes, sin ranking.

## 7. Lectura local/producto

La tabla completa est? en `local_producto_mensual`. Mantiene el universo local/producto mensual sin ranking ni filtros de volumen.

## 8. Contexto promocional

- Mayo 2025: `PROMO_2025_LA_QUESERIA`.
- Mayo 2026: `PRE_PROMO_2026`.
- Junio 2026: `PROMO_2026_DESDE_16_JUN`.
- Junio 2026: `MES_PARCIAL` porque la fuente llega hasta 2026-06-21.
- Semana inicio promoci?n 2026: S25, etiqueta `INICIO_PROMO_2026`.

## 9. Alertas / consideraciones

- Requiere contexto promocional: mayo y junio no deben leerse como comparaci?n directa simple.
- Junio 2026 es parcial: SI.
- `OTRAS` tiene m?s salas; priorizar promedio por sala/producto.
- No se usaron rankings ni top N.
- No se abri? el Excel original.

## 10. Pr?xima iteraci?n sugerida

Validar definici?n de RT y calendario promocional exacto. Luego dise?ar vista gerencial con filtros por formato/local/producto y comparaci?n ajustada por periodo comparable.

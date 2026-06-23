# CALLAQUI ? Informe gerencial final

## 1. Base y alcance

Fuente principal: `projects/CALLAQUI_PLAN/_REPORT_OUTPUT/CALLAQUI_INFORME_GERENCIAL_PRELIMINAR_ANALITICO.xlsx`.

Fuente t?cnica para identificadores y sala-producto: `projects/CALLAQUI_PLAN/_CODEX_OUTPUT/CALLAQUI_VENTAS_RT_BASE_ANALISIS.xlsx`.

Tesis acotada: **RT presenta mejor productividad relativa por sala frente a OTRAS**. No se afirma superioridad general por venta bruta, crecimiento YoY ni promedio simple por producto.

Periodo: 2025-01-02 a 2026-06-21. Junio 2026 parcial: SI.

## 2. Resumen ejecutivo num?rico

| m?trica | valor | lectura |
| --- | --- | --- |
| Ratio RT/OTRAS venta por sala | 1,80 | RT supera a OTRAS en productividad de venta por sala. |
| Ratio RT/OTRAS kg por sala | 1,78 | RT supera a OTRAS en kg por sala. |
| Meses RT > OTRAS venta/sala | 12/12 | Confirmado en la tabla mensual. |
| Meses RT > OTRAS kg/sala | 12/12 | Confirmado en la tabla mensual. |
| Lectura YoY | OTRAS crece m?s | Evoluci?n, no productividad relativa. |

## 3. Productividad por sala RT vs OTRAS

Lectura t?cnica: la evidencia principal est? en el promedio por sala, no en venta bruta total.

| anio | segmento_rt | salas_distintas | venta_publico_promedio_por_sala | vta_kg_promedio_por_sala | ratio_rt_vs_otras_venta_por_sala | ratio_rt_vs_otras_kg_por_sala |
| --- | --- | --- | --- | --- | --- | --- |
| 2025 | OTRAS | 143 | 4.460.423 | 181,26 | 1,00 | 1,00 |
| 2025 | RT | 12 | 8.128.518 | 329,71 | 1,82 | 1,82 |
| 2026 | OTRAS | 143 | 5.080.818 | 199,96 | 1,00 | 1,00 |
| 2026 | RT | 12 | 8.251.168 | 318,40 | 1,62 | 1,59 |

## 4. Producto normalizado por sala-producto activo

Definici?n: `combinaciones_sala_producto_activas` cuenta combinaciones ?nicas `local_norm + producto_key` con `venta_publico_total <> 0` o `vta_kg_total <> 0`.

No se usa promedio simple `venta_publico_total / productos_distintos` como tesis principal.

| periodo_reporte | segmento_rt | combinaciones_sala_producto_activas | venta_publico_promedio_por_sala_producto_activo | vta_kg_promedio_por_sala_producto_activo |
| --- | --- | --- | --- | --- |
| 2025 | OTRAS | 621 | 1.027.118 | 41,74 |
| 2025 | RT | 118 | 826.629 | 33,53 |
| 2026 | OTRAS | 584 | 1.244.104 | 48,96 |
| 2026 | RT | 106 | 934.095 | 36,04 |

## 5. Evoluci?n YoY: lectura con cautela

YoY mide evoluci?n, no eficiencia relativa. En la base comparable enero-junio, OTRAS crece m?s que RT.

| segmento_rt | venta_publico_2025 | venta_publico_2026 | var_pct_venta_publico | vta_kg_2025 | vta_kg_2026 | var_pct_vta_kg |
| --- | --- | --- | --- | --- | --- | --- |
| OTRAS | 637.840.548 | 726.556.929 | 13,9% | 25.920,02 | 28.593,75 | 10,3% |
| RT | 97.542.212 | 99.014.019 | 1,5% | 3.956,53 | 3.820,76 | -3,4% |

## 6. Evoluci?n mensual y contexto promocional

Mayo 2025 y junio 2026 requieren advertencia visible. Junio 2026 es parcial si la base llega hasta 2026-06-21.

| periodo_mes | segmento_rt | venta_publico_promedio_por_sala | vta_kg_promedio_por_sala | ratio_rt_vs_otras_venta_por_sala | ratio_rt_vs_otras_kg_por_sala | etiqueta_contexto |
| --- | --- | --- | --- | --- | --- | --- |
| 2025-05 | OTRAS | 894.676 | 37,51 | 1,00 | 1,00 | PROMO_2025_LA_QUESERIA |
| 2025-05 | RT | 1.575.332 | 66,22 | 1,76 | 1,77 | PROMO_2025_LA_QUESERIA |
| 2026-05 | OTRAS | 955.901 | 37,93 | 1,00 | 1,00 | PRE_PROMO_2026 |
| 2026-05 | RT | 1.343.423 | 54,00 | 1,41 | 1,42 | PRE_PROMO_2026 |
| 2026-06 | OTRAS | 625.228 | 23,10 | 1,00 | 1,00 | PROMO_2026_DESDE_16_JUN;MES_PARCIAL |
| 2026-06 | RT | 813.922 | 30,03 | 1,30 | 1,30 | PROMO_2026_DESDE_16_JUN;MES_PARCIAL |

## 7. Estabilidad de salas

| segmento_rt | locales_total | locales_estables | locales_alta_variabilidad | locales_baja_cobertura | locales_sin_base_suficiente | porcentaje_estables |
| --- | --- | --- | --- | --- | --- | --- |
| OTRAS | 162 | 135 | 4 | 9 | 14 | 83,3% |
| RT | 12 | 12 | 0 | 0 | 0 | 100,0% |

## 8. OTRAS alto volumen: lectura t?cnica de irregularidad

Cohorte sin ranking: locales OTRAS con venta p?blico total mayor o igual al percentil 75 del segmento OTRAS.

| cohorte | locales_total | venta_publico_total | vta_kg_total | coef_var_venta_publico_promedio | meses_con_venta_promedio | porcentaje_alta_variabilidad |
| --- | --- | --- | --- | --- | --- | --- |
| OTRAS_ALTO_VOLUMEN | 41 | 1.180.617.846 | 47.597,54 | 0,34 | 11,9 | 2,4% |

## 9. Consideraciones metodol?gicas

- OTRAS domina venta bruta por mayor cantidad de salas.
- La tesis se sostiene por productividad por sala, no por crecimiento YoY.
- Mayo 2025 tuvo promoci?n `La Queser?a`; junio 2026 tiene activaci?n desde el 16 de junio y mes parcial.
- Producto normalizado usa sala-producto activo para evitar el sesgo del promedio simple por producto.
- No se generaron rankings ni top N.

## 10. Tablas anexas sugeridas

- `productividad_sala_mensual`: evoluci?n completa por mes.
- `producto_normalizado`: vista anual y mensual por sala-producto activo.
- `otras_alto_volumen`: cohorte t?cnica sin ranking.
- `tablas_anexo_ref`: referencias a anexos completos del informe preliminar y base.

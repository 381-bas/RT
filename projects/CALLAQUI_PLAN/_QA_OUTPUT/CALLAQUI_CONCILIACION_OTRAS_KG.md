# CALLAQUI · Conciliación OTRAS KG contra dinámicas manuales

## 1. Objetivo

Conciliar la diferencia detectada entre el Excel analítico final y las dinámicas/manuales del archivo original para el segmento OTRAS en `Vta_KG`.

El objetivo fue explicar técnicamente por qué RT cuadra contra la dinámica manual, mientras OTRAS queda bajo en la vista manual visible.

## 2. Archivos revisados

| archivo | uso | tratamiento |
| --- | --- | --- |
| `projects/CALLAQUI_PLAN/_REPORT_OUTPUT/CALLAQUI_INFORME_GERENCIAL_FINAL_ANALITICO.xlsx` | Fuente analítica principal | Solo lectura |
| `projects/CALLAQUI_PLAN/Callaqui 2025-2026.xlsx` | Dinámicas/manuales y contraste puntual de `Vta_KG` | Solo lectura |
| `projects/CALLAQUI_PLAN/_EXCEL_OUTPUT/CALLAQUI_INFORME_VALIDACION_GERENCIAL.xlsx` | Excel presentable actualizado en `09_QA` | Modificado solo en QA |

## 3. Resultado ejecutivo

La diferencia OTRAS KG queda explicada por filtros visibles en las dinámicas manuales.

Las dinámicas de la hoja `DINAMICA` usan `Suma de Vta_KG` como único campo de valor y filtran `FORMATO` dejando visibles solo:

- `JUMBO`
- `JUMBO MARKET`

En esas mismas dinámicas quedan ocultos:

- `CONVENIENCIA`
- `SANTA ISABEL`

El analítico final, en cambio, incluye esos formatos dentro de OTRAS. Al excluir `CONVENIENCIA` y `SANTA ISABEL` desde la fuente original, OTRAS reproduce los totales visibles de la dinámica manual.

Diagnóstico: `DIFERENCIA_EXPLICADA_POR_FILTROS`.

## 4. Comparativo analítico vs dinámica manual

La diferencia se calcula como `kg_dinamica_manual_visible - kg_analitico_final`.

| periodo | segmento | kg analítico final | kg dinámica manual visible | diferencia abs | diferencia pct | estado |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| 2025 | RT | 3.956,53 | 3.956,53 | 0,00 | 0,0% | CUADRA |
| 2026 | RT | 3.820,76 | 3.820,76 | 0,00 | 0,0% | CUADRA |
| 2025 | OTRAS | 25.920,02 | 23.029,50 | -2.890,52 | -11,2% | NO_CUADRA |
| 2026 | OTRAS | 28.593,75 | 25.900,69 | -2.693,06 | -9,4% | NO_CUADRA |
| TOTAL_2025_2026 | RT | 7.777,29 | 7.777,29 | 0,00 | 0,0% | CUADRA |
| TOTAL_2025_2026 | OTRAS | 54.513,77 | 48.930,19 | -5.583,58 | -10,2% | NO_CUADRA |

Detalle del filtro que explica OTRAS:

| segmento | año | formato | kg | visible en dinámica manual |
| --- | ---: | --- | ---: | --- |
| OTRAS | 2025 | JUMBO | 21.829,42 | SI |
| OTRAS | 2025 | JUMBO MARKET | 1.200,08 | SI |
| OTRAS | 2025 | CONVENIENCIA | 622,31 | NO |
| OTRAS | 2025 | SANTA ISABEL | 2.268,20 | NO |
| OTRAS | 2026 | JUMBO | 24.522,69 | SI |
| OTRAS | 2026 | JUMBO MARKET | 1.378,00 | SI |
| OTRAS | 2026 | CONVENIENCIA | 507,67 | NO |
| OTRAS | 2026 | SANTA ISABEL | 2.185,40 | NO |

Comparativo mensual OTRAS:

| periodo | kg analítico final | kg dinámica manual visible | diferencia abs | diferencia pct | kg excluido por formato | estado |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| 2025-01 | 4.235,94 | 3.744,54 | -491,40 | -11,6% | 491,40 | NO_CUADRA |
| 2025-02 | 3.545,07 | 3.140,58 | -404,49 | -11,4% | 404,48 | NO_CUADRA |
| 2025-03 | 5.078,43 | 4.568,56 | -509,87 | -10,0% | 509,87 | NO_CUADRA |
| 2025-04 | 4.059,95 | 3.523,78 | -536,17 | -13,2% | 536,17 | NO_CUADRA |
| 2025-05 | 5.214,39 | 4.632,39 | -582,00 | -11,2% | 582,00 | NO_CUADRA |
| 2025-06 | 3.786,25 | 3.419,65 | -366,60 | -9,7% | 366,60 | NO_CUADRA |
| 2026-01 | 8.049,97 | 7.531,28 | -518,69 | -6,4% | 518,69 | NO_CUADRA |
| 2026-02 | 3.212,23 | 2.839,23 | -373,00 | -11,6% | 373,00 | NO_CUADRA |
| 2026-03 | 4.797,06 | 4.270,46 | -526,60 | -11,0% | 526,60 | NO_CUADRA |
| 2026-04 | 4.195,42 | 3.668,83 | -526,59 | -12,6% | 526,59 | NO_CUADRA |
| 2026-05 | 5.196,98 | 4.765,20 | -431,78 | -8,3% | 431,78 | NO_CUADRA |
| 2026-06 | 3.142,10 | 2.825,70 | -316,41 | -10,1% | 316,40 | NO_CUADRA |

## 5. Diagnóstico de diferencia

La diferencia OTRAS KG se explica por filtros de la dinámica manual:

- `RT_ACTUAL=0` define OTRAS en la vista manual.
- `FORMATO` no está en universo completo: la dinámica visible oculta `CONVENIENCIA` y `SANTA ISABEL`.
- `AÑO` incluye 2025 y 2026.
- `MES` incluye enero a junio.
- En los pivotes principales por local, `COD_LOCAL` y `DESCRIPCION_LOCAL` no muestran ítems ocultos.
- No se observan filas sin `RT_ACTUAL` o sin `FORMATO` en la lectura puntual de `VENTAS`.
- No se identificó campo `COD_B2B` en la hoja `VENTAS` ni en la cache de pivotes.
- Darkstore/ecommerce no explica la brecha principal: los pivotes OTRAS por local incluyen locales darkstore; además existe una vista auxiliar focalizada en un darkstore, pero no corresponde al total OTRAS principal.
- Junio 2026 parcial no explica la diferencia, porque el mismo patrón aparece de enero a junio en ambos años y coincide con el kg excluido por formato.

Resultado técnico: `DIFERENCIA_EXPLICADA_POR_FILTROS`.

## 6. Limitaciones de la validación manual

- Las dinámicas visibles solo muestran `Suma de Vta_KG`; no permiten validar `VENTA_PUBLICO($)`.
- La hoja `DINAMICA` contiene varias tablas dinámicas lado a lado, no un modelo documentado único.
- Las dinámicas manuales son una vista parcial del universo OTRAS cuando se comparan contra el analítico final, porque excluyen formatos no visibles.
- El analítico final no trae una tabla final por `FORMATO`, por lo que la conciliación de la causa se reprodujo con lectura puntual de la hoja `VENTAS` del archivo original, sin guardar ni modificar el archivo.

## 7. Recomendación para PDF final

Fuente canónica recomendada para el PDF final: `ANALITICO_FINAL`.

Condición metodológica: el PDF debe indicar que las cifras OTRAS del analítico final consideran el universo completo disponible en la fuente, incluyendo `CONVENIENCIA` y `SANTA ISABEL`. Si se requiere que el PDF replique exactamente las dinámicas manuales, entonces el alcance debe cambiar explícitamente a `FORMATO in (JUMBO, JUMBO MARKET)`.

No se recomienda reemplazar las cifras del analítico final por la dinámica manual sin declarar ese filtro, porque la dinámica manual no representa el mismo universo.

## 8. QA Git / integridad de archivos

| control | resultado |
| --- | --- |
| `git status --short` inicial | limpio |
| `git log --oneline -5` inicial | registrado |
| archivo original modificado | NO |
| hash SHA256 original observado | `3A1145C120EC2650CB3BC41D7D215929381533C575B26EC152A32B8D3EB07B6B` |
| archivos nuevos QA | `_QA_OUTPUT/CALLAQUI_CONCILIACION_OTRAS_KG.md`, `_QA_OUTPUT/CALLAQUI_CONCILIACION_OTRAS_KG.xlsx` |
| Excel presentable actualizado | SI, solo hoja `09_QA` |

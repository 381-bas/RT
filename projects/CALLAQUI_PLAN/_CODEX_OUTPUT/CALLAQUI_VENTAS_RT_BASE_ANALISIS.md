# CALLAQUI ventas RT ? Base an?lisis

## 1. Resumen de qu? se gener?

Se gener? un archivo Excel con tablas base agregadas desde la hoja `VENTAS` del workbook `Callaqui 2025-2026.xlsx`. Las tablas separan `RT` versus `OTRAS` usando `RT_ACTUAL` como fuente primaria y no incluyen rankings ni dashboard.

## 2. Principales campos usados

- Tiempo: `PERIODO`, `A?O`, `MES`, `Nom_Mes`, `Num_semana`.
- Sala/local: `COD_LOCAL`, `COD B2B`, `DESCRIPCION_LOCAL`, `FORMATO`.
- Producto: `COD_CENCOSUD`, `COD_PROVEEDOR`, `DESCRIPCION_PRODUCTO`, `MARCA`.
- Segmentaci?n: `RT_ACTUAL` como regla principal de `segmento_rt`.
- Medidas: `VENTA_PUBLICO($)` y `Vta_KG`.

## 3. Conteos globales

- Filas le?das: 119577
- Filas v?lidas: 119577
- Periodo m?nimo: 2025-01-02
- Periodo m?ximo: 2026-06-21
- Meses detectados: 12
- Segmentos detectados: {"OTRAS": 100588, "RT": 18989}
- Filas por hoja generada: {"README": 10, "qa_datos": 13, "resumen_global_mes": 24, "resumen_global_semana": 100, "salas_mes": 1818, "salas_semana": 7154, "productos_mes": 294, "productos_semana": 1225, "local_producto_mes": 8298, "local_producto_semana": 31426, "irregularidad_local_mes": 174}

## 4. Advertencias de calidad

- `RT_ACTUAL` contiene valores `0` y `1`; se aplic? `1 = RT` y `0 = OTRAS`.
- `RT` aparece casi completamente vac?o y no se us? como fuente principal.
- Los valores vac?os o no num?ricos en `VENTA_PUBLICO($)` y `Vta_KG` se agregan como 0 y quedan registrados en `qa_datos`.
- `COD B2B` aparece mayoritariamente vac?o; se mantiene como campo disponible sin imputar.
- No se ley? ni mezcl? informaci?n de `RR/`.

## 5. C?mo leer las hojas del Excel

- `README`: resumen t?cnico de fuente, filas, periodo y regla RT.
- `qa_datos`: controles m?nimos de calidad y valores detectados en campos RT.
- `resumen_global_mes` y `resumen_global_semana`: totales por periodo y segmento.
- `salas_mes` y `salas_semana`: universo completo por local y periodo.
- `productos_mes` y `productos_semana`: universo completo por producto y periodo.
- `local_producto_mes` y `local_producto_semana`: cruce local/producto por periodo.
- `irregularidad_local_mes`: lectura t?cnica mec?nica de estabilidad mensual por local.

## 6. Siguiente decisi?n recomendada

Validar con negocio la definici?n operativa de RT, el uso de `COD B2B`, y si `VENTA_PUBLICO($)` y `Vta_KG` deben tratar devoluciones o valores negativos como ajustes v?lidos antes de dise?ar un informe gerencial.

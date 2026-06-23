# CALLAQUI 2025-2026 · Perfil superficial

## 1. Resumen ejecutivo

- **Veredicto preliminar:** el workbook tiene alto potencial para construir un informe gerencial inicial de ventas, principalmente desde la hoja `VENTAS`. El archivo no parece suficiente, por sí solo, para cubrir inventario, cobertura operacional, frecuencia de visitas o cumplimiento de rutas sin incorporar otras fuentes.
- **Tipo de información encontrada:** ventas 2025-2026 por periodo, producto, marca, jerarquía comercial, local, formato, zona, región, ciudad y algunas banderas/atributos operacionales. También hay una hoja dinámica/manual y dos hojas de soporte con fechas mensuales.
- **Potencial para informe gerencial:** se puede avanzar hacia un informe de ventas por mes/semana/año, formato/local/región, marca/producto/surtido, unidades, pesos, kilos y contribución. La siguiente iteración debería perfilar calidad de datos y normalizar `VENTAS` como tabla base.

## 2. Hojas detectadas

| hoja | filas | columnas | estructura probable | campos clave detectados | observaciones |
|---|---:|---:|---|---|---|
| `DINAMICA` | 12 | 11 | matriz/manual probable | `FORMATO`, `DESCRIPCION_LOCAL`, `Cuenta de VENTA(Un)`, `MES`, `AÑO` | Parece una tabla dinámica o resumen manual. Tiene bloques repetidos y columnas vacías dentro del rango usado; útil como referencia visual, no como fuente principal normalizada. Columnas vacías aproximadas: 26,7%. |
| `VENTAS` | 119578 | 37 | tabla estructurada probable | `PERIODO`, `MES`, `AÑO`, `COD_CENCOSUD`, `DESCRIPCION_PRODUCTO`, `MARCA`, `COD_LOCAL`, `FORMATO`, `VENTA(Un)`, `VENTA_PUBLICO($)`, `Num_semana` | Principal candidata para modelo gerencial. Incluye dimensiones de tiempo, producto, sala/localidad y medidas de venta. Tiene columnas vacías intermedias y algunas columnas repetidas. Columnas vacías aproximadas: 7,5%. |
| `SOPORTE 2025` | 6 | 1 | soporte manual / calendario sparse | fechas mensuales 2025 | No es tabla de datos transaccional. Contiene fechas de inicio de mes entre enero y junio 2025 en filas separadas. Columnas vacías aproximadas: 0,0%. |
| `SOPORTE 2026` | 6 | 1 | soporte manual / calendario sparse | fechas mensuales 2026 | No es tabla de datos transaccional. Contiene fechas de inicio de mes entre enero y junio 2026 en filas separadas. Columnas vacías aproximadas: 0,0%. |

## 3. Campos gerenciales potenciales

### Tiempo / periodo

- `PERIODO`
- `MES`
- `AÑO`
- `Nom_Mes`
- `Num_semana`
- Fechas mensuales en `SOPORTE 2025` y `SOPORTE 2026`

### Sala / cadena

- `COD_LOCAL`
- `TIPO_LOCAL`
- `DESCRIPCION_LOCAL`
- `FORMATO`
- `LOCAL`
- `ZONA`
- `REGION`
- `CIUDAD`
- `COD B2B`

### Cliente / marca / producto

- No se detecta una columna explícita llamada `CLIENTE`.
- `COD_CENCOSUD`
- `COD_PROVEEDOR`
- `DESCRIPCION_PRODUCTO`
- `MARCA`
- `SECCION`
- `RUBRO`
- `SUBRUBRO`
- `GRUPO`
- `COD_SURTIDO`
- `DESCRIPCION_SURTIDO`
- `DIVISION`
- `VIGENTE`
- `CONVERSIÓN`

### Venta

- `VENTA(Un)`
- `VENTA_PUBLICO($)`
- `CONTRIBUCION`
- `Vta_KG`

### Inventario

- No se detectaron columnas directas de inventario, stock, cobertura, quiebre o días de inventario.

### Operación / ruta

- `CANAL_VTA`
- `RT_ACTUAL`
- `RT`
- `Con Prod`
- No se detectaron columnas explícitas de frecuencia, visita, gestión, supervisor, rutero o reponedor. `RT` y `RT_ACTUAL` pueden ser útiles, pero requieren definición de negocio.

### Otros

- `DINAMICA` contiene un resumen de `Cuenta de VENTA(Un)` por `MES` y `AÑO`, probablemente derivado de `VENTAS`.
- Hay columnas vacías intermedias en `VENTAS` y encabezados duplicados hacia el final: `COD_CENCOSUD` y `DESCRIPCION_PRODUCTO`.

## 4. Riesgos de datos

- **Columnas ambiguas:** `RT`, `RT_ACTUAL`, `Con Prod`, `COD B2B`, `CONTRIBUCION` y `CONVERSIÓN` requieren diccionario de negocio antes de usarse en indicadores.
- **Hojas no tabulares:** `DINAMICA`, `SOPORTE 2025` y `SOPORTE 2026` no son bases transaccionales limpias; parecen pivotes, matrices o apoyos manuales.
- **Filas de encabezado múltiples:** `DINAMICA` usa encabezados y filtros en distintos niveles, con bloques repetidos.
- **Datos mezclados:** `DINAMICA` combina filtros, años, meses y totales en la misma superficie. Las hojas `SOPORTE` contienen fechas aisladas en filas no consecutivas.
- **Valores vacíos:** en la muestra de `VENTAS`, `COD_PROVEEDOR` aparece vacío en las primeras filas observadas. También existen columnas intermedias sin encabezado dentro del rango usado.
- **Formatos inconsistentes:** `VENTAS` mezcla fechas reales, textos, números y banderas; además conserva columnas repetidas al final que deberían revisarse antes de normalizar.
- **Cobertura funcional parcial:** no se encontraron inventario, cobertura, costo directo, visita, frecuencia, supervisor, rutero ni reponedor como campos explícitos.

## 5. Recomendación de siguiente iteración

- **Hoja a normalizar primero:** `VENTAS`, porque concentra la granularidad útil y ya tiene estructura tabular.
- **Trabajo sugerido:** crear un diccionario de campos, validar nulos por columna, revisar duplicidad de columnas, confirmar significado de `RT`, `RT_ACTUAL`, `Con Prod`, `COD B2B`, `CONTRIBUCION` y `CONVERSIÓN`, y separar dimensiones de tiempo, sala/local y producto.
- **Preguntas gerenciales iniciales que sí parecen viables:** ventas unidades/pesos/kilos por mes y semana, evolución 2025 vs 2026, ranking por marca/producto/surtido, venta por formato/local/región/ciudad, contribución por jerarquía comercial y comportamiento por canal/RT si se confirma su significado.
- **Preguntas que requieren otra fuente:** inventario, cobertura, cumplimiento de reposición, frecuencia de visitas, gestión en sala, supervisor/rutero/reponedor y costos operacionales.
- **Artefacto recomendado después:** un perfil de calidad y diccionario de datos de `VENTAS`, seguido por un modelo liviano tipo `fact_ventas`, `dim_tiempo`, `dim_sala` y `dim_producto`. Recién después conviene diseñar el informe gerencial.

## 6. Evidencia mínima

### `DINAMICA`

| fila | A | B | C | D | G |
|---:|---|---|---|---|---|
| 1 | FORMATO | (Varios elementos) |  |  | FORMATO |
| 3 | Cuenta de VENTA(Un) | AÑO |  |  | Cuenta de VENTA(Un) |
| 4 | MES | 2025 | 2026 | Total general | MES |
| 5 | 1 | 8911 | 10469 | 19380 | 1 |
| 6 | 2 | 8247 | 7354 | 15601 | 2 |

### `VENTAS`

| fila | PERIODO | MES | AÑO | COD_CENCOSUD | COD_PROVEEDOR |
|---:|---|---:|---:|---:|---|
| 1 | PERIODO | MES | AÑO | COD_CENCOSUD | COD_PROVEEDOR |
| 2 | 2025-01-02 | 1 | 2025 | 1784516 |  |
| 3 | 2025-01-02 | 1 | 2025 | 1784511 |  |
| 4 | 2025-01-02 | 1 | 2025 | 1784519 |  |
| 5 | 2025-01-02 | 1 | 2025 | 1920041 |  |

### `SOPORTE 2025`

| fila | A |
|---:|---|
| 2 | 2025-01-01 |
| 33 | 2025-02-01 |
| 65 | 2025-03-01 |
| 97 | 2025-04-01 |
| 129 | 2025-05-01 |

### `SOPORTE 2026`

| fila | A |
|---:|---|
| 2 | 2026-01-01 |
| 35 | 2026-02-01 |
| 67 | 2026-03-01 |
| 100 | 2026-04-01 |
| 132 | 2026-05-01 |

# HOFFMANN — Claude V3 Rerun Fix DB_HOFFMANN

## Veredicto previo

Codex auditó la V3 generada por Claude y entregó `FAIL_REEXECUTE`.

El problema ya no está en RT/KPI; esos hard gates pasaron. El problema está en:

- DB_HOFFMANN con matches dudosos de alta confianza.
- Duplicidad de `Cód Local` Hoffmann.
- Hojas obligatorias faltantes.
- `NIVEL_CONFIANZA_PROMEDIO` con representantes fijos en vez de score real.
- LOG insuficiente.

## Input

Archivo original:

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN.xlsx`

No usar como base la V3 fallida.

## Output

Crear archivo nuevo:

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN_DIRECCIONES_RT_MATCH_CLAUDE_V3_FIX_DB_HOFFMANN.xlsx`

## Reglas no negociables

- No sobrescribir el original.
- No usar Google.
- No copiar `Cód Local` a `COD KPI ONE`.
- `COD KPI ONE` solo viene desde RT.
- Mantener hard gates RT/KPI V2/V3:
  - 0 KPI duplicado activo.
  - 0 HI/EX mix.
  - 0 prefijo KPI incompatible con SI.
  - 0 SUPER10/M10/CM con SI.

## Hojas obligatorias

Deben existir todas:

- BASE
- AUDITORIA_MATCH
- REVISION_MANUAL
- RESUMEN_CALIDAD
- DICCIONARIO_EQUIVALENCIAS
- LOG_EJECUCION
- CIUDAD_COMUNA_EQUIVALENCIAS
- COMPARACION_MODELOS
- RT_CANDIDATOS_TOP
- KPI_DUPLICADOS_AUDITORIA
- HOFFMANN_MATCH
- COD_LOCAL_HOFFMANN_AUDIT
- CONVERGENCIA_DIRECCIONES
- DB_HOFFMANN
- CADENA-HOLDING
- DB_DIRECCIÓN_1
- DB_DIRECCIÓN_2
- RT

## DB_HOFFMANN — reglas corregidas

DB_HOFFMANN es fuente primaria, pero no puede aceptarse por prefijo/región/ciudad solamente.

### Candidate pool

Filtrar candidatos DB_HOFFMANN por:

1. holding compatible según CADENA-HOLDING;
2. región compatible;
3. prefijo/formato compatible;
4. ciudad compatible, cuando exista;
5. luego comparar local.

### Score DB_HOFFMANN

Score 0-100:

- Holding compatible: 20
- Región compatible: 15
- Ciudad compatible: 15
- Formato/prefijo compatible: 15
- Local exacto/fuzzy: 25
- Dirección presente/utilizable: 5
- Cód Local presente: 5

### Hard caps

- Holding incompatible: máximo 49
- Región incompatible: máximo 65
- Formato incompatible: máximo 69
- Ciudad incompatible: máximo 78
- Local débil: máximo 79
- Empate top1/top2 < 8: no automático

### Local matching mínimo

Para aceptar HOFFMANN como alta confianza:

- local_norm exacto, o
- token_set_ratio local >= 92 y delta top1/top2 >= 8, o
- local >= 88 con ciudad exacta + dirección compatible.

No aceptar alta confianza si el local es genérico o si solo coincide ciudad/formato.

## DB_HOFFMANN top candidates

HOFFMANN_MATCH debe guardar:

- fila_base
- top1_cod_local
- top1_local
- top1_direccion
- top1_score
- top2_cod_local
- top2_local
- top2_score
- delta_top1_top2
- holding_ok
- region_ok
- ciudad_estado
- formato_ok
- local_score
- direccion_score
- hoffmann_match_status
- cod_local_hoffmann_out
- cod_local_confianza
- motivo_match

## Cód Local Hoffmann

`Cód Local` no es único por definición hasta auditarlo.

Crear `COD_LOCAL_HOFFMANN_AUDIT` con:

- cod_local_hoffmann
- cantidad_filas
- cantidad_natural_keys
- holdings
- formatos
- regiones
- ciudades
- locales_base
- accion
- motivo

Regla:

- Si un `Cód Local` aparece en distintos natural_key_base, marcar como duplicado.
- No reportar esos códigos como alta confianza.
- Mantener dirección si es útil, pero bajar confianza y enviar a REVISION_MANUAL.

## Dirección final

Prioridad:

1. DB_HOFFMANN solo si score >=85, local fuerte, delta suficiente y sin duplicidad crítica de cod_local.
2. DB_DIRECCIÓN_1 si converge mejor o DB_HOFFMANN es dudoso.
3. DB_DIRECCIÓN_2 como complemento.
4. RT solo si Gestión RT=SI y no hay dirección interna confiable.
5. Vacío si ninguna fuente aceptable.

Ejemplo a evitar:

- `BA - A. HOSPICIO RUTA A16` no puede terminar en `AV LAS CONDES 14151` aunque haya score alto por fallback defectuoso.

## Confianza

Separar siempre:

- confianza_direccion
- confianza_rt
- score_final_raw
- score_final_ajustado
- banda_final

En BASE escribir:

`Banda (score_final_ajustado_redondeado)`

Prohibido usar representantes fijos:

- Alta (91)
- Media-Alta (82)
- Media (73)
- Baja (61)

El número debe ser el score real por fila.

Si hay `RT_EMPATE`, `delta crítico`, `cod_local_duplicado`, `local genérico`, `direccion divergente`, no puede haber lectura limpia de Alta sin flag o degradación.

## CONVERGENCIA_DIRECCIONES

Debe comparar por fila:

- DB_HOFFMANN dirección
- DB1 dirección
- DB2 dirección
- RT dirección
- dirección final
- fuente final
- estado convergencia: CONVERGEN / PARCIAL / DIVERGEN / SIN_DATOS
- motivo convergencia
- requiere_google

## COMPARACION_MODELOS

Debe existir con 1.112 filas.

`decision_final` y `motivo_decision_final` deben quedar vacíos.

## RT_CANDIDATOS_TOP

Debe existir con top 3 candidatos RT por fila, o menos si no hay 3.

Debe incluir top1/top2/delta.

## LOG_EJECUCION

Debe registrar:

- timestamp
- archivo input
- archivo output
- scorer usado (rapidfuzz/difflib)
- versión/librería si aplica
- pesos DB_HOFFMANN
- pesos RT
- umbrales
- hard gates aplicados
- warnings
- QA final

## QA obligatorio

- BASE = 1.112 filas.
- Columnas A:I intactas.
- Solo J:M modificadas.
- 0 COD KPI ONE con Gestión RT=NO.
- 0 KPI duplicado activo.
- 0 HI/EX mix.
- 0 prefijo KPI incompatible con SI.
- 0 SUPER10/M10/CM con SI.
- 0 Cód Local copiado a COD KPI ONE.
- Todas las hojas obligatorias existen.
- HOFFMANN_MATCH = 1.112 filas.
- COMPARACION_MODELOS = 1.112 filas.
- RT_CANDIDATOS_TOP existe.
- LOG_EJECUCION suficiente.
- NIVEL_CONFIANZA_PROMEDIO usa score real.
- No aparece texto `validado por Google`.

## Reporte final esperado

- Ruta archivo generado.
- QA PASS/FAIL.
- Gestión RT SI/NO.
- KPI poblados.
- Cód Local Hoffmann confiables/tentativos/duplicados.
- Direcciones por fuente.
- Confianza por banda.
- Revisión manual.
- Recuperados vs V2.
- Errores corregidos vs V3 fallida.
- Confirmación de original intacto.

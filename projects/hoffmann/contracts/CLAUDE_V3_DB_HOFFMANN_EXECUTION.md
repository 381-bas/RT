# Contrato Claude — HOFFMANN V3 DB_HOFFMANN

## Rol

Ejecutar una V3 controlada desde el archivo original `EVALUACIÓN_HOFFMANN.xlsx`, incorporando `DB_HOFFMANN` y `CADENA-HOLDING` como nuevas fuentes internas.

## Input

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN.xlsx`

Hojas esperadas:

- BASE
- DB_HOFFMANN
- CADENA-HOLDING
- DB_DIRECCIÓN_1
- DB_DIRECCIÓN_2
- RT

## Output

Crear archivo nuevo:

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN_DIRECCIONES_RT_MATCH_CLAUDE_V3_DB_HOFFMANN.xlsx`

No modificar archivo original.

## Objetivo V3

Completar BASE como salida limpia, escribiendo solo:

- GESTIÓN RT (SI/NO)
- COD KPI ONE
- DIRECCIÓN
- NIVEL_CONFIANZA_PROMEDIO

Agregar trazabilidad completa en hojas auxiliares.

## Fuente primaria nueva

`DB_HOFFMANN` pasa a ser la primera fuente de dirección y código local HOFFMANN.

Usar:

- `Cód Local`
- `Local`
- `Dirección`
- `Ciudad`
- `Región`
- `Holding`

`CADENA-HOLDING` debe mapear `BASE.Cadena` hacia `Holding` esperado.

## Regla crítica

`Cód Local` de DB_HOFFMANN NO es `COD KPI ONE`.

- `Cód Local` debe guardarse como `cod_local_hoffmann` en auditoría.
- `COD KPI ONE` solo se pobla desde RT si Gestión RT=SI.

## Matching DB_HOFFMANN

Score 0-100:

- Holding compatible: 20
- Región compatible: 15
- Ciudad compatible: 20
- Local exacto/fuzzy: 30
- Dirección presente/utilizable: 10
- Cód Local presente: 5

Caps:

- Holding incompatible: máximo 49
- Región incompatible: máximo 65
- Ciudad incompatible: máximo 78
- Empate top1/top2 < 8: no automático

## Dirección final

Prioridad:

1. DB_HOFFMANN si score >= 85 y sin conflicto.
2. DB_DIRECCIÓN_1 si cruza mejor por código/local/holding/región.
3. DB_DIRECCIÓN_2 si mejora o complementa sin contradicción.
4. RT solo si Gestión RT=SI y no hay fuente interna confiable.
5. Vacío si ninguna fuente aceptable.

No escribir `PENDIENTE` en BASE.

## Gestión RT

Mantener contrato V2:

- candidate pool estricto por cadena/formato/región.
- prefijo KPI compatible observado desde RT.
- top1/top2 delta >= 8 o candidato único real.
- 0 KPI duplicado activo entre natural keys distintas.
- 0 mezcla HI/EX.

DB_HOFFMANN puede ayudar a resolver local genérico, pero no puede saltarse los hard gates RT.

## Confianza

Separar:

- `confianza_direccion`
- `confianza_rt`
- `score_final_ajustado`
- `banda_final`

En BASE escribir `NIVEL_CONFIANZA_PROMEDIO` como `Banda (score_real_redondeado)`.

## Hojas auxiliares obligatorias

Además de las V2:

- HOFFMANN_MATCH
- COD_LOCAL_HOFFMANN_AUDIT
- CONVERGENCIA_DIRECCIONES

Mantener:

- AUDITORIA_MATCH
- REVISION_MANUAL
- RESUMEN_CALIDAD
- DICCIONARIO_EQUIVALENCIAS
- LOG_EJECUCION
- CIUDAD_COMUNA_EQUIVALENCIAS
- COMPARACION_MODELOS
- RT_CANDIDATOS_TOP
- KPI_DUPLICADOS_AUDITORIA

## QA obligatorio

- BASE 1.112 filas.
- Columnas originales A:I intactas.
- Solo J:M modificadas.
- COD KPI ONE solo con Gestión RT=SI.
- 0 KPI duplicado activo.
- 0 HI/EX mix.
- 0 prefijo KPI incompatible con SI.
- `cod_local_hoffmann` trazado en auditoría.
- `Cód Local` no se copia a `COD KPI ONE`.
- HOFFMANN_MATCH tiene 1.112 filas.
- CONVERGENCIA_DIRECCIONES compara DB_HOFFMANN / DB1 / DB2 / RT.
- Google no validado.
- LOG registra pesos, scorer y umbrales.

## Output esperado en respuesta

- Ruta archivo V3.
- QA PASS/FAIL.
- Gestión RT SI/NO.
- KPI poblados.
- cod_local_hoffmann poblados.
- direcciones pobladas/vacías.
- confianza por banda.
- revisión manual.
- recuperados vs V2.
- nuevos conflictos.
- archivo original intacto.

# HOFFMANN — Plan V3 con DB_HOFFMANN

## Diagnóstico inicial

Nuevo archivo `EVALUACIÓN_HOFFMANN.xlsx` contiene:

| Hoja | Filas | Rol |
|---|---:|---|
| BASE | 1.112 | Salida limpia del cliente |
| DB_HOFFMANN | 1.121 | Nueva fuente primaria para local/código/dirección/holding |
| CADENA-HOLDING | 10 | Homologación cadena comercial → holding |
| DB_DIRECCIÓN_1 | 955 | Fuente histórica direcciones + COD. LOCAL |
| DB_DIRECCIÓN_2 | 485 | Fuente histórica complementaria |
| RT | 859 | Operatividad RT para Gestión RT y COD KPI ONE |

## Cambio respecto a V2

`DB_HOFFMANN` pasa a ser **primera fuente de matching** porque trae:

- `Cód Local`
- `Local`
- `Dirección`
- `Ciudad`
- `Región`
- `Holding`

`CADENA-HOLDING` permite cruzar `BASE.Cadena` contra `DB_HOFFMANN.Holding` y contra las DB de dirección.

## Cobertura preliminar observada

Cruce exacto normalizado preliminar entre `BASE.Local` y `DB_HOFFMANN.Local`:

| Criterio | Único | Múltiple | Sin match |
|---|---:|---:|---:|
| local_norm | 672 | 307 | 133 |
| local_norm + región | 672 | 303 | 137 |
| local_norm + región + holding | 888 | 76 | 148 |
| local_norm + región + holding + ciudad | 824 | 61 | 227 |

Conclusión: `DB_HOFFMANN` debe aumentar la confianza, pero todavía requiere reglas para duplicados y locales genéricos.

## Reglas V3 propuestas

### 1. Fuente primaria: DB_HOFFMANN

Para cada fila BASE:

1. Derivar `holding_expected` usando `CADENA-HOLDING`.
2. Normalizar `BASE.Local` removiendo prefijo operativo.
3. Buscar candidatos en `DB_HOFFMANN` filtrando por:
   - región código 2 dígitos,
   - holding,
   - formato/cadena inferida cuando aplique,
   - ciudad si es compatible.
4. Calcular score Hoffmann.

### 2. Score DB_HOFFMANN

| Señal | Peso |
|---|---:|
| Holding compatible | 20 |
| Región compatible | 15 |
| Ciudad compatible | 20 |
| Local exacto/fuzzy | 30 |
| Dirección presente y utilizable | 10 |
| Cód Local presente | 5 |

Hard caps:

- Holding incompatible: score máximo 49.
- Región incompatible: score máximo 65.
- Ciudad incompatible: score máximo 78.
- Empate top1/top2 bajo 8 puntos: no automático.

### 3. Cód Local

`DB_HOFFMANN.Cód Local` debe poblar una columna auxiliar, no `COD KPI ONE`.

Diferenciar:

- `cod_local_hoffmann`: código local del cliente/fuente HOFFMANN.
- `COD KPI ONE`: código operacional RT, solo desde hoja RT.

### 4. Dirección final

Prioridad:

1. `DB_HOFFMANN.Dirección` si score Hoffmann alto y sin conflicto.
2. `DB_DIRECCIÓN_1` si cruza por `Cód Local` / holding / región / ciudad.
3. `DB_DIRECCIÓN_2` si complementa dirección y no contradice.
4. `RT.DIRECCIÓN` solo si Gestión RT=SI y no hay dirección confiable interna.
5. Google solo en fase posterior.

### 5. Gestión RT

Se mantiene el contrato V2 endurecido:

- candidate pool RT por cadena/formato/región.
- prefijo KPI compatible.
- top1/top2 delta suficiente.
- 0 KPI duplicado activo entre natural keys distintas.

DB_HOFFMANN puede ayudar a resolver `local_generico` y recuperar falsos negativos, pero no reemplaza el hard gate RT.

## Nuevas hojas auxiliares V3

Agregar:

- `HOFFMANN_MATCH`: detalle del cruce BASE vs DB_HOFFMANN.
- `COD_LOCAL_HOFFMANN_AUDIT`: duplicidad/conflictos de `Cód Local`.
- `CONVERGENCIA_DIRECCIONES`: comparación DB_HOFFMANN vs DB1 vs DB2 vs RT.

## Objetivo de V3

1. Recuperar SI confiables bloqueados en V2 por local genérico.
2. Mejorar dirección final con fuente directa DB_HOFFMANN.
3. Reducir REVISION_MANUAL antes de Google.
4. Separar confianza de dirección vs confianza RT.

## No negociables

- `COD KPI ONE` sigue viniendo solo de RT.
- `Cód Local` de DB_HOFFMANN no equivale a `COD KPI ONE`.
- No se usa Google como validado en V3 inicial.
- No se sobrescribe el archivo original.

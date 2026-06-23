# Proyecto HOFFMANN

## Objetivo

Construir una base validada de locales HOFFMANN con:

- `GESTIÓN RT (SI/NO)`
- `COD KPI ONE`
- `DIRECCIÓN`
- `NIVEL_CONFIANZA_PROMEDIO`

La hoja `BASE` del Excel del cliente es la salida limpia. Las hojas auxiliares mantienen trazabilidad y auditoría.

## Fuentes disponibles

1. `BASE`: locales objetivo del cliente.
2. `DB_HOFFMANN`: nueva fuente primaria para `Cód Local`, local, dirección, ciudad, región y holding.
3. `CADENA-HOLDING`: mapa de cadena comercial a holding.
4. `DB_DIRECCIÓN_1`: fuente histórica de direcciones con `COD. LOCAL`.
5. `DB_DIRECCIÓN_2`: fuente histórica complementaria de direcciones.
6. `RT`: operación actual Retail Trust para validar Gestión RT y COD KPI ONE.

## Estado de iteraciones

- V1: fallida para RT/KPI por sobre-matching.
- V2: aprobada como versión interna conservadora; no cliente.
- V3 propuesta: incorporar `DB_HOFFMANN` como fuente primaria y `CADENA-HOLDING` como llave de homologación.

## Principio de calidad

La V3 no debe maximizar cantidad de `SI`; debe maximizar certeza operacional y trazabilidad de dirección.

# Contrato Codex — Auditoría HOFFMANN V3

## Rol

Auditar en solo lectura la V3 generada con `DB_HOFFMANN`.

No modificar archivos.  
No generar Excel.  
No usar Google.  
No inventar direcciones.

## Archivos

Original:

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN.xlsx`

V2 aprobada interna:

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN_DIRECCIONES_RT_MATCH_CLAUDE_V2.xlsx`

V3 a auditar:

`C:\Users\basti\Desktop\HOFFMANN\EVALUACIÓN_HOFFMANN_DIRECCIONES_RT_MATCH_CLAUDE_V3_DB_HOFFMANN.xlsx`

## Objetivo

Validar si V3 mejora V2 usando `DB_HOFFMANN` sin romper los hard gates de RT/KPI.

## QA duro

1. BASE = 1.112 filas.
2. Columnas A:I intactas contra original.
3. Solo J:M modificadas.
4. `COD KPI ONE` solo con `GESTIÓN RT=SI`.
5. 0 KPI duplicado activo entre natural keys distintas.
6. 0 HI/EX mix.
7. 0 prefijo KPI incompatible con SI.
8. 0 `Cód Local` copiado erróneamente a `COD KPI ONE`.
9. `HOFFMANN_MATCH` existe y tiene 1.112 filas.
10. `COD_LOCAL_HOFFMANN_AUDIT` existe.
11. `CONVERGENCIA_DIRECCIONES` existe.
12. `COMPARACION_MODELOS.decision_final` sigue vacío.
13. No aparece texto `validado por Google`.
14. URLs Google son solo sugeridas.

## Auditoría específica

### DB_HOFFMANN

Revisar:

- match por Local + Holding + Región + Ciudad.
- duplicidad de `Cód Local`.
- conflictos entre `DB_HOFFMANN.Dirección` y DB1/DB2.
- casos con local genérico recuperados.

### RT/KPI

Comparar V2 vs V3:

- SI V2.
- SI V3.
- SI nuevos.
- SI perdidos.
- KPI duplicados activos.
- prefijos incompatibles.
- mezcla HI/EX.

### Dirección

Clasificar:

- dirección tomada de DB_HOFFMANN.
- dirección tomada de DB1.
- dirección tomada de DB2.
- dirección tomada de RT.
- vacía.
- conflicto entre fuentes.

## Veredicto

Entregar uno de:

- `PASS_INTERNAL_V3`
- `PASS_WITH_MINOR_FIXES`
- `FAIL_REEXECUTE`

## JSON final

```json
{
  "verdict": "PASS_INTERNAL_V3 | PASS_WITH_MINOR_FIXES | FAIL_REEXECUTE",
  "client_ready": false,
  "internal_ready": true,
  "db_hoffmann_improves_v2": true,
  "blocking_findings": [],
  "high_risk_findings": [],
  "medium_risk_findings": [],
  "qa_passed": true,
  "v2_vs_v3_summary": {},
  "required_fixes_before_client_delivery": [],
  "next_step": ""
}
```

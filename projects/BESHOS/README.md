# BESHOS

**Cliente**: BESHOS (bebidas/néctar)  
**Período**: Agosto 2026 (S26-S33)

## Entrada

Lee primero: `_memoria/AGENT.md` — contexto, decisiones codificadas, archivos críticos.

## Estructura

```
_memoria/            ← Punto de entrada (AGENT.md)
datos/               ← Fuentes compartidas (ANÁLISIS_BESHOS.xlsx)
tareas/              ← Iteraciones cronológicas
  ├── resumen_retail_26_08/     ← Análisis completado (scripts + output)
  └── Cierre_Piloto_31_08/      ← Validación con cliente (en progreso)
```

## Tareas activas

| Tarea | Estado | Descripción |
|---|---|---|
| `resumen_retail_26_08` | ✅ Completada | Extracción 5 retailers, pre-post takeover, Excel cliente |
| `Cierre_Piloto_31_08` | 🔄 En progreso | Validación con cliente, feedback, decisiones escalado |

## Entregables

- `tareas/resumen_retail_26_08/_generated/BESHOS - Resumen Retail (Cliente).xlsx`
- `tareas/Cierre_Piloto_31_08/Analisis_cierre_cliente.xlsx` (cuando esté listo)

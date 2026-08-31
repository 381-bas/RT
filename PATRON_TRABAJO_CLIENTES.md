# Patrón de trabajo con clientes — RT

**Última actualización**: 2026-08-31  
**Objetivo**: Flujo limpio, seguro, documentado.

---

## Estructura

Cada cliente vive en `projects/<NOMBRE>/`:

```
projects/CLIENTE/
│
├── README.md                  ← Resumen ejecutivo (qué es, tareas activas)
├── _memoria/AGENT.md          ← ENTRADA DE CLAUDE (contexto, decisiones)
│
├── datos/                     ← Archivos fuente (compartidos, no duplicados)
│   ├── archivo_fuente_1.xlsx
│   ├── archivo_fuente_2.xlsx
│   └── ...
│
└── tareas/TAREA_FECHA/        ← Iteraciones cronológicas
    ├── README.md              ← Descripción tarea, entrada, checklist
    ├── archivo_entrada.xlsx   ← Proporciona el usuario
    ├── archivo_salida.xlsx    ← Entrega al cliente (versión actual)
    ├── archivo_salida_v_ensayo.xlsx  ← Backup defensivo (último buen estado)
    │
    └── _generated/            ← Tu carpeta (Claude)
        ├── script.py          ← Scripts reutilizables
        ├── trazabilidad.json  ← Qué se procesó/cambió
        └── notas_cambios.md   ← Documentación técnica
```

---

## Nombre de tareas

Formato: **`<descripcion>_<DD_MM>`** (día-mes al final)

Ejemplos:
- `resumen_retail_26_08` (26 de agosto)
- `Cierre_Piloto_31_08` (31 de agosto)
- `locales_trebol_31_08` (31 de agosto)
- `automatizacion_semanal_15_09` (15 de septiembre)

**Ventajas:**
- Ordenan cronológicamente automáticamente (`ls` / `find`)
- Autoexplicativos (qué es + cuándo)
- Fácil de detectar sin mirar fechas en archivo

---

## Flujo de trabajo

### 1. Entrada
- Usuario proporciona `archivo_entrada.xlsx` en la tarea
- Lees `README.md` de la tarea (descripción, checklist)

### 2. Análisis y procesamiento
- Analizas entrada
- Generas `archivo_salida.xlsx` con transformaciones
- Documentas en `_generated/` (scripts, JSON, notas)

### 3. Cambios mayores (defensivo)
Antes de cambios grandes:
```
1. Copia: archivo_salida.xlsx → archivo_salida_v_ensayo.xlsx
2. Trabaja en archivo_salida.xlsx
3. Si falla: recupera desde v_ensayo.xlsx
4. Si ok: v_ensayo.xlsx queda como último buen estado
```

### 4. Documentación
En `_generated/`:
- **script.py**: código reutilizable
- **trazabilidad.json**: `{ entrada, transformaciones, salida, fecha }`
- **notas.md**: qué se hizo, por qué, decisiones

### 5. Sin cementerio de versiones
❌ NO: `archivo_v1.xlsx`, `archivo_v2.xlsx`, `archivo_backup_1.xlsx`, ...  
✅ SÍ: un `v_ensayo.xlsx` defensivo, resto en documentación

---

## JSONs y MDs

Patrón de `trazabilidad.json`:

```json
{
  "tarea": "resumen_retail_26_08",
  "fecha_ejecucion": "2026-08-26",
  "entrada": "ANÁLISIS_BESHOS.xlsx",
  "transformaciones": [
    {
      "paso": 1,
      "descripcion": "Filtrar Walmart a 4 SKUs activos",
      "resultado_filas": 1024
    }
  ],
  "salida": "BESHOS - Resumen Retail (Cliente).xlsx",
  "notas": "Pre-post takeover S26 vs S1-S25"
}
```

Patrón de `notas.md`:

```markdown
# Notas de cambios — TAREA_FECHA

**Ejecutado**: YYYY-MM-DD  
**Usuario**: (si aplica)

## Cambios realizados

1. Filtraje de datos
2. Cálculos pre/post
3. Formato de salida

## Decisiones

- Mantener solo Walmart 4 SKUs
- Excluir período anterior a S26

## Siguiente iteración

- Validar con cliente
- Posibles ajustes formato
```

---

## Referencia rápida

| Carpeta/Archivo | Rol | Modificable |
|---|---|---|
| `_memoria/AGENT.md` | Contexto cliente | Sí (si cambian reglas) |
| `datos/*.xlsx` | Fuentes compartidas | No (lectura) |
| `tareas/TAREA/README.md` | Descripción tarea | Sí |
| `tareas/TAREA/*.xlsx` | Entrada/salida | Sí (trabajo normal) |
| `tareas/TAREA/v_ensayo.xlsx` | Backup defensivo | Solo si falla |
| `tareas/TAREA/_generated/*` | Scripts, JSON, notas | Sí (documentación) |

---

## Principios

✅ **Limpio**: una entrada, una salida (+ v_ensayo defensivo)  
✅ **Seguro**: backup defensivo ante fallos  
✅ **Documentado**: JSON + MDs en `_generated/`  
✅ **Reutilizable**: scripts en `_generated/` para futuras tareas  
✅ **Histórico**: todas las tareas viven en `tareas/YYYY-MM_*/` cronológicamente

---

## Ver también

- `projects/BESHOS/_memoria/AGENT.md` — Ejemplo BESHOS
- `projects/BESHOS/tareas/resumen_retail_26_08/README.md` — Ejemplo tarea completada
- `projects/BESHOS/tareas/Cierre_Piloto_31_08/README.md` — Ejemplo tarea en progreso
- `memoria_claude/AGENT.md` — Proyecto principal (motor de minutas)

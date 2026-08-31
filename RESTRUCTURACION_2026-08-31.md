# Restructuración de proyecto RT — 2026-08-31

## ✅ Completado

### 1. Limpieza de `memoria_claude/`
- ✅ Backups centralizados en `_backups/` (14 archivos: BACKUP_02_MINUTAS_CLIENTES_*.xlsm/xlsx, BACKUP_02_SIN_REPORTE_*.xlsx)
- ✅ memoria_claude/ ahora limpia: solo archivos del proyecto principal (AGENT.md, 00-05_*.json, scripts/, memoria_proyecto.json)
- ✅ AGENT.md actualizado con referencia a estructura de clientes en `projects/`

### 2. Reorganización de BESHOS
Antes:
```
projects/BESHOS_resumen_retail/
├── 1_fuente/ (ANÁLISIS_BESHOS.xlsx)
├── 2_trabajo/ (scripts + JSON)
└── 3_entrega/ (Excel cliente)
```

Ahora:
```
projects/BESHOS_resumen_retail/
├── _memoria/AGENT.md                          ← contexto del cliente
├── datos/ANÁLISIS_BESHOS.xlsx                 ← fuente compartida
└── tareas/
    ├── 2026-08_resumen_retail/                ← iteración S26-S33
    │   ├── extract_full.py
    │   ├── build_client_report.py
    │   ├── analisis_beshos_retail.json
    │   ├── resumen_retail_chart.html
    │   └── BESHOS - Resumen Retail (Cliente).xlsx
    └── Cierre_Piloto/                         ← próxima fase (planificación)
        └── README.md
```

- ✅ Commit: `refactor: Reorganizar estructura de carpetas según nuevo modelo`
- ✅ AGENT.md de cliente creado con decisiones codificadas y referencias a archivos críticos
- ✅ Cierre_Piloto iniciado con checklist y decisiones pendientes

### 3. Estructura de carpetas RT
```
RT/
├── _backups/                    ← centralizados (14 BACKUP_*.xlsm/xlsx)
├── _memoria/                    ← proyecto principal (Motor + Plantilla)
├── projects/
│   ├── BESHOS_resumen_retail/   ← restructurado ✅
│   ├── hoffmann/                ← pendiente
│   └── TREBOL/                  ← pendiente
├── CLIENTE_GG/                  ← solo queda TREBOL sin organizar
├── PILAR/                       ← cliente histórico (evaluar si va a projects/)
├── DB_/                         ← datos de trabajo, no cliente
└── ... (otros: CC/, HISTORIA/, HojasMaps/, RR/, RRHH/, scripts/)
```

---

## ⏳ Pendiente

### Corto plazo (requiere decisión):
1. **TREBOL**: ¿crear `projects/TREBOL/` y mover de `CLIENTE_GG/TREBOL/`?
2. **hoffmann**: revisar si está en git o solo en una carpeta local
3. **PILAR**: ¿mantener en `RT/PILAR/` o mover a `projects/PILAR/` si es cliente activo?

### Mediano plazo (después de Cierre_Piloto):
- Aplicar el modelo de estructura (datos + tareas + _memoria) a otros clientes
- Crear AGENT.md por cliente (hoffmann, TREBOL, PILAR, etc.)
- Evaluar si CLIENTE_GG se elimina o se usa para trabajo puntual

### Nuevo flujo para Claude:
1. Cuando trabajes con cliente X, lee **primero**: `projects/X/_memoria/AGENT.md`
2. Datos fuente: `projects/X/datos/`
3. Tarea actual: `projects/X/tareas/YYYY-MM_nombre/`
4. Histórico: todas las carpetas `tareas/YYYY-MM_*` quedan para referencia

---

## Notas técnicas

- Nombres de carpetas tareas: **ISO fecha (YYYY-MM)** + descripción breve, sin espacios
  - Ejemplo: `2026-08_resumen_retail`, `2026-09_dashboard_interactivo`
  - Ventaja: ordenan cronológicamente (ls, find, git log)
  
- Archivos Excel grandes (fuente): van en `datos/`, **nunca** se duplican en tareas
  - Si tarea necesita versión modificada, crear BACKUP_YYYY-MM-DD_*.xlsx en la tarea misma
  
- JSON intermedios (trazabilidad): generados automáticamente, viven en `tareas/YYYY-MM_*/`
  - Nunca editar a mano (salvo para debugging puntual)

---

## Próximo paso

**Bastian**: ¿Reorganizamos TREBOL ahora o después del Cierre_Piloto de BESHOS?

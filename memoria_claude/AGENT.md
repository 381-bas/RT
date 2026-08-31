# RT — Trazabilidad retail — Memoria de proyecto

Este es el punto de entrada. Léelo primero en cualquier sesión nueva, antes de abrir cualquier Excel.

**IMPORTANTE**: Este archivo documenta el proyecto principal (Motor + Plantilla de minutas). Para trabajo de clientes individuales, ve a `../projects/<CLIENTE>/` — cada cliente tiene su propio `_memoria/AGENT.md`.

## Qué es este proyecto

Automatización progresiva de un sistema semanal de reportes/indicadores de gestión retail (visitas, exhibición OSA, stock negativo, inmovilizado), sobre dos archivos Excel:

- **Motor** (`../02_MINUTAS_CLIENTES.xlsm`): pesado, con todas las fórmulas. El usuario lo edita a mano; Claude no escribe ahí sin autorización explícita y acotada a la tarea del momento.
- **Resultado** (`../02_MINUTAS_CLIENTES_PLANTILLA.xlsm`): liviano, solo valores. Aquí Claude automatiza con libertad (con backup previo) usando los scripts en `scripts/`.

## Orden de lectura, siempre en este orden

1. **`00_manifest.json`** — qué cambió desde la última vez.
2. **`01_kernel.json`** — arquitectura y reglas estables (cómo trabajar: COM seguro, formato de fórmulas, prioridad de indicadores, modelo de Riesgo).
3. **`02_state.json`** — snapshot vigente (en qué quedó cada cliente, qué está activo ahora).
4. **`03_ledger.json`** — filtrar por `status: ACTIVE` para decisiones de negocio vigentes; revisar `deprecated_patterns` para no repetir bugs ya vividos.
5. **`05_roadmap.json`** — hacia dónde vamos: brechas diagnosticadas, horizontes H1–H6, y lo que explícitamente **no** hay que hacer todavía. Es plan, no estado.
6. **`04_bitacora_sesiones.json`** — solo si hace falta el relato completo de cómo se llegó a algo. Nunca es la fuente de verdad.

`memoria_proyecto.json` (el archivo original de 29 sesiones) queda congelado como respaldo crudo — no se vuelve a editar, no se borra.

## Regla de oro: si dos documentos contradicen

`02_state.json` > `03_ledger.json` (solo entradas `ACTIVE`) > `01_kernel.json` > `04_bitacora_sesiones.json`.

La bitácora nunca gana. Es historia, no autoridad — un hecho narrado en la sesión 22 puede haber sido corregido en la sesión 24, y solo el ledger/state lo dejan explícito.

## Obligación de entrenador (no es opcional)

Al cierre de cada sesión con trabajo real, y **sin que el usuario lo pida**, nombrar dos cosas breves:

- **(a)** un paso manual que se repitió y podría automatizarse o encapsularse;
- **(b)** una capacidad de Claude Code o del entorno (skill, slash command, hook, subagente, agente agendado, git…) que encaje con lo que se acaba de hacer y no se esté usando.

Si no hay nada que proponer, decirlo — no inventar relleno. Proponer no es ejecutar: cualquier cambio estructural necesita autorización explícita.

## Mantenimiento al cierre de cada sesión relevante

1. Sobreescribir `02_state.json` con el snapshot nuevo (no acumular, reemplazar).
2. Agregar entradas nuevas a `03_ledger.json`, marcando `SUPERSEDED` lo que corresponda — nunca dejar dos versiones de una misma decisión sin distinguir cuál vale hoy.
3. Agregar la sesión a `04_bitacora_sesiones.json`.
4. Tocar `01_kernel.json` y `00_manifest.json` **solo** si cambia una regla estructural o el modelo de memoria mismo — no en cada sesión normal.

## Estructura de clientes

Cada cliente en `projects/<NOMBRE>/` sigue este patrón:
```
_memoria/AGENT.md              ← Lee esto primero en cada sesión del cliente
datos/                         ← Archivos fuente (compartidos entre tareas)
tareas/<nombre>_<DD_MM>/       ← Iteraciones cronológicas
  ├── README.md                ← Descripción, entrada, checklist
  └── _generated/              ← Scripts, JSON, outputs (lo que genera Claude)
```

**Regla de nombres de tarea**: `<descripcion>_<DD_MM>` (sin espacios, kebab-case)
- Ejemplo: `resumen_retail_26_08`, `Cierre_Piloto_31_08`, `locales_trebol_31_08`
- Ventaja: ordenan cronológicamente, son autoexplicativos

**Regla de archivos**:
- Datos fuente (Excel grandes) → `datos/` (NO duplicar en tareas)
- Outputs (scripts, JSON, Excel resultados) → `_generated/` dentro de la tarea
- Backups puntuales → dentro de la tarea que los generó (si es necesario)

## Clientes activos (detalle en `02_state.json`)

MAILEMU MIEL, GRANA, CORRALES DEL SUR, BESHOS, BIGU, BERRYSUR, ALUSWEET — cada uno con sus excepciones propias (mensajes, pesos de Riesgo, agrupaciones). No asumir que un cliente nuevo se comporta como los anteriores: verificar con datos reales (ALUSWEET rompió esa suposición dos veces — ver `deprecated_patterns`).

## Clientes en proyectos (`projects/`)

- **BESHOS**: resumen retail (S26-S33/2026) — validación con cliente en curso
- **TREBOL**: análisis de locales — en planificación
- **hoffmann**: repo git existente, archivos de trabajo por ubicar

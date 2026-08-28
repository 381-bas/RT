# RT

Repositorio operativo para proyectos de Retail Trust.

## Estructura

```text
projects/
  hoffmann/
    README.md
    docs/
    contracts/
    outputs/        # referencias / manifiestos; no binarios pesados por defecto
```

## Convenciones

- Cada proyecto debe mantener contrato, auditoría y salidas separadas.
- Los archivos originales de cliente no se sobrescriben.
- Las salidas Excel se versionan por iteración: V1, V2, V3, etc.
- Toda validación externa, como Google Maps, debe quedar trazada con fuente y nivel de confianza.

## Proyecto activo

- `projects/hoffmann`: matching de locales, direcciones, Gestión RT y COD KPI ONE.

## Encargos de análisis (informes por cliente, agregado 2026-08-27)

Un segundo patrón, para encargos de análisis comercial/de datos (no desarrollo de sistema):
distinto ciclo de vida, misma disciplina de separar fuente, trabajo y entrega.

```text
projects/
  <CLIENTE>_<ENCARGO>/
    README.md      # que se pidio, cuando, estado, cual es el entregable vigente
    1_fuente/       # datos crudos recibidos -- nunca se editan, fuera de git
    2_trabajo/      # scripts y exploracion -- los .py van a git, los intermedios no
    3_entrega/      # el archivo limpio a presentar -- fuera de git
```

- `1_fuente/` y `3_entrega/` quedan fuera de git (ver `.gitignore`) por ser datos, no razonamiento.
- Al cerrar un encargo: `3_entrega/` queda con el archivo final, `2_trabajo/` conserva solo los
  scripts reproducibles (se borran los dumps/JSON intermedios que ya no aportan), y el README
  anota qué se borró y por qué.
- Ver `memoria_claude/05_roadmap.json` (horizonte H1) para el diagnóstico completo que originó
  este patrón.

## Seguridad: hook pre-commit (agregado 2026-08-27)

El repo trae un hook de git versionado que bloquea automáticamente cualquier intento de
comitear datos (`.xlsx`, `.xlsm`, archivos grandes sin extensión) o texto que parezca una
credencial (connection strings, claves privadas, API keys). Se activa una vez por máquina:

```bash
git config core.hooksPath scripts/git-hooks
```

Sin ese `git config`, el hook existe en el repo pero **no corre** — es la única parte de
esta protección que no viaja sola con `git clone`.

## Encargos activos con este patrón

- `projects/BESHOS_resumen_retail`: informe de rendimiento retail para BESHOS.

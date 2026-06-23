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

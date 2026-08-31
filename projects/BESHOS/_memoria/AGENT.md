# BESHOS — Análisis retail & gestión retail

Cliente: BESHOS (bebidas/néctar, retail pasivo)  
Período base: S26/2026 (22-06-2026) — fecha de takeover de la cuenta  
Última actualización: 2026-08-31

## Estado actual

**Fase 1 (completada): Resumen Retail**
- Extracción de datos de 5 retailers: Walmart, Jumbo, Santa Isabel, SMU, Tottus
- Filtrado de Walmart a 4 SKUs activos BESHOS: 700419, 700420, 700421, 5064681
- JSON de trazabilidad: `tareas/2026-08_resumen_retail/analisis_beshos_retail.json`
- Reporte cliente limpio (solo números, gráficos, sin warnings): `tareas/2026-08_resumen_retail/BESHOS - Resumen Retail (Cliente).xlsx`
- Pre-post takeover (S1-S25 vs S26-S33):
  - **Walmart**: +38,4% (740 → 1024 unidades/sem promedio)
  - **Jumbo**: +2,3% (1130 → 1156)
  - **Santa Isabel**: +13,1% (272 → 307)
  - **SMU**: -17,2% (53 → 44)
  - **Tottus**: -3,8% (142 → 137)

**Decisiones codificadas:**
- Walmart = 4 SKUs (no 3, no total reportado)
- Gráficos de evolución: S26-S33 únicamente, layout vertical limpio (una tarjeta por retailer, no grilla 2D)
- Removidas todas las notas sobre "alcance restringido" de Walmart — el cliente solo ve números reales

## Próximas fases

Ver `tareas/Cierre_Piloto/README.md` — hoja de ruta de cierre y mejoras secundarias.

## Archivos críticos

| Ruta | Rol | Modificable por Claude |
|---|---|---|
| `datos/ANÁLISIS_BESHOS.xlsx` | Fuente cruda (7.2MB, todas las hojas) | No — solo lectura |
| `tareas/2026-08_resumen_retail/extract_full.py` | Extractor de datos + cálculo pre/post | Sí, con validación |
| `tareas/2026-08_resumen_retail/build_client_report.py` | Generador del Excel cliente | Sí, con autorización por cambios UX |
| `tareas/2026-08_resumen_retail/analisis_beshos_retail.json` | Datos intermedios (trazabilidad) | Autogenerado, no editar |
| `tareas/2026-08_resumen_retail/BESHOS - Resumen Retail (Cliente).xlsx` | Entrega al cliente | Autogenerado |

## Reglas de trabajo

1. **Datos**: siempre filtrar Walmart a 4 SKUs en ingesta (nunca mostrar total reportado)
2. **Charts**: `visible_cells_only = False` (fuentes pueden estar ocultas/collapsed)
3. **Backups**: si se modifica un Excel, crear BACKUP_*.xlsx antes
4. **JSON**: mantener estructura de `weekly_series`, `pre_post_takeover`, `top_skus`, `top_stores` para trazabilidad

Léeme siempre al empezar una sesión nueva en este cliente.

# BESHOS — Resumen Retail 26-08

**Período**: S26-S33 / 2026  
**Fecha**: Completado 26-08-2026  
**Estado**: Listo para presentación al cliente

## Descripción

Análisis de desempeño retail por cadena (Walmart, Jumbo, Santa Isabel, SMU, Tottus) desde takeover de cuenta (S26/22-06-2026) vs. período anterior.

## Entrada (datos compartidos)

- Fuente: `../../datos/ANÁLISIS_BESHOS.xlsx` (7.2MB)
  - Hojas: RES_WALMART, RES_JUMBO, RES_SMU, RES_TOTTUS, etc.

## Salida (`_generated/`)

| Archivo | Descripción |
|---|---|
| `extract_full.py` | Script extractor (filtra Walmart a 4 SKUs, calcula pre/post) |
| `build_client_report.py` | Constructor del Excel cliente |
| `analisis_beshos_retail.json` | Datos intermedios (trazabilidad) |
| `BESHOS - Resumen Retail (Cliente).xlsx` | **Entrega cliente** (19KB) |
| `resumen_retail_chart.html` | Preview de gráficos (verificación visual) |

## Clave de decisiones

- Walmart: 4 SKUs activos (700419, 700420, 700421, 5064681) — filtrado en ingesta
- Período: S26-S33 únicamente en gráficos
- Métricas: unidades/semana promedio, % cambio pre-post takeover
- Formato: limpio (sin warnings), orientado a cliente

## Ejecutar

```bash
python extract_full.py         # → analisis_beshos_retail.json
python build_client_report.py  # → BESHOS - Resumen Retail (Cliente).xlsx
```

Ambos scripts leen de `../../datos/ANÁLISIS_BESHOS.xlsx`.

## Próximo paso

Validación con cliente → ir a `../Cierre_Piloto_31_08/`

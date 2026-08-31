# BESHOS — Cierre Piloto 31-08

**Inicio**: 2026-08-31  
**Estado**: Planificación  
**Propósito**: Validar reporte con cliente, documentar feedback, decidir escalado

## Entrada

- Reporte cliente: `../resumen_retail_26_08/_generated/BESHOS - Resumen Retail (Cliente).xlsx`
- Datos fuente: `../../datos/ANÁLISIS_BESHOS.xlsx`

## Checklist

- [ ] Presentación al cliente (Bastian)
- [ ] Capturar feedback (formato, números, gráficos)
- [ ] Corregir si hay desvíos vs. Excel fuente
- [ ] Documentar preguntas → definen nuevas tareas

## Decisiones clave

**¿Es recurrente o puntual?**
- Recurrente → automatizar (nueva tarea `2026-09_automatizacion_semanal_*`)
- Puntual → cerrar y documentar como case study

**¿Se replica a otros clientes?**
- HOFFMANN, TREBOL, GRANA usan patrón similar
- Scripts (`extract_full.py` + `build_client_report.py`) son base reutilizable

## Mejoras secundarias (si cliente lo pide)

- Gráficos: ajustar último label (si se corta)
- Hoja DETALLE SEMANAL (desglose por tienda)
- Dashboard interactivo (Tableau/Power BI)

## Salida (`_generated/`)

Aquí se documentan cambios solicitados, versiones nuevas del reporte, etc.

# Cierre Piloto — BESHOS

**Inicio**: 2026-08-31  
**Estado**: Planificación  
**Propósito**: Validar con cliente, documentar lecciones, decidir escalado

---

## Checklist de cierre

- [ ] Presentación del reporte "BESHOS - Resumen Retail" al cliente (Bastian)
- [ ] Feedback del cliente sobre formato, números, gráficos
- [ ] Corregir cualquier número fuera de rango (validar Excel fuente)
- [ ] Documentar qué preguntas hizo el cliente → pueden derivar en nuevas tareas

## Posibles mejoras secundarias (no bloqueantes)

- [ ] Gráficos: ajustar ancho si último label se corta ("1090" → "1C")
- [ ] Añadir hoja de DETALLE SEMANAL (semana a semana, breakdown de tiendas)
- [ ] Dashboard interactivo en Tableau/Power BI (si cliente lo pide)
- [ ] Automatizar lectura de ANÁLISIS_BESHOS.xlsx cada semana (si cliente sigue el patrón)

## Decisiones pendientes

**¿Es recurrente?**
- Si el cliente dice "hazme esto cada semana", pasar a fase 2: automatización semanal con alertas
- Si dice "fue puntual", cerrar y documentar como case study

**¿Escalamos a otros clientes?**
- HOFFMANN, TREBOL, GRANA — ¿tienen necesidad similar de análisis retail?
- El modelo de `extract_full.py` + `build_client_report.py` es replicable

## Notas

Archivo fuente es pesado (~7.2MB) porque contiene todas las dinámicas del negocio (hojas RES_WALMART, RES_JUMBO, etc.). En próximas iteraciones considerar:
- Exportar solo las hojas que usamos (más rápido)
- Cache del JSON si no cambió la fuente (evitar re-cálculo)

Próximo commit: después de validación del cliente.

# -*- coding: utf-8 -*-
"""
build_observaciones.py (sesion 27)
Crea/actualiza la hoja OBSERVACIONES en 02_MINUTAS_CLIENTES_PLANTILLA.xlsm: tabla
generica y reutilizable (Cliente | Fecha | Categoria | Observacion | Prioridad) para
que Claude deje hallazgos de analisis de ventas/negocio, pensada para varios clientes
en el tiempo -- NO se wipea completa cada corrida, solo se reemplazan las filas del
CLIENTE que se esta procesando en esa corrida (las de otros clientes se conservan).

Fuente de datos: hoja VENTAS (CLIENTE, MES, COD KPI One, Cod. Cadena, Descripcion
Producto, 2025_un, 2025_pesos, 2026_un, 2026_pesos).

IMPORTANTE: normaliza el comparativo a VENTA PROMEDIO SEMANAL (2025 entre SEMANAS_2025,
2026 entre SEMANAS_2026_CERRADAS) -- comparar totales crudos es enganoso cuando los
periodos tienen distinto largo (2025=mes completo/4 semanas, 2026=semanas cerradas del
mes en curso). ACTUALIZAR SEMANAS_2026_CERRADAS cada semana que cierre una nueva.

NUEVO (sesion 31, "otra vuelta mas larga con mas analisis"): SEMANAS_2026_CERRADAS 2->3
(el usuario confirmo 3 semanas cerradas a la fecha de esta corrida). Se agrega una categoria
nueva de hallazgo, CRUCE_VENTA_OPERACION: cruza las peores salas de CADENA_FOCO por caida
YoY (ya calculadas en salas_var) contra PLANTILLA (leida de la misma corrida) para ver
cuales de esas salas TAMBIEN tienen un indicador operativo activo esta semana (OSA,
INMOVILIZADO, SEGUIMIENTO, STOCK NEGATIVO) -- responde la pregunta que el usuario dejo
pendiente en sesion 29 ("esto coincide con la PLANTILLA? los niveles de OSA, de
INCUMPLIMIENTO?") con datos reales en vez de dejarla como consulta aislada. Requiere que
PLANTILLA ya este actualizada (correr build_plantilla.py ANTES que este script).

NUEVO (sesion 34, refactor de mantenimiento): norm_code/fmt_num, inferir_cadena/
normalizar (Tabla6), y la apertura segura de Excel via COM se sacaron a rt_common.py
-- estaban copiados LITERALMENTE de build_plantilla.py, con riesgo real de que un fix
en un archivo no llegara al otro (ya paso en sesion 28). Sin cambio de comportamiento.
"""
import datetime
from collections import defaultdict, Counter

import rt_common as rc

FILE_PATH = r"C:\Users\basti\Desktop\RT\02_MINUTAS_CLIENTES_PLANTILLA.xlsm"

CLIENTE = "ALUSWEET"
SEMANAS_2025 = 4
SEMANAS_2026_CERRADAS = 3  # <-- ACTUALIZAR cada semana que cierre una nueva

CADENA_FOCO = "WALMART"  # cadena de mayor prioridad operativa para este cliente
TOP_N_PEORES_SALAS = 10

OBS_HEADERS = ["Cliente", "Fecha", "Categoria", "Observacion", "Prioridad"]

def log(msg):
    print(msg, flush=True)

# Compartidas con build_plantilla.py -- ver memoria_claude/scripts/rt_common.py
norm_code = rc.norm_code
fmt_num = rc.fmt_num

def main():
    with rc.abrir_excel_com(FILE_PATH) as (app, wb):
        log("Abriendo workbook liviano (instancia independiente)...")

        ws_v = wb.Worksheets("VENTAS")
        ws_m = wb.Worksheets("MINUTA")

        # --- Tabla6 (prefijos), igual que en build_plantilla -- ver rt_common.py ---
        lo6 = ws_m.ListObjects("Tabla6")
        tabla6 = rc.cargar_tabla6(lo6)

        def inferir_cadena(codigo_crudo):
            return rc.inferir_cadena(codigo_crudo, tabla6)

        def normalizar(codigo_crudo, cadena):
            return rc.normalizar(codigo_crudo, cadena, tabla6)

        # --- Leer VENTAS ---
        hdr_v = ws_v.Range(ws_v.Cells(1, 1), ws_v.Cells(1, ws_v.UsedRange.Columns.Count)).Value[0]
        idxv = {h: i for i, h in enumerate(hdr_v)}
        last_row_v = ws_v.UsedRange.Rows.Count
        data_v = ws_v.Range(ws_v.Cells(2, 1), ws_v.Cells(last_row_v, len(hdr_v))).Value

        filas = []
        for row in data_v:
            if not row[idxv["CLIENTE"]] or row[idxv["CLIENTE"]] != CLIENTE:
                continue
            cod = row[idxv["COD KPI One"]]
            sku = row[idxv["C\u00f3d. Cadena"]]
            prod = row[idxv["Descripci\u00f3n Producto"]]
            u25 = row[idxv["2025_un"]] or 0
            p25 = row[idxv["2025_pesos"]] or 0
            u26 = row[idxv["2026_un"]]
            p26 = row[idxv["2026_pesos"]]
            cadena = inferir_cadena(cod)
            cod_norm = normalizar(cod, cadena) if cadena else norm_code(cod)
            filas.append({"cod": cod_norm, "cadena": cadena or "SIN_CADENA", "sku": sku, "prod": prod,
                          "u25": u25, "p25": p25, "u26": (u26 or 0), "p26": (p26 or 0),
                          "inhabilitado": (u26 is None and u25 > 0)})
        log(f"Filas VENTAS {CLIENTE}: {len(filas)}")

        # --- 1) Rollup por cadena ---
        por_cadena = defaultdict(lambda: {"u25": 0, "p25": 0, "u26": 0, "p26": 0, "n": 0})
        for f in filas:
            d = por_cadena[f["cadena"]]
            d["u25"] += f["u25"]; d["p25"] += f["p25"]; d["u26"] += f["u26"]; d["p26"] += f["p26"]; d["n"] += 1

        # --- 2) Inhabilitados por cadena ---
        inhab = [f for f in filas if f["inhabilitado"]]
        inhab_por_cadena = Counter(f["cadena"] for f in inhab)
        inhab_top5 = sorted(inhab, key=lambda x: -x["p25"])[:5]

        # --- 3) Peores salas del CADENA_FOCO ---
        por_sala = defaultdict(lambda: {"u25": 0, "p25": 0, "u26": 0, "p26": 0, "cadena": None, "n_sku": 0})
        for f in filas:
            k = (f["cadena"], f["cod"])
            d = por_sala[k]
            d["u25"] += f["u25"]; d["p25"] += f["p25"]; d["u26"] += f["u26"]; d["p26"] += f["p26"]
            d["cadena"] = f["cadena"]; d["n_sku"] += 1

        salas_foco = [(k[1], d) for k, d in por_sala.items() if d["cadena"] == CADENA_FOCO and d["p25"] > 0]
        salas_var = []
        for cod, d in salas_foco:
            prom25 = d["p25"] / SEMANAS_2025
            prom26 = d["p26"] / SEMANAS_2026_CERRADAS
            var = (prom26 - prom25) / prom25 * 100
            salas_var.append((cod, prom25, prom26, var, d["n_sku"]))
        salas_var.sort(key=lambda x: x[3])
        n_mejoran = sum(1 for x in salas_var if x[3] > 0)
        n_empeoran = sum(1 for x in salas_var if x[3] <= 0)

        # --- 4) Cruce con PLANTILLA (sesion 31): de las peores salas de CADENA_FOCO por
        # caida YoY, cuales tienen ADEMAS un indicador operativo activo esta semana. ---
        cruce = []
        try:
            ws_p = wb.Worksheets("PLANTILLA")
            hdr_p = ws_p.Range(ws_p.Cells(1, 1), ws_p.Cells(1, ws_p.UsedRange.Columns.Count)).Value[0]
            idxp = {h: i for i, h in enumerate(hdr_p)}
            last_row_p = ws_p.UsedRange.Rows.Count
            data_p = ws_p.Range(ws_p.Cells(2, 1), ws_p.Cells(last_row_p, len(hdr_p))).Value
            if not isinstance(data_p, tuple):
                data_p = (data_p,)
            plantilla_por_sala = {}
            for r in data_p:
                if not r[idxp["Marca"]] or str(r[idxp["Marca"]]).strip() != CLIENTE:
                    continue
                if r[idxp["Cadena"]] != CADENA_FOCO:
                    continue
                cod_p = norm_code(r[idxp["Cod Local"]])
                plantilla_por_sala[cod_p] = r[idxp["Observacion"]]
            peores_codigos = {cod for cod, _p25, _p26, _var, _n in salas_var[:TOP_N_PEORES_SALAS]}
            for cod in peores_codigos:
                obs_plantilla = plantilla_por_sala.get(cod)
                if obs_plantilla:
                    cruce.append((cod, obs_plantilla))
            log(f"Cruce venta/operacion: {len(cruce)} de las {len(peores_codigos)} peores salas de "
                f"{CADENA_FOCO} tienen tambien indicador operativo activo en PLANTILLA esta semana.")
        except Exception as e:
            log(f"Aviso: no se pudo cruzar con PLANTILLA ({e}) -- se omite CRUCE_VENTA_OPERACION.")

        # --- Construir filas de observaciones ---
        hoy_date = datetime.date.today()
        hoy = datetime.datetime(hoy_date.year, hoy_date.month, hoy_date.day)  # COM necesita datetime, no date
        filas_obs = []

        def add(categoria, obs, prioridad):
            filas_obs.append([CLIENTE, hoy, categoria, obs, prioridad])

        add("METODOLOGIA",
            f"Comparativo YoY normalizado a venta PROMEDIO SEMANAL (2025={SEMANAS_2025} semanas cerradas, "
            f"2026={SEMANAS_2026_CERRADAS} semanas cerradas -- ACTUALIZAR este divisor cada semana que cierre "
            f"una nueva). El comparativo de TOTALES crudos sin normalizar sugeria caidas de -50% o mas que "
            f"eran en realidad artefacto de comparar periodos de distinto largo, no una caida real.",
            "Alta")

        for c, d in sorted(por_cadena.items(), key=lambda x: -x[1]["p25"]):
            prom25 = d["p25"] / SEMANAS_2025
            prom26 = d["p26"] / SEMANAS_2026_CERRADAS
            var = (prom26 - prom25) / prom25 * 100 if prom25 else None
            if var is None:
                continue
            foco_txt = " (CADENA FOCO -- mayor prioridad operativa del equipo)" if c == CADENA_FOCO else ""
            prioridad = "Alta" if var < -10 else ("Media" if var < 5 else "Baja")
            add("RENDIMIENTO_CADENA",
                f"{c}{foco_txt}: venta promedio semanal {'estable' if abs(var)<3 else ('cae' if var<0 else 'mejora')} "
                f"{var:+.1f}% vs 2025 (${prom25:,.0f} -> ${prom26:,.0f} por semana, {d['n']} combinaciones sala+SKU).",
                prioridad)

        add("RENDIMIENTO_CADENA",
            f"{CADENA_FOCO}: el promedio esconde dispersion real -- de {len(salas_var)} salas con base 2025, "
            f"{n_mejoran} mejoran y {n_empeoran} empeoran YoY. No asumir que 'promedio estable' = 'sin problemas'.",
            "Media")

        for cod, prom25, prom26, var, n_sku in salas_var[:TOP_N_PEORES_SALAS]:
            prioridad = "Alta" if var < -30 else "Media"
            add(f"VENTA_SALA_{CADENA_FOCO}",
                f"Sala {cod} ({CADENA_FOCO}): caida de {var:.1f}% en venta promedio semanal "
                f"(${prom25:,.0f} -> ${prom26:,.0f}/semana, {n_sku} SKU). Priorizar revision/visita.",
                prioridad)

        if cruce:
            add("CRUCE_VENTA_OPERACION",
                f"{len(cruce)} de las {TOP_N_PEORES_SALAS} peores salas de {CADENA_FOCO} por caida YoY "
                f"tienen ADEMAS un indicador operativo activo esta semana en PLANTILLA (OSA/INMOVILIZADO/"
                f"SEGUIMIENTO/STOCK NEGATIVO) -- la caida de venta no es solo un numero, tiene una causa "
                f"gestionable identificada en terreno.",
                "Alta")
            for cod, obs_plantilla in cruce:
                add("CRUCE_VENTA_OPERACION",
                    f"Sala {cod} ({CADENA_FOCO}): entre las peores por caida YoY, y esta semana PLANTILLA "
                    f"reporta -> {obs_plantilla}",
                    "Alta")
        else:
            add("CRUCE_VENTA_OPERACION",
                f"Ninguna de las {TOP_N_PEORES_SALAS} peores salas de {CADENA_FOCO} por caida YoY tiene "
                f"un indicador operativo activo esta semana en PLANTILLA -- la caida de venta en esas "
                f"salas no se explica (por ahora) por incumplimiento, OSA o inmovilizado detectado.",
                "Media")

        add("PRODUCTO_INHABILITADO",
            f"{len(inhab)} combinaciones SKU+local con venta 2025>0 y venta 2026 nula "
            f"({dict(inhab_por_cadena)}) -- probable baja de catalogo. Validar contra listado de "
            f"productos habilitados vigente antes de reportarlo como quiebre de gestion.",
            "Media")
        for f in inhab_top5:
            add("PRODUCTO_INHABILITADO",
                f"{f['cadena']} {f['cod']} SKU {fmt_num(f['sku'])} \"{f['prod']}\": vendia {fmt_num(f['u25'])}u "
                f"(${f['p25']:,.0f}) en 2025, sin venta registrada en 2026.",
                "Media")

        add("SUGERENCIA_INDICADOR",
            "Indice de dispersion de desempeno por cadena: reportar no solo el promedio (que puede verse "
            "'estable') sino el % de salas en caida vs en mejora -- evita subestimar el problema real "
            "cuando salas buenas compensan a las malas en el promedio.",
            "Media")
        add("SUGERENCIA_INDICADOR",
            "Cobertura de catalogo activo: trackear semana a semana cuantos SKU siguen con venta vs se "
            "van cayendo a cero, como alerta temprana de perdida de espacio o relacion comercial "
            "(complementario a 'producto inhabilitado', pero en tendencia, no solo foto puntual).",
            "Media")
        add("SUGERENCIA_INDICADOR",
            "Tendencia semana a semana DENTRO del año actual (no solo YoY): con 2+ semanas cerradas de "
            "2026 ya se puede ver si el ritmo semanal esta acelerando o desacelerando, sin esperar el "
            "cierre completo del mes para reaccionar.",
            "Baja")

        log(f"Total filas de observaciones a escribir para {CLIENTE}: {len(filas_obs)}")

        # --- Escribir en hoja OBSERVACIONES (crear si no existe) ---
        nombres_hojas = [s.Name for s in wb.Worksheets]
        if "OBSERVACIONES" not in nombres_hojas:
            ws_obs = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws_obs.Name = "OBSERVACIONES"
            ws_obs.Range("A1:E1").Value = OBS_HEADERS
            log("Hoja OBSERVACIONES creada.")
        else:
            ws_obs = wb.Worksheets("OBSERVACIONES")
            log("Hoja OBSERVACIONES ya existia, se reutiliza.")

        last_row_obs = ws_obs.UsedRange.Rows.Count
        existentes = []
        if last_row_obs > 1:
            existentes = ws_obs.Range(f"A2:E{last_row_obs}").Value
            # sesion 34: mismo fix que en build_plantilla.py -- si hay 1 sola fila, COM
            # devuelve una tupla PLANA (una fila, no una tupla-de-tuplas), y envolverla mal
            # ("[existentes]") la trataria como una fila de 1 sola columna en vez de 5.
            if not isinstance(existentes, tuple) or (existentes and not isinstance(existentes[0], tuple)):
                existentes = (existentes,)

        # Conservar filas de OTROS clientes, descartar las de este CLIENTE (se reemplazan)
        conservar = [r for r in existentes if r and r[0] and str(r[0]).strip() != CLIENTE]
        nuevas = conservar + filas_obs
        log(f"Filas conservadas de otros clientes: {len(conservar)} | Filas nuevas de {CLIENTE}: {len(filas_obs)} | Total final: {len(nuevas)}")

        # Limpiar y reescribir todo el bloque de datos -- fila por fila + verificacion,
        # ver rt_common.py (mismo patron de seguridad que build_plantilla.py).
        if last_row_obs > 1:
            ws_obs.Range(f"A2:E{last_row_obs}").ClearContents()
        rc.escribir_filas_verificado(ws_obs, fila_inicio=2, num_columnas=5, out_rows=nuevas, log=log)

        log("Guardando...")
        wb.Save()
        log("Guardado.")

        log(f"OBSERVACIONES UsedRange filas: {ws_obs.UsedRange.Rows.Count}")
        log("PROCESO_OK")
    # cierre de Excel a cargo de rc.abrir_excel_com() -- ver rt_common.py.

if __name__ == "__main__":
    main()

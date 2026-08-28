# -*- coding: utf-8 -*-
"""
build_resumen_gestores.py (sesion 34) -- genera 02_SIN_REPORTE_OSA_Y_VISITA_GESTORES.xlsx
completo (LOCALES_SIN_VISITA, OSA_SIN_EXHIBIR, RESUMEN) a partir de las hojas REPORTES y
OSA del archivo liviano (02_MINUTAS_CLIENTES_PLANTILLA.xlsm).

Antes de este script, el usuario copiaba a mano las hojas del archivo liviano hacia este
archivo, y el RESUMEN (dashboard por gestor) se armaba con formulas de Excel encima.
Ese proceso manual tenia 2 problemas reales, encontrados en sesion 34 y corregidos aca:

1) Gestor "#N/A" visible en el reporte que reciben los supervisores -- 3 filas donde el
   cruce de gestor fallaba quedaban con el error de Excel crudo, incluida una fila entera
   de "Gestor: #N/A" en el panel de OSA del RESUMEN. Aca: si GESTORES viene vacio/None,
   se usa el texto explicito GESTOR_SIN_ASIGNAR -- nunca un error crudo.

2) Las columnas de fecha (17 al 26 de agosto, ej.) mostraban el stock/foco de ese dia
   FORMATEADO COMO FECHA en vez de numero (12 se veia como "12-ene-1900"). Aca: las hojas
   de destino se BORRAN Y RECREAN de cero en cada corrida (Worksheets.Delete + Add), asi
   no heredan ningun formato de celda de una corrida anterior -- nunca hay arrastre.

Parametro a actualizar CADA SEMANA (igual criterio que SEMANAS_2026_CERRADAS en
build_observaciones.py): SEMANA_ACTUAL_DESDE, la fecha desde la que se considera que un
reporte SI se actualizo esta semana. Un reporte cuya ULTIMA fecha con dato sea ANTERIOR a
esta fecha cuenta como "atrasado".
"""
import datetime
from collections import defaultdict, Counter

import rt_common as rc

FILE_ORIGEN = r"C:\Users\basti\Desktop\RT\02_MINUTAS_CLIENTES_PLANTILLA.xlsm"
FILE_DESTINO = r"C:\Users\basti\Desktop\RT\02_SIN_REPORTE_OSA_Y_VISITA_GESTORES.xlsx"

# <-- ACTUALIZAR cada semana: primer dia de la semana EN CURSO (no cerrada).
SEMANA_ACTUAL_DESDE = datetime.date(2026, 8, 25)

GESTOR_SIN_ASIGNAR = "SIN ASIGNAR"
TOP_N_CLIENTES_FOCO = 3

MESES_ES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
            "septiembre", "octubre", "noviembre", "diciembre"]


def log(msg):
    print(msg, flush=True)


def parse_fecha_header(valor):
    """Los headers de columnas de fecha en OSA vienen como texto 'DD-MM-AAAA'."""
    if isinstance(valor, datetime.datetime):
        return valor.date()
    if isinstance(valor, datetime.date):
        return valor
    if isinstance(valor, str):
        try:
            return datetime.datetime.strptime(valor.strip(), "%d-%m-%Y").date()
        except ValueError:
            return None
    return None


def leer_tabla_com(ws):
    """Lee una hoja completa (headers en fila 1, datos desde fila 2) via COM, devolviendo
    (headers, filas-como-tuplas). Robusto al caso de 0 o 1 fila de datos."""
    n_cols = ws.UsedRange.Columns.Count
    n_rows = ws.UsedRange.Rows.Count
    headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, n_cols)).Value
    if not isinstance(headers, tuple):
        headers = (headers,)
    elif isinstance(headers[0], tuple):
        headers = headers[0]

    if n_rows <= 1:
        return list(headers), []

    data = ws.Range(ws.Cells(2, 1), ws.Cells(n_rows, n_cols)).Value
    if not isinstance(data, tuple) or (data and not isinstance(data[0], tuple)):
        data = (data,)
    return list(headers), list(data)


def obtener_o_recrear_hoja(wb, nombre):
    """Borra la hoja si ya existe y la vuelve a crear vacia -- asi nunca hereda formato
    de celda de una corrida anterior (esto es lo que evita el bug de fechas)."""
    nombres_hojas = [s.Name for s in wb.Worksheets]
    if nombre in nombres_hojas:
        wb.Worksheets(nombre).Delete()
    return wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))


def main():
    with rc.abrir_excel_com(FILE_ORIGEN) as (app, wb_o):
        log(f"Leyendo fuente: {FILE_ORIGEN}")
        ws_r = wb_o.Worksheets("REPORTES")
        ws_o = wb_o.Worksheets("OSA")

        r_headers, r_data = leer_tabla_com(ws_r)
        r_idx = {h: i for i, h in enumerate(r_headers)}

        locales_sin_visita = []
        for row in r_data:
            gestor_raw = rc.valor_o_default(row[r_idx["GESTORES"]], None)
            gestor = str(gestor_raw).strip() if gestor_raw else GESTOR_SIN_ASIGNAR
            locales_sin_visita.append({
                "gestor": gestor,
                "cod_kpi": row[r_idx["COD KPI ONE"]],
                "local": row[r_idx["LOCAL"]],
                "cliente": row[r_idx["CLIENTE"]],
                "veces_semana": row[r_idx["VECES POR SEMANA"]] or 0,
                "faltante": row[r_idx["FALTANTE"]],
            })
        log(f"REPORTES -> {len(locales_sin_visita)} filas (LOCALES_SIN_VISITA).")

        o_headers, o_data = leer_tabla_com(ws_o)
        o_idx = {h: i for i, h in enumerate(o_headers)}
        fecha_cols = [(i, parse_fecha_header(h)) for i, h in enumerate(o_headers)]
        fecha_cols = [(i, f) for i, f in fecha_cols if f is not None]
        col_total = len(o_headers) - 1  # "Total general" es siempre la ultima columna

        osa_sin_exhibir = []
        for row in o_data:
            gestor_raw = rc.valor_o_default(row[o_idx["GESTORES"]], None)
            gestor = str(gestor_raw).strip() if gestor_raw else GESTOR_SIN_ASIGNAR
            ultima_fecha = None
            valores_fecha = {}
            for i, fecha in fecha_cols:
                if row[i] is not None:
                    valores_fecha[fecha] = row[i]
                    if ultima_fecha is None or fecha > ultima_fecha:
                        ultima_fecha = fecha
            osa_sin_exhibir.append({
                "gestor": gestor,
                "cliente": row[o_idx["CLIENTE"]],
                "cod_kpi": row[o_idx["COD KPI ONE"]],
                "local": row[o_idx["Local"]],
                "producto": row[o_idx["Producto"]],
                "valores_fecha": valores_fecha,
                "stock_total": rc.valor_o_default(row[col_total], 0) or 0,
                "ultima_fecha": ultima_fecha,
                "atrasado": bool(ultima_fecha and ultima_fecha < SEMANA_ACTUAL_DESDE),
            })
        log(f"OSA -> {len(osa_sin_exhibir)} filas (OSA_SIN_EXHIBIR).")

        fechas_ordenadas = sorted({f for _i, f in fecha_cols})

        # --- Panel 1: visitas sin reporte, por gestor ---
        por_gestor_v = defaultdict(lambda: {"locales": 0, "visitas": 0, "clientes": Counter()})
        for r in locales_sin_visita:
            g = por_gestor_v[r["gestor"]]
            g["locales"] += 1
            g["visitas"] += r["veces_semana"]
            if r["cliente"]:
                g["clientes"][str(r["cliente"]).strip()] += 1
        total_visitas = sum(v["visitas"] for v in por_gestor_v.values())

        panel1 = []
        for gestor, v in sorted(por_gestor_v.items(), key=lambda kv: -kv[1]["visitas"]):
            clientes_foco = ", ".join(c for c, _n in v["clientes"].most_common(TOP_N_CLIENTES_FOCO))
            pct = (v["visitas"] / total_visitas) if total_visitas else 0
            panel1.append([gestor, v["locales"], v["visitas"], pct, clientes_foco])
        panel1.append(["TOTAL", sum(v["locales"] for v in por_gestor_v.values()),
                        total_visitas, 1.0 if total_visitas else 0, None])

        # --- Panel 2: OSA -- quiebre y recurrencia, por gestor ---
        por_gestor_o = defaultdict(lambda: {"lineas": 0, "stock": 0.0, "atrasado": 0})
        for r in osa_sin_exhibir:
            g = por_gestor_o[r["gestor"]]
            g["lineas"] += 1
            g["stock"] += r["stock_total"]
            if r["atrasado"]:
                g["atrasado"] += 1
        total_lineas = sum(v["lineas"] for v in por_gestor_o.values())
        total_stock = sum(v["stock"] for v in por_gestor_o.values())
        total_atrasado = sum(v["atrasado"] for v in por_gestor_o.values())

        panel2 = []
        for gestor, v in sorted(por_gestor_o.items(), key=lambda kv: -kv[1]["lineas"]):
            pct_atr = (v["atrasado"] / v["lineas"]) if v["lineas"] else 0
            panel2.append([gestor, v["lineas"], v["stock"], v["atrasado"], pct_atr])
        panel2.append(["TOTAL", total_lineas, total_stock, total_atrasado,
                        (total_atrasado / total_lineas) if total_lineas else 0])

        log(f"Panel visitas: {len(panel1) - 1} gestores, {total_visitas} visitas sin reporte.")
        log(f"Panel OSA: {len(panel2) - 1} gestores, {total_lineas} lineas quiebre, "
            f"{total_atrasado} atrasadas ({(total_atrasado/total_lineas*100 if total_lineas else 0):.1f}%).")

        # ============================ ESCRITURA EN DESTINO ============================
        log(f"Abriendo destino: {FILE_DESTINO}")
        wb_d = app.Workbooks.Open(FILE_DESTINO, UpdateLinks=0, ReadOnly=False)
        try:
            # --- LOCALES_SIN_VISITA ---
            ws_lsv = obtener_o_recrear_hoja(wb_d, "LOCALES_SIN_VISITA")
            ws_lsv.Name = "LOCALES_SIN_VISITA"
            headers_lsv = ["GESTORES", "COD KPI ONE", "LOCAL", "CLIENTE", "VECES POR SEMANA", "FALTANTE"]
            ws_lsv.Range("A1:F1").Value = headers_lsv
            filas_lsv = [[r["gestor"], r["cod_kpi"], r["local"], r["cliente"], r["veces_semana"], r["faltante"]]
                         for r in locales_sin_visita]
            rc.escribir_filas_verificado(ws_lsv, fila_inicio=2, num_columnas=6, out_rows=filas_lsv, log=log)

            # --- OSA_SIN_EXHIBIR ---
            ws_ose = obtener_o_recrear_hoja(wb_d, "OSA_SIN_EXHIBIR")
            ws_ose.Name = "OSA_SIN_EXHIBIR"
            headers_ose = ["GESTOR", "CLIENTE", "COD KPI ONE", "Local", "Producto"] + \
                          [datetime.datetime(f.year, f.month, f.day) for f in fechas_ordenadas] + \
                          ["Stock Total"]
            n_cols_ose = len(headers_ose)
            ws_ose.Range(ws_ose.Cells(1, 1), ws_ose.Cells(1, n_cols_ose)).Value = headers_ose
            filas_ose = []
            for r in osa_sin_exhibir:
                fila = [r["gestor"], r["cliente"], r["cod_kpi"], r["local"], r["producto"]]
                fila += [r["valores_fecha"].get(f) for f in fechas_ordenadas]
                fila.append(r["stock_total"])
                filas_ose.append(fila)
            rc.escribir_filas_verificado(ws_ose, fila_inicio=2, num_columnas=n_cols_ose, out_rows=filas_ose, log=log)
            # Formato explicito "General" en las columnas de fecha/stock -- hoja recien
            # creada ya deberia venir en General, esto es cinturon-y-tirantes.
            if filas_ose:
                col_ini = 6
                col_fin = n_cols_ose
                ws_ose.Range(ws_ose.Cells(2, col_ini), ws_ose.Cells(1 + len(filas_ose), col_fin)).NumberFormat = "General"

            # --- RESUMEN ---
            ws_res = obtener_o_recrear_hoja(wb_d, "RESUMEN")
            ws_res.Name = "RESUMEN"
            if fechas_ordenadas:
                periodo = (f"{fechas_ordenadas[0].day} al {fechas_ordenadas[-1].day} de "
                           f"{MESES_ES[fechas_ordenadas[-1].month]} {fechas_ordenadas[-1].year}")
            else:
                periodo = "(sin fechas detectadas)"
            hoy = datetime.date.today()

            filas_resumen = []
            filas_resumen.append([None, f"RESUMEN Y ALERTAS  ·  Periodo {periodo}", None, None, None, None])
            filas_resumen.append([None,
                f"Generado automaticamente el {hoy.day:02d}-{hoy.month:02d}-{hoy.year} desde REPORTES/OSA de "
                f"02_MINUTAS_CLIENTES_PLANTILLA.xlsm. Reporte atrasado = ultima fecha con dato anterior a "
                f"{SEMANA_ACTUAL_DESDE.strftime('%d-%m-%Y')} (semana en curso).",
                None, None, None, None])
            filas_resumen.append([None, None, None, None, None, None])
            filas_resumen.append([None, "VISITAS SIN REPORTE", None, None, None, None])
            filas_resumen.append([None, "Gestor", "Locales", "Visitas no report.", "% incump.", "Clientes foco"])
            for fila in panel1:
                filas_resumen.append([None] + fila)
            filas_resumen.append([None, None, None, None, None, None])
            filas_resumen.append([None, "OSA — QUIEBRE Y RECURRENCIA DE REPORTE", None, None, None, None])
            filas_resumen.append([None, "Gestor", "Líneas quiebre", "Stock (u)", "Reporte atrasado", "% atrasado"])
            for fila in panel2:
                filas_resumen.append([None] + fila)
            filas_resumen.append([None, None, None, None, None, None])
            filas_resumen.append([None,
                f"Reporte atrasado = ultimo reporte del local anterior a {SEMANA_ACTUAL_DESDE.strftime('%d-%m-%Y')} "
                f"(no actualizado en la semana en curso). Gestor sin cruce -> '{GESTOR_SIN_ASIGNAR}' (nunca #N/A).",
                None, None, None, None])

            rc.escribir_filas_verificado(ws_res, fila_inicio=1, num_columnas=6, out_rows=filas_resumen, log=log)

            log("Guardando destino...")
            wb_d.Save()
            log("Guardado.")
        finally:
            wb_d.Close(SaveChanges=False)

    log("PROCESO_OK")


if __name__ == "__main__":
    main()

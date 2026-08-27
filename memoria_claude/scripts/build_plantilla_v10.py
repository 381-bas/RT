# -*- coding: utf-8 -*-
"""
build_plantilla_v7.py (sesion 24)
Version consolidada final para el archivo LIVIANO 02_MINUTAS_CLIENTES_PLANTILLA.xlsm.

Hereda de versiones anteriores:
  v3: deteccion de columnas por NOMBRE (ListObjects) en vez de posicion fija.
  v4: SEGUIMIENTO/OSA/NEGATIVO/INMOVILIZADO -- OSA consolida productos en 1 linea sin
      Local/STOCK; INMOVILIZADO se descarta COMPLETO a nivel de SALA si hay cualquier OSA;
      NEGATIVO nunca se suprime; exclusion 'LIBRO' para TNOGAL (en MINUTA y en OSA).
  v5: reglas especiales por cliente -- CLIENTES_SIN_REPORTES / CLIENTES_MENSAJE_GENERAL
      (hoy: CORRALES DEL SUR -- sin SEGUIMIENTO, 1 sola linea 'REVISION GENERAL' si hay
      cualquier OSA/NEGATIVO/INMOVILIZADO).
  v6: columna Riesgo (0 a 1) en PLANTILLA, columna M.

NUEVO v7 (sesion 24): Riesgo ponderado por MAGNITUD DE STOCK, no solo presencia/ausencia.
  Hallazgo del usuario (sesion 23, confirmado con datos reales de BIGU en sesion 24):
  INMOVILIZADO siempre tiene venta=0 (tautologico, confirmado con VTA_S-1/VTA_S-2 de MINUTA),
  asi que el dato de venta no aporta calibracion ahi -- pero el STOCK si: stock bajo (ej. 4u)
  probablemente es ruido/desajuste fisico-vs-B2B, stock alto es capital muerto real.
  Mismo criterio se aplica a OSA usando STOCK_B2B (columna 'Total general' del panel OSA).
  peso_stock(stock, umbral) = min(1, stock/umbral) -- continuo, no binario.
  NEGATIVO se deja SIN ponderar (cualquier stock negativo ya es anomalia real en si misma).

NUEVO v8 (sesion 25): texto de INMOVILIZADO renovado ('buscar mas caras' generaba
  resistencia del equipo -- si el producto no vende, 'vender mas' no es una accion clara).
  Nuevo texto por defecto: enfocado en LAYOUT/EXHIBICION, no en volumen de venta.
  Se agrega ACCIONES_POR_CLIENTE para permitir mensajes personalizados por cliente a
  futuro, sin tocar el default compartido.

NUEVO v9 (sesion 26, cliente BERRYSUR): 2 cambios.
  1) INMOVILIZADO ahora se CONSOLIDA en 1 sola linea por sala (igual que OSA desde sesion
     17), agrupando pares SKU+Producto en vez de 1 linea por producto -- para cuentas con
     muchas combinaciones (BERRYSUR: hasta 5 en una sola sala) donde itemizar es pesado de
     leer. Aplica a TODOS los clientes (mismo criterio de consistencia que OSA).
  2) PESOS_POR_CLIENTE: permite subir/bajar la influencia relativa de cada tipo de
     indicador en el calculo de Riesgo, por cliente. BERRYSUR: OSA x1.5, INMOVILIZADO x0.5
     -- porque en esta cuenta el producto es de baja rotacion por naturaleza (INMOVILIZADO
     es casi siempre 1, poco informativo como diferenciador) y OSA es el foco operativo
     real a gestionar.
"""
import datetime
from collections import defaultdict
import win32com.client as win32
import pythoncom

FILE_PATH = r"C:\Users\basti\Desktop\RT\02_MINUTAS_CLIENTES_PLANTILLA.xlsm"

ACCIONES = {
    "OSA": "Verificar visita y stock actualizado.",
    "STOCK NEGATIVO": "Realizar ajuste.",
    "INMOVILIZADO": "Ajustar layout de sala y buscar oportunidades de mejora en exhibición.",
    "SEGUIMIENTO": "Verificar disponibilidad en local.",
    "REVISION_GENERAL": "Realizar barrido general de la sala: ajustar inventario y verificar exhibición de productos.",
}

# sesion 25: mensajes personalizados por cliente (opcional). Si un cliente tiene una
# entrada aqui para un tipo de indicador, se usa esa en vez del default de ACCIONES.
ACCIONES_POR_CLIENTE = {
    # "BIGU": {"INMOVILIZADO": "texto especifico de BIGU si se define a futuro"},
}

def get_accion(tipo, cliente):
    return ACCIONES_POR_CLIENTE.get(cliente, {}).get(tipo, ACCIONES[tipo])

CLIENTES_SIN_REPORTES = {"CORRALES DEL SUR"}
CLIENTES_MENSAJE_GENERAL = {"CORRALES DEL SUR"}

# sesion 24: umbrales de "stock significativo" para ponderar severidad (unidades)
UMBRAL_STOCK_OSA = 15
UMBRAL_STOCK_INMOV = 15

# sesion 26: peso relativo de cada tipo de indicador en el calculo de Riesgo, por cliente.
# Default: todos pesan 1.0 (igual que antes). Un cliente puede subir/bajar la influencia
# de un tipo especifico sin afectar a los demas clientes.
PESOS_DEFAULT = {"SEGUIMIENTO": 1.0, "OSA": 1.0, "STOCK NEGATIVO": 1.0, "INMOVILIZADO": 1.0}
PESOS_POR_CLIENTE = {
    "BERRYSUR": {"OSA": 1.5, "INMOVILIZADO": 0.5},  # producto de baja rotacion: OSA es el foco real
}

def get_pesos(cliente):
    p = dict(PESOS_DEFAULT)
    p.update(PESOS_POR_CLIENTE.get(cliente, {}))
    return p

# sesion 27: multiplicador de Riesgo por CADENA, por cliente (distinto del peso por tipo
# de indicador arriba). Se aplica al final, sobre el Riesgo ya calculado, y se cachea a 1.0.
PESOS_CADENA_DEFAULT = 1.0
PESOS_CADENA_POR_CLIENTE = {
    "ALUSWEET": {"WALMART": 1.3},  # foco operativo del equipo va a WALMART (mas salas, mas medicion)
}

def get_peso_cadena(cliente, cadena):
    return PESOS_CADENA_POR_CLIENTE.get(cliente, {}).get(cadena, PESOS_CADENA_DEFAULT)

def log(msg):
    print(msg, flush=True)

def norm_code(x):
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return str(x).strip()

def fmt_num(x):
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x)

def es_libro(texto):
    return bool(texto) and "LIBRO" in str(texto).upper()

def peso_stock(stock, umbral):
    """0 a 1: severidad continua segun magnitud de stock. Stock bajo -> peso bajo
    (probable ruido/desajuste), stock >= umbral -> peso maximo 1.0."""
    try:
        s = float(stock)
    except (TypeError, ValueError):
        s = 0.0
    if umbral <= 0:
        return 1.0
    return max(0.0, min(1.0, s / umbral))

def calcular_riesgo(datos, cliente, cadena):
    """Riesgo (0 a 1) = 0.5*(tipos/max_tipos) + 0.5*(min(focos,6)/6), con OSA e
    INMOVILIZADO ponderados por magnitud de stock (peso_stock), cada tipo ademas
    escalado por el peso relativo del cliente (PESOS_POR_CLIENTE, default 1.0 todos),
    y el resultado final multiplicado por el peso de CADENA del cliente (capado a 1.0)."""
    pesos = get_pesos(cliente)

    peso_seguimiento = (1.0 if datos["sin_visita"] else 0.0) * pesos["SEGUIMIENTO"]

    pesos_osa = [peso_stock(stock, UMBRAL_STOCK_OSA) for (_prod, stock) in datos["osa_productos"]]
    osa_presente = (1.0 if pesos_osa else 0.0) * pesos["OSA"]
    osa_peso_total = sum(pesos_osa) * pesos["OSA"]

    peso_negativo = (1.0 if datos["negativo"] else 0.0) * pesos["STOCK NEGATIVO"]
    negativo_focos = len(datos["negativo"]) * pesos["STOCK NEGATIVO"]  # conteo sin ponderar por stock, si por cliente

    pesos_inmov = [peso_stock(stock, UMBRAL_STOCK_INMOV) for (_sku, _prod, stock) in datos["inmov"]]
    inmov_severidad = (max(pesos_inmov) if pesos_inmov else 0.0) * pesos["INMOVILIZADO"]
    inmov_peso_total = sum(pesos_inmov) * pesos["INMOVILIZADO"]

    max_tipos = sum(pesos.values())
    tipos = peso_seguimiento + osa_presente + peso_negativo + inmov_severidad
    focos = peso_seguimiento + osa_peso_total + negativo_focos + inmov_peso_total

    riesgo_base = 0.5 * (tipos / max_tipos) + 0.5 * (min(focos, 6) / 6)
    riesgo_final = min(1.0, riesgo_base * get_peso_cadena(cliente, cadena))
    return round(riesgo_final, 2)

def find_header_row_and_cols(ws, needed_names, max_scan_rows=10, max_scan_cols=30):
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(max_scan_rows, max_scan_cols)).Value
    for r_idx, row in enumerate(rng, start=1):
        found = {}
        for c_idx, val in enumerate(row, start=1):
            if val in needed_names:
                found.setdefault(val, []).append(c_idx)
        if all(n in found for n in needed_names):
            return r_idx, found
    raise ValueError(f"No se encontro fila de encabezado con: {needed_names}")

def main():
    pythoncom.CoInitialize()
    app = win32.gencache.EnsureDispatch(win32.DispatchEx('Excel.Application'))
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    app.ScreenUpdating = False
    wb = None
    try:
        log("Abriendo workbook liviano (instancia independiente)...")
        wb = app.Workbooks.Open(FILE_PATH, UpdateLinks=0, ReadOnly=False)

        ws_m = wb.Worksheets("MINUTA")
        ws_p = wb.Worksheets("PLANTILLA")

        # Detectar si REPORTES/OSA son Tablas (ListObjects, archivo liviano clasico) o
        # rangos con pivot-like headers (fallback). Preferimos ListObjects si existen.
        ws_r = wb.Worksheets("REPORTES")
        ws_o = wb.Worksheets("OSA")

        # --- Tabla6 (prefijos) ---
        lo6 = ws_m.ListObjects("Tabla6")
        tabla6 = {}
        for row in lo6.DataBodyRange.Value:
            cadena, suf1, suf2, obs = row
            if not cadena:
                continue
            prefijos = [p for p in (suf1, suf2) if p]
            accion = "conservar" if obs and "conservar" in str(obs).lower() else "numerico"
            tabla6[str(cadena).strip().upper()] = {"prefijos": prefijos, "accion": accion}
        log(f"Tabla6: {tabla6}")

        def inferir_cadena(codigo_crudo):
            s = str(codigo_crudo).strip().upper()
            for cadena, info in tabla6.items():
                for p in info["prefijos"]:
                    if s.startswith(p.upper()):
                        resto = s[len(p):]
                        if resto.isdigit() or info["accion"] == "conservar":
                            return cadena
            return None

        def normalizar(codigo_crudo, cadena):
            s = str(codigo_crudo).strip()
            info = tabla6.get(cadena)
            if not info:
                return norm_code(s)
            if info["accion"] == "conservar":
                return norm_code(s)
            for p in info["prefijos"]:
                if s.upper().startswith(p.upper()):
                    resto = s[len(p):]
                    if resto.isdigit():
                        return norm_code(int(resto))
            return norm_code(s)

        # --- MINUTA (Tabla4) ---
        lo4 = ws_m.ListObjects("Tabla4")
        names4 = [lo4.ListColumns(i).Name for i in range(1, lo4.ListColumns.Count + 1)]
        idx4 = {n: i for i, n in enumerate(names4)}
        minuta_data = lo4.DataBodyRange.Value
        clientes_foco = sorted({str(r[idx4["CLIENTE"]]).strip() for r in minuta_data if r[idx4["CLIENTE"]]})
        log(f"Clientes en MINUTA (este archivo): {clientes_foco}")

        minuta_rows = []
        for r in minuta_data:
            cliente = r[idx4["CLIENTE"]]
            if not cliente:
                continue
            cliente = str(cliente).strip()
            desc_prod = r[idx4["DESCRIPCION_PRODUCTO"]]
            if cliente.upper() == "TNOGAL" and es_libro(desc_prod):
                continue
            retail_raw = r[idx4["RETAIL"]]
            minuta_rows.append({
                # sesion 27: normalizar mayusculas -- se detecto RETAIL con casing
                # inconsistente en ALUSWEET ('Walmart'/'Tottus'/'Cencosud'), lo que partia
                # una misma sala en 2 grupos distintos al no calzar con la cadena en
                # mayusculas que usan REPORTES/OSA (via inferir_cadena, siempre upper).
                "cadena": str(retail_raw).strip().upper() if retail_raw else None, "cliente": cliente,
                "cod_local_norm": norm_code(r[idx4["COD_LOCAL"]]),
                "sku": r[idx4["COD_CENCOSUD"]], "desc_prod": desc_prod,
                "stock": r[idx4["STOCK(Un)"]],
                "negativo": r[idx4["NEGATIVO"]], "inmov": r[idx4["INMOVILZIADO"]],
            })
        log(f"MINUTA filas utilizables (tras exclusion LIBRO/TNOGAL): {len(minuta_rows)}")

        # --- REPORTES = CUMPLIMIENTO (Tabla1) ---
        lo1 = ws_r.ListObjects("Tabla1")
        names1 = [lo1.ListColumns(i).Name for i in range(1, lo1.ListColumns.Count + 1)]
        idx1 = {n: i for i, n in enumerate(names1)}
        visita0 = []
        for r in lo1.DataBodyRange.Value:
            cliente = r[idx1["CLIENTE"]]
            if not cliente or str(cliente).strip() not in clientes_foco:
                continue
            if str(cliente).strip() in CLIENTES_SIN_REPORTES:
                continue
            cod_kpi = r[idx1["COD KPI ONE"]]
            local = r[idx1["LOCAL"]]
            cadena = inferir_cadena(cod_kpi)
            cod_local_norm = normalizar(cod_kpi, cadena) if cadena else norm_code(cod_kpi)
            visita0.append({"cliente": str(cliente).strip(), "cadena": cadena,
                             "cod_local_norm": cod_local_norm, "local_texto": local})
        log(f"REPORTES/CUMPLIMIENTO filas (filtradas a clientes de MINUTA): {len(visita0)}")

        # --- OSA (Tabla3) ---
        lo3 = ws_o.ListObjects("Tabla3")
        names3 = [lo3.ListColumns(i).Name for i in range(1, lo3.ListColumns.Count + 1)]
        idx3 = {n: i for i, n in enumerate(names3)}
        c_total_name = names3[-1]  # 'Total general' es la ultima columna
        osa_rows = []
        for r in lo3.DataBodyRange.Value:
            cliente = r[idx3["CLIENTE"]]
            if not cliente or str(cliente).strip() not in clientes_foco:
                continue
            cliente = str(cliente).strip()
            cod_kpi = r[idx3["COD KPI ONE"]]
            local = r[idx3["Local"]]
            producto = r[idx3["Producto"]]
            total = r[idx3[c_total_name]]
            if cliente.upper() == "TNOGAL" and es_libro(producto):
                continue
            cadena = inferir_cadena(cod_kpi)
            cod_local_norm = normalizar(cod_kpi, cadena) if cadena else norm_code(cod_kpi)
            osa_rows.append({"cliente": cliente, "cadena": cadena, "cod_local_norm": cod_local_norm,
                              "local_texto": local, "producto": producto, "stock_total": total})
        log(f"OSA filas (filtradas a clientes de MINUTA, tras exclusion LIBRO/TNOGAL): {len(osa_rows)}")

        # --- Consolidar por (cliente, cadena, cod_local_norm) ---
        # osa_productos: lista de (producto, stock) -- se usa tanto para el texto (solo nombre,
        # deduplicado) como para el Riesgo (peso por stock).
        grupos = defaultdict(lambda: {"sin_visita": False, "osa_productos": [], "negativo": [], "inmov": []})
        for v in visita0:
            k = (v["cliente"], v["cadena"], v["cod_local_norm"])
            grupos[k]["sin_visita"] = True

        for o in osa_rows:
            k = (o["cliente"], o["cadena"], o["cod_local_norm"])
            nombres_ya = [p for p, _s in grupos[k]["osa_productos"]]
            if o["producto"] not in nombres_ya:
                grupos[k]["osa_productos"].append((o["producto"], o["stock_total"]))

        for m in minuta_rows:
            k = (m["cliente"], m["cadena"], m["cod_local_norm"])
            if m["negativo"]:
                grupos[k]["negativo"].append((m["sku"], m["desc_prod"], m["stock"]))
            if m["inmov"]:
                grupos[k]["inmov"].append((m["sku"], m["desc_prod"], m["stock"]))

        filas_finales = []
        for (cliente, cadena, cod_local), datos in grupos.items():
            partes = []
            if datos["sin_visita"]:
                partes.append(f"SEGUIMIENTO | {get_accion('SEGUIMIENTO', cliente)}")

            if cliente in CLIENTES_MENSAJE_GENERAL:
                if datos["osa_productos"] or datos["negativo"] or datos["inmov"]:
                    partes.append(f"REVISION GENERAL | {get_accion('REVISION_GENERAL', cliente)}")
            else:
                if datos["osa_productos"]:
                    productos_txt = ", ".join(p for p, _s in datos["osa_productos"])
                    partes.append(f"OSA | {productos_txt} | {get_accion('OSA', cliente)}")
                for sku, prod, stock in datos["negativo"]:
                    partes.append(f"STOCK NEGATIVO | SKU: {fmt_num(sku)} | {prod} | STOCK: {fmt_num(stock)} | {get_accion('STOCK NEGATIVO', cliente)}")
                if not datos["osa_productos"] and datos["inmov"]:
                    # sesion 26: consolidado en 1 sola linea (igual criterio que OSA),
                    # pares SKU+Producto separados por coma, sin STOCK individual.
                    pares_txt = ", ".join(f"{fmt_num(sku)} {prod}" for sku, prod, _stock in datos["inmov"])
                    partes.append(f"INMOVILIZADO | {pares_txt} | {get_accion('INMOVILIZADO', cliente)}")

            if not partes:
                continue
            riesgo = calcular_riesgo(datos, cliente, cadena)
            filas_finales.append((cliente, cadena, cod_local, " / ".join(partes), riesgo))

        log(f"TOTAL filas a escribir en PLANTILLA: {len(filas_finales)}")
        por_cliente = defaultdict(int)
        for f in filas_finales:
            por_cliente[f[0]] += 1
        log(f"Por cliente: {dict(por_cliente)}")

        hoy_date = datetime.date.today()
        semana_iso = hoy_date.isocalendar()[1]
        hoy = datetime.datetime(hoy_date.year, hoy_date.month, hoy_date.day)
        log(f"Fecha de ejecucion: {hoy.isoformat()} | Semana ISO: {semana_iso}")

        if ws_p.Range("M1").Value != "Riesgo":
            ws_p.Range("M1").Value = "Riesgo"
            log("PLANTILLA: columna M1='Riesgo' agregada.")

        last_row_p = ws_p.UsedRange.Rows.Count
        if last_row_p > 1:
            ws_p.Range(f"A2:M{last_row_p}").ClearContents()
            log(f"PLANTILLA: filas 2:{last_row_p} limpiadas.")

        n = len(filas_finales)
        if n > 0:
            out_rows = []
            for cliente, cadena, cod_local, observacion, riesgo in filas_finales:
                out_rows.append([semana_iso, hoy, cliente, cadena, cod_local, None,
                                  None, None, observacion, None, None, None, riesgo])
            ws_p.Range(f"A2:M{1+n}").Value = out_rows
            log(f"PLANTILLA: {n} filas escritas (A2:M{1+n}).")

        log("Guardando (archivo liviano, sin formulas, sin recalculo)...")
        wb.Save()
        log("Guardado.")

        final_used = ws_p.UsedRange.Rows.Count
        log(f"PLANTILLA UsedRange filas: {final_used}")
        sample = ws_p.Range("A1:M8").Value
        for r in sample:
            log(str(r))

        log("PROCESO_OK")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception as e:
            log(f"Aviso cerrando wb: {e}")
        try:
            app.Quit()
        except Exception as e:
            log(f"Aviso cerrando app: {e}")
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()

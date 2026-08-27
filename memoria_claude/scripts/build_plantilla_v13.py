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

NUEVO v12 (sesion 30, nueva semana de gestion -- TNOGAL, MAILEMU MIEL, CORRALES DEL SUR):
  El usuario informo que las 3 hojas fuente (REPORTES, OSA y MINUTA) quedaron normalizadas
  para cruzar SIEMPRE por "COD KPI ONE" + "CLIENTE" -- MINUTA ahora trae su propia columna
  "COD KPI ONE" (ya con el prefijo de cadena correcto, igual formato que REPORTES/OSA), asi
  que ya NO hace falta el rodeo anterior de leer COD_LOCAL + RETAIL y adivinar la cadena por
  cliente. Cambios:
  1) MINUTA: la cadena y el codigo de local normalizado se derivan de idx4["COD KPI ONE"]
     con inferir_cadena()+normalizar() -- EXACTAMENTE el mismo camino que REPORTES/OSA (antes
     se usaba RETAIL uppercased + COD_LOCAL). RETAIL se conserva solo como fallback de cadena
     si inferir_cadena() no reconoce el prefijo (no deberia pasar con datos ya normalizados).
  2) QUIEBRES sigue sin usarse en ninguna parte del pipeline (confirmado explicitamente por
     el usuario: el indicador de quiebres NUNCA debe aparecer en PLANTILLA, para ningun
     cliente, porque ese dato ya se le presenta al cliente por otra via). minuta_rows no lee
     ni expone esa columna -- se deja este comentario como guardrail para no reintroducirla.
  3) Alcance de esta corrida: los unicos clientes con datos de MINUTA en este archivo son
     TNOGAL, MAILEMU MIEL y CORRALES DEL SUR (nueva semana). Como el pipeline ya filtraba
     REPORTES/OSA a `clientes_foco` (los clientes presentes en MINUTA), el resultado queda
     acotado automaticamente a estos 3 -- sin necesidad de una lista hardcodeada.

NUEVO v13 (sesion 30, continuacion -- "modelo mas eficiente" + autocompletar Local/Gestor/
  Supervisor):
  1) YA NO se hace CLEAR total de PLANTILLA. Se reemplazan SOLO las filas de los clientes
     presentes en `clientes_foco` (los que trae la MINUTA de esta corrida); las filas de
     cualquier otro cliente que ya estuviera en PLANTILLA (de una semana/corrida anterior)
     se preservan intactas, incluida cualquier edicion manual en Estatus/Respuesta
     Supervisor/Fecha Respuesta que el usuario haya hecho sobre esas filas. Esto es lo que
     permite trabajar "solo a los clientes nuevos que vayan apareciendo, o lo que se indique
     en cada mensaje" sin perder el trabajo ya cargado de los demas.
  2) Local/Gestor/Supervisor (columnas F/G/H, antes siempre en blanco) ahora se autocompletan
     cruzando por (CLIENTE, cadena, cod_local_norm) contra RUTA_RUTERO (Tabla7), que trae
     CADENA/LOCAL/GESTORES/SUPERVISOR por sala. Si una sala no aparece en RUTA_RUTERO (caso
     raro), se usa como respaldo el texto de Local/Gestor visto en la propia fila de
     REPORTES u OSA para esa sala; si tampoco hay eso, queda en blanco (nunca se inventa).
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
            # sesion 28 fix: limpiar el float ANTES de convertir a texto (norm_code primero),
            # si no, un codigo numerico limpio tipo 84.0 (WALMART/TOTTUS/SMU desde MINUTA)
            # se stringifica a "84.0" y se pierde el ".0" -- rompe el calce con OSA/REPORTES
            # (que dan "84"), duplicando la sala en 2 grupos distintos.
            s = norm_code(codigo_crudo)
            info = tabla6.get(cadena)
            if not info:
                return s
            if info["accion"] == "conservar":
                return s
            for p in info["prefijos"]:
                if s.upper().startswith(p.upper()):
                    resto = s[len(p):]
                    if resto.isdigit():
                        return norm_code(int(resto))
            return s

        # --- RUTA_RUTERO (Tabla7) -- sesion 30 (v13): fuente de Local/Gestor/Supervisor ---
        # Universo completo de salas x cliente (no solo las que tienen foco esta semana).
        # Se cruza por (CLIENTE, cadena, cod_local_norm) -- misma llave que usa `grupos` mas
        # abajo -- para autocompletar PLANTILLA sin inventar nada: si una sala no aparece
        # aqui, se intenta con el texto de Local/Gestor visto en REPORTES/OSA, y si tampoco
        # hay, queda en blanco.
        ruta_lookup = {}
        try:
            ws_ru = wb.Worksheets("RUTA_RUTERO")
            lo7 = ws_ru.ListObjects("Tabla7")
            names7 = [lo7.ListColumns(i).Name for i in range(1, lo7.ListColumns.Count + 1)]
            idx7 = {n: i for i, n in enumerate(names7)}
            n_ruta = 0
            for r in lo7.DataBodyRange.Value:
                cliente_ru = r[idx7["CLIENTE"]]
                cod_kpi_ru = r[idx7["COD KPI ONE"]]
                if not cliente_ru or not cod_kpi_ru:
                    continue
                cliente_ru = str(cliente_ru).strip()
                cadena_ru = inferir_cadena(cod_kpi_ru) or (str(r[idx7["CADENA"]]).strip().upper() if r[idx7["CADENA"]] else None)
                cod_local_ru = normalizar(cod_kpi_ru, cadena_ru) if cadena_ru else norm_code(cod_kpi_ru)
                k = (cliente_ru, cadena_ru, cod_local_ru)
                if k not in ruta_lookup:
                    ruta_lookup[k] = {
                        "local": r[idx7["LOCAL"]],
                        "gestor": r[idx7["GESTORES"]],
                        "supervisor": r[idx7["SUPERVISOR"]],
                    }
                    n_ruta += 1
            log(f"RUTA_RUTERO: {n_ruta} salas indexadas para autocompletar Local/Gestor/Supervisor.")
        except Exception as e:
            log(f"Aviso: no se pudo leer RUTA_RUTERO/Tabla7 ({e}) -- Local/Gestor/Supervisor quedaran en blanco salvo respaldo de REPORTES/OSA.")

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
            # sesion 30: MINUTA ya trae su propia columna "COD KPI ONE" (mismo formato,
            # con prefijo de cadena, que REPORTES/OSA). Se deriva cadena+codigo normalizado
            # por el MISMO camino que las otras 2 hojas fuente (inferir_cadena+normalizar),
            # en vez del rodeo anterior via RETAIL+COD_LOCAL. RETAIL queda solo de fallback
            # si por algun motivo el prefijo no calza con ninguna cadena de Tabla6.
            cod_kpi_minuta = r[idx4["COD KPI ONE"]]
            cadena_minuta = inferir_cadena(cod_kpi_minuta)
            if not cadena_minuta:
                retail_raw = r[idx4["RETAIL"]]
                cadena_minuta = str(retail_raw).strip().upper() if retail_raw else None
            cod_local_minuta = normalizar(cod_kpi_minuta, cadena_minuta) if cadena_minuta else norm_code(cod_kpi_minuta)
            minuta_rows.append({
                "cadena": cadena_minuta, "cliente": cliente,
                "cod_local_norm": cod_local_minuta,
                # QUIEBRES (idx4["QUIEBRES"]) deliberadamente NO se lee aqui -- el
                # indicador de quiebres nunca debe llegar a PLANTILLA (sesion 30).
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
            gestor = r[idx1["GESTORES"]] if "GESTORES" in idx1 else None
            cadena = inferir_cadena(cod_kpi)
            cod_local_norm = normalizar(cod_kpi, cadena) if cadena else norm_code(cod_kpi)
            visita0.append({"cliente": str(cliente).strip(), "cadena": cadena,
                             "cod_local_norm": cod_local_norm, "local_texto": local, "gestor_texto": gestor})
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
            gestor = r[idx3["GESTORES"]] if "GESTORES" in idx3 else None
            producto = r[idx3["Producto"]]
            total = r[idx3[c_total_name]]
            if cliente.upper() == "TNOGAL" and es_libro(producto):
                continue
            cadena = inferir_cadena(cod_kpi)
            cod_local_norm = normalizar(cod_kpi, cadena) if cadena else norm_code(cod_kpi)
            osa_rows.append({"cliente": cliente, "cadena": cadena, "cod_local_norm": cod_local_norm,
                              "local_texto": local, "gestor_texto": gestor, "producto": producto, "stock_total": total})
        log(f"OSA filas (filtradas a clientes de MINUTA, tras exclusion LIBRO/TNOGAL): {len(osa_rows)}")

        # --- Consolidar por (cliente, cadena, cod_local_norm) ---
        # osa_productos: lista de (producto, stock) -- se usa tanto para el texto (solo nombre,
        # deduplicado) como para el Riesgo (peso por stock).
        grupos = defaultdict(lambda: {"sin_visita": False, "osa_productos": [], "negativo": [], "inmov": [],
                                       "local_fallback": None, "gestor_fallback": None})
        for v in visita0:
            k = (v["cliente"], v["cadena"], v["cod_local_norm"])
            grupos[k]["sin_visita"] = True
            if not grupos[k]["local_fallback"] and v["local_texto"]:
                grupos[k]["local_fallback"] = v["local_texto"]
            if not grupos[k]["gestor_fallback"] and v["gestor_texto"]:
                grupos[k]["gestor_fallback"] = v["gestor_texto"]

        for o in osa_rows:
            k = (o["cliente"], o["cadena"], o["cod_local_norm"])
            nombres_ya = [p for p, _s in grupos[k]["osa_productos"]]
            if o["producto"] not in nombres_ya:
                grupos[k]["osa_productos"].append((o["producto"], o["stock_total"]))
            if not grupos[k]["local_fallback"] and o["local_texto"]:
                grupos[k]["local_fallback"] = o["local_texto"]
            if not grupos[k]["gestor_fallback"] and o["gestor_texto"]:
                grupos[k]["gestor_fallback"] = o["gestor_texto"]

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

            # sesion 30 (v13): resolver Local/Gestor/Supervisor -- prioridad RUTA_RUTERO,
            # respaldo el texto visto en REPORTES/OSA para esa misma sala, si no hay nada
            # se deja en blanco (nunca se inventa un valor).
            info_ruta = ruta_lookup.get((cliente, cadena, cod_local), {})
            local_final = info_ruta.get("local") or datos["local_fallback"]
            gestor_final = info_ruta.get("gestor") or datos["gestor_fallback"]
            supervisor_final = info_ruta.get("supervisor")

            filas_finales.append((cliente, cadena, cod_local, local_final, gestor_final,
                                   supervisor_final, " / ".join(partes), riesgo))

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

        # sesion 30 (v13): "modelo eficiente" -- ya NO se limpia toda PLANTILLA. Se leen
        # las filas existentes, se conservan tal cual las de clientes que NO estan en esta
        # corrida (`clientes_foco`), y solo se reemplazan las de los clientes que si estan.
        last_row_p = ws_p.UsedRange.Rows.Count
        filas_preservadas = []
        if last_row_p > 1:
            existentes = ws_p.Range(f"A2:M{last_row_p}").Value
            if not isinstance(existentes, tuple) or (existentes and not isinstance(existentes[0], tuple)):
                existentes = (existentes,)
            for fila in existentes:
                marca = fila[2]  # columna C = Marca (cliente)
                if marca is None:
                    continue
                if str(marca).strip() in clientes_foco:
                    continue  # se reemplaza por la version nueva calculada en esta corrida
                filas_preservadas.append(fila)
        log(f"PLANTILLA: {len(filas_preservadas)} filas preservadas de clientes fuera de esta corrida "
            f"(clientes_foco={sorted(clientes_foco)}).")

        out_rows_nuevas = []
        for cliente, cadena, cod_local, local_final, gestor_final, supervisor_final, observacion, riesgo in filas_finales:
            out_rows_nuevas.append([semana_iso, hoy, cliente, cadena, cod_local, local_final,
                                     gestor_final, supervisor_final, observacion, None, None, None, riesgo])

        out_rows = list(filas_preservadas) + out_rows_nuevas
        n = len(out_rows)

        if last_row_p > 1:
            ws_p.Range(f"A2:M{last_row_p}").ClearContents()
        if n > 0:
            ws_p.Range(f"A2:M{1+n}").Value = out_rows
        log(f"PLANTILLA: {len(filas_finales)} filas nuevas/actualizadas + {len(filas_preservadas)} preservadas "
            f"= {n} filas totales escritas (A2:M{1+n}).")

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

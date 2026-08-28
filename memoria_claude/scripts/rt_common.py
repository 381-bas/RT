# -*- coding: utf-8 -*-
"""
rt_common.py -- funciones compartidas entre los scripts de RT (build_plantilla.py,
build_observaciones.py, y cualquier script nuevo que procese los mismos datos).

Por que existe (sesion 34, 2026-08-27): norm_code(), inferir_cadena() y normalizar()
estaban copiadas LITERALMENTE en build_plantilla.py y build_observaciones.py -- si se
corregia un bug en una copia, la otra quedaba con el bug viejo sin que nadie lo notara
(ya paso una vez, sesion 28, con el orden de operaciones de normalizar()). Este modulo
es la unica fuente de verdad para esa logica.

Tambien centraliza el patron de escritura segura por COM (fila por fila + verificacion),
encontrado en sesion 32 tras corromperse filas al escribir arrays grandes de una vez.
Ver memoria_claude/01_kernel.json -> "escritura_COM_de_muchas_filas_ojo" para el detalle
completo de por que existe.

Uso tipico en un script nuevo:

    import rt_common as rc

    tabla6 = rc.cargar_tabla6(ws_minuta.ListObjects("Tabla6"))
    cadena = rc.inferir_cadena(cod_kpi, tabla6)
    cod_local = rc.normalizar(cod_kpi, cadena, tabla6)

    with rc.abrir_excel_com(FILE_PATH) as (app, wb):
        ws = wb.Worksheets("PLANTILLA")
        rc.escribir_filas_verificado(ws, fila_inicio=2, num_columnas=13, out_rows=filas, log=print)
        wb.Save()
"""
import contextlib

import pythoncom
import win32com.client as win32


# ---------------------------------------------------------------------------
# Normalizacion de codigos y cadenas (Tabla6: CADENA | SUFIJO | SUFIJO2 | OBSERVACION)
# ---------------------------------------------------------------------------

def norm_code(x):
    """Convierte un codigo leido de Excel/COM a texto limpio, sin arrastrar el '.0'
    de un float (ej. 84.0 -> '84', no '84.0'). Llamar SIEMPRE antes de cualquier logica
    de prefijo -- hacerlo despues rompe el join entre hojas (sesion 28)."""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return str(x).strip()


def fmt_num(x):
    """Igual que norm_code pero pensado para mostrar en texto (SKU, stock), no como
    llave de cruce."""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x)



# Cuando una celda de Excel contiene un error (#N/A, #REF!, #VALUE!, etc.) y se lee su
# .Value via COM, NO llega como texto "#N/A" -- llega como un entero negativo especifico
# (el HRESULT del error, ej. -2146826246 para #N/A). Un simple "if valor:" no lo detecta
# porque es un numero truthy como cualquier otro. Encontrado en sesion 34 en la columna
# GESTORES de OSA (un cruce roto en el archivo fuente).
ERRORES_COM_EXCEL = {
    -2146826288,  # #NULL!
    -2146826281,  # #DIV/0!
    -2146826273,  # #VALUE!
    -2146826265,  # #REF!
    -2146826259,  # #NAME?
    -2146826252,  # #NUM!
    -2146826246,  # #N/A
}


def es_error_com(valor):
    """True si `valor` es un error de Excel leido via COM (entero HRESULT) o el texto
    literal de un error (#N/A, #REF!, etc. -- por si llega ya como string, ej. al leer
    con openpyxl en vez de COM)."""
    if isinstance(valor, int) and valor in ERRORES_COM_EXCEL:
        return True
    if isinstance(valor, str) and valor.strip().startswith("#"):
        return True
    return False


def valor_o_default(valor, default):
    """Devuelve `default` si `valor` es None, vacio, o un error de Excel (ver
    es_error_com) -- nunca deja pasar un error de Excel como si fuera un dato real."""
    if valor is None or es_error_com(valor):
        return default
    if isinstance(valor, str) and not valor.strip():
        return default
    return valor


def es_libro(texto):
    """TNOGAL: excluir productos de la linea 'LIBRO' (no es foco de gestion de este
    pipeline). Aplica tanto a MINUTA como a OSA."""
    return bool(texto) and "LIBRO" in str(texto).upper()


def peso_stock(stock, umbral):
    """0 a 1: severidad continua segun magnitud de stock. Stock bajo -> peso bajo
    (probable ruido/desajuste fisico-vs-B2B), stock >= umbral -> peso maximo 1.0."""
    try:
        s = float(stock)
    except (TypeError, ValueError):
        s = 0.0
    if umbral <= 0:
        return 1.0
    return max(0.0, min(1.0, s / umbral))


def cargar_tabla6(list_object):
    """Lee la tabla de prefijos por cadena (CADENA | SUFIJO | SUFIJO2 | OBSERVACION,
    hoja MINUTA) y la deja como dict {CADENA: {"prefijos": [...], "accion": "conservar"|"numerico"}}.
    `list_object` es un ListObject de win32com, ej. ws_minuta.ListObjects("Tabla6")."""
    tabla6 = {}
    for row in list_object.DataBodyRange.Value:
        cadena, suf1, suf2, obs = row
        if not cadena:
            continue
        prefijos = [p for p in (suf1, suf2) if p]
        accion = "conservar" if obs and "conservar" in str(obs).lower() else "numerico"
        tabla6[str(cadena).strip().upper()] = {"prefijos": prefijos, "accion": accion}
    return tabla6


def inferir_cadena(codigo_crudo, tabla6):
    """Deduce la cadena (WALMART/CENCOSUD/TOTTUS/SMU/...) a partir del PREFIJO del
    codigo (ej. 'EX142' -> WALMART), sin depender de un campo RETAIL/CADENA aparte --
    ese campo puede venir con casing inconsistente o simplemente no existir en la hoja
    (REPORTES y OSA no tienen columna de cadena propia)."""
    s = str(codigo_crudo).strip().upper()
    for cadena, info in tabla6.items():
        for p in info["prefijos"]:
            if s.startswith(p.upper()):
                resto = s[len(p):]
                if resto.isdigit() or info["accion"] == "conservar":
                    return cadena
    return None


def normalizar(codigo_crudo, cadena, tabla6):
    """Codigo de local NORMALIZADO para mostrar/cruzar: si la cadena usa 'conservar'
    (ej. CENCOSUD) deja el codigo tal cual; si usa 'numerico' (ej. WALMART/TOTTUS/SMU)
    quita el prefijo y deja solo la parte numerica.

    orden_de_operaciones_critico (sesion 28): norm_code() SIEMPRE primero, antes de
    cualquier logica de prefijo. Si se hace al reves, un codigo numerico limpio tipo
    84.0 se stringifica a '84.0' en vez de '84', rompe el calce con otras hojas y
    duplica la sala en 2 grupos distintos."""
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


# ---------------------------------------------------------------------------
# Automatizacion COM segura (abrir/cerrar Excel real sin arriesgar la sesion del usuario)
# ---------------------------------------------------------------------------

@contextlib.contextmanager
def abrir_excel_com(file_path, visible=False):
    """Context manager que abre Excel de forma segura para automatizacion:
    - DispatchEx (instancia INDEPENDIENTE) en vez de Dispatch simple -- Dispatch()
      puede engancharse a la sesion de Excel que el usuario tiene abierta y arrastrarla
      en un crash si la automatizacion falla (paso de verdad, sesion 6).
    - AutomationSecurity=3 (sin macros), DisplayAlerts=False, ScreenUpdating=False.
    - Cierra SIEMPRE el workbook (sin guardar) y la app al salir, incluso si hay error.

    Uso:
        with abrir_excel_com(FILE_PATH) as (app, wb):
            ws = wb.Worksheets("PLANTILLA")
            ...
            wb.Save()   # guardar es responsabilidad de quien llama, explicito
    """
    pythoncom.CoInitialize()
    app = win32.gencache.EnsureDispatch(win32.DispatchEx('Excel.Application'))
    app.Visible = visible
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    app.ScreenUpdating = False
    wb = None
    try:
        wb = app.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
        yield app, wb
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Escritura segura de muchas filas via COM (fila por fila + verificacion)
# ---------------------------------------------------------------------------

def _valores_iguales(a, b):
    """Compara 2 valores de celda tolerando las conversiones que Excel hace solo:
    - texto numerico '613' se relee como float 613.0 -- sin tolerar esto, la
      verificacion marca falsos positivos en cualquier columna de codigos.
    - una fecha escrita como datetime.datetime de Python se relee como
      pywintypes.datetime CON zona horaria (ej. GMT Standard Time) -- comparados
      como texto nunca calzan, aunque sea literalmente el mismo dia. Bug real
      encontrado en sesion 34: al generalizar esta funcion se perdio la exclusion
      que build_plantilla.py tenia a mano para su columna de fecha, y la primera
      corrida marco las 372 filas como 'corruptas' (falso positivo, no error real)."""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False

    if hasattr(a, "year") and hasattr(b, "year"):
        try:
            return (a.year, a.month, a.day) == (b.year, b.month, b.day)
        except AttributeError:
            pass

    try:
        fa, fb = float(a), float(b)
        return abs(fa - fb) < 0.005
    except (TypeError, ValueError):
        pass

    def _norm(x):
        if isinstance(x, float) and x.is_integer():
            return str(int(x))
        return str(x).strip()

    return _norm(a) == _norm(b)


def escribir_filas_verificado(ws, fila_inicio, num_columnas, out_rows, log=print):
    """Escribe `out_rows` (lista de listas, todas de largo `num_columnas`) a partir de
    `ws.Cells(fila_inicio, 1)`, FILA POR FILA -- no como un solo array 2D grande.

    Por que fila por fila (sesion 32): escribir un array 2D de varios cientos de filas
    en una sola asignacion via COM produjo, en este entorno, filas corruptas al azar
    (texto '#N/A' literal, filas en blanco, o el contenido de OTRA fila del mismo
    cliente duplicado en una posicion vecina) -- confirmado que los datos en Python
    eran correctos fila por fila justo antes de escribir. Ver 01_kernel.json,
    "escritura_COM_de_muchas_filas_ojo".

    Despues de escribir, relee todo el bloque y repara (reescribe individualmente)
    cualquier fila que no calce exactamente contra lo esperado -- capa de seguridad
    adicional, no solo mas lenta que confiar ciegamente.

    Devuelve la cantidad de filas que hubo que reparar (0 en el caso normal)."""
    n = len(out_rows)
    if n == 0:
        return 0

    for i, fila in enumerate(out_rows):
        r = fila_inicio + i
        ws.Range(ws.Cells(r, 1), ws.Cells(r, num_columnas)).Value = fila

    releido = ws.Range(
        ws.Cells(fila_inicio, 1), ws.Cells(fila_inicio + n - 1, num_columnas)
    ).Value
    if not isinstance(releido, tuple) or (releido and not isinstance(releido[0], tuple)):
        releido = (releido,)

    n_reparadas = 0
    for i, (esperado, actual) in enumerate(zip(out_rows, releido)):
        if not all(_valores_iguales(a, e) for a, e in zip(actual, esperado)):
            r = fila_inicio + i
            ws.Range(ws.Cells(r, 1), ws.Cells(r, num_columnas)).Value = esperado
            n_reparadas += 1

    if n_reparadas:
        log(f"escribir_filas_verificado: {n_reparadas} fila(s) llegaron corruptas de la "
            f"escritura y se repararon individualmente.")
    return n_reparadas

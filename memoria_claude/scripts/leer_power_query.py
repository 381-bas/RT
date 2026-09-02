# -*- coding: utf-8 -*-
"""
leer_power_query.py (sesion 38) -- extrae y muestra el codigo M real de TODAS las
Power Query de un .xlsx/.xlsm, leyendo directo del archivo (SOLO LECTURA, nunca escribe
ni modifica nada).

Por que existe: Excel guarda las consultas de Power Query como un blob binario
comprimido dentro de customXml/item1.xml (formato "DataMashup"). openpyxl no lo expone.
Sin esto, para saber que hace exactamente una query hay que confiar en lo que el
usuario describe de memoria o en capturas de pantalla del editor. Con esto, Claude
puede leer el codigo M EXACTO -- los nombres reales de los pasos, los filtros, las
columnas -- y dar instrucciones de clics precisas en vez de genericas.

Formato DataMashup (no documentado oficialmente, pero estable):
    bytes 0-3:  version (uint32 LE)
    bytes 4-7:  longitud N del paquete (uint32 LE)
    bytes 8-(8+N): el paquete mismo, que es un .zip valido con
        Config/Package.xml, [Content_Types].xml, y Formulas/Section1.m
        (el codigo M real, codificado en UTF-16).

Uso:
    python leer_power_query.py archivo.xlsm
    python leer_power_query.py archivo.xlsm --salida resultado.txt

LIMITACION CONOCIDA (sesion 38): funciono perfecto para leer las 4 queries de
02_MINUTAS_CLIENTES_PLANTILLA.xlsm ANTES de que el usuario editara OSA_ALERTA en el
editor de Power Query. Despues de ese guardado de Excel, el mismo archivo empezo a dar
texto ilegible (mojibake tipo CJK) en el .m extraido, incluso probando las 3 variantes
de UTF-16 (LE/BE/generico). La causa exacta no se investigo a fondo -- no bloqueaba
nada (la validacion de datos real se hizo leyendo la hoja OSA_ALERTA con openpyxl/
inspect_excel.py, no con esta herramienta). Si vuelve a fallar, la via mas rapida de
diagnostico es probablemente inspeccionar el zip interno byte a byte en vez de asumir
que el problema es de endianness.
"""
import argparse
import base64
import io
import re
import struct
import zipfile


def _decodificar_utf16_robusto(datos_bytes):
    """Decodifica bytes UTF-16 probando ambas endianness y quedandose con la que
    produce texto plausible (mayoria ASCII imprimible) -- se encontro en sesion 38 que
    el mismo archivo, guardado de nuevo por Excel tras editar la Power Query, puede
    cambiar de BOM/endianness entre una corrida y otra, produciendo mojibake tipo CJK
    si se asume una sola endianness fija."""
    candidatos = []
    for enc in ('utf-16', 'utf-16-le', 'utf-16-be'):
        try:
            candidatos.append(datos_bytes.decode(enc))
        except UnicodeError:
            pass
    if not candidatos:
        return datos_bytes.decode('utf-8', errors='replace')

    def _puntaje_ascii(s):
        if not s:
            return 0
        imprimibles = sum(1 for ch in s if 32 <= ord(ch) < 127 or ch in '\n\r\t')
        return imprimibles / len(s)

    return max(candidatos, key=_puntaje_ascii)


def extraer_m_code(path):
    with zipfile.ZipFile(path) as z:
        nombres = z.namelist()
        if 'customXml/item1.xml' not in nombres:
            return None
        raw = z.read('customXml/item1.xml')

    # el item1.xml puede no ser el DataMashup si hay mas de un customXml -- probar
    # item1, item2, etc. hasta encontrar uno que matchee <DataMashup ...>
    texto = None
    for enc in ('utf-16', 'utf-8'):
        try:
            candidato = raw.decode(enc)
        except UnicodeError:
            continue
        if '<DataMashup' in candidato:
            texto = candidato
            break
    if texto is None:
        return None

    m = re.search(r'<DataMashup[^>]*>([A-Za-z0-9+/=]+)</DataMashup>', texto)
    if not m:
        return None

    blob = base64.b64decode(m.group(1))
    n1 = struct.unpack_from('<I', blob, 4)[0]
    package_zip = blob[8:8 + n1]

    resultado = {}
    with zipfile.ZipFile(io.BytesIO(package_zip)) as iz:
        for nombre in iz.namelist():
            if nombre.lower().endswith('.m'):
                resultado[nombre] = _decodificar_utf16_robusto(iz.read(nombre))
    return resultado


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('archivo', help='Ruta al .xlsx/.xlsm con Power Query a inspeccionar')
    ap.add_argument('--salida', help='Si se indica, escribe el resultado a este archivo '
                     '(UTF-8 explicito) en vez de imprimir por consola -- evita el mojibake '
                     'que produce PowerShell/consola con texto UTF-16 en print().')
    args = ap.parse_args()

    m_files = extraer_m_code(args.archivo)
    if not m_files:
        print(f"No se encontro un bloque DataMashup (Power Query) en '{args.archivo}', "
              f"o el archivo no tiene ninguna consulta.")
        return

    bloques = []
    for nombre, codigo in m_files.items():
        bloques.append(f"=== {nombre} ===\n{codigo}\n")
    texto_final = "\n".join(bloques)

    if args.salida:
        with open(args.salida, "w", encoding="utf-8") as f:
            f.write(texto_final)
        print(f"Escrito en {args.salida} ({len(texto_final)} caracteres).")
    else:
        # sys.stdout con encoding explicito UTF-8 -- evita el mojibake de PowerShell
        import sys
        sys.stdout.buffer.write(texto_final.encode("utf-8", errors="replace"))
        sys.stdout.buffer.write(b"\n")


if __name__ == '__main__':
    main()

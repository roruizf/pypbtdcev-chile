"""PBTD-01 v2.2 ↔ VCEV: from_pbtd01 (lectura con openpyxl) y to_pbtd01 (escritura por XML, ver xlsm_xml)."""

import datetime as dt
import hashlib
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple, get_column_letter
from openpyxl.utils.datetime import to_excel

from .celdas import cargar_celdas
from .vcev import VCEV, Incidencia, Procedencia, asignar

TIPOS_SIMPLES = (int, float, str)


def sha256_archivo(ruta) -> str:
    resumen = hashlib.sha256()
    with open(ruta, "rb") as archivo:
        for bloque in iter(lambda: archivo.read(1 << 20), b""):
            resumen.update(bloque)
    return resumen.hexdigest()


def normalizar(valor):
    """Valor de celda → str | float | int | None (bool → int, fecha/hora → serial Excel, "" → None)."""
    if valor is None or (isinstance(valor, str) and valor == ""):
        return None
    if isinstance(valor, bool):
        return int(valor)
    if isinstance(valor, dt.timedelta):
        return valor.total_seconds() / 86400
    if isinstance(valor, (dt.datetime, dt.date, dt.time)):
        return to_excel(valor)
    if isinstance(valor, TIPOS_SIMPLES):
        return valor
    return str(valor)


def leer_celdas(ruta, data_only: bool, celdas=None) -> dict:
    """{(hoja, celda): (valor, es_formula)} de las celdas del mapeo, con openpyxl en modo solo lectura."""
    celdas = cargar_celdas() if celdas is None else celdas
    por_hoja: dict[str, set[str]] = {}
    for c in celdas:
        por_hoja.setdefault(c["hoja"], set()).add(c["celda"])
    libro = load_workbook(ruta, read_only=True, data_only=data_only, keep_links=False)
    try:
        resultado = {}
        for hoja, refs in por_hoja.items():
            ws = libro[hoja]
            posiciones = [coordinate_to_tuple(ref) for ref in refs]
            fila_min = min(f for f, _ in posiciones)
            fila_max = max(f for f, _ in posiciones)
            col_max = max(c for _, c in posiciones)
            filas = ws.iter_rows(min_row=fila_min, max_row=fila_max, min_col=1, max_col=col_max)
            for i, fila in enumerate(filas):
                for j, celda in enumerate(fila):
                    ref = f"{get_column_letter(j + 1)}{fila_min + i}"
                    if ref in refs:
                        resultado[(hoja, ref)] = (celda.value, getattr(celda, "data_type", None) == "f")
        return resultado
    finally:
        libro.close()


def _texto_formula(valor) -> str:
    texto = str(getattr(valor, "text", valor))
    return texto if texto.startswith("=") else "=" + texto


def from_pbtd01(ruta) -> VCEV:
    """Lee las 1324 celdas editables de un PBTD-01 v2.2: valor en caché + fórmula de usuario en procedencia."""
    ruta = Path(ruta)
    celdas = cargar_celdas()
    sha = sha256_archivo(ruta)
    valores = leer_celdas(ruta, True, celdas)
    formulas = leer_celdas(ruta, False, celdas)
    vcev = VCEV()
    for c in celdas:
        clave = (c["hoja"], c["celda"])
        crudo = valores.get(clave, (None, False))[0]
        texto, es_formula = formulas.get(clave, (None, False))
        formula = _texto_formula(texto) if es_formula else None
        valor = normalizar(crudo)
        if formula is not None and crudo is None:
            vcev.incidencias.append(
                Incidencia(hoja=c["hoja"], celda=c["celda"], campo=c["campo"], valor=None,
                           regla="formula_sin_cache", detalle=formula)
            )
        elif crudo is not None and not isinstance(crudo, str) and type(crudo) not in (int, float):
            vcev.incidencias.append(
                Incidencia(hoja=c["hoja"], celda=c["celda"], campo=c["campo"], valor=valor,
                           regla="tipo_convertido", detalle=type(crudo).__name__)
            )
        asignar(vcev, c["campo"], valor)
        vcev.procedencia[c["campo"]] = Procedencia(
            libro_sha256=sha, hoja=c["hoja"], celda=c["celda"], formula_usuario=formula
        )
    return vcev

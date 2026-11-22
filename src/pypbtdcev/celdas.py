"""Mapeo de las 1324 celdas editables de PBTD-01 v2.2 y plantilla oficial empaquetada."""

import json
from functools import cache
from importlib.resources import files
from pathlib import Path

ARCHIVO_CELDAS = "pbtd01-v2.2-celdas.json"
NOMBRE_PLANTILLA = "01.-PBTD-Datos-de-Arquitectura-v2.2.xlsm"
SHA256_PLANTILLA = "dec31e42ad0d49fe62f9e521067bd4ddce7ab1a5d583aa104865302188b5de9b"
HOJAS = ("CEV-CEVE", "3. Tablas Envolvente")
TOTALES = {"CEV-CEVE": 944, "3. Tablas Envolvente": 380}
PLANTILLA_PBTD01 = Path(str(files("pypbtdcev") / "plantillas" / NOMBRE_PLANTILLA))


@cache
def _texto() -> str:
    return (files("pypbtdcev") / "datos" / ARCHIVO_CELDAS).read_text(encoding="utf-8")


def cargar_mapeo() -> dict:
    """El archivo de mapeo completo (objeto nuevo en cada llamada)."""
    return json.loads(_texto())


def cargar_celdas() -> list[dict]:
    """Las 1324 entradas del mapeo, en el orden del generador (copias: se pueden modificar)."""
    return cargar_mapeo()["celdas"]

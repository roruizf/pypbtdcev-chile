"""Ida y vuelta PBTD-01 → VCEV → PBTD-01: comparación de las 1324 celdas y validación de dominio."""

from .celdas import cargar_celdas
from .pbtd01 import leer_celdas, normalizar
from .vcev import VCEV, Incidencia, leer

OPERADORES = {
    "between": lambda v, a, b: a <= v <= b,
    "notBetween": lambda v, a, b: not (a <= v <= b),
    "greaterThan": lambda v, a, b: v > a,
    "greaterThanOrEqual": lambda v, a, b: v >= a,
    "lessThan": lambda v, a, b: v < a,
    "lessThanOrEqual": lambda v, a, b: v <= a,
    "equal": lambda v, a, b: v == a,
    "notEqual": lambda v, a, b: v != a,
}
DOS_LIMITES = {"between", "notBetween"}


def _es_numero(valor) -> bool:
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def iguales(a, b) -> bool:
    """int y float iguales si numéricamente iguales; None == ""; textos exactos; número ≠ texto."""
    a, b = normalizar(a), normalizar(b)
    if a is None or b is None:
        return a is None and b is None
    if _es_numero(a) and _es_numero(b):
        return a == b
    if _es_numero(a) or _es_numero(b):
        return False
    return a == b


def comparar(origen, salida) -> list[dict]:
    """Diferencias en las 1324 celdas editables entre dos libros PBTD-01 (valores en caché)."""
    celdas = cargar_celdas()
    a = leer_celdas(origen, True, celdas)
    b = leer_celdas(salida, True, celdas)
    diferencias = []
    for c in celdas:
        clave = (c["hoja"], c["celda"])
        valor_a = a.get(clave, (None, False))[0]
        valor_b = b.get(clave, (None, False))[0]
        if not iguales(valor_a, valor_b):
            diferencias.append({"hoja": c["hoja"], "celda": c["celda"], "campo": c["campo"],
                                "origen": normalizar(valor_a), "salida": normalizar(valor_b)})
    return diferencias


def _regla(valor, dominio: dict):
    tipo = dominio.get("tipo")
    if tipo == "lista":
        return None if any(iguales(valor, x) for x in dominio["valores"]) else "dominio_lista"
    if tipo in ("decimal", "entero"):
        if not _es_numero(valor):
            return "dominio_tipo"
        if tipo == "entero" and not float(valor).is_integer():
            return "dominio_entero"
        operador = dominio.get("operador") or "between"
        minimo, maximo = dominio.get("min"), dominio.get("max")
        comparador = OPERADORES.get(operador)
        if comparador is None or minimo is None or (operador in DOS_LIMITES and maximo is None):
            return None
        return None if comparador(valor, minimo, maximo) else "dominio_rango"
    return None  # lista_dinamica, texto_libre, otro: no se valida


def validar(vcev: VCEV) -> list[Incidencia]:
    """Incidencias de dominio (informativas: nunca bloquean la ida y vuelta)."""
    incidencias = []
    for c in cargar_celdas():
        valor = leer(vcev, c["campo"])
        if valor is None:
            continue
        regla = _regla(valor, c["dominio"])
        if regla:
            incidencias.append(Incidencia(hoja=c["hoja"], celda=c["celda"], campo=c["campo"], valor=valor, regla=regla))
    return incidencias

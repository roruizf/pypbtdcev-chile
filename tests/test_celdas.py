"""Pruebas del mapeo de las 1324 celdas editables de PBTD-01 v2.2 y de su generador."""

import importlib.util
from collections import Counter
from functools import cache
from importlib.resources import files
from pathlib import Path

import openpyxl
import pytest

from pypbtdcev.celdas import ARCHIVO_CELDAS, PLANTILLA_PBTD01, cargar_celdas, cargar_mapeo

CEVE = "CEV-CEVE"
TABLAS = "3. Tablas Envolvente"
GENERADOR = Path(__file__).resolve().parents[1] / "herramientas" / "generar_celdas.py"

CLAVES = [
    "hoja",
    "celda",
    "familia",
    "campo",
    "significado",
    "seccion_c1",
    "dominio",
    "valor_plantilla",
    "en_xml_plantilla",
]
TIPOS = {"lista", "lista_dinamica", "decimal", "entero", "texto_libre", "otro"}

CONTEOS = {
    (CEVE, "Datos generales (validados)"): 11,
    (CEVE, "Datos generales (texto libre)"): 5,
    (CEVE, "Certificado (texto/descripción)"): 18,
    (CEVE, "Dimensiones"): 6,
    (CEVE, "Muros"): 96,
    (CEVE, "Puentes térmicos particulares"): 25,
    (CEVE, "Puertas"): 45,
    (CEVE, "Ventanas"): 408,
    (CEVE, "FAR (azimut)"): 8,
    (CEVE, "FAR (División)"): 68,
    (CEVE, "FAR (A/B/D)"): 192,
    (CEVE, "Techos"): 25,
    (CEVE, "Pisos"): 24,
    (CEVE, "Infiltración/ventilación"): 8,
    (CEVE, "Perfil de uso"): 5,
    (TABLAS, "Puertas"): 84,
    (TABLAS, "Vidrios"): 59,
    (TABLAS, "Marcos"): 48,
    (TABLAS, "Muros"): 105,
    (TABLAS, "Techos"): 24,
    (TABLAS, "Pisos"): 60,
}


@cache
def _generador():
    spec = importlib.util.spec_from_file_location("generar_celdas", GENERADOR)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def _indice() -> dict:
    return {(c["hoja"], c["celda"]): c for c in cargar_celdas()}


def test_totales():
    celdas = cargar_celdas()
    por_hoja = Counter(c["hoja"] for c in celdas)
    assert por_hoja == {CEVE: 944, TABLAS: 380}
    assert len(celdas) == 1324
    assert cargar_mapeo()["totales"] == {CEVE: 944, TABLAS: 380, "total": 1324}


def test_celdas_y_campos_unicos():
    celdas = cargar_celdas()
    assert len({(c["hoja"], c["celda"]) for c in celdas}) == 1324
    assert len({c["campo"] for c in celdas}) == 1324


def test_conteo_por_familia():
    conteo = Counter((c["hoja"], c["familia"]) for c in cargar_celdas())
    assert len(CONTEOS) == 21
    assert dict(conteo) == CONTEOS


def test_claves_de_cada_entrada():
    for entrada in cargar_celdas():
        assert list(entrada) == CLAVES


def test_dominios_permitidos():
    for entrada in cargar_celdas():
        dominio = entrada["dominio"]
        assert dominio["tipo"] in TIPOS, entrada
        if dominio["tipo"] == "lista":
            assert dominio["valores"], entrada


def test_puntos_de_control_de_dominio():
    indice = _indice()
    h103 = indice[(CEVE, "H103")]["dominio"]
    assert h103["tipo"] == "lista"
    assert {"Corredera", "Fija"} <= set(h103["valores"])
    l158 = indice[(CEVE, "L158")]["dominio"]
    assert l158["tipo"] == "lista"
    assert {"Ventilado", "No Ventilado"} <= set(l158["valores"])
    h62 = indice[(TABLAS, "H62")]["dominio"]
    assert h62["tipo"] == "lista"
    assert {"Sin", "Interior", "Exterior"} <= set(h62["valores"])
    assert indice[(CEVE, "D66")]["dominio"]["tipo"] == "lista_dinamica"
    assert indice[(CEVE, "E13")]["dominio"]["tipo"] == "texto_libre"


@pytest.mark.filterwarnings("ignore::UserWarning")
def test_plantilla_sin_formulas_en_celdas_mapeadas():
    libro = openpyxl.load_workbook(PLANTILLA_PBTD01, data_only=False)
    con_formula = [f"{c['hoja']}!{c['celda']}" for c in cargar_celdas() if libro[c["hoja"]][c["celda"]].data_type == "f"]
    assert con_formula == []


@pytest.mark.filterwarnings("ignore::UserWarning")
def test_generacion_reproducible(tmp_path):
    salida = tmp_path / "x.json"
    total, _ = _generador().generar(salida)
    assert total == 1324
    comiteado = (files("pypbtdcev") / "datos" / ARCHIVO_CELDAS).read_bytes()
    assert salida.read_bytes() == comiteado


def test_comprobar_rechaza_formula():
    generador = _generador()
    libro = openpyxl.Workbook()
    libro.active.title = CEVE
    libro.create_sheet(TABLAS)
    libro[CEVE]["E7"] = "=1+1"
    with pytest.raises(generador.GeneracionError, match="fórmula"):
        generador.comprobar(generador.tabla(), libro)


def test_comprobar_rechaza_totales():
    generador = _generador()
    libro = openpyxl.Workbook()
    libro.active.title = CEVE
    libro.create_sheet(TABLAS)
    with pytest.raises(generador.GeneracionError, match="totales"):
        generador.comprobar(generador.tabla()[:-1], libro)

"""Pruebas de los modelos VCEV: cobertura contra el mapeo, acceso por camino y JSON."""

import pydantic
import pytest

from pypbtdcev.celdas import cargar_celdas
from pypbtdcev.vcev import VCEV, CampoDesconocido, asignar, leer, rutas_hoja


def test_rutas_coinciden_con_el_mapeo():
    rutas = rutas_hoja()
    assert len(rutas) == len(set(rutas)) == 1324
    assert set(rutas) == {c["campo"] for c in cargar_celdas()}


def test_asignar_y_leer():
    vcev = VCEV()
    asignar(vcev, "ceve.ventanas[20].tipo_vidrio", "V1")
    assert leer(vcev, "ceve.ventanas[20].tipo_vidrio") == "V1"


@pytest.mark.parametrize("campo", ["ceve.ventanas[20].alto", "ceve.inexistente"])
def test_campo_desconocido(campo):
    vcev = VCEV()
    with pytest.raises(CampoDesconocido):
        leer(vcev, campo)
    with pytest.raises(CampoDesconocido):
        asignar(vcev, campo, 1.0)


def test_ida_y_vuelta_json_conserva_tipos():
    valores = {
        "ceve.datos_generales.dormitorios": 3,
        "ceve.dimensiones[0].area": 3.0,
        "ceve.datos_generales.identificacion_vivienda": "3",
        "tablas.vidrios[15].fs": None,
    }
    vcev = VCEV()
    for campo, valor in valores.items():
        asignar(vcev, campo, valor)
    vuelta = VCEV.model_validate_json(vcev.model_dump_json())
    assert vuelta == vcev
    for campo, valor in valores.items():
        leido = leer(vuelta, campo)
        assert leido == valor
        assert type(leido) is type(valor)


def test_lista_en_campo_de_valor_falla():
    vcev = VCEV()
    with pytest.raises(pydantic.ValidationError):
        asignar(vcev, "ceve.muros[0].area", [1, 2])

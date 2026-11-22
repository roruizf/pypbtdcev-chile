"""Ida y vuelta PBTD-01 → VCEV → PBTD-01 del libro de ejemplo, comparación y validación de dominio."""

from pathlib import Path

import openpyxl
import pytest

import pypbtdcev
from pypbtdcev import asignar, comparar, from_pbtd01, iguales, to_pbtd01, validar

EJEMPLO = Path(__file__).resolve().parents[1] / "ejemplos" / "01.-Ejamplo-1.xlsm"


@pytest.fixture(scope="module")
def salida_ejemplo(tmp_path_factory):
    salida = tmp_path_factory.mktemp("ida_vuelta") / "salida.xlsm"
    to_pbtd01(from_pbtd01(EJEMPLO), salida)
    return salida


def test_ida_y_vuelta_sin_diferencias(salida_ejemplo):
    assert comparar(EJEMPLO, salida_ejemplo) == []


def test_un_cambio_una_diferencia(tmp_path):
    vcev = from_pbtd01(EJEMPLO)
    asignar(vcev, "ceve.datos_generales.identificacion_vivienda", "Vivienda cambiada por la prueba")
    salida = tmp_path / "cambio.xlsm"
    to_pbtd01(vcev, salida)
    diferencias = comparar(EJEMPLO, salida)
    assert len(diferencias) == 1
    assert (diferencias[0]["hoja"], diferencias[0]["celda"]) == ("CEV-CEVE", "E13")


def test_iguales():
    assert iguales(1, 1.0) is True
    assert iguales(None, "") is True
    assert iguales("1", 1) is False


def test_validar_no_bloquea(tmp_path):
    vcev = from_pbtd01(EJEMPLO)
    asignar(vcev, "ceve.ventanas[0].tipo_cierre", "Inventado")
    incidencias = [i for i in validar(vcev) if i.celda == "H103"]
    assert len(incidencias) == 1
    assert incidencias[0].regla == "dominio_lista"
    assert incidencias[0].hoja == "CEV-CEVE"
    salida = tmp_path / "inventado.xlsm"
    to_pbtd01(vcev, salida)
    assert salida.exists()
    diferencias = comparar(EJEMPLO, salida)
    assert [(d["celda"], d["salida"]) for d in diferencias] == [("H103", "Inventado")]


@pytest.mark.filterwarnings("ignore::UserWarning")
def test_formula_de_usuario_no_se_restaura(salida_ejemplo):
    celda = openpyxl.load_workbook(salida_ejemplo, data_only=False)["CEV-CEVE"]["D52"]
    assert celda.data_type != "f"
    assert isinstance(celda.value, float)
    assert celda.value == pytest.approx(56.88, abs=1e-9)


def test_exports():
    assert pypbtdcev.__version__ == "0.2.0"
    for nombre in ("VCEV", "from_pbtd01", "to_pbtd01", "comparar", "validar", "cargar_celdas"):
        assert hasattr(pypbtdcev, nombre), nombre
        assert nombre in pypbtdcev.__all__

"""Pruebas de from_pbtd01: valores en caché, fórmulas de usuario, procedencia y normalización."""

import datetime as dt
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils.datetime import to_excel

from pypbtdcev.celdas import cargar_celdas
from pypbtdcev.pbtd01 import from_pbtd01, normalizar
from pypbtdcev.vcev import leer

EJEMPLO = Path(__file__).resolve().parents[1] / "ejemplos" / "01.-Ejamplo-1.xlsm"
SHA_EJEMPLO = "54670b8c09dd307e7a52032c3b013843f31da9c5648d61582a9b5fc380830e92"


@pytest.fixture(scope="module")
def vcev_ejemplo():
    return from_pbtd01(EJEMPLO)


def test_valores_y_formulas_de_usuario(vcev_ejemplo):
    v = vcev_ejemplo
    proc = v.procedencia

    assert leer(v, "ceve.dimensiones[0].area") == pytest.approx(56.88, abs=1e-9)
    assert proc["ceve.dimensiones[0].area"].formula_usuario == "=4.5*12.64"

    assert repr(leer(v, "ceve.dimensiones[0].altura")) == "3.6500000000000004"
    assert proc["ceve.dimensiones[0].altura"].formula_usuario == "=2.9+(4.4-2.9)/2"

    assert leer(v, "ceve.pisos[0].area") == pytest.approx(56.88, abs=1e-9)
    assert proc["ceve.pisos[0].area"].formula_usuario == "=4.5*12.64"

    assert leer(v, "tablas.muros[1].u") == 0.88
    assert proc["tablas.muros[1].u"].formula_usuario == "=ROUND(1/(0.17+(0.1/0.104)),2)"


def test_procedencia_completa(vcev_ejemplo):
    celdas = cargar_celdas()
    assert len(vcev_ejemplo.procedencia) == 1324
    for c in celdas:
        p = vcev_ejemplo.procedencia[c["campo"]]
        assert p.fuente == "pbtd01"
        assert p.libro_sha256 == SHA_EJEMPLO
        assert (p.hoja, p.celda) == (c["hoja"], c["celda"])


@pytest.mark.filterwarnings("ignore::UserWarning")
def test_formula_sin_cache(tmp_path):
    libro = openpyxl.load_workbook(EJEMPLO, keep_vba=True)
    libro["CEV-CEVE"]["D52"] = "=1+1"
    sintetico = tmp_path / "sin-cache.xlsm"
    libro.save(sintetico)  # libro sintético de prueba: openpyxl no escribe valores en caché

    vcev = from_pbtd01(sintetico)
    assert leer(vcev, "ceve.dimensiones[0].area") is None
    incidencias = [i for i in vcev.incidencias if i.celda == "D52"]
    assert len(incidencias) == 1
    assert incidencias[0].regla == "formula_sin_cache"
    assert incidencias[0].hoja == "CEV-CEVE"


def test_normalizar():
    assert normalizar(True) == 1
    assert type(normalizar(True)) is int
    assert normalizar("") is None
    assert repr(normalizar(3.6500000000000004)) == "3.6500000000000004"
    fecha = dt.datetime(2026, 1, 1)
    assert normalizar(fecha) == to_excel(fecha)

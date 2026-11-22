"""Pruebas de la escritura PBTD-01 por edición directa del XML (xlsm_xml)."""

import hashlib
import re
from pathlib import Path
from zipfile import ZipFile

import openpyxl
import pytest

from pypbtdcev.celdas import PLANTILLA_PBTD01
from pypbtdcev.xlsm_xml import escribir_libro, rutas_hojas, verificar_preservacion

EJEMPLO = Path(__file__).resolve().parents[1] / "ejemplos" / "01.-Ejamplo-1.xlsm"
CAMBIOS = {"CEV-CEVE": {"E13": "Depto 101", "D52": 0.1 + 0.2, "F230": None}}


@pytest.fixture(scope="module")
def escrito(tmp_path_factory):
    destino = tmp_path_factory.mktemp("xml") / "salida.xlsm"
    informe = escribir_libro(PLANTILLA_PBTD01, destino, CAMBIOS)
    return destino, informe


@pytest.mark.filterwarnings("ignore::UserWarning")
def test_valores_escritos(escrito):
    destino, _ = escrito
    ws = openpyxl.load_workbook(destino, data_only=False)["CEV-CEVE"]
    assert ws["E13"].value == "Depto 101"
    assert ws["D52"].value == 0.30000000000000004
    assert repr(ws["D52"].value) == "0.30000000000000004"
    assert ws["F230"].value is None
    with ZipFile(destino) as z:
        xml = z.read(rutas_hojas(z)["CEV-CEVE"]).decode("utf-8")
    assert "<v>0.30000000000000004</v>" in xml


@pytest.mark.parametrize(
    ("plantilla", "cambios"),
    [
        (PLANTILLA_PBTD01, {"CEV-CEVE": {"F52": 1.0}}),
        (PLANTILLA_PBTD01, {"CEV-CEVE": {"ZZ9999": 1.0}}),
        (EJEMPLO, {"CEV-CEVE": {"E13": "x"}}),
    ],
    ids=["formula", "ausente", "otra_plantilla"],
)
def test_rechazos_sin_destino(tmp_path, plantilla, cambios):
    destino = tmp_path / "rechazo.xlsm"
    with pytest.raises(ValueError):
        escribir_libro(plantilla, destino, cambios)
    assert not destino.exists()


def test_informe_y_estructura(escrito):
    destino, informe = escrito
    with ZipFile(PLANTILLA_PBTD01) as a:
        miembros_plantilla = a.namelist()
        # La plantilla v2.2 empaquetada no trae xl/vbaProject.bin: el SHA-256 esperado es None en ese caso.
        vba_plantilla = (
            hashlib.sha256(a.read("xl/vbaProject.bin")).hexdigest() if "xl/vbaProject.bin" in miembros_plantilla else None
        )
    assert informe["formulas_preservadas"] > 0
    assert informe["vba_sha256"] == vba_plantilla
    with ZipFile(destino) as b:
        assert b.namelist() == miembros_plantilla
        workbook = b.read("xl/workbook.xml").decode("utf-8")
    assert re.search(r'<calcPr\b[^>]*\bfullCalcOnLoad="1"', workbook)


def test_verificar_preservacion_detecta_estilos(tmp_path):
    copia = tmp_path / "estilos.xlsm"
    with ZipFile(PLANTILLA_PBTD01) as fuente, ZipFile(copia, "w") as salida:
        for item in fuente.infolist():
            datos = fuente.read(item.filename)
            if item.filename == "xl/styles.xml":
                datos += b" "
            salida.writestr(item, datos)
    with pytest.raises(AssertionError):
        verificar_preservacion(PLANTILLA_PBTD01, copia)

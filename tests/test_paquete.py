"""Pruebas básicas del paquete pypbtdcev 0.2.0."""

import hashlib
import importlib.metadata
import importlib.resources
from pathlib import Path

import pypbtdcev

SHA_PLANTILLA = "dec31e42ad0d49fe62f9e521067bd4ddce7ab1a5d583aa104865302188b5de9b"
SHA_EJEMPLO = "54670b8c09dd307e7a52032c3b013843f31da9c5648d61582a9b5fc380830e92"


def _sha256(datos: bytes) -> str:
    return hashlib.sha256(datos).hexdigest()


def test_version():
    assert pypbtdcev.__version__ == "0.2.0"
    assert importlib.metadata.version("pypbtdcev") == "0.2.0"


def test_clases_heredadas():
    from pypbtdcev.escritor import EscritorPBTD01_v2
    from pypbtdcev.lector import LectorPBTD01_v2, LectorPBTD03_v2

    assert issubclass(LectorPBTD03_v2, LectorPBTD01_v2)
    assert EscritorPBTD01_v2 is not None


def test_plantilla_empaquetada():
    plantilla = importlib.resources.files("pypbtdcev") / "plantillas" / "01.-PBTD-Datos-de-Arquitectura-v2.2.xlsm"
    assert _sha256(plantilla.read_bytes()) == SHA_PLANTILLA


def test_libro_ejemplo():
    ejemplo = Path(__file__).resolve().parents[1] / "ejemplos" / "01.-Ejamplo-1.xlsm"
    assert _sha256(ejemplo.read_bytes()) == SHA_EJEMPLO

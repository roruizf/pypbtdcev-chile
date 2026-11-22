"""La rueda de pypbtdcev empaqueta el mapeo, la plantilla y los módulos, y nada de tests/ ni herramientas/."""

import subprocess
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
MIEMBROS = [
    "pypbtdcev/datos/pbtd01-v2.2-celdas.json",
    "pypbtdcev/plantillas/01.-PBTD-Datos-de-Arquitectura-v2.2.xlsm",
    "pypbtdcev/vcev.py",
    "pypbtdcev/lector.py",
    "pypbtdcev/escritor.py",
]


def test_contenido_de_la_rueda(tmp_path):
    subprocess.run(["uv", "build", "--wheel", "--out-dir", str(tmp_path)], cwd=RAIZ, check=True)
    ruedas = list(tmp_path.glob("*.whl"))
    assert len(ruedas) == 1
    with zipfile.ZipFile(ruedas[0]) as rueda:
        nombres = rueda.namelist()
    for miembro in MIEMBROS:
        assert miembro in nombres, miembro
    assert not [n for n in nombres if n.startswith(("tests/", "herramientas/"))]

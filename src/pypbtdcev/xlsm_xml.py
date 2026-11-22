"""Escritura PBTD-01 por edición directa del XML del .xlsm (port de cevstudio pbtd01_writer, E2).

Copia todos los miembros ZIP de la plantilla oficial y parcha solo las hojas pedidas, ubicadas vía xl/workbook.xml
y xl/_rels/workbook.xml.rels. Conserva el atributo de estilo de cada celda; números como <v> con repr() (ida y
vuelta exacta; cevstudio usa '%.12g'), textos como inlineStr, None vacía la celda. Rechaza escribir sobre una
fórmula, sobre una celda ausente del XML y con una plantilla de SHA-256 distinto del fijado. Borra los cachés de
fórmulas, fija fullCalcOnLoad="1" y verifica la preservación. openpyxl no se usa para escribir.
"""

import hashlib
import math
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

from .celdas import SHA256_PLANTILLA

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
CELDA = re.compile(r'<c\b[^>]*\br="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)
FORMULA = re.compile(r"<f(?:\s|>|/)")
FORMULA_COMPLETA = re.compile(rb"<f\b[^>]*?(?:/>|>.*?</f>)", re.DOTALL)
CALC_PR = re.compile(r"<calcPr\b[^>]*/>")


def sha256(ruta) -> str:
    return hashlib.sha256(Path(ruta).read_bytes()).hexdigest()


def rutas_hojas(archivo: ZipFile) -> dict[str, str]:
    """{nombre de hoja: miembro ZIP de su XML}."""
    rels = ET.fromstring(archivo.read("xl/_rels/workbook.xml.rels"))
    destinos = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    resultado = {}
    for hoja in ET.fromstring(archivo.read("xl/workbook.xml")).find("s:sheets", NS):
        destino = destinos[hoja.attrib[REL_ID]]
        resultado[hoja.attrib["name"]] = destino.lstrip("/") if destino.startswith("/") else "xl/" + destino
    return resultado


def parchar_hoja(xml: str, cambios: dict) -> str:
    vistas = set()

    def reemplazar(coincidencia):
        ref = coincidencia.group(1)
        if ref not in cambios:
            return coincidencia.group(0)
        original = coincidencia.group(0)
        if FORMULA.search(original):
            raise ValueError(f"se rechaza escribir sobre una fórmula: {ref}")
        vistas.add(ref)
        inicio = original[: original.index(">") + 1]
        if inicio.endswith("/>"):
            inicio = inicio[:-2].rstrip() + ">"
        inicio = re.sub(r'\st="[^"]*"', "", inicio)
        valor = cambios[ref]
        if valor is None:
            return inicio + "</c>"
        if isinstance(valor, bool):
            valor = int(valor)
        if isinstance(valor, (int, float)):
            if isinstance(valor, float) and not math.isfinite(valor):
                raise ValueError(f"valor no finito en {ref}")
            texto = repr(valor) if isinstance(valor, float) else str(valor)
            return f"{inicio}<v>{texto}</v></c>"
        texto = escape(str(valor)).replace("\r", "&#13;")
        return inicio[:-1] + ' t="inlineStr"><is><t xml:space="preserve">' + texto + "</t></is></c>"

    nuevo = CELDA.sub(reemplazar, xml)
    faltan = sorted(set(cambios) - vistas)
    if faltan:
        raise ValueError(f"celdas del mapeo ausentes del XML de la plantilla: {faltan}")
    return nuevo


def borrar_caches(xml: str) -> str:
    def limpiar(coincidencia):
        celda = coincidencia.group(0)
        if not FORMULA.search(celda):
            return celda
        return re.sub(r"<v(?:\s[^>]*)?>.*?</v>|<v\s*/>", "", celda, flags=re.DOTALL)

    return CELDA.sub(limpiar, xml)


def fijar_calculo(xml: str) -> str:
    """calcPr con fullCalcOnLoad="1" (Excel recalcula todo al abrir)."""

    def tocar(coincidencia):
        etiqueta = re.sub(r'\s+fullCalcOnLoad="[^"]*"', "", coincidencia.group(0))
        return etiqueta[:-2].rstrip() + ' fullCalcOnLoad="1"/>'

    nuevo, cuantos = CALC_PR.subn(tocar, xml)
    if cuantos:
        return nuevo
    if "<calcPr" in xml:
        raise ValueError("calcPr no autocontenido en xl/workbook.xml")
    ancla = "</definedNames>" if "</definedNames>" in xml else "</sheets>"
    return xml.replace(ancla, ancla + '<calcPr fullCalcOnLoad="1"/>', 1)


def verificar_preservacion(plantilla, destino) -> dict:
    with ZipFile(plantilla) as a, ZipFile(destino) as b:
        if a.namelist() != b.namelist():
            raise AssertionError("cambiaron los miembros del ZIP")
        formulas = 0
        for nombre in a.namelist():
            parte_a, parte_b = a.read(nombre), b.read(nombre)
            if nombre.startswith("xl/worksheets/") and nombre.endswith(".xml"):
                if FORMULA_COMPLETA.findall(parte_a) != FORMULA_COMPLETA.findall(parte_b):
                    raise AssertionError(f"cambió una fórmula: {nombre}")
                formulas += len(FORMULA_COMPLETA.findall(parte_a))
                ET.fromstring(parte_b)
            elif nombre != "xl/workbook.xml" and parte_a != parte_b:
                raise AssertionError(f"cambió una parte que no es hoja: {nombre}")
        ET.fromstring(b.read("xl/workbook.xml"))
        vba = hashlib.sha256(b.read("xl/vbaProject.bin")).hexdigest() if "xl/vbaProject.bin" in b.namelist() else None
    return {
        "formulas_preservadas": formulas,
        "vba_sha256": vba,
        "partes_no_hoja_identicas": True,
        "caches_de_formulas": "borrados; Excel recalcula al abrir (fullCalcOnLoad=1)",
    }


def escribir_libro(plantilla, destino, cambios: dict[str, dict]) -> dict:
    """Copia la plantilla en `destino` aplicando {hoja: {celda: valor}}; devuelve el informe de preservación."""
    plantilla, destino = Path(plantilla), Path(destino)
    if sha256(plantilla) != SHA256_PLANTILLA:
        raise ValueError(f"SHA-256 de plantilla desconocido: {plantilla} (se exige {SHA256_PLANTILLA})")
    if destino.resolve() == plantilla.resolve():
        raise ValueError("el destino no puede ser la plantilla")
    destino.parent.mkdir(parents=True, exist_ok=True)
    try:
        with ZipFile(plantilla) as fuente:
            rutas = rutas_hojas(fuente)
            desconocidas = sorted(set(cambios) - set(rutas))
            if desconocidas:
                raise ValueError(f"hojas inexistentes en la plantilla: {desconocidas}")
            parches = {rutas[hoja]: valores for hoja, valores in cambios.items() if valores}
            with ZipFile(destino, "w") as salida:
                for item in fuente.infolist():
                    datos = fuente.read(item.filename)
                    if item.filename in parches:
                        datos = parchar_hoja(datos.decode("utf-8"), parches[item.filename]).encode("utf-8")
                    if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                        datos = borrar_caches(datos.decode("utf-8")).encode("utf-8")
                    if item.filename == "xl/workbook.xml":
                        datos = fijar_calculo(datos.decode("utf-8")).encode("utf-8")
                    salida.writestr(item, datos)
        informe = verificar_preservacion(plantilla, destino)
    except Exception:
        destino.unlink(missing_ok=True)
        raise
    informe["celdas_escritas"] = sum(len(valores) for valores in cambios.values())
    return informe

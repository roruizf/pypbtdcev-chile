"""Genera src/pypbtdcev/datos/pbtd01-v2.2-celdas.json: las 1324 celdas editables de PBTD-01 v2.2.

Fuente de la lista: tabla §11 "Exhaustividad workbook" del diccionario C1 de cevstudio
(docs/model-methods/pbtd01-celda-a-celda.md), codificada literalmente en tabla(). Lee la plantilla empaquetada
(SHA-256 fijo) para: rechazar celdas listadas con fórmula, resolver el dominio de cada celda desde sus validaciones
de datos (rangos con nombre de '0. Tablas referencia'), guardar el valor que trae la plantilla y si la celda existe
como elemento <c> en el XML de su hoja.

Uso:     uv run python herramientas/generar_celdas.py [--salida RUTA]
Salida:  CELDAS OK 944 + 380 = 1324 (celdas sin elemento <c> en la plantilla: N) · exit 0
Errores: plantilla desconocida, celda listada con fórmula o totales distintos → "error: …" en stderr, exit 1.
Re-ejecutable: dos corridas producen el mismo archivo byte a byte (no escribe fechas).
"""

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook

from pypbtdcev.celdas import ARCHIVO_CELDAS, HOJAS, NOMBRE_PLANTILLA, SHA256_PLANTILLA, TOTALES

RAIZ = Path(__file__).resolve().parents[1]
PLANTILLA = RAIZ / "src" / "pypbtdcev" / "plantillas" / NOMBRE_PLANTILLA
DESTINO = RAIZ / "src" / "pypbtdcev" / "datos" / ARCHIVO_CELDAS
CEVE, TABLAS = HOJAS
REFERENCIA = "0. Tablas referencia"
CELDA_XML = re.compile(r'<c\b[^>]*\br="([A-Z]+[0-9]+)"')
NOMBRE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")
RANGO = re.compile(r"^(?:'?(?P<hoja>[^'!]+)'?!)?(?P<ref>\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?)$")
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

FAV = [
    ("P", "fav1_d", "FAV1: distancia vertical alero–vano D [m]"),
    ("Q", "fav1_l", "FAV1: proyección del alero L [m]"),
    ("R", "fav2_izq_profundidad", "FAV2 izquierda: profundidad [m]"),
    ("S", "fav2_izq_separacion", "FAV2 izquierda: separación [m]"),
    ("T", "fav2_der_profundidad", "FAV2 derecha: profundidad [m]"),
    ("U", "fav2_der_separacion", "FAV2 derecha: separación [m]"),
    ("V", "fav3_espesor", "FAV3: espesor de lamas [m]"),
    ("W", "fav3_separacion", "FAV3: separación de lamas [m]"),
    ("X", "fav3_beta", "FAV3: ángulo β de lamas"),
    ("Y", "fav3_alfa", "FAV3: ángulo α de lamas"),
]
FAR_BLOQUES = [
    # (clave, celda azimut, primera fila, última fila División, última fila A/B/D, col. División, A, B, D, nombre)
    ("n", "G125", 127, 135, 134, "E", "F", "G", "H", "Norte"),
    ("e", "L125", 127, 135, 134, "J", "K", "L", "M", "Este"),
    ("s", "Q125", 127, 135, 134, "O", "P", "Q", "R", "Sur"),
    ("o", "V125", 127, 135, 134, "T", "U", "V", "W", "Oeste"),
    ("ne", "G137", 139, 146, 146, "E", "F", "G", "H", "Noreste"),
    ("se", "L137", 139, 146, 146, "J", "K", "L", "M", "Sureste"),
    ("so", "Q137", 139, 146, 146, "O", "P", "Q", "R", "Suroeste"),
    ("no", "V137", 139, 146, 146, "T", "U", "V", "W", "Noroeste"),
]


class GeneracionError(Exception):
    """Plantilla desconocida, celda listada con fórmula o totales distintos de 944/380/1324."""


def _sueltas(hoja, familia, seccion, prefijo, celdas):
    return [(hoja, celda, familia, f"{prefijo}.{campo}", significado, seccion) for celda, campo, significado in celdas]


def _filas(hoja, familia, seccion, prefijo, primera, filas, columnas):
    return [
        (hoja, f"{col}{fila}", familia, f"{prefijo}[{fila - primera}].{campo}", f"{significado} (fila {fila})", seccion)
        for fila in filas
        for col, campo, significado in columnas
    ]


def tabla() -> list[tuple[str, str, str, str, str, str]]:
    """(hoja, celda, familia, campo, significado, seccion_c1) de las 1324 celdas, en orden de la tabla §11."""
    e = []
    e += _sueltas(CEVE, "Datos generales (validados)", "C1 §1", "ceve.datos_generales", [
        ("E7", "tipo_calificacion", "Tipo de calificación (Pre calificación / Calificación)"),
        ("G7", "vivienda_nueva_antigua", "Vivienda nueva o antigua"),
        ("E8", "region", "Región"),
        ("F8", "region_f8", "Región: celda F8 del rango validado E8:G8"),
        ("G8", "region_g8", "Región: celda G8 del rango validado E8:G8"),
        ("E9", "comuna", "Comuna (lista dependiente de la región)"),
        ("E10", "zona_termica", "Zona térmica del proyecto (NCh1079, A-I)"),
        ("E11", "dormitorios", "Dormitorios de la vivienda"),
        ("E16", "tipo_vivienda", "Tipo de vivienda (casa aislada, pareada o continua, departamento)"),
        ("F16", "tipo_vivienda_f16", "Tipo de vivienda: celda F16 del rango validado E16:G16"),
        ("G16", "tipo_vivienda_g16", "Tipo de vivienda: celda G16 del rango validado E16:G16"),
    ])
    e += _sueltas(CEVE, "Datos generales (texto libre)", "C1 §1", "ceve.datos_generales", [
        ("E13", "identificacion_vivienda", "Identificación de la vivienda a evaluar"),
        ("E14", "nombre_proyecto", "Nombre del proyecto"),
        ("E15", "direccion", "Dirección de la vivienda"),
        ("E25", "caso_interno_evaluador", "Caso interno del evaluador"),
        ("E26", "iteracion_evaluador", "Iteración del evaluador"),
    ])
    e += _sueltas(CEVE, "Certificado (texto/descripción)", "C1 §1", "ceve.certificado", [
        ("E19", "rol_vivienda", "Rol de la vivienda (solo certificado)"),
        ("E20", "evaluador", "Evaluador (solo certificado)"),
        ("E21", "rol_registro_evaluador", "Rol de registro del evaluador (solo certificado)"),
        ("E22", "rut_evaluador", "RUT del evaluador (solo certificado)"),
        ("E28", "solicitado_por", "Solicitado por (solo certificado)"),
        ("E29", "rut_mandante", "RUT del mandante (solo certificado)"),
        ("E33", "descripcion_muro_principal", "Descripción del muro principal (certificado)"),
        ("E34", "descripcion_muro_secundario", "Descripción del muro secundario (certificado)"),
        ("E35", "descripcion_piso", "Descripción del piso (certificado)"),
        ("E36", "descripcion_techo_principal", "Descripción del techo principal (certificado)"),
        ("E37", "descripcion_techo_secundario", "Descripción del techo secundario (certificado)"),
        ("E38", "descripcion_vidrio_principal", "Descripción del vidrio principal (certificado)"),
        ("E39", "descripcion_vidrio_secundario", "Descripción del vidrio secundario (certificado)"),
        ("E40", "descripcion_puerta", "Descripción de la puerta (certificado)"),
        ("O38", "descripcion_marco_principal", "Descripción del marco principal (certificado)"),
        ("O39", "descripcion_marco_secundario", "Descripción del marco secundario (certificado)"),
        ("E44", "sistema_calefaccion", "Sistema de calefacción (solo certificado)"),
        ("E45", "sistema_agua_caliente", "Sistema de agua caliente sanitaria (solo certificado)"),
    ])
    e += _filas(CEVE, "Dimensiones", "C1 §2", "ceve.dimensiones", 52, range(52, 55), [
        ("D", "area", "Área del piso [m²]"),
        ("E", "altura", "Altura del piso [m]"),
    ])
    e += _filas(CEVE, "Muros", "C1 §3", "ceve.muros", 66, range(66, 82), [
        ("D", "construccion", "Construcción del muro (referencia a la tabla Muros)"),
        ("E", "azimut", "Azimut del muro (rango)"),
        ("H", "area", "Área del muro [m²]"),
        ("K", "p01", "Puente térmico P01, esquina saliente [ml]"),
        ("L", "p02", "Puente térmico P02, esquina entrante [ml]"),
        ("M", "p03", "Puente térmico P03, encuentro con elemento interior [ml]"),
    ])
    e += _filas(CEVE, "Puentes térmicos particulares", "C1 §3.1", "ceve.puentes_particulares", 87, range(87, 92), [
        ("D", "alojada_en", "Muro o techo que aloja el puente particular"),
        ("E", "azimut", "Azimut (rango)"),
        ("G", "elemento_perpendicular", "Elemento perpendicular (terraza)"),
        ("H", "aislacion", "Aislación del puente (terraza)"),
        ("I", "longitud", "Longitud [m]"),
    ])
    e += _filas(CEVE, "Puertas", "C1 §4", "ceve.puertas", 96, range(96, 99), [
        ("D", "tipo", "Tipo de puerta (referencia a la tabla Puertas)"),
        ("E", "azimut", "Azimut (rango)"),
        ("H", "categoria_infiltracion", "Categoría de infiltraciones (materialidad)"),
        ("K", "alto", "Alto [m]"),
        ("L", "ancho", "Ancho [m]"),
        *FAV,
    ])
    ventana = [
        ("D", "tipo_vidrio", "Tipo de ventana (referencia a la tabla Vidrios)"),
        ("E", "azimut", "Azimut (rango)"),
        ("G", "elemento_envolvente", "Elemento de envolvente que contiene la ventana"),
        ("H", "tipo_cierre", "Tipo de cierre"),
        ("I", "posicion", "Posición del ventanal"),
        ("J", "aislacion_marco_vano", "Aislación marco/vano"),
        ("M", "material_marco", "Material del marco"),
        ("N", "marco", "Marco (referencia a la tabla Marcos)"),
    ]
    e += _filas(CEVE, "Ventanas", "C1 §5", "ceve.ventanas", 103, range(103, 124), ventana)
    e += _filas(CEVE, "Ventanas", "C1 §5", "ceve.ventanas", 103, range(103, 123), [
        ("K", "alto", "Alto [m]"),
        ("L", "ancho", "Ancho [m]"),
        *FAV,
    ])
    e += _sueltas(CEVE, "FAR (azimut)", "C1 §6", "ceve.far", [
        (celda, f"{clave}.azimut", f"Azimut de las obstrucciones, bloque {nombre}")
        for clave, celda, *_, nombre in FAR_BLOQUES
    ])
    for clave, _, primera, ultima_div, _, col_div, _, _, _, nombre in FAR_BLOQUES:
        e += _filas(CEVE, "FAR (División)", "C1 §6", f"ceve.far.{clave}.obstrucciones", primera,
                    range(primera, ultima_div + 1),
                    [(col_div, "division", f"División de la obstrucción (sector de 36°), bloque {nombre}")])
    for clave, _, primera, _, ultima_abd, _, col_a, col_b, col_d, nombre in FAR_BLOQUES:
        e += _filas(CEVE, "FAR (A/B/D)", "C1 §6", f"ceve.far.{clave}.obstrucciones", primera,
                    range(primera, ultima_abd + 1), [
                        (col_a, "a", f"A: diferencia de altura de la obstrucción [m], bloque {nombre}"),
                        (col_b, "b", f"B: distancia horizontal a la obstrucción [m], bloque {nombre}"),
                        (col_d, "d", f"D: ancho de la obstrucción [m], bloque {nombre}"),
                    ])
    e += _filas(CEVE, "Techos", "C1 §7", "ceve.techos", 150, range(150, 155), [
        ("D", "construccion", "Construcción del techo (referencia a la tabla Techos)"),
        ("F", "densidad", "Densidad del techo"),
        ("G", "area", "Área del techo [m²]"),
        ("K", "camaras_aire", "Cámaras de aire"),
        ("L", "tipo_cubierta", "Tipo de cubierta"),
    ])
    e += _filas(CEVE, "Pisos", "C1 §8", "ceve.pisos", 158, range(158, 163), [
        ("D", "construccion", "Construcción del piso (referencia a la tabla Pisos)"),
        ("F", "densidad", "Densidad del piso"),
        ("L", "tipo_piso", "Piso ventilado o no ventilado"),
    ])
    e += _filas(CEVE, "Pisos", "C1 §8", "ceve.pisos", 158, range(158, 162), [
        ("G", "area", "Área del piso [m²]"),
        ("K", "perimetro_terreno", "Perímetro en contacto con el terreno [m]"),
    ])
    e += _filas(CEVE, "Pisos", "C1 §8", "ceve.pisos", 158, range(162, 163), [
        ("N", "columna_n", "Celda editable N162 listada en la tabla §11 de C1 (familia Pisos)"),
    ])
    e += _sueltas(CEVE, "Infiltración/ventilación", "C1 §9", "ceve.infiltracion", [
        ("F217", "ensayo_presurizacion", "¿Cuenta con ensayo de presurización?"),
        ("F219", "valor_ensayo_rah50", "Valor del ensayo [RAH] a 50 Pa"),
        ("F221", "ductos_ventilacion", "Número de ductos de ventilación"),
        ("F223", "celosias", "Número de celosías"),
        ("E227", "ventilacion_mecanica", "Ventilación mecánica (identificación)"),
        ("F229", "eficiencia_recuperador", "Eficiencia del recuperador de calor [%]"),
        ("F232", "sensor_co2", "¿Tiene sensor de CO2?"),
        ("F234", "rah_memoria", "RAH según memoria de cálculo"),
    ])
    e += _sueltas(CEVE, "Perfil de uso", "C1 §9.1", "ceve.perfil_uso", [
        ("H188", "contador_hora_cargas", "Contador de hora de cargas internas (semilla)"),
        ("AQ189", "cargas_internas_variables", "¿Cargas internas variables?"),
        ("H215", "contador_hora_infiltraciones", "Contador de hora de infiltraciones (semilla)"),
        ("AP225", "perfil_infiltraciones_variable", "¿Perfil de infiltraciones variable?"),
        ("F230", "eficiencia_por_defecto", "Eficiencia por defecto [%]"),
    ])
    e += _filas(TABLAS, "Puertas", "C1 §10.1", "tablas.puertas", 12, range(12, 24), [
        ("B", "nombre", "Nombre de la puerta"),
        ("C", "abreviatura", "Abreviatura (referencia usada en CEV-CEVE)"),
        ("D", "u_opaca", "U puerta opaca [W/m²K]"),
        ("E", "vidrio", "Vidrio de la puerta (referencia a la tabla Vidrios)"),
        ("F", "fraccion_vidrio", "Fracción de vidrio (0-1)"),
        ("G", "u_marco", "U marco [W/m²K]"),
        ("H", "fraccion_marco", "Fracción de marco (0-1)"),
    ])
    e += _filas(TABLAS, "Vidrios", "C1 §10.2", "tablas.vidrios", 27, range(27, 43), [
        ("B", "nombre", "Nombre del vidrio"),
        ("C", "abreviatura", "Abreviatura (referencia usada en CEV-CEVE)"),
        ("D", "u_vidrio", "U vidrio [W/m²K]"),
    ])
    e += _filas(TABLAS, "Vidrios", "C1 §10.2", "tablas.vidrios", 27, range(32, 43), [
        ("E", "fs", "Factor solar del vidrio (fila libre)"),
    ])
    e += _filas(TABLAS, "Marcos", "C1 §10.3", "tablas.marcos", 46, range(46, 58), [
        ("B", "nombre", "Nombre del marco"),
        ("C", "abreviatura", "Abreviatura (referencia usada en CEV-CEVE)"),
        ("D", "u_marco", "U marco [W/m²K]"),
        ("E", "fraccion_marco", "Fracción de marco FM (0-1)"),
    ])
    e += _filas(TABLAS, "Muros", "C1 §10.4", "tablas.muros", 61, range(61, 76), [
        ("B", "nombre", "Nombre del muro"),
        ("C", "abreviatura", "Abreviatura (referencia usada en CEV-CEVE)"),
        ("D", "densidad", "Tipología/materialidad (densidad)"),
        ("E", "u", "U [W/m²K]"),
        ("F", "espesor_solido_cm", "Espesor del muro sólido [cm]"),
        ("G", "espesor_aislante_cm", "Espesor del aislante [cm]"),
        ("H", "posicion_aislacion", "Posición de la aislación"),
    ])
    e += _filas(TABLAS, "Techos", "C1 §10.5", "tablas.techos", 79, range(79, 83), [
        ("B", "nombre", "Nombre del techo"),
        ("C", "abreviatura", "Abreviatura (referencia usada en CEV-CEVE)"),
        ("D", "u", "U [W/m²K]"),
        ("F", "espesor_solido_cm", "Espesor del techo sólido [cm]"),
        ("G", "espesor_aislante_cm", "Espesor del aislante [cm]"),
        ("H", "posicion_aislacion", "Posición de la aislación"),
    ])
    e += _filas(TABLAS, "Pisos", "C1 §10.6", "tablas.pisos", 87, range(87, 92), [
        ("B", "nombre", "Nombre del piso"),
        ("C", "abreviatura", "Abreviatura (referencia usada en CEV-CEVE)"),
        ("D", "u_ventilado", "U piso ventilado [W/m²K]"),
        ("E", "lambda_aislacion", "λ de la aislación bajo piso [W/mK]"),
        ("F", "espesor_aislante_cm", "Espesor del aislante [cm]"),
        ("G", "refuerzo_vertical_lambda", "Refuerzo vertical: λ [W/mK]"),
        ("H", "refuerzo_vertical_espesor_cm", "Refuerzo vertical: espesor [cm]"),
        ("I", "refuerzo_vertical_d", "Refuerzo vertical: D [m]"),
        ("J", "refuerzo_horizontal_lambda", "Refuerzo horizontal: λ [W/mK]"),
        ("K", "refuerzo_horizontal_espesor_cm", "Refuerzo horizontal: espesor [cm]"),
        ("L", "refuerzo_horizontal_d", "Refuerzo horizontal: D [m]"),
        ("M", "posicion_aislacion", "Posición de la aislación"),
    ])
    return e


def comprobar(entradas, libro) -> None:
    """Totales 944/380/1324, celdas y campos únicos, y ninguna celda listada con fórmula en `libro`."""
    por_hoja: dict[str, int] = {}
    for hoja, *_ in entradas:
        por_hoja[hoja] = por_hoja.get(hoja, 0) + 1
    if por_hoja != TOTALES or len(entradas) != sum(TOTALES.values()):
        raise GeneracionError(f"totales {por_hoja} (se esperan {TOTALES}, total {sum(TOTALES.values())})")
    if len({(h, c) for h, c, *_ in entradas}) != len(entradas):
        raise GeneracionError("hay celdas repetidas en la tabla")
    if len({x[3] for x in entradas}) != len(entradas):
        raise GeneracionError("hay campos VCEV repetidos en la tabla")
    con_formula = [f"{h}!{c}" for h, c, *_ in entradas if libro[h][c].data_type == "f"]
    if con_formula:
        raise GeneracionError(f"celdas listadas con fórmula en la plantilla: {', '.join(con_formula[:20])}")


def _sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def _celdas_xml(ruta: Path) -> dict[str, set[str]]:
    """{hoja: referencias que existen como elemento <c> en el XML de esa hoja}."""
    with zipfile.ZipFile(ruta) as archivo:
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in ET.fromstring(archivo.read("xl/_rels/workbook.xml.rels"))}
        resultado: dict[str, set[str]] = {}
        for hoja in ET.fromstring(archivo.read("xl/workbook.xml")).find("s:sheets", NS):
            nombre = hoja.attrib["name"]
            if nombre not in HOJAS:
                continue
            destino = rels[hoja.attrib[REL_ID]]
            parte = destino.lstrip("/") if destino.startswith("/") else "xl/" + destino
            resultado[nombre] = set(CELDA_XML.findall(archivo.read(parte).decode("utf-8")))
    return resultado


def _valor_json(valor):
    if valor is None or isinstance(valor, str):
        return valor
    if isinstance(valor, bool):
        return int(valor)
    if isinstance(valor, (int, float)):
        return valor
    return str(valor)


def _numero(formula):
    try:
        return float(str(formula).lstrip("="))
    except (TypeError, ValueError):
        return None


def _valores(libro_v, destinos) -> list:
    valores = []
    for hoja, ref in destinos:
        celdas = libro_v[hoja][ref]
        if not isinstance(celdas, tuple):
            celdas = ((celdas,),)
        for fila in celdas:
            for celda in fila if isinstance(fila, tuple) else (fila,):
                valor = _valor_json(celda.value)
                if valor is not None and valor != "":
                    valores.append(valor)
    return valores


def _destinos(formula: str, hoja: str, libro_f):
    """[(hoja, ref)] a los que apunta la fórmula de una lista, o None si no es un rango fijo resoluble."""
    if NOMBRE.match(formula):
        definido = libro_f[hoja].defined_names.get(formula) or libro_f.defined_names.get(formula)
        if definido is None:
            return None
        try:
            return [(h, r) for h, r in definido.destinations]
        except Exception:  # noqa: BLE001 — nombre definido por fórmula (OFFSET, INDIRECT…): dinámico
            return None
    coincidencia = RANGO.match(formula)
    if coincidencia:
        return [(coincidencia.group("hoja") or hoja, coincidencia.group("ref"))]
    return None


def dominio(hoja: str, celda: str, libro_f, libro_v, validaciones) -> dict:
    validacion = next((dv for dv in validaciones if celda in dv.sqref), None)
    if validacion is None or validacion.type is None:
        return {"tipo": "texto_libre"}
    formula1 = (validacion.formula1 or "").lstrip("=")
    if validacion.type == "list":
        if formula1.startswith('"') and formula1.endswith('"'):
            return {"tipo": "lista", "origen": formula1, "valores": [v.strip() for v in formula1[1:-1].split(",")]}
        destinos = _destinos(formula1, hoja, libro_f)
        if destinos and all(h == REFERENCIA for h, _ in destinos):
            valores = _valores(libro_v, destinos)
            if valores:
                return {"tipo": "lista", "origen": formula1, "valores": valores}
        return {"tipo": "lista_dinamica", "origen": formula1}
    if validacion.type in ("decimal", "whole"):
        return {
            "tipo": "decimal" if validacion.type == "decimal" else "entero",
            "operador": validacion.operator or "between",
            "min": _numero(validacion.formula1),
            "max": _numero(validacion.formula2),
            "formula1": validacion.formula1,
            "formula2": validacion.formula2,
        }
    return {"tipo": "otro", "validacion": validacion.type, "formula1": validacion.formula1}


def generar(salida: Path) -> tuple[int, int]:
    """Escribe el mapeo en `salida`; devuelve (celdas, celdas sin elemento <c> en la plantilla)."""
    if _sha256(PLANTILLA) != SHA256_PLANTILLA:
        raise GeneracionError(f"SHA-256 de plantilla desconocido: {PLANTILLA} (se exige {SHA256_PLANTILLA})")
    libro_f = load_workbook(PLANTILLA, data_only=False)
    libro_v = load_workbook(PLANTILLA, data_only=True)
    entradas = tabla()
    comprobar(entradas, libro_f)
    en_xml = _celdas_xml(PLANTILLA)
    validaciones = {hoja: list(libro_f[hoja].data_validations.dataValidation) for hoja in HOJAS}
    celdas = [
        {
            "hoja": hoja,
            "celda": celda,
            "familia": familia,
            "campo": campo,
            "significado": significado,
            "seccion_c1": seccion,
            "dominio": dominio(hoja, celda, libro_f, libro_v, validaciones[hoja]),
            "valor_plantilla": _valor_json(libro_f[hoja][celda].value),
            "en_xml_plantilla": celda in en_xml[hoja],
        }
        for hoja, celda, familia, campo, significado, seccion in entradas
    ]
    datos = {
        "plantilla": NOMBRE_PLANTILLA,
        "plantilla_sha256": SHA256_PLANTILLA,
        "generador": "herramientas/generar_celdas.py",
        "fuente": "diccionario C1 de cevstudio, tabla §11 (docs/model-methods/pbtd01-celda-a-celda.md)",
        "totales": {**TOTALES, "total": len(celdas)},
        "celdas": celdas,
    }
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text(json.dumps(datos, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return len(celdas), sum(1 for c in celdas if not c["en_xml_plantilla"])


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--salida", type=Path, default=DESTINO)
    args = parser.parse_args(argv)
    try:
        total, ausentes = generar(args.salida)
    except GeneracionError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(f"CELDAS OK {TOTALES[CEVE]} + {TOTALES[TABLAS]} = {total} (celdas sin elemento <c> en la plantilla: {ausentes})")
    return 0


if __name__ == "__main__":
    sys.exit(main())

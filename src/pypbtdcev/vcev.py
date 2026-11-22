"""VCEV: la vivienda CEV como datos, estructurada por las familias del diccionario C1 (PBTD-01 v2.2).

Cada hoja de valor corresponde a exactamente una celda editable de la plantilla PBTD-01 v2.2 (1324 en total); el
mapeo está en pypbtdcev/datos/pbtd01-v2.2-celdas.json (clave "campo"). Los valores son permisivos
(str | float | int | None) para que cualquier libro real cargue; el dominio se revisa aparte con validar().
Las filas de cada familia son tuplas de largo fijo: el índice es la fila menos la primera fila del bloque.
"""

import re
from functools import cache
from typing import Literal, get_args, get_origin

from pydantic import BaseModel, ConfigDict, Field

Valor = str | float | int | None


class _Base(BaseModel):
    model_config = ConfigDict(extra="forbid", validate_assignment=True)


def _tupla(*tipos):
    """Anotación tuple[A, A, …, B] de largo fijo (una posición por fila de la planilla)."""
    return tuple[tipos]


def _filas(*tipos):
    """Valor por defecto: una instancia vacía por fila."""
    return Field(default_factory=lambda: tuple(tipo() for tipo in tipos))


class DatosGenerales(_Base):
    tipo_calificacion: Valor = None  # E7
    vivienda_nueva_antigua: Valor = None  # G7
    region: Valor = None  # E8
    region_f8: Valor = None  # F8
    region_g8: Valor = None  # G8
    comuna: Valor = None  # E9
    zona_termica: Valor = None  # E10
    dormitorios: Valor = None  # E11
    identificacion_vivienda: Valor = None  # E13
    nombre_proyecto: Valor = None  # E14
    direccion: Valor = None  # E15
    tipo_vivienda: Valor = None  # E16
    tipo_vivienda_f16: Valor = None  # F16
    tipo_vivienda_g16: Valor = None  # G16
    caso_interno_evaluador: Valor = None  # E25
    iteracion_evaluador: Valor = None  # E26


class Certificado(_Base):
    rol_vivienda: Valor = None  # E19
    evaluador: Valor = None  # E20
    rol_registro_evaluador: Valor = None  # E21
    rut_evaluador: Valor = None  # E22
    solicitado_por: Valor = None  # E28
    rut_mandante: Valor = None  # E29
    descripcion_muro_principal: Valor = None  # E33
    descripcion_muro_secundario: Valor = None  # E34
    descripcion_piso: Valor = None  # E35
    descripcion_techo_principal: Valor = None  # E36
    descripcion_techo_secundario: Valor = None  # E37
    descripcion_vidrio_principal: Valor = None  # E38
    descripcion_vidrio_secundario: Valor = None  # E39
    descripcion_puerta: Valor = None  # E40
    descripcion_marco_principal: Valor = None  # O38
    descripcion_marco_secundario: Valor = None  # O39
    sistema_calefaccion: Valor = None  # E44
    sistema_agua_caliente: Valor = None  # E45


class Dimension(_Base):  # filas 52-54
    area: Valor = None  # D
    altura: Valor = None  # E


class Muro(_Base):  # filas 66-81
    construccion: Valor = None  # D
    azimut: Valor = None  # E
    area: Valor = None  # H
    p01: Valor = None  # K
    p02: Valor = None  # L
    p03: Valor = None  # M


class PuenteParticular(_Base):  # filas 87-91
    alojada_en: Valor = None  # D
    azimut: Valor = None  # E
    elemento_perpendicular: Valor = None  # G
    aislacion: Valor = None  # H
    longitud: Valor = None  # I


class Fav(_Base):  # columnas P-Y de puertas y ventanas
    fav1_d: Valor = None  # P
    fav1_l: Valor = None  # Q
    fav2_izq_profundidad: Valor = None  # R
    fav2_izq_separacion: Valor = None  # S
    fav2_der_profundidad: Valor = None  # T
    fav2_der_separacion: Valor = None  # U
    fav3_espesor: Valor = None  # V
    fav3_separacion: Valor = None  # W
    fav3_beta: Valor = None  # X
    fav3_alfa: Valor = None  # Y


class Puerta(Fav):  # filas 96-98
    tipo: Valor = None  # D
    azimut: Valor = None  # E
    categoria_infiltracion: Valor = None  # H
    alto: Valor = None  # K
    ancho: Valor = None  # L


class VentanaFila123(_Base):  # fila 123: sin alto, ancho ni FAV
    tipo_vidrio: Valor = None  # D
    azimut: Valor = None  # E
    elemento_envolvente: Valor = None  # G
    tipo_cierre: Valor = None  # H
    posicion: Valor = None  # I
    aislacion_marco_vano: Valor = None  # J
    material_marco: Valor = None  # M
    marco: Valor = None  # N


class Ventana(VentanaFila123, Fav):  # filas 103-122
    alto: Valor = None  # K
    ancho: Valor = None  # L


class ObstruccionSoloDivision(_Base):  # fila 135 de los bloques N/E/S/O
    division: Valor = None


class Obstruccion(ObstruccionSoloDivision):  # filas 127-134 y 139-146
    a: Valor = None
    b: Valor = None
    d: Valor = None


OBSTRUCCIONES_CARDINAL = (Obstruccion,) * 8 + (ObstruccionSoloDivision,)
OBSTRUCCIONES_DIAGONAL = (Obstruccion,) * 8


class FarCardinal(_Base):
    azimut: Valor = None
    obstrucciones: _tupla(*OBSTRUCCIONES_CARDINAL) = _filas(*OBSTRUCCIONES_CARDINAL)


class FarDiagonal(_Base):
    azimut: Valor = None
    obstrucciones: _tupla(*OBSTRUCCIONES_DIAGONAL) = _filas(*OBSTRUCCIONES_DIAGONAL)


class Far(_Base):
    n: FarCardinal = Field(default_factory=FarCardinal)  # G125; E/F/G/H 127-135
    e: FarCardinal = Field(default_factory=FarCardinal)  # L125; J/K/L/M
    s: FarCardinal = Field(default_factory=FarCardinal)  # Q125; O/P/Q/R
    o: FarCardinal = Field(default_factory=FarCardinal)  # V125; T/U/V/W
    ne: FarDiagonal = Field(default_factory=FarDiagonal)  # G137; E/F/G/H 139-146
    se: FarDiagonal = Field(default_factory=FarDiagonal)  # L137; J/K/L/M
    so: FarDiagonal = Field(default_factory=FarDiagonal)  # Q137; O/P/Q/R
    no: FarDiagonal = Field(default_factory=FarDiagonal)  # V137; T/U/V/W


class Techo(_Base):  # filas 150-154
    construccion: Valor = None  # D
    densidad: Valor = None  # F
    area: Valor = None  # G
    camaras_aire: Valor = None  # K
    tipo_cubierta: Valor = None  # L


class PisoFila162(_Base):  # fila 162: sin área ni perímetro, con N
    construccion: Valor = None  # D
    densidad: Valor = None  # F
    tipo_piso: Valor = None  # L
    columna_n: Valor = None  # N


class Piso(_Base):  # filas 158-161
    construccion: Valor = None  # D
    densidad: Valor = None  # F
    area: Valor = None  # G
    perimetro_terreno: Valor = None  # K
    tipo_piso: Valor = None  # L


class Infiltracion(_Base):
    ensayo_presurizacion: Valor = None  # F217
    valor_ensayo_rah50: Valor = None  # F219
    ductos_ventilacion: Valor = None  # F221
    celosias: Valor = None  # F223
    ventilacion_mecanica: Valor = None  # E227
    eficiencia_recuperador: Valor = None  # F229
    sensor_co2: Valor = None  # F232
    rah_memoria: Valor = None  # F234


class PerfilUso(_Base):
    contador_hora_cargas: Valor = None  # H188
    cargas_internas_variables: Valor = None  # AQ189
    contador_hora_infiltraciones: Valor = None  # H215
    perfil_infiltraciones_variable: Valor = None  # AP225
    eficiencia_por_defecto: Valor = None  # F230


DIMENSIONES = (Dimension,) * 3
MUROS = (Muro,) * 16
PUENTES = (PuenteParticular,) * 5
PUERTAS = (Puerta,) * 3
VENTANAS = (Ventana,) * 20 + (VentanaFila123,)
TECHOS = (Techo,) * 5
PISOS = (Piso,) * 4 + (PisoFila162,)


class HojaCEVE(_Base):
    datos_generales: DatosGenerales = Field(default_factory=DatosGenerales)
    certificado: Certificado = Field(default_factory=Certificado)
    dimensiones: _tupla(*DIMENSIONES) = _filas(*DIMENSIONES)
    muros: _tupla(*MUROS) = _filas(*MUROS)
    puentes_particulares: _tupla(*PUENTES) = _filas(*PUENTES)
    puertas: _tupla(*PUERTAS) = _filas(*PUERTAS)
    ventanas: _tupla(*VENTANAS) = _filas(*VENTANAS)
    far: Far = Field(default_factory=Far)
    techos: _tupla(*TECHOS) = _filas(*TECHOS)
    pisos: _tupla(*PISOS) = _filas(*PISOS)
    infiltracion: Infiltracion = Field(default_factory=Infiltracion)
    perfil_uso: PerfilUso = Field(default_factory=PerfilUso)


class TablaPuerta(_Base):  # filas 12-23
    nombre: Valor = None  # B
    abreviatura: Valor = None  # C
    u_opaca: Valor = None  # D
    vidrio: Valor = None  # E
    fraccion_vidrio: Valor = None  # F
    u_marco: Valor = None  # G
    fraccion_marco: Valor = None  # H


class TablaVidrioPrellenado(_Base):  # filas 27-31 (FS es fórmula del template)
    nombre: Valor = None  # B
    abreviatura: Valor = None  # C
    u_vidrio: Valor = None  # D


class TablaVidrio(TablaVidrioPrellenado):  # filas 32-42
    fs: Valor = None  # E


class TablaMarco(_Base):  # filas 46-57
    nombre: Valor = None  # B
    abreviatura: Valor = None  # C
    u_marco: Valor = None  # D
    fraccion_marco: Valor = None  # E


class TablaMuro(_Base):  # filas 61-75
    nombre: Valor = None  # B
    abreviatura: Valor = None  # C
    densidad: Valor = None  # D
    u: Valor = None  # E
    espesor_solido_cm: Valor = None  # F
    espesor_aislante_cm: Valor = None  # G
    posicion_aislacion: Valor = None  # H


class TablaTecho(_Base):  # filas 79-82
    nombre: Valor = None  # B
    abreviatura: Valor = None  # C
    u: Valor = None  # D
    espesor_solido_cm: Valor = None  # F
    espesor_aislante_cm: Valor = None  # G
    posicion_aislacion: Valor = None  # H


class TablaPiso(_Base):  # filas 87-91
    nombre: Valor = None  # B
    abreviatura: Valor = None  # C
    u_ventilado: Valor = None  # D
    lambda_aislacion: Valor = None  # E
    espesor_aislante_cm: Valor = None  # F
    refuerzo_vertical_lambda: Valor = None  # G
    refuerzo_vertical_espesor_cm: Valor = None  # H
    refuerzo_vertical_d: Valor = None  # I
    refuerzo_horizontal_lambda: Valor = None  # J
    refuerzo_horizontal_espesor_cm: Valor = None  # K
    refuerzo_horizontal_d: Valor = None  # L
    posicion_aislacion: Valor = None  # M


T_PUERTAS = (TablaPuerta,) * 12
T_VIDRIOS = (TablaVidrioPrellenado,) * 5 + (TablaVidrio,) * 11
T_MARCOS = (TablaMarco,) * 12
T_MUROS = (TablaMuro,) * 15
T_TECHOS = (TablaTecho,) * 4
T_PISOS = (TablaPiso,) * 5


class TablasEnvolvente(_Base):
    puertas: _tupla(*T_PUERTAS) = _filas(*T_PUERTAS)
    vidrios: _tupla(*T_VIDRIOS) = _filas(*T_VIDRIOS)
    marcos: _tupla(*T_MARCOS) = _filas(*T_MARCOS)
    muros: _tupla(*T_MUROS) = _filas(*T_MUROS)
    techos: _tupla(*T_TECHOS) = _filas(*T_TECHOS)
    pisos: _tupla(*T_PISOS) = _filas(*T_PISOS)


class Procedencia(_Base):
    fuente: Literal["pbtd01"] = "pbtd01"
    libro_sha256: str
    hoja: str
    celda: str
    formula_usuario: str | None = None


class Incidencia(_Base):
    hoja: str
    celda: str
    campo: str
    valor: Valor = None
    regla: str
    detalle: str = ""


class VCEV(_Base):
    version_esquema: Literal["1"] = "1"
    ceve: HojaCEVE = Field(default_factory=HojaCEVE)
    tablas: TablasEnvolvente = Field(default_factory=TablasEnvolvente)
    procedencia: dict[str, Procedencia] = Field(default_factory=dict)
    incidencias: list[Incidencia] = Field(default_factory=list)


class CampoDesconocido(KeyError):
    """Camino que no es una hoja de valor del VCEV."""


_PASO = re.compile(r"^([a-z_][a-z0-9_]*)(?:\[([0-9]+)\])?$")


def _rutas_de(anotacion, ruta):
    if get_origin(anotacion) is tuple:
        for indice, sub in enumerate(get_args(anotacion)):
            yield from _rutas_de(sub, f"{ruta}[{indice}]")
    elif isinstance(anotacion, type) and issubclass(anotacion, BaseModel):
        for nombre, info in anotacion.model_fields.items():
            yield from _rutas_de(info.annotation, f"{ruta}.{nombre}")
    else:
        yield ruta


@cache
def _todas() -> tuple[str, ...]:
    return tuple(_rutas_de(HojaCEVE, "ceve")) + tuple(_rutas_de(TablasEnvolvente, "tablas"))


@cache
def _conjunto() -> frozenset[str]:
    return frozenset(_todas())


def rutas_hoja() -> list[str]:
    """Los caminos de todas las hojas de valor (1324), en orden de declaración."""
    return list(_todas())


def _ubicar(vcev: VCEV, campo: str):
    if campo not in _conjunto():
        raise CampoDesconocido(campo)
    partes = campo.split(".")
    objeto = vcev
    for parte in partes[:-1]:
        nombre, indice = _PASO.match(parte).groups()
        objeto = getattr(objeto, nombre)
        if indice is not None:
            objeto = objeto[int(indice)]
    return objeto, partes[-1]


def leer(vcev: VCEV, campo: str) -> Valor:
    objeto, nombre = _ubicar(vcev, campo)
    return getattr(objeto, nombre)


def asignar(vcev: VCEV, campo: str, valor: Valor) -> None:
    objeto, nombre = _ubicar(vcev, campo)
    setattr(objeto, nombre, valor)

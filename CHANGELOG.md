# Changelog

## [0.2.0] — 2026-11-22

### Agregado

- Esquema VCEV (`pypbtdcev.vcev`, pydantic v2) estructurado por las familias del diccionario C1 de PBTD-01 v2.2.
- Mapeo de las 1324 celdas editables (`pypbtdcev/datos/pbtd01-v2.2-celdas.json`) con significado, dominio y
  sección C1, generado por `herramientas/generar_celdas.py`.
- `from_pbtd01`, `to_pbtd01` (escritura por edición directa del XML del `.xlsm`), `comparar` y `validar`.

### Cambiado

- `requires-python` pasa a `>=3.10`; nueva dependencia `pydantic>=2.14,<3`.

### Sin cambios

- `LectorPBTD01_v2`, `EscritorPBTD01_v2`, `LectorPBTD03_v2` y las plantillas (heredados).

# pypbtdcev

Una librería para leer y escribir planillas de PBTD de Calificación Energética de Viviendas en Chile.

## Instalación

```bash
pip install pypbtdcev
```

## Uso

*Próximamente: ejemplos de cómo usar la librería.*

## Desarrollo y Pruebas

Para ejecutar los scripts (como `clonar_planilla.py`) que se encuentran en la carpeta `pruebas/`, debes posicionarte en la **raíz del proyecto (la carpeta que contiene el archivo `pyproject.toml` y el directorio `src`)** y ejecutarlos como un módulo de Python. Esto asegura que las importaciones del paquete `pypbtdcev` y las rutas a los archivos funcionen correctamente.

Por ejemplo, para ejecutar el script de clonación:
```bash
python -m pruebas.clonar_planilla
```

## Uso (VCEV)

Desde 0.2.0, `pypbtdcev` lee y escribe la planilla PBTD-01 v2.2 a través del esquema VCEV (las 1324 celdas
editables de `CEV-CEVE` y `3. Tablas Envolvente`, mapeadas en `pypbtdcev/datos/pbtd01-v2.2-celdas.json`):

```python
from pypbtdcev import comparar, from_pbtd01, to_pbtd01, validar

vcev = from_pbtd01("ejemplos/01.-Ejamplo-1.xlsm")      # valores en caché + procedencia por celda
incidencias = validar(vcev)                            # dominio de cada celda (informativo, no bloquea)
informe = to_pbtd01(vcev, "salida.xlsm")               # copia de la plantilla oficial, editada por XML
assert comparar("ejemplos/01.-Ejamplo-1.xlsm", "salida.xlsm") == []
```

- `to_pbtd01` exige la plantilla oficial empaquetada (SHA-256 `dec31e42…b5de9b`), conserva VBA, estilos y
  fórmulas, rechaza escribir sobre una fórmula y pide a Excel recalcular al abrir.
- Una fórmula de usuario en una celda editable (p. ej. `=4.5*12.64`) se lee por su valor en caché y queda en
  `vcev.procedencia[campo].formula_usuario`; al escribir se escribe el valor, no la fórmula.
- `LectorPBTD01_v2`, `EscritorPBTD01_v2` y `LectorPBTD03_v2` siguen disponibles sin cambios (heredados).

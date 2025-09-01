# Archivo: pruebas/clonar_planilla.py

# Importamos las clases desde nuestro paquete pypbtdcev
from pypbtdcev.lector import LectorPBTD01_v2
from pypbtdcev.escritor import EscritorPBTD01_v2


def main():
    """
    Flujo de trabajo para clonar los datos de una planilla a otra.
    1. Lee una planilla Excel existente que ya tiene datos.
    2. Usa una planilla limpia como plantilla.
    3. Escribe los datos leídos en la plantilla para crear un nuevo archivo.
    """
    # --- Definir Rutas de Archivos ---
    # NOTA: Las rutas se han actualizado a la nueva estructura de carpetas.
    # Para que funcionen, este script debe ejecutarse desde el directorio raíz del proyecto.
    # La forma recomendada de ejecutarlo es: python -m pruebas.clonar_planilla
    
    # Archivo del cual leeremos los datos (debe estar ya rellenado)
    archivo_origen_completo = 'ejemplos/01.-Ejamplo-1.xlsm'
    
    # Archivo que usaremos como base limpia para escribir los datos
    archivo_plantilla_limpia = 'src/pypbtdcev/plantillas/01.-PBTD-Datos-de-Arquitectura-v2.2.xlsm'
    
    # Nombre del nuevo archivo que se va a crear con los datos clonados
    archivo_clonado_salida = 'planilla_clonada.xlsm'

    print("--- INICIANDO PROCESO DE CLONACIÓN DE DATOS ---")

    # --- PASO 1: LEER ---
    print(f"\n[PASO 1] Leyendo datos desde '{archivo_origen_completo}'...")
    lector = LectorPBTD01_v2(archivo_origen_completo)
    datos_leidos = lector.datos_extraidos

    if not datos_leidos:
        print("Proceso abortado debido a un error de lectura.")
        return
    
    print(" -> Lectura completada.")

    # --- PASO 2: MODIFICAR ---
    # Para esta prueba, no hacemos ninguna modificación.
    # Los datos leídos se pasarán directamente al escritor.
    print("\n[PASO 2] Omitiendo modificación de datos para la prueba de clonación.")

    # --- PASO 3: ESCRIBIR ---
    print("\n[PASO 3] Creando nueva planilla clonada...")
    escritor = EscritorPBTD01_v2()
    escritor.crear_nueva_planilla(
        ruta_plantilla=archivo_plantilla_limpia,
        ruta_salida=archivo_clonado_salida,
        datos=datos_leidos
    )

if __name__ == '__main__':
    main()
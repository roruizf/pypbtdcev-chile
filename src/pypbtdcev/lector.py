import pandas as pd
import numpy as np
import json


# -------------------------------------------
# --- 01.-PBTD-Datos-de-Arquitectura-v2.2 ---
# -------------------------------------------


class LectorPBTD01_v2:
    def __init__(self, filepath):
        try:
            self.xl_file_data = pd.read_excel(
                filepath, sheet_name=None, header=None)
            self.datos_extraidos = self._parse_all_sheets()
        except FileNotFoundError:
            print(
                f"❌ Error: No se encontró el archivo en la ruta '{filepath}'.")
            self.xl_file_data = None
            self.datos_extraidos = None
        except Exception as e:
            print(f"Ha ocurrido un error inesperado al leer el archivo: {e}")
            import traceback
            traceback.print_exc()
            self.xl_file_data = None
            self.datos_extraidos = None

    def _get_cell_value(self, df, cell_coord):
        """
        Función auxiliar para obtener el valor de una celda usando notación Excel (ej: 'C7').
        """
        col_str = ''.join(filter(str.isalpha, cell_coord))
        row_idx = int(''.join(filter(str.isdigit, cell_coord))) - 1

        # Convertir la letra de la columna a un índice numérico (A=0, B=1, ...)
        col_idx = sum([(ord(char) - ord('A') + 1) * (26 ** i)
                      for i, char in enumerate(reversed(col_str.upper()))]) - 1

        # .iat es un método muy rápido para acceder a un valor por sus índices numéricos
        return df.iat[row_idx, col_idx]

    def _convertir_decimales_a_float(self, df, columnas):
        """
        Función auxiliar para convertir columnas con comas decimales a números (float).
        """
        for col in columnas:
            if col in df.columns:
                # Convierte la columna a string, reemplaza la coma y luego convierte a numérico.
                # errors='coerce' convertirá cualquier valor no numérico en NaN (Not a Number).
                df[col] = pd.to_numeric(df[col].astype(
                    str).str.replace(',', '.'), errors='coerce')
        return df

    def _extraer_bloque_obstruccion(self, df, ancla_fila, ancla_col):
        """
        Extrae toda la información de un bloque de obstrucción a partir de una celda ancla.
        El ancla es la celda que contiene la orientación (ej: 'N').
        """
        # --- Lectura de datos individuales ---
        porcentaje_numpy = df.iat[ancla_fila - 1, ancla_col + 3]
        porcentaje = float(porcentaje_numpy) if pd.notna(
            porcentaje_numpy) else None

        azimut_numpy = df.iat[ancla_fila, ancla_col + 2]
        azimut_rango = str(azimut_numpy) if pd.notna(azimut_numpy) else None

        bloque = {
            'anual_referencial_rad_directa_porc': porcentaje,
            'azimut_rango': azimut_rango
        }

        # Extraer la tabla de 8 obstrucciones
        inicio_tabla = ancla_fila + 2
        fin_tabla = inicio_tabla + 8

        # Seleccionamos las columnas específicas que necesitamos para la tabla
        # column_indices = [2, 4, 5, 6, 7] # C, E, F, G, H
        column_indices = [
            2,              # Columna id_obstruccion (Siempre es la columna C)
            ancla_col,      # Columna division (ej: E para el bloque N)
            ancla_col + 1,  # Columna a_m (ej: F para el bloque N)
            ancla_col + 2,  # Columna b_m (ej: G para el bloque N)
            ancla_col + 3   # Columna d_m (ej: H para el bloque N)
        ]
        df_tabla = df.iloc[inicio_tabla:fin_tabla, column_indices].copy()

        df_tabla.columns = ['id_obstruccion', 'division', 'a_m', 'b_m', 'd_m']

        df_tabla = self._convertir_decimales_a_float(
            df_tabla, ['a_m', 'b_m', 'd_m'])

        json_string = df_tabla.replace(
            {np.nan: None}).to_json(orient='records')
        bloque['obstrucciones_detalle'] = json.loads(json_string)

        return bloque

    def _limpiar_dict_nan(self, d):
        """
        Recorre un diccionario simple y reemplaza los valores NaN por None.
        """
        for key, value in d.items():
            # pd.isna() detecta correctamente los valores nulos de pandas/numpy
            if pd.isna(value):
                d[key] = None
        return d

    def _parse_all_sheets(self):
        """
        Método principal que orquesta el parseo de todas las hojas de interés.
        """
        datos_completos = {}

        # Llama a la función de parseo para cada hoja y guarda su resultado
        datos_completos['CEV-CEVE'] = self._parsear_hoja_cev_ceve()
        datos_completos['3. Tablas Envolvente'] = self._parsear_hoja_tablas_envolvente(
        )

        return datos_completos

    def _parsear_hoja_cev_ceve(self):
        """
        Parsea la hoja 'CEV-CEVE' y extrae sus tres secciones principales.
        """
        sheet_name = 'CEV-CEVE'
        if sheet_name not in self.xl_file_data:
            print(
                f"Advertencia: No se encontró la hoja '{sheet_name}' en el archivo.")
            return

        df = self.xl_file_data[sheet_name]
        hoja_dict = {}
        # ---------------------------------------------------------
        # --- 1.1 Datos Generales (usando coordenadas de Excel) ---
        # ---------------------------------------------------------
        hoja_dict['datos_generales_proyecto'] = {
            'tipo_de_calificacion': self._get_cell_value(df, 'E7'),
            'tipo_de_vivienda_calificacion': self._get_cell_value(df, 'G7'),
            'region': self._get_cell_value(df, 'E8'),
            'comuna': self._get_cell_value(df, 'E9'),
            'zona_termica_proyecto': self._get_cell_value(df, 'E10'),
            'dormitorios_de_la_vivienda': self._get_cell_value(df, 'E11'),
            'identificacion_de_la_vivienda_a_evaluar': self._get_cell_value(df, 'E13'),
            'nombre_del_proyecto': self._get_cell_value(df, 'E14'),
            'direccion_de_la_vivienda': self._get_cell_value(df, 'E15'),
            'tipo_de_vivienda': self._get_cell_value(df, 'E16'),
            'rol_vivienda': self._get_cell_value(df, 'E19'),
            'evaluador_energetico': self._get_cell_value(df, 'E20'),
            'rol_registro_de_evaluadores': self._get_cell_value(df, 'E21'),
            'rut_evaluador': self._get_cell_value(df, 'E22'),
            'version_planilla': self._get_cell_value(df, 'E24'),
            'caso_interno_evaluador': self._get_cell_value(df, 'E25'),
            'iteracion_evaluador': self._get_cell_value(df, 'E26'),
            'solicitado_por': self._get_cell_value(df, 'E28'),
            'rut_mandante': self._get_cell_value(df, 'E29')
        }
        # Aplicamos la limpieza
        hoja_dict['datos_generales_proyecto'] = self._limpiar_dict_nan(
            hoja_dict['datos_generales_proyecto'])

        # --------------------------------------
        # --- 1.2 Elementos de la Envolvente ---
        # --------------------------------------
        hoja_dict['elementos_de_la_envolvente'] = {
            'muro_principal': self._get_cell_value(df, 'E33'),
            'muro_secundario': self._get_cell_value(df, 'E34'),
            'piso_principal': self._get_cell_value(df, 'E35'),
            'techo_principal': self._get_cell_value(df, 'E36'),
            'techo_secundario': self._get_cell_value(df, 'E37'),
            'ventana_principal_vidrio': self._get_cell_value(df, 'E38'),
            'ventana_principal_marco': self._get_cell_value(df, 'O38'),
            'ventana_secundaria_vidrio': self._get_cell_value(df, 'E39'),
            'ventana_secundaria_marco': self._get_cell_value(df, 'O39'),
            'puerta_principal': self._get_cell_value(df, 'E40')
        }
        # Aplicamos la limpieza
        hoja_dict['elementos_de_la_envolvente'] = self._limpiar_dict_nan(
            hoja_dict['elementos_de_la_envolvente'])

        # --- 1.3 Calefacción y ACS ---
        hoja_dict['calefaccion_y_acs'] = {
            'sistema_de_calefaccion': self._get_cell_value(df, 'E44'),
            'sistema_de_agua_caliente': self._get_cell_value(df, 'E45')
        }
        # Aplicamos la limpieza
        hoja_dict['calefaccion_y_acs'] = self._limpiar_dict_nan(
            hoja_dict['calefaccion_y_acs'])

        # --- 2. Dimensiones de la vivienda (Método de Coordenadas Fijas) ---
        # Extraemos el bloque de la tabla de pisos (Filas 52 a 54, Columnas C a F)
        df_pisos = df.iloc[51:54, 2:6].copy()

        # Asignamos nombres de columna para fácil acceso
        df_pisos.columns = ['piso', 'area_m2', 'altura_m', 'volumen_m3']

        # Extraemos los totales de sus celdas específicas
        totales = {
            'area_total_m2': df.iat[55, 3],  # Celda D56
            'volumen_total_m3': df.iat[55, 5]  # Celda F56
        }

        # Limpiamos los valores vacíos (NaN) convirtiéndolos a None
        df_pisos.replace({np.nan: None}, inplace=True)
        for key, value in totales.items():
            if pd.isna(value):
                totales[key] = None

        # Construimos el diccionario final para esta sección
        hoja_dict['dimensiones_de_la_vivienda'] = {
            'pisos': df_pisos.to_dict(orient='records'),
            'totales': totales
        }

        # ----------------------------------------------------------------------------------
        # --- 3.1 Area y coeficiente de transferencia de calor por elemento constructivo ---
        # ----------------------------------------------------------------------------------

        # -------------------
        # --- 3.1.1 Muros ---
        # -------------------
        # Definimos el bloque de la tabla de Muros (Filas 66 a 81)
        # Seleccionamos solo las columnas con datos: C, D, F, G, H, I, L, M, N, O
        column_indices = [2, 3, 4, 5, 6, 7, 8, 10,  11, 12, 14]
        df_muros = df.iloc[65:81, column_indices].copy()

        # Asignamos nombres de columna programáticos y limpios
        df_muros.columns = [
            'muro', 'nombre_muro', 'angulo_azimut', 'orientacion', 'densidad_muro',
            'area_m2', 'u_w_m2k', 'puente_termico_p01', 'puente_termico_p02',
            'puente_termico_p03', 'posicion_aislacion'
        ]

        # Usamos nuestra nueva función para corregir los decimales con coma
        df_muros = self._convertir_decimales_a_float(
            df_muros, ['area_m2', 'u_w_m2k'])

        # Reemplazamos todos los NaN restantes por None para un JSON limpio
        df_muros.replace({np.nan: None}, inplace=True)

        hoja_dict['area_y_coeficiente_muros'] = df_muros.to_dict(
            orient='records')

        # -------------------------------------------
        # --- 3.1.2 Puentes térmicos particulares ---
        # -------------------------------------------

        # Extraemos el bloque de la tabla (Filas 87 a 91, Columnas C a I)
        df_puentes = df.iloc[86:91, 2:9].copy()

        # Asignamos nombres de columna
        df_puentes.columns = [
            'id_puente_termico', 'alojada_en_muro', 'azimut', 'orientacion',
            'elemento_perpendicular', 'aislacion', 'longitud_m'
        ]

        # Convertimos la columna de longitud a número, manejando comas
        df_puentes = self._convertir_decimales_a_float(
            df_puentes, ['longitud_m'])

        # Limpiamos vacíos y añadimos al diccionario principal
        df_puentes.replace({np.nan: None}, inplace=True)
        hoja_dict['puentes_termicos_particulares'] = df_puentes.to_dict(
            orient='records')

        # ---------------------
        # --- 3.1.3 Puertas ---
        # ---------------------

        # Extraemos el bloque de la tabla (Filas 96 a 98) y seleccionamos las columnas con datos.
        column_indices_puertas = [2, 3, 4, 5, 7, 10, 11,
                                  13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]
        df_puertas = df.iloc[95:98, column_indices_puertas].copy()

        # Asignamos nombres de columna para aplanar la cabecera compleja del Excel
        df_puertas.columns = [
            'id_puerta', 'tipo_puerta', 'azimut', 'orientacion', 'categoria_infiltracion',
            'alto_m', 'ancho_m', 'area_vidrio_m2',
            'fav1_d', 'fav1_l', 'fav2_izquierda_p', 'fav2_izquierda_s',
            'fav2_derecha_p', 'fav2_derecha_s', 'fav3_e', 'fav3_t',
            'fav3_beta', 'fav3_alpha'
        ]
        # Definimos qué columnas deben ser numéricas
        columnas_numericas = [
            'alto_m', 'ancho_m', 'area_vidrio_m2', 'fav1_d', 'fav1_l',
            'fav2_izquierda_p', 'fav2_izquierda_s', 'fav2_derecha_p', 'fav2_derecha_s',
            'fav3_e', 'fav3_t', 'fav3_beta', 'fav3_alpha'
        ]
        df_puertas = self._convertir_decimales_a_float(
            df_puertas, columnas_numericas)

        # Limpiamos vacíos y añadimos al diccionario principal
        df_puertas.replace({np.nan: None}, inplace=True)
        hoja_dict['puertas'] = df_puertas.to_dict(orient='records')

        # ----------------------
        # --- 3.1.4 Ventanas ---
        # ----------------------

        # Extraemos el bloque de la tabla (Filas 103 a 122) y seleccionamos las columnas con datos.
        column_indices_ventanas = [
            2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13,
            15, 16, 17, 18, 19, 20, 21, 22, 23, 24
        ]
        df_ventanas = df.iloc[102:122, column_indices_ventanas].copy()

        # Asignamos nombres de columna para aplanar la cabecera compleja del Excel
        df_ventanas.columns = [
            'id_ventana', 'tipo_ventana', 'azimut', 'orientacion', 'elemento_envolvente',
            'tipo_cierre', 'posicion_ventanal', 'aislacion_con_sin_retorno', 'alto_m', 'ancho_m',
            'categoria_para_pt_y_infilt', 'tipo_marco', 'fav1_d', 'fav1_l', 'fav2_izquierda_p',
            'fav2_izquierda_s', 'fav2_derecha_p', 'fav2_derecha_s', 'fav3_e', 'fav3_t',
            'fav3_beta', 'fav3_alpha'
        ]

        # Definimos qué columnas deben ser numéricas
        columnas_numericas = [
            'alto_m', 'ancho_m', 'fav1_d', 'fav1_l', 'fav2_izquierda_p',
            'fav2_izquierda_s', 'fav2_derecha_p', 'fav2_derecha_s', 'fav3_e', 'fav3_t',
            'fav3_beta', 'fav3_alpha'
        ]
        df_ventanas = self._convertir_decimales_a_float(
            df_ventanas, columnas_numericas)

        # Limpiamos vacíos y añadimos al diccionario principal
        df_ventanas.replace({np.nan: None}, inplace=True)
        hoja_dict['ventanas'] = df_ventanas.to_dict(orient='records')

        # ---------------------------
        # --- 3.1.5 Obstrucciones ---
        # ---------------------------

        mapa_obstrucciones = {
            # Orientacion: (fila_ancla, col_ancla) -> índices numéricos
            'N': (124, 4),  # Celda E125
            'E': (124, 9),  # Celda J125
            'S': (124, 14),  # Celda O125
            'O': (124, 19),  # Celda T125
            'NE': (136, 4),  # Celda E137
            'SE': (136, 9),  # Celda J137
            'SO': (136, 14),  # Celda O137
            'NO': (136, 19)  # Celda T137
        }

        datos_obstrucciones = {}
        for orientacion, anclas in mapa_obstrucciones.items():
            fila, col = anclas
            datos_obstrucciones[orientacion] = self._extraer_bloque_obstruccion(
                df, fila, col)

        hoja_dict['obstrucciones'] = datos_obstrucciones

        # --- 3.1.6 Techos ---
        # Extraemos el bloque de la tabla (Filas 150 a 154) y seleccionamos las columnas con datos.
        column_indices_techos = [2, 3, 5, 6, 7,
                                 10, 11, 13]  # C, F, G, H, J, K, M
        df_techos = df.iloc[149:154, column_indices_techos].copy()

        # Asignamos nombres de columna para unificar la tabla
        df_techos.columns = [
            'id_techo', 'techos', 'densidad_techo', 'area_m2', 'u_w_m2k',
            'camaras_de_aire', 'tipo_de_cubierta', 'posicion_aislacion'
        ]

        # Convertimos las columnas numéricas
        df_techos = self._convertir_decimales_a_float(
            df_techos, ['area_m2', 'u_w_m2k'])

        # Limpiamos vacíos y convertimos a formato JSON-nativo
        json_string = df_techos.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['techos'] = json.loads(json_string)

        # -------------------
        # --- 3.1.7 Pisos ---
        # -------------------

        # Extraemos el bloque de la tabla (Filas 158 a 161)
        column_indices_pisos = [2, 3, 5, 6, 7, 10,
                                11, 13, 17]  # C, D, F, G, H, K, L, N, R
        df_pisos = df.iloc[157:161, column_indices_pisos].copy()

        # Asignamos nombres de columna
        df_pisos.columns = [
            'id_piso', 'piso', 'densidad_piso', 'area_m2', 'u_w_m2k',
            'perimetro_contacto_terreno_m', 'piso_ventilado',
            'posicion_aislacion', 'ls_w_k'
        ]

        # Convertimos las columnas numéricas
        columnas_numericas = ['area_m2', 'u_w_m2k',
                              'perimetro_contacto_terreno_m', 'ls_w_k']
        df_pisos = self._convertir_decimales_a_float(
            df_pisos, columnas_numericas)

        # Limpiamos vacíos y convertimos a formato JSON-nativo
        json_string = df_pisos.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['pisos'] = json.loads(json_string)

        # --------------------------------
        # --- 3.1.8 Resumen Envolvente ---
        # --------------------------------
        # 1. Extraer los bloques de datos en DataFrames individuales
        df_opacos = df.iloc[168:178, 2:6].copy()
        df_opacos.columns = ['orientacion', 'opacos_area_total_m2',
                             'opacos_area_efectiva_m2', 'opacos_u_w_m2k']
        df_opacos['orientacion'] = df_opacos['orientacion'].str.strip()
        df_opacos.set_index('orientacion', inplace=True)

        df_traslucidos = df.iloc[168:178, 7:9].copy()
        df_traslucidos.columns = ['traslucidos_area_m2', 'traslucidos_u_w_m2k']
        df_traslucidos.index = df_opacos.index  # Usar el mismo índice que opacos

        df_puentes_raw = df.iloc[168:178, 10:16].copy()
        df_puentes_raw.columns = ['pt_p01', 'pt_p02',
                                  'pt_p03', 'pt_p04', 'pt_p05', 'pt_total']
        df_puentes_raw.index = df_opacos.index

        df_sigma = df.iloc[168:178, 18:19].copy()
        df_sigma.columns = ['sigma_ua_alpha_l_w_k']
        df_sigma.index = df_opacos.index

        # 2. Tratar datos de Puentes Térmicos
        # Asigna el valor None de 'Piso Ls' a la fila 'Pisos', columna 'pt_p05'
        df_puentes_raw.loc['Pisos', 'pt_p05'] = np.nan

        # 3. Unir todas las tablas en una sola
        df_resumen = pd.concat(
            [df_opacos, df_traslucidos, df_puentes_raw, df_sigma], axis=1)

        # 4. Convertir todas las columnas a numérico
        df_resumen = self._convertir_decimales_a_float(
            df_resumen, df_resumen.columns)

        # 5. Estructurar el diccionario final
        resumen_dict = {}
        # Extraer el valor total final
        total_envolvente = self._get_cell_value(df, 'U168')
        resumen_dict['total_envolvente_no_adiabatica_m2'] = float(
            total_envolvente) if pd.notna(total_envolvente) else None

        # Convertir la tabla final a formato JSON-nativo
        json_string = df_resumen.reset_index().replace(
            {np.nan: None}).to_json(orient='records')
        resumen_dict['tabla_resumen'] = json.loads(json_string)

        hoja_dict['resumen_envolvente'] = resumen_dict

        # ---------------------------------------------
        # --- 4.1 Condiciones de uso de la vivienda ---
        # ---------------------------------------------
        condiciones_uso_dict = {}

        # -------------------------------------------
        # --- Parte 1: Ganancias internas por uso ---
        # -------------------------------------------

        # Leemos los valores directamente de las celdas para mayor precisión
        usuarios_diurna_raw = self._get_cell_value(df, 'E194')
        usuarios_nocturna_raw = self._get_cell_value(df, 'F194')
        iluminacion_diurna_raw = self._get_cell_value(df, 'E195')
        iluminacion_nocturna_raw = self._get_cell_value(df, 'F195')

        condiciones_uso_dict['ganancias_internas_por_uso'] = {
            "usuarios_w_m2": {
                "diurna": float(usuarios_diurna_raw) if pd.notna(usuarios_diurna_raw) else None,
                "nocturna": float(usuarios_nocturna_raw) if pd.notna(usuarios_nocturna_raw) else None
            },
            "iluminacion_w_m2": {
                "diurna": float(iluminacion_diurna_raw) if pd.notna(iluminacion_diurna_raw) else None,
                "nocturna": float(iluminacion_nocturna_raw) if pd.notna(iluminacion_nocturna_raw) else None
            }
        }

        # -----------------------------------------
        # --- Parte 2: Cargas internas horarias ---
        # -----------------------------------------

        # Extraemos el bloque de la tabla horaria (Filas 188 a 211, Columnas H a T)
        df_cargas = df.iloc[187:211,
                            7:20].copy()

        # Asignamos nombres de columna
        df_cargas.columns = [
            'hora', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
            'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
        ]

        # Convertimos todas las columnas de meses a valores numéricos
        columnas_meses = df_cargas.columns.drop('hora')
        df_cargas = self._convertir_decimales_a_float(
            df_cargas, columnas_meses)

        # Limpiamos y convertimos a formato JSON-nativo
        json_string = df_cargas.replace(
            {np.nan: None}).to_json(orient='records')
        condiciones_uso_dict['cargas_internas_horarias_w_m2'] = json.loads(
            json_string)

        # -------------------------------
        # --- Parte 3: Infiltraciones ---
        # -------------------------------

        infiltraciones_dict = {
            'cuenta_con_ensayo_presurizacion': self._get_cell_value(df, 'F217'),
            'valor_ensayo_presurizacion_rah_a_50pa': self._get_cell_value(df, 'F219'),
            'cantidad_ductos_ventilacion': self._get_cell_value(df, 'F221'),
            'cantidad_celosias': self._get_cell_value(df, 'F223')
        }
        for key, value in infiltraciones_dict.items():
            valor_numerico = pd.to_numeric(
                str(value).replace(',', '.'), errors='coerce')
            infiltraciones_dict[key] = float(
                valor_numerico) if pd.notna(valor_numerico) else value

        # Aplicamos la limpieza
        infiltraciones_dict = self._limpiar_dict_nan(infiltraciones_dict)

        condiciones_uso_dict['infiltraciones'] = infiltraciones_dict

        # ----------------------------
        # --- Parte 4: Ventilación ---
        # ----------------------------

        ventilacion_dict = {
            'ventilacion_mecanica_vm': self._get_cell_value(df, 'E227'),
            'eficiencia_recuperador_calor_porc': self._get_cell_value(df, 'F229'),
            'eficiencia_por_defecto_porc': self._get_cell_value(df, 'F230'),
            'tiene_sensor_co2': self._get_cell_value(df, 'F232'),
            'rah_segun_memoria_calculo': self._get_cell_value(df, 'F234')
        }
        for key, value in ventilacion_dict.items():
            valor_numerico = pd.to_numeric(
                str(value).replace(',', '.'), errors='coerce')
            ventilacion_dict[key] = float(
                valor_numerico) if pd.notna(valor_numerico) else value

        # Aplicamos la limpieza
        ventilacion_dict = self._limpiar_dict_nan(ventilacion_dict)

        condiciones_uso_dict['ventilacion'] = ventilacion_dict

        # -------------------------------------
        # --- Parte 5: Renovaciones de aire ---
        # -------------------------------------

        df_renovaciones = df.iloc[214:238, 7:20].copy()
        df_renovaciones.columns = [
            'hora', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
            'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
        ]
        columnas_meses = df_renovaciones.columns.drop('hora')
        df_renovaciones = self._convertir_decimales_a_float(
            df_renovaciones, columnas_meses)

        json_string = df_renovaciones.replace(
            {np.nan: None}).to_json(orient='records')
        condiciones_uso_dict['renovaciones_aire_por_hora'] = json.loads(
            json_string)

        hoja_dict['condiciones_de_uso'] = condiciones_uso_dict

        # self.datos_extraidos[sheet_name] = hoja_dict
        return hoja_dict

    def _parsear_hoja_tablas_envolvente(self):
        """
        Parsea la hoja '3. Tablas Envolvente' y extrae todas sus tablas.
        """
        sheet_name = '3. Tablas Envolvente'
        if sheet_name not in self.xl_file_data:
            print(
                f"Advertencia: No se encontró la hoja '{sheet_name}' en el archivo.")
            return

        df = self.xl_file_data[sheet_name]
        hoja_dict = {}

        # ----------------------
        # --- Tabla: Puertas ---
        # ----------------------
        # Extraemos el bloque de la tabla (Filas 12 a 23, Columnas B a K)
        df_puertas = df.iloc[11:23, 1:11].copy()

        # Asignamos nombres de columna
        df_puertas.columns = [
            'nombre', 'abreviatura', 'u_puerta_opaca_w_m2k', 'vidrio',
            'porcentaje_vidrio', 'u_marco_w_m2k', 'porcentaje_marco',
            'u_ponderado', 'u_ponderado_opaco', 'u_vidrio_w_m2k'
        ]

        # Convertimos las columnas numéricas (incluyendo los porcentajes por seguridad)
        columnas_numericas = [
            'u_puerta_opaca_w_m2k', 'porcentaje_vidrio', 'u_marco_w_m2k',
            'porcentaje_marco', 'u_ponderado', 'u_ponderado_opaco', 'u_vidrio_w_m2k'
        ]
        df_puertas = self._convertir_decimales_a_float(
            df_puertas, columnas_numericas)

        # --- Limpieza de Datos ---
        # Define las columnas de entrada y de salida
        trigger_keys = [
            'nombre', 'abreviatura', 'u_puerta_opaca_w_m2k', 'vidrio',
            'porcentaje_vidrio', 'u_marco_w_m2k', 'porcentaje_marco'
        ]
        target_keys = ['u_ponderado', 'u_ponderado_opaco', 'u_vidrio_w_m2k']

        # Crea una máscara booleana: será True para las filas donde TODAS las trigger_keys son NaN
        mask = df_puertas[trigger_keys].isna().all(axis=1)

        # Usa la máscara para asignar NaN a las target_keys en esas filas específicas
        df_puertas.loc[mask, target_keys] = np.nan

        # Limpiamos y convertimos a formato JSON-nativo
        json_string = df_puertas.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['puertas'] = json.loads(json_string)

        # ---------------------
        # --- Tabla Vidrios ---
        # ---------------------
        # Extraemos el bloque de la tabla (Filas 27 a 42, Columnas B a E)
        df_vidrios = df.iloc[26:42, 1:5].copy()

        # Asignamos nombres de columna
        df_vidrios.columns = [
            'nombre', 'abreviatura', 'u_vidrio_w_m2k', 'fs_vidrio'
        ]

        # Convertimos las columnas numéricas
        columnas_numericas_vidrios = ['u_vidrio_w_m2k', 'fs_vidrio']
        df_vidrios = self._convertir_decimales_a_float(
            df_vidrios, columnas_numericas_vidrios)

        # Limpiamos y convertimos a formato JSON-nativo
        json_string_vidrios = df_vidrios.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['vidrios'] = json.loads(json_string_vidrios)

        # ----------------------------
        # --- Tabla Marcos Ventana ---
        # ----------------------------

        # Extraemos el bloque de la tabla (Filas 46 a 57, Columnas B a E)
        df_marcos = df.iloc[45:57, 1:5].copy()

        # Asignamos nombres de columna
        df_marcos.columns = [
            'nombre_tipo_marcos', 'abreviatura', 'ufr_w_m2k', 'fm'
        ]

        # Convertimos las columnas numéricas
        # La función maneja tanto las comas en 'ufr' como los porcentajes en 'fm'
        columnas_numericas_marcos = ['ufr_w_m2k', 'fm']
        df_marcos = self._convertir_decimales_a_float(
            df_marcos, columnas_numericas_marcos)

        # Limpiamos y convertimos a formato JSON-nativo
        json_string_marcos = df_marcos.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['marcos_ventana'] = json.loads(json_string_marcos)

        # ---------------------------------
        # --- Tabla Muros transmitancia ---
        # ---------------------------------

        # Extraemos el bloque de la tabla (Filas 61 a 75, Columnas B a H)
        df_muros = df.iloc[60:75, 1:8].copy()

        # Asignamos nombres de columna
        df_muros.columns = [
            'nombre', 'abreviatura', 'tipologia_materialidad', 'u_w_m2k',
            'espesor_muro_solido_cm', 'espesor_aislante_cm', 'posicion_aislacion'
        ]

        # Convertimos las columnas numéricas
        columnas_numericas_muros = [
            'u_w_m2k', 'espesor_muro_solido_cm', 'espesor_aislante_cm']
        df_muros = self._convertir_decimales_a_float(
            df_muros, columnas_numericas_muros)

        # Limpiamos y convertimos a formato JSON-nativo
        json_string_muros = df_muros.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['muros_transmitancia'] = json.loads(json_string_muros)

        # ----------------------------------
        # --- Tabla Techos transmitancia ---
        # ----------------------------------

        # Extraemos el bloque de la tabla (Filas 79 a 83, Columnas B, C, D, F, G, H)
        df_techos = df.iloc[78:82, [1, 2, 3, 5, 6, 7]].copy()

        # Asignamos nombres de columna
        df_techos.columns = [
            'nombre', 'abreviatura', 'u_w_m2k',
            'espesor_techo_solido_cm', 'espesor_aislante_cm', 'posicion_aislacion'
        ]

        # Convertimos las columnas numéricas
        columnas_numericas_techos = [
            'u_w_m2k', 'espesor_techo_solido_cm', 'espesor_aislante_cm']
        df_techos = self._convertir_decimales_a_float(
            df_techos, columnas_numericas_techos)

        # Limpiamos y convertimos a formato JSON-nativo
        json_string_techos = df_techos.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['techos_transmitancia'] = json.loads(json_string_techos)

        # ---------------------------------
        # --- Tabla Pisos transmitancia ---
        # ---------------------------------

        # Extraemos el bloque de la tabla (Filas 87 a 100, Columnas B a M)
        df_pisos = df.iloc[86:100, 1:13].copy()

        # Asignamos nombres de columna, aplanando los encabezados multinivel
        df_pisos.columns = [
            'nombre', 'abreviatura', 'u_piso_ventilado_w_m2k',
            'aislacion_terreno_lambda_w_mk', 'aislacion_terreno_e_aislante_cm',
            'refuerzo_vert_lambda_w_mk', 'refuerzo_vert_e_aislante_cm', 'refuerzo_vert_d_cm',
            'refuerzo_horiz_lambda_w_mk', 'refuerzo_horiz_e_aislante_cm', 'refuerzo_horiz_d_cm',
            'posicion_aislacion'
        ]

        # Convertimos todas las columnas que deben ser numéricas
        columnas_numericas_pisos = [
            'u_piso_ventilado_w_m2k', 'aislacion_terreno_lambda_w_mk',
            'aislacion_terreno_e_aislante_cm', 'refuerzo_vert_lambda_w_mk',
            'refuerzo_vert_e_aislante_cm', 'refuerzo_vert_d_cm', 'refuerzo_horiz_lambda_w_mk',
            'refuerzo_horiz_e_aislante_cm', 'refuerzo_horiz_d_cm'
        ]
        df_pisos = self._convertir_decimales_a_float(
            df_pisos, columnas_numericas_pisos)

        # Limpiamos y convertimos a formato JSON-nativo
        json_string_pisos = df_pisos.replace(
            {np.nan: None}).to_json(orient='records')
        hoja_dict['pisos_transmitancia'] = json.loads(json_string_pisos)

        return hoja_dict


# ---------------------------------------------------
# --- 03.-PBTD-Datos-de-Equipos-y-Resultados-v2.2 ---
# ---------------------------------------------------

class LectorPBTD03_v2(LectorPBTD01_v2):
    def __init__(self, filepath):
        super().__init__(filepath)

    def _parse_all_sheets(self):
        """
        Orquestador principal para PBTD 03.
        """
        datos_completos = {}

        # 1. CEV-CEVE (Heredado)
        datos_completos['CEV-CEVE'] = self._parsear_hoja_cev_ceve()

        # 2. Resumen (Dashboard completo)
        datos_completos['Resumen'] = self._parsear_hoja_resumen()

        # 3. Resultados (Tabla horaria)
        datos_completos['Resultados'] = self._parsear_hoja_resultados()

        # 4. Anexo Calculos (Pendiente para futuro)
        datos_completos['Anexo Calculos'] = None

        return datos_completos

    def _limpiar_dict_recursivo(self, d):
        """
        Auxiliar para limpiar diccionarios anidados, convirtiendo NaN a None
        y strings numéricos a float (manejando coma decimal).
        """
        if not isinstance(d, dict):
            return d
        
        for k, v in d.items():
            if isinstance(v, dict):
                d[k] = self._limpiar_dict_recursivo(v)
            elif pd.isna(v):
                d[k] = None
            else:
                try:
                    if isinstance(v, str):
                        val_clean = v.replace(',', '.')
                        d[k] = float(val_clean)
                    else:
                        d[k] = float(v)
                except (ValueError, TypeError):
                    d[k] = v
        return d

    def _parsear_hoja_resultados(self):
        """
        Lee la tabla horaria de la hoja 'Resultados'.
        """
        target_name = 'resultados'
        sheet_found = None
        
        for sheet in self.xl_file_data.keys():
            if target_name in sheet.lower().strip():
                sheet_found = sheet
                break
        
        if sheet_found is None:
            print(f"❌ Error: No se encontró ninguna hoja que coincida con '{target_name}'")
            return None

        df = self.xl_file_data[sheet_found]

        # Definir Rangos
        col_start = 2  # Columna C
        col_end = 61   # Columna BI (hasta 61 exclusivo)
        header_row_idx = 4       # Fila 5 Excel
        data_start_idx = 5       # Fila 6 Excel
        data_end_idx = 3124      # Fila 3124 Excel

        # Extraer Encabezados
        try:
            raw_headers = df.iloc[header_row_idx, col_start:col_end].tolist()
            headers = []
            IDX_AE = 30
            IDX_BG = 58
            
            for i, h in enumerate(raw_headers):
                current_global_idx = col_start + i
                if current_global_idx == IDX_AE and (pd.isna(h) or str(h).strip() == ""):
                    headers.append("Desconocido")
                    continue
                
                if pd.isna(h):
                    headers.append(f"col_{len(headers)}")
                else:
                    h_str = str(h).strip().lower()
                    h_str = h_str.replace(' ', '_').replace('.', '').replace('\n', '')
                    headers.append(h_str)
        except IndexError:
            return None

        # Extraer Datos
        try:
            df_tabla = df.iloc[data_start_idx:data_end_idx, col_start:col_end].copy()
            if len(df_tabla.columns) == len(headers):
                df_tabla.columns = headers
            else:
                min_len = min(len(df_tabla.columns), len(headers))
                df_tabla.columns = headers[:min_len] + [f"extra_{i}" for i in range(len(df_tabla.columns) - min_len)]
        except IndexError:
             return None

        # Convertir a números (Excluyendo BG)
        idx_relativo_bg = IDX_BG - col_start
        cols_to_convert = list(df_tabla.columns)
        if idx_relativo_bg < len(headers):
            col_bg_name = headers[idx_relativo_bg]
            if col_bg_name in cols_to_convert:
                cols_to_convert.remove(col_bg_name)
        
        df_tabla = self._convertir_decimales_a_float(df_tabla, cols_to_convert)

        try:
            json_string = df_tabla.replace({np.nan: None}).to_json(orient='records')
            return json.loads(json_string)
        except Exception:
            return None

    def _parsear_hoja_resumen(self):
        """
        Lee la hoja 'Resumen' completa.
        Incluye: Demanda, Confort, Consumos, Tablas Mensuales, Flujos.
        """
        sheet_name = None
        for nombre in ['Resumen', 'Resultados', 'Resumen Resultados']:
            if nombre in self.xl_file_data:
                sheet_name = nombre
                break
        if sheet_name is None:
            for hoja in self.xl_file_data.keys():
                if "resumen" in hoja.lower():
                    sheet_name = hoja
                    break
        
        if sheet_name is None:
            print("Aviso: No se encontró la hoja 'Resumen'.")
            return None

        df = self.xl_file_data[sheet_name]

        # --- Estructuras Base ---
        demanda_energetica = {}
        confort_termico = {}
        consumos = {}

        # =================================================================
        # 1. DEMANDA ENERGÉTICA
        # =================================================================
        try:
            cols_indices = range(1, 9) 
            nombres_cols = []
            for col_idx in cols_indices:
                partes = []
                for row_idx in [3, 4, 5]:
                    val = df.iat[row_idx, col_idx]
                    if pd.notna(val): partes.append(str(val).strip())
                nombre_limpio = "_".join(partes).lower()
                nombre_limpio = (nombre_limpio.replace(' ', '_').replace('.', '').replace('[', '').replace(']', '')
                                 .replace('-', '_').replace('/', '_por_').replace('%', 'porc').replace('__', '_'))
                nombres_cols.append(nombre_limpio)

            datos_base = {}
            datos_propuesto = {}
            for i, nombre in enumerate(nombres_cols):
                col_abs = cols_indices[i]
                datos_base[nombre] = df.iat[6, col_abs]
                datos_propuesto[nombre] = df.iat[7, col_abs]

            val_ahorro = df.iat[6, 9] if pd.notna(df.iat[6, 9]) else df.iat[7, 9]
            val_letra = df.iat[6, 10] if pd.notna(df.iat[6, 10]) else df.iat[7, 10]

            demanda_energetica = {
                'tabla_demanda': {
                    'caso_base': self._limpiar_dict_nan(datos_base),
                    'caso_propuesto': self._limpiar_dict_nan(datos_propuesto)
                },
                'comparativa_casos': {
                    'ahorro_total_porc': val_ahorro,
                    'letra_calificacion': val_letra
                }
            }
        except Exception: pass

        # =================================================================
        # 2. CONFORT TÉRMICO
        # =================================================================
        try:
            cols_indices = range(1, 6)
            nombres_cols = []
            for col_idx in cols_indices:
                partes = []
                for row_idx in [10, 11, 12]: 
                    val = df.iat[row_idx, col_idx]
                    if pd.notna(val): partes.append(str(val).strip())
                nombre_limpio = "_".join(partes).lower()
                nombre_limpio = (nombre_limpio.replace(' ', '_').replace('.', '').replace('[', '').replace(']', '')
                                 .replace('(', '').replace(')', '').replace('+', '_mas').replace('-', '_menos')
                                 .replace('%', 'porc').replace('__', '_'))
                nombres_cols.append(nombre_limpio)

            datos_base = {}
            datos_propuesto = {}
            for i, nombre in enumerate(nombres_cols):
                col_abs = cols_indices[i]
                datos_base[nombre] = df.iat[13, col_abs]
                datos_propuesto[nombre] = df.iat[14, col_abs]

            confort_termico = {
                'caso_base': self._limpiar_dict_nan(datos_base),
                'caso_propuesto': self._limpiar_dict_nan(datos_propuesto)
            }
        except Exception: pass

        # =================================================================
        # 3. CONSUMOS
        # =================================================================
        try:
            def _leer_par(fila_idx):
                try:
                    return {'kwh_ano': df.iat[fila_idx, 4], 'kwh_m2_ano': df.iat[fila_idx, 6]}
                except IndexError: return {'kwh_ano': None, 'kwh_m2_ano': None}

            solar_termica = {'aporte_calefaccion': _leer_par(22), 'aporte_acs': _leer_par(23)}
            energia_primaria = {'calefaccion': _leer_par(26), 'acs': _leer_par(27), 'iluminacion': _leer_par(28), 'ventiladores': _leer_par(29)}
            generacion_pv = {'generacion_total': _leer_par(31), 'aporte_consumos_basicos': _leer_par(32), 'aporte_consumos_electrodomesticos_o_red': _leer_par(33)}
            balance = {'consumo_total_antes_pv': _leer_par(36), 'aporte_pv_consumos_basicos': _leer_par(37), 'consumos_basicos_a_suplir': _leer_par(38),
                       'consumo_total_final': _leer_par(40), 'consumo_referencia': _leer_par(41),
                       'indicadores': {'coeficiente_c': df.iat[42, 4], 'ahorro_total_porc': df.iat[42, 9]}}
            pv_electro = {'aporte_kwh_ano': df.iat[45, 4], 'porcentaje_consumo_medio': df.iat[45, 6]}

            consumos = {
                '1_aporte_solar_termica': self._limpiar_dict_recursivo(solar_termica),
                '2_consumos_energia_primaria_base': self._limpiar_dict_recursivo(energia_primaria),
                '3_generacion_fotovoltaica': self._limpiar_dict_recursivo(generacion_pv),
                '4_balance_general': self._limpiar_dict_recursivo(balance),
                '5_aporte_pv_electrodomesticos': self._limpiar_dict_recursivo(pv_electro)
            }
        except Exception: pass

        # =================================================================
        # 4. TABLAS MENSUALES
        # =================================================================
        def _extraer_tabla_mensual(fila_inicio_excel, fila_fin_excel, limpieza_simple=False):
            tabla = {}
            meses_keys = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
                          'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
            start_idx = fila_inicio_excel - 1
            end_idx = fila_fin_excel 
            for row_idx in range(start_idx, end_idx):
                try:
                    raw_label = df.iat[row_idx, 13]
                    if pd.isna(raw_label):
                        label = f"fila_{row_idx}"
                    else:
                        label = str(raw_label).lower().strip()
                        if limpieza_simple:
                            label = (label.replace(' ', '_').replace('-', '_').replace(':', '')
                                          .replace('(', '').replace(')', '').replace('+', '')
                                          .replace('.', '').replace('ñ', 'n'))
                        else:
                            label = (label.replace(' ', '_').replace('°', '_deg').replace('+', '_mas')
                                          .replace('-', '_menos').replace('.', '').replace('ñ', 'n'))
                        while '__' in label: label = label.replace('__', '_')

                    datos_fila = {}
                    for m_idx, mes in enumerate(meses_keys):
                        col_idx = 14 + m_idx
                        val = df.iat[row_idx, col_idx]
                        datos_fila[mes] = float(val) if pd.notna(val) else None
                    val_total = df.iat[row_idx, 26]
                    datos_fila['anual'] = float(val_total) if pd.notna(val_total) else None
                    tabla[label] = datos_fila
                except IndexError: continue
            return tabla

        tablas_mensuales = {}
        try:
            tablas_mensuales['1_demanda_calefaccion_comparativa'] = _extraer_tabla_mensual(7, 8, True)
            tablas_mensuales['2_demanda_refrigeracion_comparativa'] = _extraer_tabla_mensual(9, 10, True)
            tablas_mensuales['3_hd_mas_comparativa'] = _extraer_tabla_mensual(12, 13, True)
            tablas_mensuales['4_hd_menos_comparativa'] = _extraer_tabla_mensual(14, 15, True)
            tablas_mensuales['5_demanda_calefaccion_escenarios'] = _extraer_tabla_mensual(19, 23, False)
            tablas_mensuales['6_demanda_refrigeracion_escenarios'] = _extraer_tabla_mensual(24, 28, False)
            tablas_mensuales['7_hd_menos_escenarios'] = _extraer_tabla_mensual(30, 34, False)
            tablas_mensuales['8_hd_mas_escenarios'] = _extraer_tabla_mensual(35, 39, False)
        except Exception: pass

        # =================================================================
        # 5. FLUJOS
        # =================================================================
        def _extraer_tabla_flujos(col_idx_start, col_idx_end, row_idx_start, row_idx_end, tipo_fila='meses'):
            tabla = {}
            try:
                headers = []
                for c in range(col_idx_start, col_idx_end):
                    val = df.iat[2, c]
                    if pd.notna(val):
                        h_str = str(val).strip().lower()
                        h_str = (h_str.replace(' ', '_').replace('.', '')
                                      .replace('[', '').replace(']', '')
                                      .replace('(', '').replace(')', ''))
                        headers.append(h_str)
                    else:
                        headers.append(f"col_{c}")

                meses_keys = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
                              'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
                count = 0
                for r in range(row_idx_start, row_idx_end):
                    datos_fila = {}
                    for i, header in enumerate(headers):
                        val = df.iat[r, col_idx_start + i]
                        if pd.isna(val):
                            datos_fila[header] = None
                        else:
                            try:
                                if isinstance(val, str): val = val.replace(',', '.')
                                datos_fila[header] = float(val)
                            except (ValueError, TypeError):
                                datos_fila[header] = val
                    
                    if tipo_fila == 'meses':
                        key = meses_keys[count] if count < len(meses_keys) else f"fila_{count+1}"
                    else:
                        key = f"hora_{count}"
                    tabla[key] = datos_fila
                    count += 1
            except Exception: return {}
            return tabla

        flujos = {}
        try:
            flujos['1_promedio_mensual_climatizacion'] = _extraer_tabla_flujos(57, 73, 3, 15, 'meses')
            flujos['2_diario_enero'] = _extraer_tabla_flujos(74, 90, 3, 27, 'horas')
            flujos['3_diario_julio'] = _extraer_tabla_flujos(91, 107, 3, 27, 'horas')
        except Exception: pass

        # RETORNO FINAL
        return {
            'demanda_energetica': demanda_energetica,
            'confort_termico': self._limpiar_dict_recursivo(confort_termico),
            'consumos': consumos,
            'tablas_mensuales': self._limpiar_dict_recursivo(tablas_mensuales),
            'flujos': self._limpiar_dict_recursivo(flujos)
        }
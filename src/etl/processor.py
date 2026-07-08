# src/etl/processor.py
"""
Módulo ETL para procesamiento de archivos PLE Ventas
- Usa POSICIONES FIJAS de columnas (estándar PLE)
- Columna G (índice 6) = TipoDoc
- Columna H (índice 7) = CodigoEstablecimiento
- Columna E (índice 4) = FechaEmision
- Columna I (índice 8) = NumeroCorrelativo
- Columna T (índice 19) = MontoOtrosConceptos
- Columna C (índice 2) = IDComprobante
- Columna D (índice 3) = Serie
"""

import pandas as pd
import io
import time

# ============================================================================
# CONFIGURACIÓN DE COLUMNAS POR POSICIÓN (0-index)
# ============================================================================

# Posiciones fijas estándar PLE VENTAS (0-index)
POSICION_PERIODO = 1   
POSICION_TIPO_DOC = 6        # Columna G (7ma columna)
POSICION_CODIGO_ESTABLECIMIENTO = 7  # Columna H (8va columna)
POSICION_FECHA_EMISION = 4   # Columna E (5ta columna)
POSICION_NUMERO_CORRELATIVO = 8  # Columna I (9na columna)
POSICION_MONTO_OTROS = 19    # Columna T (20va columna)
POSICION_ID_COMPROBANTE = 2  # Columna C (3ra columna)
POSICION_SERIE = 3           # Columna D (4ta columna)

# Nombres de columnas para el DataFrame (se asignan por posición)
COLUMNAS_PLE = [
    'col_0', 'Periodo', 'IDComprobante', 'Serie', 'FechaEmision',  # 0-4
    'col_5', 'TipoDoc', 'CodigoEstablecimiento', 'NumeroCorrelativo', 'col_9',  # 5-9
    'col_10', 'col_11', 'col_12', 'col_13', 'col_14',  # 10-14
    'col_15', 'col_16', 'col_17', 'col_18', 'MontoOtrosConceptos',  # 15-19
    'col_20', 'col_21', 'col_22', 'col_23', 'col_24',  # 20-24
    'col_25', 'col_26', 'col_27', 'col_28', 'col_29',  # 25-29
    'col_30', 'col_31', 'col_32', 'col_33', 'col_34', 'col_35'  # 30-35
]

# ============================================================================
# REGLAS DE NEGOCIO
# ============================================================================

PREFIX_RULES = {
    "01": "F",
    "03": "B"
}

# ============================================================================
# SERIES DE MONEDA EXTRANJERA (EXCLUIDAS)
# ============================================================================

SERIES_MONEDA_EXTRANJERA = ['B015', 'B041']

def es_serie_moneda_extranjera(codigo_establecimiento: str) -> bool:
    """Determina si un código de establecimiento es de moneda extranjera."""
    if pd.isna(codigo_establecimiento):
        return False
    codigo_str = str(codigo_establecimiento).strip().upper()
    return codigo_str in SERIES_MONEDA_EXTRANJERA

def filtrar_boletas_validas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtra las boletas válidas para el procesamiento.
    Excluye series de moneda extranjera (B015, B041)
    """
    # Filtrar solo tipo 03
    df_filtrado = df[df['TipoDoc'] == '03'].copy()
    
    # Excluir series de moneda extranjera
    mask_excluir = df_filtrado['CodigoEstablecimiento'].apply(es_serie_moneda_extranjera)
    df_filtrado = df_filtrado[~mask_excluir].copy()
    
    return df_filtrado

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def normalizar_tipo_doc(valor):
    """Convierte 1 -> 01, 3 -> 03."""
    try:
        return str(int(valor)).zfill(2)
    except:
        return str(valor)

def normalizar_codigo(tipo_doc, codigo_original):
    """Normaliza el código de establecimiento añadiendo prefijos si no existen."""
    if tipo_doc not in PREFIX_RULES:
        return str(codigo_original)

    prefijo_necesario = PREFIX_RULES[tipo_doc]
    codigo_str = str(codigo_original).strip()

    if codigo_str.startswith(prefijo_necesario):
        return codigo_str
    
    return prefijo_necesario + codigo_str

# ============================================================================
# LECTURA DE EXCEL CON POSICIONES FIJAS (CORREGIDO)
# ============================================================================

def leer_excel_por_posiciones(file_stream, hoja_nombre, fila_inicio=0):
    """
    Lee un Excel usando posiciones fijas (no nombres de columnas).
    
    Args:
        file_stream: Archivo Excel
        hoja_nombre: Nombre de la hoja
        fila_inicio: Fila donde empiezan los datos (0-index)
    
    Returns:
        pd.DataFrame: DataFrame con columnas nombradas según estándar PLE
    """
    # Leer sin encabezados (header=None)
    df_raw = pd.read_excel(
        file_stream, 
        sheet_name=hoja_nombre, 
        header=None,
        dtype=str,
        engine='openpyxl'
    )
    
    if df_raw.empty:
        print(f"⚠️ Hoja '{hoja_nombre}': Archivo vacío")
        return pd.DataFrame()
    
    # Asegurar que tenemos suficientes columnas
    max_col = max(POSICION_ID_COMPROBANTE, POSICION_SERIE, POSICION_FECHA_EMISION,
                  POSICION_TIPO_DOC, POSICION_CODIGO_ESTABLECIMIENTO, 
                  POSICION_NUMERO_CORRELATIVO, POSICION_MONTO_OTROS,
                  POSICION_PERIODO)
    
    if len(df_raw.columns) <= max_col:
        print(f"⚠️ Hoja '{hoja_nombre}': Solo {len(df_raw.columns)} columnas, se necesitan {max_col + 1}")
        return pd.DataFrame()
    
    # Filtrar filas vacías
    df_raw = df_raw.dropna(how='all')
    
    if df_raw.empty:
        print(f"⚠️ Hoja '{hoja_nombre}': Sin datos después de limpiar vacíos")
        return pd.DataFrame()
    
    # Saltar filas si es necesario
    if fila_inicio > 0:
        df_raw = df_raw.iloc[fila_inicio:]
    
    # Resetear índices
    df_raw = df_raw.reset_index(drop=True)
    
    # ✅ CREAR UN NUEVO DATAFRAME CON LAS COLUMNAS NECESARIAS
    # Mapeo de posiciones a nombres
    mapeo_posiciones = {
        POSICION_PERIODO: 'Periodo',
        POSICION_ID_COMPROBANTE: 'IDComprobante',
        POSICION_SERIE: 'Serie',
        POSICION_FECHA_EMISION: 'FechaEmision',
        POSICION_TIPO_DOC: 'TipoDoc',
        POSICION_CODIGO_ESTABLECIMIENTO: 'CodigoEstablecimiento',
        POSICION_NUMERO_CORRELATIVO: 'NumeroCorrelativo',
        POSICION_MONTO_OTROS: 'MontoOtrosConceptos'
    }
    
    # Crear diccionario con los datos
    datos_columnas = {}
    for pos, nombre in mapeo_posiciones.items():
        if pos < len(df_raw.columns):
            datos_columnas[nombre] = df_raw[pos]
        else:
            datos_columnas[nombre] = ''  # Columna vacía si no existe
    
    # Crear el DataFrame
    df_resultado = pd.DataFrame(datos_columnas)
    
    # 🔍 DEBUG: Mostrar información
    print(f"\n🔍 DEBUG - Hoja '{hoja_nombre}':")
    print(f"  ✅ Filas leídas: {len(df_raw)}")
    print(f"  ✅ DataFrame creado: {len(df_resultado)} filas, {len(df_resultado.columns)} columnas")
    print(f"  📊 Columnas: {list(df_resultado.columns)}")
    
    # Mostrar el Periodo si existe
    if 'Periodo' in df_resultado.columns and not df_resultado.empty:
        print(f"  📌 Periodo (columna B): '{df_resultado['Periodo'].iloc[0]}'")
    
    # Mostrar primeras filas
    print(f"\n  📋 Primeras 3 filas:")
    print(df_resultado.head(3).to_string())
    print("-"*60)
    
    return df_resultado
# ============================================================================
# MÓDULO 2: VALIDACIÓN DOCUMENTAL
# ============================================================================

def procesar_excel(file_stream, progress_callback=None):
    """
    Orquestador principal: Extrae, Normaliza y Detecta Duplicados.
    Usa POSICIONES FIJAS para las columnas.
    """
    start_time = time.perf_counter()
    try:
        excel_file = pd.ExcelFile(file_stream)
        hojas = excel_file.sheet_names
        total_hojas = len(hojas)
        
        # 🔥 DEBUG: Mostrar información del archivo
        print("\n" + "="*60)
        print("🔍 VALIDADOR DOCUMENTAL - DIAGNÓSTICO")
        print("="*60)
        print(f"📄 Archivo: {file_stream.name if hasattr(file_stream, 'name') else 'archivo_subido'}")
        print(f"📋 Hojas encontradas: {hojas}")
        print("="*60)
        
        df_limpio = pd.DataFrame()
        lista_errores = []
        
        total_filas = 0
        duplicados_encontrados = 0

        for idx, hoja in enumerate(hojas):
            print(f"\n🔎 Procesando hoja: '{hoja}' ({(idx+1)}/{total_hojas})")
            print("-"*40)
            
            try:
                # 🔥 USAR POSICIONES FIJAS
                df = leer_excel_por_posiciones(excel_file, hoja, fila_inicio=0)
                
                if df.empty:
                    lista_errores.append(f"Hoja '{hoja}' está vacía o no tiene suficientes columnas.")
                    print(f"  ⚠️ Hoja vacía - SALTANDO")
                    continue

                # Mostrar primeras filas
                print(f"  📋 Primeras 3 filas:")
                print(df.head(3).to_string())
                print("-"*40)

                # Ajustar índices de fila para auditoría
                df['_ExcelRow'] = df.index + 2 

                # Normalización de Tipo de Documento
                if 'TipoDoc' in df.columns:
                    df['TipoDoc_Norm'] = df['TipoDoc'].apply(normalizar_tipo_doc)
                    print(f"  📌 Tipos de documento encontrados: {df['TipoDoc_Norm'].unique().tolist()}")
                else:
                    lista_errores.append(f"Hoja '{hoja}': Falta columna 'TipoDoc'")
                    print(f"  ❌ Falta columna 'TipoDoc'")
                    continue

                # Normalización de Código de Establecimiento
                if 'CodigoEstablecimiento' in df.columns:
                    df['CodigoEstablecimiento_Norm'] = df.apply(
                        lambda x: normalizar_codigo(x['TipoDoc_Norm'], x['CodigoEstablecimiento']), axis=1
                    )
                    print(f"  📌 Códigos encontrados: {df['CodigoEstablecimiento_Norm'].unique().tolist()[:5]}...")
                else:
                    lista_errores.append(f"Hoja '{hoja}': Falta columna 'CodigoEstablecimiento'")
                    print(f"  ❌ Falta columna 'CodigoEstablecimiento'")
                    continue

                df['Hoja_Nombre'] = hoja
                df_limpio = pd.concat([df_limpio, df], ignore_index=True)

                if progress_callback:
                    progress_callback((idx + 1) / max(total_hojas, 1), f"procesando hoja {idx + 1} de {total_hojas}: {hoja}")

            except Exception as e:
                print(f"  ❌ Error en hoja '{hoja}': {str(e)}")
                lista_errores.append(f"Hoja '{hoja}': Error {str(e)}")
                continue

        # Detección de duplicados para documentos 01 y 03
        df_solo_validos = df_limpio[df_limpio['TipoDoc_Norm'].isin(['01', '03'])].copy()

        if 'NumeroCorrelativo' in df_solo_validos.columns:
            df_solo_validos['Clave_Unica'] = (
                df_solo_validos['CodigoEstablecimiento_Norm'].astype(str) + '-' + 
                df_solo_validos['NumeroCorrelativo'].astype(str)
            )

            duplicated_mask = df_solo_validos.duplicated(subset=['Clave_Unica'], keep=False)
            df_duplicados = df_solo_validos[duplicated_mask].copy()
            
            df_trazabilidad = df_duplicados[['Hoja_Nombre', '_ExcelRow', 'TipoDoc_Norm', 'CodigoEstablecimiento_Norm', 'NumeroCorrelativo', 'Clave_Unica']].sort_values(by='Clave_Unica')
            
            duplicados_encontrados = len(df_duplicados)
        else:
            df_trazabilidad = pd.DataFrame()
            lista_errores.append("Falta columna 'NumeroCorrelativo' para generar claves.")

        total_filas = len(df_limpio)
        elapsed = time.perf_counter() - start_time
        
        # 🔍 RESUMEN FINAL
        print("\n" + "="*60)
        print("📊 RESUMEN DEL DIAGNÓSTICO")
        print("="*60)
        print(f"  📄 Total hojas: {total_hojas}")
        print(f"  ✅ Registros procesados: {total_filas:,}")
        print(f"  🔍 Duplicados encontrados: {duplicados_encontrados}")
        print("="*60 + "\n")
        
        if progress_callback:
            progress_callback(1.0, f"finalizado en {elapsed:.2f} segundos")

        return {
            "success": True,
            "hojas_procesadas": hojas,
            "total_registros": total_filas,
            "duplicados_count": duplicados_encontrados,
            "df_duplicados": df_trazabilidad,
            "df_limpio": df_limpio, 
            "errores": lista_errores,
            "elapsed_seconds": elapsed
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"Error fatal en el procesamiento: {str(e)}",
            "errores": []
        }

# ============================================================================
# MÓDULO 3: ORDENAR BOLETAS POR DÍA (CON POSICIONES FIJAS)
# ============================================================================

def ordenar_boletas(file_stream, progress_callback=None):
    """
    Procesa un archivo Excel de boletas (PLE Ventas) y genera un resumen con columnas fijas.
    Usa POSICIONES FIJAS para las columnas.
    EXCLUYE series de moneda extranjera: B015, B041
    """
    start_time = time.perf_counter()
    try:
        # =========================================================================
        # 1. DEFINICIÓN DE LAS COLUMNAS FIJAS DE SALIDA
        # =========================================================================
        COLUMNAS_FIJAS = [
            'Tipo', 'Periodo', 'IDComprobante', 'Serie', 'FechaEmision',
            'FechaVencimiento', 'TipoDoc', 'CodigoEstablecimiento', 'NumeroCorrelativo',
            'campo_10', 'TipoDocCliente', 'NumeroDocCliente', 'RazonSocialCliente',
            'MontoOperacionesExoneradas', 'MontoOperacionesGravadas', 'MontoIGV',
            'CodigoCentroCosto', 'MontoIGVRetenido', 'MontoReversiones',
            'MontoOtrosConceptos', 'MontoPercepciones', 'MontoDetraccion',
            'MontoTotal', 'campo_23', 'campo_24', 'campo_25', 'campo_26',
            'campo_27', 'campo_28', 'campo_29', 'campo_30', 'campo_31',
            'campo_32', 'campo_33', 'campo_34', 'campo_35'
        ]
        
        # =========================================================================
        # 2. LECTURA Y PROCESAMIENTO HOJA POR HOJA
        # =========================================================================
        excel_file = pd.ExcelFile(file_stream, engine='openpyxl')
        hojas = excel_file.sheet_names
        total_hojas = len(hojas)
        
        # 🔍 DEBUG
        print("\n" + "="*60)
        print("🔍 ORDENAR BOLETAS - DIAGNÓSTICO")
        print("="*60)
        print(f"📄 Archivo: {file_stream.name if hasattr(file_stream, 'name') else 'archivo_subido'}")
        print(f"📋 Hojas encontradas: {hojas}")
        print("="*60)
        
        if progress_callback:
            progress_callback(0.05, f'📊 Leyendo {total_hojas} hojas...')
        
        dfs_validos = []
        hojas_procesadas = 0
        hojas_saltadas = 0
        series_excluidas = []
        
        for idx, hoja_nombre in enumerate(hojas):
            print(f"\n🔎 Procesando hoja: '{hoja_nombre}' ({(idx+1)}/{total_hojas})")
            print("-"*40)
            
            try:
                # 🔥 USAR POSICIONES FIJAS - leer la hoja
                df_hoja = leer_excel_por_posiciones(excel_file, hoja_nombre, fila_inicio=0)
                
                if df_hoja.empty:
                    print(f"  ⚠️ Hoja vacía o sin columnas suficientes - SALTANDO")
                    hojas_saltadas += 1
                    continue
                
                # Verificar columnas críticas
                tiene_tipodoc = 'TipoDoc' in df_hoja.columns
                tiene_codigo = 'CodigoEstablecimiento' in df_hoja.columns
                
                print(f"  🔍 Columnas críticas:")
                print(f"    - 'TipoDoc': {'✅ SI' if tiene_tipodoc else '❌ NO'}")
                print(f"    - 'CodigoEstablecimiento': {'✅ SI' if tiene_codigo else '❌ NO'}")
                print(f"  📊 Total columnas: {len(df_hoja.columns)}")
                
                if not tiene_tipodoc or not tiene_codigo:
                    print(f"  ❌ Faltan columnas críticas - SALTANDO")
                    hojas_saltadas += 1
                    continue
                
                # 🔍 DEBUG: Mostrar valores únicos
                tipos_unicos = df_hoja['TipoDoc'].unique()
                print(f"\n  📌 Valores en 'TipoDoc': {list(tipos_unicos)}")
                
                codigos_unicos = df_hoja['CodigoEstablecimiento'].unique()
                print(f"  📌 Valores en 'CodigoEstablecimiento': {list(codigos_unicos)[:10]}...")
                
                # 🔥 FILTRAR: solo TipoDoc = '03' y excluir B015/B041
                df_hoja_filtrado = filtrar_boletas_validas(df_hoja)
                print(f"\n  🎯 Después de filtrar: {len(df_hoja_filtrado)} filas")
                
                # Registrar series excluidas
                mask_excluidas = df_hoja['CodigoEstablecimiento'].apply(es_serie_moneda_extranjera)
                if mask_excluidas.any():
                    excluidas = df_hoja[mask_excluidas]['CodigoEstablecimiento'].unique().tolist()
                    series_excluidas.extend(excluidas)
                    print(f"  🚫 Series excluidas (moneda extranjera): {excluidas}")
                
                if df_hoja_filtrado.empty:
                    print(f"  ⚠️ No hay datos válidos después del filtro - SALTANDO")
                    hojas_saltadas += 1
                    continue
                
                dfs_validos.append(df_hoja_filtrado)
                hojas_procesadas += 1
                
                if progress_callback:
                    progress_callback(0.05 + 0.05 * ((idx + 1) / total_hojas), 
                                    f'Hoja {idx+1}/{total_hojas}: "{hoja_nombre}" - {len(df_hoja_filtrado):,} filas válidas')
                
            except Exception as e:
                print(f"  ❌ Error en hoja '{hoja_nombre}': {str(e)}")
                hojas_saltadas += 1
                continue
        
        # 🔍 RESUMEN FINAL
        print("\n" + "="*60)
        print("📊 RESUMEN DEL DIAGNÓSTICO")
        print("="*60)
        print(f"  📄 Total hojas: {total_hojas}")
        print(f"  ✅ Hojas procesadas: {hojas_procesadas}")
        print(f"  ⚠️ Hojas saltadas: {hojas_saltadas}")
        print(f"  📊 DataFrames válidos: {len(dfs_validos)}")
        if series_excluidas:
            print(f"  🚫 Series excluidas: {set(series_excluidas)}")
        print("="*60 + "\n")
        
        # Verificar si se encontraron datos válidos
        if not dfs_validos:
            mensaje = f'No se encontraron datos válidos en ninguna de las {total_hojas} hojas.'
            if series_excluidas:
                mensaje += f' Se excluyeron las series de moneda extranjera: {", ".join(set(series_excluidas))}'
            return {
                'success': False, 
                'error': mensaje
            }
        
        # =========================================================================
        # 3. CONTINUAR CON EL PROCESAMIENTO NORMAL
        # =========================================================================
        df = pd.concat(dfs_validos, ignore_index=True)
        
        if progress_callback:
            progress_callback(0.12, f'✅ Hojas procesadas: {hojas_procesadas} | Hojas omitidas: {hojas_saltadas}')
            if series_excluidas:
                progress_callback(0.12, f'⚠️ Series excluidas: {", ".join(set(series_excluidas))}')

        # Guardar el número original de filas
        original_rows = len(df)
        
        if progress_callback:
            progress_callback(0.15, f'Filtrando datos (total original: {original_rows:,} filas)...')
        
        # Verificar columnas requeridas
        columnas_requeridas = ['FechaEmision', 'NumeroCorrelativo', 'MontoOtrosConceptos', 'IDComprobante', 'Serie']
        missing = [c for c in columnas_requeridas if c not in df.columns]
        if missing:
            return {
                'success': False,
                'error': f"El archivo no tiene las columnas requeridas: {', '.join(missing)}"
            }
        
        # Convertir columnas numéricas
        if progress_callback:
            progress_callback(0.20, 'Convirtiendo columnas numéricas...')

        df['NumeroCorrelativo'] = pd.to_numeric(df['NumeroCorrelativo'], errors='coerce')
        df['MontoOtrosConceptos'] = pd.to_numeric(df['MontoOtrosConceptos'], errors='coerce').fillna(0)
        
        # Agrupar
        if progress_callback:
            progress_callback(0.25, 'Agrupando por FechaEmision y CodigoEstablecimiento...')
            
        df_sorted = df.sort_values(['FechaEmision', 'CodigoEstablecimiento', 'NumeroCorrelativo'])
        grupos = df_sorted.groupby(['FechaEmision', 'CodigoEstablecimiento'])
        total_grupos = grupos.ngroups if hasattr(grupos, 'ngroups') else 1
        
        output_rows = []
        id_counter = 1
        serie_counter = 1
        
        if progress_callback:
            progress_callback(0.30, f'Generando {total_grupos} grupos...')
            
        for idx, ((fecha, establecimiento), grupo) in enumerate(grupos):
            primera = grupo.iloc[0]

            # 🔥 OBTENER EL PERIODO DE LA COLUMNA B (posición 1)
            periodo = primera['Periodo'] if 'Periodo' in primera else ''
            
            # 🔍 DEBUG: Verificar que el Periodo se está tomando
            print(f"  📌 Periodo del grupo: '{periodo}'")
            
            primer_correlativo = grupo['NumeroCorrelativo'].min()
            ultimo_correlativo = int(grupo['NumeroCorrelativo'].max())
            suma_total = grupo['MontoOtrosConceptos'].sum()
            
            # Generar IDComprobante
            id_original = str(primera['IDComprobante'])
            if '-' in id_original:
                id_parts = id_original.split('-')
                nuevo_id = f"{id_parts[0]}-{id_parts[1][:4]}"
            else:
                nuevo_id = id_original[:7]

            # Generar Serie
            nueva_serie = f"M123"
            serie_counter += 1
            
            # Construir diccionario
            row_dict = {}
            # 🔥 OBTENER EL PERIODO DE LA COLUMNA B (antes del bucle)
            periodo = primera['Periodo'] if 'Periodo' in primera else ''

            for col in COLUMNAS_FIJAS:
                if col == 'Tipo':
                    row_dict[col] = " "
                elif col == 'Periodo':
                    row_dict[col] = periodo  # ✅ Asignación específica
                elif col == 'IDComprobante':
                    row_dict[col] = nuevo_id
                elif col == 'Serie':
                    row_dict[col] = nueva_serie
                elif col == 'NumeroCorrelativo':
                    row_dict[col] = primer_correlativo
                elif col == 'campo_10':
                    row_dict[col] = ultimo_correlativo
                elif col == 'MontoOtrosConceptos':
                    row_dict[col] = suma_total
                elif col == 'campo_25':
                    row_dict[col] = suma_total
                elif col == 'campo_35':
                    row_dict[col] = 1
                elif col in ['campo_23', 'campo_24']:
                    row_dict[col] = ''
                elif col == 'NumeroDocCliente':
                    row_dict[col] = ''
                elif col == 'RazonSocialCliente':
                    row_dict[col] = ''
                else:
                    row_dict[col] = primera[col] if col in primera else ''
            output_rows.append(row_dict)
            id_counter += 1
            serie_counter += 1
            
            if progress_callback:
                progress_callback(0.20 + 0.70 * ((idx + 1) / max(total_grupos, 1)), 
                                f"Agrupando registro {idx + 1} de {total_grupos}")
        
        # Crear DataFrame de salida
        if progress_callback:
            progress_callback(0.95, 'Creando DataFrame de salida...')
        
        df_out = pd.DataFrame(output_rows, columns=COLUMNAS_FIJAS)
        
        # Escribir Excel
        if progress_callback:
            progress_callback(0.98, 'Escribiendo archivo Excel de salida...')
        
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            for establecimiento, grupo_out in sorted(df_out.groupby('CodigoEstablecimiento')):
                sheet_name = str(establecimiento)[:31]
                grupo_out.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Aplicar formato
                worksheet = writer.sheets[sheet_name]
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                font = Font(bold=True, color="FFFFFF", size=11)
                border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                for cell in worksheet[1]:
                    if cell.value:
                        cell.fill = fill
                        cell.font = font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = border
                
                # Ajustar columnas
                for col in worksheet.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value:
                                max_len = max(max_len, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = min(max_len + 2, 30)
                    worksheet.column_dimensions[col_letter].width = adjusted_width
        
        output_buffer.seek(0)
        elapsed = time.perf_counter() - start_time
        elapsed_minutes = elapsed / 60
        
        if progress_callback:
            progress_callback(1.0, f"finalizado en ({elapsed_minutes:.2f} minutos)")
        
        mensaje_series = ""
        if series_excluidas:
            mensaje_series = f" (Series excluidas: {', '.join(set(series_excluidas))})"
        
        return {
            'success': True,
            'message': f'Procesamiento exitoso ✨ en ({elapsed_minutes:.2f} minutos){mensaje_series}',
            'buffer': output_buffer,
            'sheets': list(df_out['CodigoEstablecimiento'].unique()),
            'total_rows': len(df_out),
            'elapsed_seconds': elapsed,
            'original_rows': original_rows,
            'filtered_rows': len(df),
            'reduction_percent': (1 - len(df) / original_rows) * 100 if original_rows > 0 else 0,
            'series_excluidas': list(set(series_excluidas)) if series_excluidas else []
        }
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        elapsed = time.perf_counter() - start_time
        return {
            'success': False,
            'error': f"Error en procesamiento 🔴: {str(e)}",
            'elapsed_seconds': elapsed
        }
# src/conciliador_.py
"""
Módulo de Conciliación para PLE VENTAS
- Optimizado para archivos masivos (500k+ registros)
- Uso de sets para comparación O(1)
- Genera reportes SIRE_BN y SIRE_SUNAT
"""

import pandas as pd
import numpy as np
import re
from typing import Tuple, Dict, List, Optional
from datetime import datetime
import os
import gc
import time
import xlsxwriter

# ============================================================================
# FUNCIÓN AUXILIAR PARA MANEJAR LÍMITES DE EXCEL
# ============================================================================

def escribir_datos_con_limite(workbook, df, nombre_base, header_format, 
                               titulo_seccion, color_fondo, color_texto='white'):
    """
    Escribe un DataFrame en una o múltiples hojas de Excel si excede el límite de filas.
    Límite de Excel: 1,048,576 filas (reservamos 100 para encabezados y formato)
    """
    LIMITE_FILAS = 1_048_000
    
    # --- CORRECCIÓN: Asegurar que titulo_seccion sea string ---
    if not isinstance(titulo_seccion, str):
        titulo_seccion = str(titulo_seccion)
    
    if df.empty:
        sheet_name = f"{nombre_base}"[:31]
        ws = workbook.add_worksheet(sheet_name)
        ws.merge_range(0, 0, 0, 6, titulo_seccion, 
                       workbook.add_format({'bold': True, 'font_color': color_texto, 
                                            'bg_color': color_fondo, 'align': 'center'}))
        ws.write(1, 0, 'No hay registros')
        return
    
    # Reordenar columnas
    cols_priority = ['ID_CONCILIACION', 'Serie_Norm', 'Numero_Norm', 'Fecha_Emision', 'Hoja_Origen']
    cols_existentes = [c for c in cols_priority if c in df.columns]
    cols_resto = [c for c in df.columns if c not in cols_existentes]
    df_ordenado = df[cols_existentes + cols_resto]
    
    total_filas = len(df_ordenado)
    num_hojas = (total_filas // LIMITE_FILAS) + 1 if total_filas > LIMITE_FILAS else 1
    
    if num_hojas == 1:
        sheet_name = f"{nombre_base}"[:31]
        ws = workbook.add_worksheet(sheet_name)
        _escribir_datos_en_hoja(ws, df_ordenado, workbook, header_format, titulo_seccion, 
                                color_fondo, color_texto)
    else:
        for parte in range(num_hojas):
            inicio = parte * LIMITE_FILAS
            fin = min((parte + 1) * LIMITE_FILAS, total_filas)
            df_parte = df_ordenado.iloc[inicio:fin]
            
            if num_hojas == 1:
                sheet_name = f"{nombre_base}"[:31]
            else:
                sheet_name = f"{nombre_base}_parte{parte + 1}"[:31]
            
            ws = workbook.add_worksheet(sheet_name)
            _escribir_datos_en_hoja(ws, df_parte, workbook, header_format, 
                                    f"{titulo_seccion} (Parte {parte + 1}/{num_hojas})",
                                    color_fondo, color_texto)


def _escribir_datos_en_hoja(ws, df, workbook, header_format, titulo, color_fondo, color_texto):
    """
    Escribe los datos en una hoja de Excel con encabezados.
    """
    row = 0
    
    # --- CORRECCIÓN: Asegurar que titulo sea string ---
    if not isinstance(titulo, str):
        titulo = str(titulo)
    
    # Título de la sección con el color correspondiente
    ws.merge_range(row, 0, row, 6, titulo,
                   workbook.add_format({'bold': True, 'font_color': color_texto,
                                        'bg_color': color_fondo, 'align': 'center'}))
    row += 1
    
    # Escribir encabezados
    for col_idx, col_name in enumerate(df.columns):
        ws.write(row, col_idx, col_name, header_format)
    row += 1
    
    # Escribir datos - Limpiar valores para evitar errores de tipo
    for _, data_row in df.iterrows():
        for col_idx, value in enumerate(data_row):
            # Limpiar el valor para evitar errores de tipo
            if value is None or pd.isna(value):
                clean_value = ''
            elif isinstance(value, (int, float)):
                clean_value = value
            else:
                clean_value = str(value)
            ws.write(row, col_idx, clean_value)
        row += 1

# ============================================================================
# CONFIGURACIÓN DE COLUMNAS FIJAS PLE VENTAS
# ============================================================================

# Posiciones fijas estándar PLE VENTAS (0-index)
# IMPORTANTE: Serie está en columna I (índice 8), Número en columna J (índice 9)
COLUMNA_SERIE = 8       # Columna I (9na columna, índice 8)
COLUMNA_NUMERO = 9      # Columna J (10ma columna, índice 9)
COLUMNA_TIPO_DOC = 7    # Columna H (8Va columna, índice 7)
COLUMNA_FECHA = 5       # Columna F (6ta columna, índice 5)

# Tipos de documento válidos para ventas
TIPOS_DOC_VENTAS = [
    '01',  # Factura
    '03',  # Boleta
    '07',  # Nota de Crédito
    '08',  # Nota de Débito
    '02',  # Nota de débito (alternativo)
    '04',  # Nota de crédito (alternativo)
    '05',  # Pasajes aéreos
    '06',  # Comprobante de retención
    '09',  # Pagos exterior
    '12',  # Ticket máquina registradora
    '14',  # Recibo de servicio público
    '16',  # Boleto de viaje / Transporte de carga
    '30',  # Documentos autorizados
    '42',  # Otros
    '46',  # No domiciliado
    '53',  # Declaración de mensajería o courier
    '87'   # Nota de crédito especial
]

# Diccionario de nombres de tipos de documento
TIPOS_DOC_NOMBRES = {
    '00': 'Sin tipo / Otros',
    '01': 'Factura',
    '02': 'Nota de débito',
    '03': 'Boleta',
    '04': 'Nota de crédito',
    '05': 'Pasajes aéreos',
    '06': 'Comprobante de retención',
    '07': 'Nota de crédito',
    '08': 'Nota de débito',
    '09': 'Pagos exterior',
    '12': 'Ticket máquina registradora',
    '14': 'Recibo de servicio público',
    '16': 'Boleto de viaje / Transporte de carga',
    '30': 'Documentos autorizados',
    '42': 'Otros',
    '46': 'No domiciliado',
    '53': 'Declaración de mensajería o courier',
    '87': 'Nota de crédito especial',
    'DESCONOCIDO': 'Tipo desconocido'
}


# ============================================================================
# 1. FUNCIONES DE NORMALIZACIÓN
# ============================================================================

def normalizar_tipo_documento(valor) -> str:
    """Normaliza el tipo de documento a formato de 2 dígitos."""
    if pd.isna(valor) or valor is None:
        return '00'
    
    try:
        valor_str = str(valor).strip()
        valor_str = re.sub(r'\.0+$', '', valor_str)
        digitos = re.sub(r'[^0-9]', '', valor_str)
        
        if len(digitos) >= 2:
            return digitos[:2].zfill(2)
        elif len(digitos) == 1:
            return digitos.zfill(2)
        else:
            return '00'
    except:
        return '00'


def normalizar_numero(valor) -> str:
    """Normaliza el número a 8 dígitos con ceros a la izquierda."""
    if pd.isna(valor) or valor is None:
        return '00000000'
    
    try:
        valor_str = str(valor).strip()
        valor_str = re.sub(r'\.0+$', '', valor_str)
        digitos = re.sub(r'[^0-9]', '', valor_str)
        return digitos.zfill(8) if digitos else '00000000'
    except:
        return '00000000'


def normalizar_serie(valor) -> str:
    """Normaliza el número de serie."""
    if pd.isna(valor) or valor is None:
        return ''
    return str(valor).strip().upper()


def extraer_mes_fecha(fecha_valor) -> Optional[str]:
    """Extrae mes y año de una fecha en formato YYYYMM."""
    if pd.isna(fecha_valor) or fecha_valor is None:
        return None
    
    try:
        if hasattr(fecha_valor, 'strftime'):
            return fecha_valor.strftime('%Y%m')
        
        fecha_str = str(fecha_valor).strip()
        formatos = ['%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y', '%Y%m%d']
        
        for fmt in formatos:
            try:
                fecha_dt = datetime.strptime(fecha_str, fmt)
                return fecha_dt.strftime('%Y%m')
            except:
                continue
        
        try:
            fecha_ts = pd.to_datetime(fecha_str, errors='coerce')
            if not pd.isna(fecha_ts):
                return fecha_ts.strftime('%Y%m')
        except:
            pass
        
        return None
    except:
        return None


# ============================================================================
# 2. LECTURA DE ARCHIVOS
# ============================================================================

def leer_hoja_conciliacion_compras(ruta_archivo: str, 
                                   hoja_nombre: str = None,
                                   fila_inicio: int = 2,
                                   progress_callback=None) -> pd.DataFrame:
    """
    Lee una hoja de Excel.
    fila_inicio = fila donde empiezan los DATOS (1-index)
    Los encabezados deben estar en la fila anterior.
    """
    start_time = time.perf_counter()
    
    try:
        if progress_callback:
            progress_callback(0.1, f"Leyendo hoja: {hoja_nombre or 'primera'}")
        
        # header = fila_inicio - 2
        # Si fila_inicio = 2 (datos empiezan en fila 2), header = 0 (fila 1 como encabezado)
        header_row = max(0, fila_inicio - 2)
        
        df = pd.read_excel(
            ruta_archivo,
            sheet_name=hoja_nombre,
            header=header_row,
            dtype=str,
            engine='openpyxl'
        )
        
        if df.empty:
            return pd.DataFrame()
        
        # Identificar columnas por posición
        columnas = df.columns.tolist()
        max_col = max(COLUMNA_TIPO_DOC, COLUMNA_FECHA, COLUMNA_SERIE, COLUMNA_NUMERO)
        
        if len(columnas) <= max_col:
            print(f"DEBUG: ERROR - Solo hay {len(columnas)} columnas, se necesitan {max_col + 1}")
            return pd.DataFrame()
        
        col_tipo = columnas[COLUMNA_TIPO_DOC]
        col_fecha = columnas[COLUMNA_FECHA]
        col_serie = columnas[COLUMNA_SERIE]
        col_numero = columnas[COLUMNA_NUMERO]
        
        # --- DEBUG: Información de columnas ---
        print(f"DEBUG: Columna Tipo = '{col_tipo}' (índice {COLUMNA_TIPO_DOC})")
        print(f"DEBUG: Columna Serie = '{col_serie}' (índice {COLUMNA_SERIE})")
        print(f"DEBUG: Columna Numero = '{col_numero}' (índice {COLUMNA_NUMERO})")
        
        if progress_callback:
            progress_callback(0.3, f"Normalizando datos ({len(df)} registros)...")
        
        df_procesado = df.copy()
        df_procesado['Tipo_Doc_Norm'] = df_procesado[col_tipo].apply(normalizar_tipo_documento)
        
        # Filtrar solo tipos de documento válidos
        mask_validos = df_procesado['Tipo_Doc_Norm'].isin(TIPOS_DOC_VENTAS)
        df_procesado = df_procesado[mask_validos].copy()
        
        print(f"DEBUG: Registros después de filtrar por tipo: {len(df_procesado)}")
        
        if df_procesado.empty:
            if progress_callback:
                progress_callback(0.5, "No hay registros con tipos válidos")
            return pd.DataFrame()
        
        df_procesado['Serie_Norm'] = df_procesado[col_serie].apply(normalizar_serie)
        df_procesado['Numero_Norm'] = df_procesado[col_numero].apply(normalizar_numero)
        df_procesado['ID_CONCILIACION'] = df_procesado['Serie_Norm'] + df_procesado['Numero_Norm']
        
        df_procesado['Mes'] = df_procesado[col_fecha].apply(extraer_mes_fecha)
        df_procesado['Fecha_Emision'] = df_procesado[col_fecha]
        
        df_procesado = df_procesado[df_procesado['ID_CONCILIACION'] != ''].copy()
        
        print(f"DEBUG: Registros finales con ID válido: {len(df_procesado)}")
        
        elapsed = time.perf_counter() - start_time
        if progress_callback:
            progress_callback(0.5, f"Hoja procesada: {len(df_procesado)} registros")
        
        return df_procesado
        
    except Exception as e:
        print(f"Error leyendo hoja {hoja_nombre}: {str(e)}")
        return pd.DataFrame()


def leer_todas_hojas_conciliacion_compras(ruta_archivo: str,
                                          fila_inicio: int = 2,
                                          progress_callback=None) -> pd.DataFrame:
    """Lee TODAS las hojas válidas de un archivo Excel y las concatena."""
    start_time = time.perf_counter()
    
    try:
        xl = pd.ExcelFile(ruta_archivo)
        hojas = xl.sheet_names
        
        if progress_callback:
            progress_callback(0.05, f"Analizando {len(hojas)} hojas...")
        
        df_total = pd.DataFrame()
        hojas_procesadas = 0
        
        for idx, hoja in enumerate(hojas):
            try:
                if progress_callback:
                    progress_callback(
                        0.05 + 0.4 * (idx / len(hojas)),
                        f"Procesando hoja {idx+1}/{len(hojas)}: {hoja}"
                    )
                
                df_hoja = leer_hoja_conciliacion_compras(
                    ruta_archivo, 
                    hoja, 
                    fila_inicio,
                    progress_callback
                )
                
                if not df_hoja.empty:
                    df_hoja['Hoja_Origen'] = hoja
                    df_total = pd.concat([df_total, df_hoja], ignore_index=True)
                    hojas_procesadas += 1
                
                del df_hoja
                gc.collect()
                
            except Exception as e:
                print(f"Error en hoja '{hoja}': {str(e)}")
                continue
        
        if df_total.empty:
            raise ValueError(f"No se encontraron datos válidos en {ruta_archivo}")
        
        df_total = df_total.drop_duplicates(subset=['ID_CONCILIACION'], keep='first')
        
        elapsed = time.perf_counter() - start_time
        if progress_callback:
            progress_callback(0.5, f"Archivo procesado: {len(df_total)} IDs únicos")
        
        return df_total
        
    except Exception as e:
        raise ValueError(f"Error procesando archivo: {str(e)}")


# ============================================================================
# 3. CONCILIACIÓN
# ============================================================================

def conciliar_archivos_compras(df1: pd.DataFrame, 
                               df2: pd.DataFrame,
                               nombre1: str = "SIRE_SUNAT",
                               nombre2: str = "SIRE_BN") -> Dict:
    """Compara dos DataFrames usando IDs de conciliación."""
    ids1 = set(df1['ID_CONCILIACION'])
    ids2 = set(df2['ID_CONCILIACION'])
    
    comunes = ids1 & ids2
    solo1_ids = ids1 - ids2
    solo2_ids = ids2 - ids1
    
    df1_solo = df1[df1['ID_CONCILIACION'].isin(solo1_ids)].copy()
    df2_solo = df2[df2['ID_CONCILIACION'].isin(solo2_ids)].copy()
    df_comunes = df1[df1['ID_CONCILIACION'].isin(comunes)].copy()
    
    resumen = {
        'archivo1': {
            'nombre': nombre1,
            'total_registros': len(df1),
            'ids_unicos': len(ids1),
            'solo_en_este': len(solo1_ids),
            'en_comun': len(comunes)
        },
        'archivo2': {
            'nombre': nombre2,
            'total_registros': len(df2),
            'ids_unicos': len(ids2),
            'solo_en_este': len(solo2_ids),
            'en_comun': len(comunes)
        },
        'coincidencias': len(comunes),
        'diferencias_totales': len(solo1_ids) + len(solo2_ids)
    }
    
    return {
        'resumen': resumen,
        'solo1': df1_solo,
        'solo2': df2_solo,
        'comunes': df_comunes
    }


# ============================================================================
# 4. REPORTE SIRE_BN
# ============================================================================

def generar_reporte_presentes_no_presentes_sire_bn(df_sire_sunat: pd.DataFrame,
                                                   df_sire_bn: pd.DataFrame,
                                                   nombre_sire_sunat: str = "SIRE_SUNAT",
                                                   nombre_sire_bn: str = "SIRE_BN",
                                                   ruta_salida: str = "reporte_presentes_bn.xlsx",
                                                   progress_callback=None) -> str:
    """
    Genera reporte de "Presentes vs No presentes" desde la perspectiva de SIRE_BN.
    """
    start_time = time.perf_counter()
    
    if progress_callback:
        progress_callback(0.6, "Generando reporte SIRE_BN...")
    
    # Normalizar tipos
    df_sire_sunat['Tipo_Doc_Norm'] = df_sire_sunat['Tipo_Doc_Norm'].apply(
        lambda x: str(x).zfill(2) if x is not None and len(str(x)) < 2 else str(x)
    )
    df_sire_bn['Tipo_Doc_Norm'] = df_sire_bn['Tipo_Doc_Norm'].apply(
        lambda x: str(x).zfill(2) if x is not None and len(str(x)) < 2 else str(x)
    )
    
    ids_sire_sunat = set(df_sire_sunat['ID_CONCILIACION'])
    tipos_sire_bn = sorted(df_sire_bn['Tipo_Doc_Norm'].unique())
    
    workbook = xlsxwriter.Workbook(ruta_salida, {'nan_inf_to_errors': False})
    
    # Formatos
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#1F4E78',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font_size': 11
    })
    
    title_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 14
    })
    
    bold_format = workbook.add_format({'bold': True})
    percent_format = workbook.add_format({'num_format': '0.00%'})
    number_format = workbook.add_format({'num_format': '#,##0'})
    
    # ==================== HOJA RESUMEN ====================

    ws_resumen = workbook.add_worksheet('Resumen')
    ws_resumen.set_column('A:A', 8)      # Tipo
    ws_resumen.set_column('B:B', 25)     # Nombre de Tipo
    ws_resumen.set_column('C:C', 18)     # Total en SIRE_BN
    ws_resumen.set_column('D:D', 22)     # Presentes en SIRE_SUNAT
    ws_resumen.set_column('E:E', 12)     # %
    ws_resumen.set_column('F:F', 22)     # No presentes en SIRE_SUNAT
    ws_resumen.set_column('G:G', 12)     # %
    
    if not isinstance(nombre_sire_sunat, str):
        nombre_sire_sunat = str(nombre_sire_sunat)
    if not isinstance(nombre_sire_bn, str):
        nombre_sire_bn = str(nombre_sire_bn)
        
    ws_resumen.merge_range('A1:G1', f'RESUMEN: REGISTROS DE {nombre_sire_bn}', title_format)

    presentes_header = f"Presentes en SIRE SUNAT"
    no_presentes_header = f"No presentes en SIRE SUNAT"

    headers = ['Tipo', 'Nombre de Tipo', 'Total en SIRE_BN', 
               presentes_header, '%', 
               no_presentes_header, '%']
    
    row = 2
    for col_idx, header in enumerate(headers):
        ws_resumen.write(row, col_idx, header, header_format)
    
    row = 3
    total_presentes = 0
    total_no_presentes = 0
    total_registros = 0
    
    for tipo in tipos_sire_bn:
        df_tipo = df_sire_bn[df_sire_bn['Tipo_Doc_Norm'] == tipo]
        total = len(df_tipo)
        
        if total == 0:
            continue
        
        mask_presente = df_tipo['ID_CONCILIACION'].isin(ids_sire_sunat)
        presentes = mask_presente.sum()
        no_presentes = total - presentes
        
        total_presentes += presentes
        total_no_presentes += no_presentes
        total_registros += total
        
        nombre_tipo = TIPOS_DOC_NOMBRES.get(tipo, f'Tipo {tipo}')
        
        pct_presentes = presentes / total if total > 0 else 0
        pct_no_presentes = no_presentes / total if total > 0 else 0
        #ws_resumen.write(row, 0, tipo) hace que se escriba el tipo en la primera columna
        ws_resumen.write(row, 0, tipo)
        ws_resumen.write(row, 1, nombre_tipo)
        ws_resumen.write(row, 2, total, number_format)
        ws_resumen.write(row, 3, presentes, number_format)
        ws_resumen.write(row, 4, pct_presentes, percent_format)
        ws_resumen.write(row, 5, no_presentes, number_format)
        ws_resumen.write(row, 6, pct_no_presentes, percent_format)
        row += 1
    
    # Fila de totales
    pct_total_presentes = total_presentes / total_registros if total_registros > 0 else 0
    pct_total_no_presentes = total_no_presentes / total_registros if total_registros > 0 else 0
    
    """
    bold_cells = ['','TOTAL', total_registros,  
                  total_presentes, pct_total_presentes, 
                  total_no_presentes, pct_total_no_presentes]
    
    for col, value in enumerate(bold_cells):
        if col in [2, 4, 6]:
            ws_resumen.write(row, col, value, number_format)
        elif col in [4, 6]:
            ws_resumen.write(row, col, value, percent_format)
        else:
            ws_resumen.write(row, col, value, bold_format)
    """
    # Escribir fila de totales
    ws_resumen.write(row, 0, ' ', bold_format)
    ws_resumen.write(row, 1, 'TOTAL', bold_format)
    ws_resumen.write(row, 2, total_registros, bold_format)# si quiero que sea negrita debo agregar bold_format en el tercer argumento ejmplo ws_resumen.write(row, 2, total_registros, bold_format)
    ws_resumen.write(row, 3, total_presentes, number_format)
    ws_resumen.write(row, 4, pct_total_presentes, percent_format)
    ws_resumen.write(row, 5, total_no_presentes, number_format)
    ws_resumen.write(row, 6, pct_total_no_presentes, percent_format)


    # Barras de datos
    if row > 3:
        """
        ws_resumen.conditional_format(3, 4, row - 1, 2, {# esta fila me indica que se aplicará un formato condicional a las celdas de la columna C (índice 2) desde la fila 4 hasta la fila anterior a la fila de totales, donde 3 indica el índice de la fila de inicio y row - 1 indica el índice de la fila final    
            # 3: me indica que se aplicará un formato condicional a las celdas de la columna C (índice 2) desde la fila 4 hasta la fila anterior a la fila de totales, donde 3 indica el índice de la fila de inicio y row - 1 indica el índice de la fila final
            # 2: me indica que se aplicará un formato condicional a las celdas de la columna C (índice 2) desde la fila 4 hasta la fila anterior a la fila de totales, donde 3 indica el índice de la fila de inicio y row - 1 indica el índice de la fila final
            # row - 1: me indica que se aplicará un formato condicional a las celdas de la columna C (índice 2) desde la fila 4 hasta la fila anterior a la fila de totales, 
            # donde 3 indica el índice de la fila de inicio y row - 1 indica el índice de la fila final
         
            'type': 'data_bar',
            'bar_color': '#10B981',
            'bar_negative_color': '#10B981',
            'data_bar_2010': True,
            'min_type': 'num',
            'min_value': 0,
            'max_type': 'num',
            'max_value': 1,
            'bar_negative_color_same': True
        })
        """
        ws_resumen.conditional_format(3, 4, row - 1, 4, {
            'type': 'data_bar',
            'bar_color': '#2563EB',
            'bar_negative_color': '#2563EB',
            'data_bar_2010': True,
            'min_type': 'num',
            'min_value': 0,
            'max_type': 'num',
            'max_value': 1,
            'bar_negative_color_same': True
        })
        
        ws_resumen.conditional_format(3, 6, row - 1, 6, {#7 me indica 
            'type': 'data_bar',
            'bar_color': '#EF4444',
            'bar_negative_color': '#EF4444',
            'data_bar_2010': True,
            'min_type': 'num',
            'min_value': 0,
            'max_type': 'num',
            'max_value': 1,
            'bar_negative_color_same': True
        })
    
    row += 2
    ws_resumen.write(row, 0, 'DESCRIPCIÓN DEL REPORTE:', bold_format)
    row += 1
    ws_resumen.write(row, 0, '• Para la presentación final agrega un grafico circular con los porcentajes de coincidencia y no coincidencia.')
    row += 1
    ws_resumen.write(row, 0, '• Y también un grafico de barras con los totales por tipo de documento.')
    
    # ==================== HOJA BD ======================
    # Crear hoja "SIRE_BD" después de Resumen
    ws_sire_bd = workbook.add_worksheet('SIRE_BD')
    
    # Título
    ws_sire_bd.merge_range('A1:L1', 'PROPUESTA SIRE BN', title_format)
    
    # Información de registros en la fila 3 (Fila 2 en índice 0-indexed)
    ws_sire_bd.write(2, 0, 'Total registros:', bold_format)
    ws_sire_bd.write(2, 1, len(df_sire_bn), bold_format)

    # Nota o instrucción en la fila 5 (Fila 4 en índice 0-indexed)
    ws_sire_bd.write(4, 0, 'Ingrese la tabla principal (Propuesta) usada para hacer la conciliación:', bold_format)

    row += 1

    # Ajustar ancho de columnas de forma eficiente (A hasta F)
    ws_sire_bd.set_column('A:M', 15)
    # ==================== HOJAS POR TIPO ====================
    for tipo in tipos_sire_bn:
        df_tipo = df_sire_bn[df_sire_bn['Tipo_Doc_Norm'] == tipo]
        total = len(df_tipo)
        
        if total == 0:
            continue
        
        mask_presente = df_tipo['ID_CONCILIACION'].isin(ids_sire_sunat)
        df_presentes = df_tipo[mask_presente].copy()
        df_no_presentes = df_tipo[~mask_presente].copy()
        
        # Limpiar valores
        df_presentes = df_presentes.fillna('')
        df_no_presentes = df_no_presentes.fillna('')
        
        # Eliminar Tipo_Doc_Norm
        if 'Tipo_Doc_Norm' in df_presentes.columns:
            df_presentes = df_presentes.drop(columns=['Tipo_Doc_Norm'])
        if 'Tipo_Doc_Norm' in df_no_presentes.columns:
            df_no_presentes = df_no_presentes.drop(columns=['Tipo_Doc_Norm'])
        
        # Convertir a string
        for col in df_presentes.columns:
            df_presentes[col] = df_presentes[col].astype(str)
        for col in df_no_presentes.columns:
            df_no_presentes[col] = df_no_presentes[col].astype(str)
        
        nombre_tipo = TIPOS_DOC_NOMBRES.get(tipo, f'Tipo {tipo}')
        base_name = f"Tipo_{tipo}"
        
        # --- PRIMERO: HOJA DE PRESENTES ---
        if not df_presentes.empty:
            # Si hay presentes, escribirlos en una hoja
            escribir_datos_con_limite(
                workbook, df_presentes,
                f"{base_name}_presentes",
                header_format,
                f'PRESENTES EN {nombre_sire_sunat}',
                '#10B981'  # Verde
            )
        else:
            # Si no hay presentes, crear una hoja con mensaje
            sheet_name = f"{base_name}_presentes"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.merge_range(0, 0, 0, 6, f'PRESENTES EN {nombre_sire_sunat}',
                        workbook.add_format({'bold': True, 'font_color': 'white', 
                                                'bg_color': '#10B981', 'align': 'center'}))
            ws.write(1, 0, 'No hay registros presentes')
        
        # --- SEGUNDO: HOJA DE NO PRESENTES ---
        if not df_no_presentes.empty:
            # Usar la función auxiliar para escribir en múltiples hojas si es necesario
            escribir_datos_con_limite(
                workbook, df_no_presentes,
                f"{base_name}_no_presentes",    
                header_format,
                f'NO PRESENTES EN {nombre_sire_sunat}',
                '#EF4444'  # Rojo
            )
        else:
            # Si no hay no presentes, crear una hoja con mensaje
            sheet_name = f"{base_name}_no_presentes"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.merge_range(0, 0, 0, 6, f'NO PRESENTES EN {nombre_sire_sunat}',
                        workbook.add_format({'bold': True, 'font_color': 'white', 
                                                'bg_color': '#EF4444', 'align': 'center'}))
            ws.write(1, 0, 'No hay registros no presentes')

    workbook.close()

    elapsed = time.perf_counter() - start_time
    if progress_callback:
        progress_callback(0.8, f"Reporte SIRE_BN generado en {elapsed:.2f}s")

    return ruta_salida


# ============================================================================
# 5. REPORTE SIRE_SUNAT
# ============================================================================

def generar_reporte_presentes_no_presentes_sire_sunat(df_sire_sunat: pd.DataFrame,
                                                     df_sire_bn: pd.DataFrame,
                                                     nombre_sire_sunat: str = "SIRE_SUNAT",
                                                     nombre_sire_bn: str = "SIRE_BN",
                                                     ruta_salida: str = "reporte_presentes_sunat.xlsx",
                                                     progress_callback=None) -> str:
    """
    Genera reporte de "Presentes vs No presentes" desde la perspectiva de SIRE_SUNAT.
    """
    start_time = time.perf_counter()
    
    if progress_callback:
        progress_callback(0.6, "Generando reporte SIRE_SUNAT...")
    
    # Normalizar tipos
    df_sire_sunat['Tipo_Doc_Norm'] = df_sire_sunat['Tipo_Doc_Norm'].apply(
        lambda x: str(x).zfill(2) if x is not None and len(str(x)) < 2 else str(x)
    )
    df_sire_bn['Tipo_Doc_Norm'] = df_sire_bn['Tipo_Doc_Norm'].apply(
        lambda x: str(x).zfill(2) if x is not None and len(str(x)) < 2 else str(x)
    )
    
    ids_sire_bn = set(df_sire_bn['ID_CONCILIACION'])
    tipos_sire_sunat = sorted(df_sire_sunat['Tipo_Doc_Norm'].unique())
    
    workbook = xlsxwriter.Workbook(ruta_salida, {'nan_inf_to_errors': False})
    
    # Formatos
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#1F4E78',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font_size': 11
    })
    
    title_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 14
    })
    
    bold_format = workbook.add_format({'bold': True})
    percent_format = workbook.add_format({'num_format': '0.00%'})
    number_format = workbook.add_format({'num_format': '#,##0'})
    
    # ==================== HOJA RESUMEN ====================
    # ws_resumen = workbook.add_worksheet('Resumen') me indica que la hoja de resumen se llama "Resumen"
    #ws_resumen.set_column('A:A', 8) establece el ancho de la columna A a 8 unidades
    ws_resumen = workbook.add_worksheet('Resumen')
    ws_resumen.set_column('A:A', 8)      # Tipo
    ws_resumen.set_column('B:B', 25)     # Nombre de Tipo
    ws_resumen.set_column('C:C', 18)     # Total en SIRE_BN
    ws_resumen.set_column('D:D', 22)     # Presentes en SIRE_SUNAT
    ws_resumen.set_column('E:E', 12)     # %
    ws_resumen.set_column('F:F', 22)     # No presentes en SIRE_SUNAT
    ws_resumen.set_column('G:G', 12)     # %
    
    ws_resumen.merge_range('A1:H1', f'RESUMEN: REGISTROS DE {nombre_sire_sunat}', title_format)
    
    headers = ['Tipo', 'Nombre de Tipo','Total en SIRE_SUNAT',   
               f'Presentes en SIRE BN', '%', 
               f'No presentes en SIRE BN', '%']
    
    row = 2
    for col_idx, header in enumerate(headers):
        ws_resumen.write(row, col_idx, header, header_format)
    
    row = 3
    total_presentes = 0
    total_no_presentes = 0
    total_registros = 0
    
    for tipo in tipos_sire_sunat:
        df_tipo = df_sire_sunat[df_sire_sunat['Tipo_Doc_Norm'] == tipo]
        total = len(df_tipo)
        
        if total == 0:
            continue
        
        mask_presente = df_tipo['ID_CONCILIACION'].isin(ids_sire_bn)
        presentes = mask_presente.sum()
        no_presentes = total - presentes
        
        total_presentes += presentes
        total_no_presentes += no_presentes
        total_registros += total
        
        nombre_tipo = TIPOS_DOC_NOMBRES.get(tipo, f'Tipo {tipo}')
        
        pct_presentes = presentes / total if total > 0 else 0
        pct_no_presentes = no_presentes / total if total > 0 else 0
        
        #ws_resumen.write(row, 0, tipo) hace que se escriba el tipo en la primera columna
        ws_resumen.write(row, 0, tipo)
        ws_resumen.write(row, 1, nombre_tipo)
        ws_resumen.write(row, 2, total, number_format)
        ws_resumen.write(row, 3, presentes, number_format)
        ws_resumen.write(row, 4, pct_presentes, percent_format)
        ws_resumen.write(row, 5, no_presentes, number_format)
        ws_resumen.write(row, 6, pct_no_presentes, percent_format)
        row += 1
    
    # Fila de totales
    pct_total_presentes = total_presentes / total_registros if total_registros > 0 else 0
    pct_total_no_presentes = total_no_presentes / total_registros if total_registros > 0 else 0
    
    
    # Escribir fila de totales
    ws_resumen.write(row, 0, ' ', bold_format)
    ws_resumen.write(row, 1, 'TOTAL', bold_format)
    ws_resumen.write(row, 2, total_registros, bold_format)# si quiero que sea negrita debo agregar bold_format en el tercer argumento ejmplo ws_resumen.write(row, 2, total_registros, bold_format)
    ws_resumen.write(row, 3, total_presentes, number_format)
    ws_resumen.write(row, 4, pct_total_presentes, percent_format)
    ws_resumen.write(row, 5, total_no_presentes, number_format)
    ws_resumen.write(row, 6, pct_total_no_presentes, percent_format)

 
    # Barras de datos

    if row > 3:
        """
        ws_resumen.conditional_format(3, 4, row - 1, 2, {
            'type': 'data_bar',
            'bar_color': '#10B981',
            'bar_negative_color': '#10B981',
            'data_bar_2010': True,
            'min_type': 'num',
            'min_value': 0,
            'max_type': 'num',
            'max_value': 1,
            'bar_negative_color_same': True
        })
        """
        
        ws_resumen.conditional_format(3, 4, row - 1, 4, {
            'type': 'data_bar',
            'bar_color': '#2563EB',
            'bar_negative_color': '#2563EB',
            'data_bar_2010': True,
            'min_type': 'num',
            'min_value': 0,
            'max_type': 'num',
            'max_value': 1,
            'bar_negative_color_same': True
        })
        

        ws_resumen.conditional_format(3, 6, row - 1, 6, {
            'type': 'data_bar',
            'bar_color': '#EF4444',
            'bar_negative_color': '#EF4444',
            'data_bar_2010': True,
            'min_type': 'num',
            'min_value': 0,
            'max_type': 'num',
            'max_value': 1,
            'bar_negative_color_same': True
        })
    
    row += 2
    ws_resumen.write(row, 0, 'DESCRIPCIÓN DEL REPORTE:', bold_format)
    row += 1
    ws_resumen.write(row, 0, '• Para la presentación final agrega un grafico circular con los porcentajes de coincidencia y no coincidencia.')
    row += 1
    ws_resumen.write(row, 0, '• Y también un grafico de barras con los totales por tipo de documento.')
    


    # ==================== HOJA BD ======================
    # Crear hoja "SIRE_BD" después de Resumen
    ws_sire_bd = workbook.add_worksheet('SIRE_BD')
    
    # Título
    ws_sire_bd.merge_range('A1:L1', 'PROPUESTA SIRE SUNAT', title_format)
    
    # Información de registros en la fila 3 (Fila 2 en índice 0-indexed)
    ws_sire_bd.write(2, 0, 'Total registros:', bold_format)
    ws_sire_bd.write(2, 1, len(df_sire_bn), bold_format)

    # Nota o instrucción en la fila 5 (Fila 4 en índice 0-indexed)
    ws_sire_bd.write(4, 0, 'Ingrese la tabla principal (Propuesta) usada para hacer la conciliación:', bold_format)

    row += 1

    # Ajustar ancho de columnas de forma eficiente (A hasta F)
    ws_sire_bd.set_column('A:M', 15)
    # ==================== HOJAS POR TIPO ====================
    for tipo in tipos_sire_sunat:
        df_tipo = df_sire_sunat[df_sire_sunat['Tipo_Doc_Norm'] == tipo]
        total = len(df_tipo)
        
        if total == 0:
            continue
        
        mask_presente = df_tipo['ID_CONCILIACION'].isin(ids_sire_bn)
        df_presentes = df_tipo[mask_presente].copy()
        df_no_presentes = df_tipo[~mask_presente].copy()
        
        # Limpiar valores
        df_presentes = df_presentes.fillna('')
        df_no_presentes = df_no_presentes.fillna('')
        
        # Eliminar Tipo_Doc_Norm
        if 'Tipo_Doc_Norm' in df_presentes.columns:
            df_presentes = df_presentes.drop(columns=['Tipo_Doc_Norm'])
        if 'Tipo_Doc_Norm' in df_no_presentes.columns:
            df_no_presentes = df_no_presentes.drop(columns=['Tipo_Doc_Norm'])
        
        # Convertir a string
        for col in df_presentes.columns:
            df_presentes[col] = df_presentes[col].astype(str)
        for col in df_no_presentes.columns:
            df_no_presentes[col] = df_no_presentes[col].astype(str)
        
        nombre_tipo = TIPOS_DOC_NOMBRES.get(tipo, f'Tipo {tipo}')
        base_name = f"Tipo_{tipo}"
        
        # --- PRIMERO: HOJA DE PRESENTES ---
        if not df_presentes.empty:
            escribir_datos_con_limite(
                workbook, df_presentes,
                f"{base_name}_presentes",
                header_format,
                f'PRESENTES EN {nombre_sire_bn}',
                '#10B981'  # Verde
            )
        else:
            sheet_name = f"{base_name}_presentes"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.merge_range(0, 0, 0, 6, f'PRESENTES EN {nombre_sire_bn}',
                        workbook.add_format({'bold': True, 'font_color': 'white', 
                                                'bg_color': '#10B981', 'align': 'center'}))
            ws.write(1, 0, 'No hay registros presentes')
        
        # --- SEGUNDO: HOJA DE NO PRESENTES ---
        if not df_no_presentes.empty:
            escribir_datos_con_limite(
                workbook, df_no_presentes,
                f"{base_name}_no_presentes",
                header_format,
                f'NO PRESENTES EN {nombre_sire_bn}',
                '#EF4444'  # Rojo
            )
        else:
            sheet_name = f"{base_name}_no_presentes"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.merge_range(0, 0, 0, 6, f' NO PRESENTES EN {nombre_sire_bn}',
                        workbook.add_format({'bold': True, 'font_color': 'white', 
                                                'bg_color': '#EF4444', 'align': 'center'}))
            ws.write(1, 0, 'No hay registros no presentes')
        
        # Ajustar ancho de columnas (opcional, se aplica después de crear todas las hojas)
        # Nota: xlsxwriter no permite ajustar columnas después de cerrar el workbook,
        # pero podemos hacerlo aquí

    workbook.close()

    elapsed = time.perf_counter() - start_time
    if progress_callback:
        progress_callback(0.8, f"Reporte SIRE_SUNAT generado en {elapsed:.2f}s")

    return ruta_salida
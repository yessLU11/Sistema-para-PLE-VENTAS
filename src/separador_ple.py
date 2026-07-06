# src/separador_ple.py
"""
Módulo Separador de PLE Ventas
- Procesa archivos TXT masivos (500MB+) y Excel con múltiples hojas
- Separa boletas (serie B) de facturas y otros
- Genera Excel optimizado con múltiples hojas
- Para Excel: SIEMPRE usa la columna G (índice 6) para identificar la serie
- Para TXT: usa la posición 7 (índice 7) donde está la serie
"""

import pandas as pd
import numpy as np
import re
from typing import Tuple, Dict, List, Optional
from datetime import datetime
import os
import gc
import time
import chardet

# ============================================================================
# CONFIGURACIÓN DE COLUMNAS
# ============================================================================

# Posiciones fijas en el TXT y excel (0-index)
COLUMNA_SERIE_TXT = 7       # Posición de la Serie (ej: B001, F001, 0128)
COLUMNA_SERIE_EXCEL = 6     # Columna G (0-index) para Excel - SIEMPRE ESTA

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def detectar_codificacion(file_path: str) -> str:
    """Detecta la codificación del archivo."""
    with open(file_path, 'rb') as f:
        raw_data = f.read(10000)
        result = chardet.detect(raw_data)
        return result['encoding'] or 'latin-1'


def limpiar_valor(valor) -> str:
    """Limpia un valor para evitar caracteres problemáticos en Excel."""
    if pd.isna(valor) or valor is None:
        return ''
    
    valor = str(valor)
    valor = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', valor)
    valor = valor.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
    valor = re.sub(r'\s+', ' ', valor)
    return valor.strip()


def debug_print(mensaje, nivel=0, mostrar=True):
    """Imprime mensajes de debug."""
    if not mostrar:
        return
    indent = "  " * nivel
    print(f"[DEBUG] {indent}{mensaje}")


# ============================================================================
# FUNCIONES DE LECTURA PARA TXT
# ============================================================================

def leer_txt_masivo(file_path: str, 
                    fila_inicio: int = 1,
                    chunk_size: int = 100000,
                    progress_callback=None) -> pd.DataFrame:
    """
    Lee un archivo TXT masivo por chunks para optimizar memoria.
    """
    start_time = time.perf_counter()
    
    if progress_callback:
        progress_callback(0.02, "Detectando codificación...")
    
    encoding = detectar_codificacion(file_path)
    
    if progress_callback:
        progress_callback(0.05, f"Codificación: {encoding}")
    
    try:
        chunks = []
        lineas_procesadas = 0
        total_lineas = 0
        
        if progress_callback:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                total_lineas = sum(1 for _ in f)
            if progress_callback:
                progress_callback(0.08, f"Total líneas: {total_lineas:,}")
        
        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            # Saltar filas según fila_inicio
            for _ in range(fila_inicio - 1):
                next(f, None)
            
            chunk = []
            for linea in f:
                linea_limpia = limpiar_valor(linea)
                if linea_limpia:
                    chunk.append(linea_limpia)
                lineas_procesadas += 1
                
                if len(chunk) >= chunk_size:
                    df_chunk = procesar_chunk_txt(chunk)
                    if not df_chunk.empty:
                        chunks.append(df_chunk)
                    chunk = []
                    
                    if progress_callback and total_lineas > 0:
                        progress_callback(
                            0.08 + 0.82 * (lineas_procesadas / total_lineas),
                            f"Procesando {lineas_procesadas:,}/{total_lineas:,} líneas"
                        )
                    gc.collect()
            
            if chunk:
                df_chunk = procesar_chunk_txt(chunk)
                if not df_chunk.empty:
                    chunks.append(df_chunk)
        
        if not chunks:
            return pd.DataFrame()
        
        df_final = pd.concat(chunks, ignore_index=True)
        
        elapsed = time.perf_counter() - start_time
        if progress_callback:
            progress_callback(0.95, f"Archivo leído en {elapsed:.2f}s - {len(df_final):,} registros")
        
        return df_final
        
    except MemoryError:
        raise MemoryError("Archivo demasiado grande para la memoria disponible.")
    except Exception as e:
        raise Exception(f"Error al leer el archivo: {str(e)}")


def procesar_chunk_txt(chunk: List[str]) -> pd.DataFrame:
    """Procesa un chunk de líneas TXT y las convierte a DataFrame."""
    if not chunk:
        return pd.DataFrame()
    
    datos = []
    for linea in chunk:
        if linea:
            # Dividir por pipe '|'
            partes = linea.split('|')
            datos.append(partes)
    
    if not datos:
        return pd.DataFrame()
    
    # Crear DataFrame con columnas genéricas
    max_cols = max(len(d) for d in datos)
    df = pd.DataFrame(datos)
    
    # Asegurar que todas las filas tengan el mismo número de columnas
    for i in range(len(df)):
        if len(df.iloc[i]) < max_cols:
            df.iloc[i] = df.iloc[i].tolist() + [''] * (max_cols - len(df.iloc[i]))
    
    return df


# ============================================================================
# FUNCIONES DE LECTURA PARA EXCEL
# ============================================================================

def leer_excel_masivo(file_path: str,
                      fila_inicio: int = 1,
                      progress_callback=None) -> pd.DataFrame:
    """
    Lee un archivo Excel con múltiples hojas y las concatena.
    SIEMPRE usa la columna G (índice 6) para identificar la serie.
    """
    start_time = time.perf_counter()
    
    if progress_callback:
        progress_callback(0.02, "Leyendo archivo Excel...")
    
    try:
        xl = pd.ExcelFile(file_path)
        hojas = xl.sheet_names
        total_hojas = len(hojas)
        
        if progress_callback:
            progress_callback(0.05, f"Total hojas: {total_hojas}")
        
        # Procesar todas las hojas
        dfs = []
        
        for idx, hoja in enumerate(hojas):
            try:
                if progress_callback:
                    progress_callback(
                        0.05 + 0.40 * ((idx + 1) / total_hojas),
                        f"Procesando hoja {idx+1}/{total_hojas}: {hoja}"
                    )
                
                # Leer la hoja - usar header=None para tomar todas las filas como datos
                df_hoja = pd.read_excel(
                    file_path,
                    sheet_name=hoja,
                    header=None,  # Sin encabezados
                    dtype=str,
                    engine='openpyxl'
                )
                
                # Eliminar filas vacías
                df_hoja = df_hoja.dropna(how='all')
                
                if not df_hoja.empty:
                    # Si fila_inicio > 1, eliminar las primeras filas
                    if fila_inicio > 1:
                        df_hoja = df_hoja.iloc[fila_inicio - 1:]
                    
                    # Resetear índices
                    df_hoja = df_hoja.reset_index(drop=True)
                    
                    # Agregar columna de origen
                    df_hoja['Hoja_Origen'] = hoja
                    
                    dfs.append(df_hoja)
                    debug_print(f"Hoja '{hoja}': {len(df_hoja)} filas, {len(df_hoja.columns)} columnas", 1, True)
                
                del df_hoja
                gc.collect()
                
            except Exception as e:
                print(f"Error en hoja '{hoja}': {str(e)}")
                continue
        
        if not dfs:
            raise ValueError("No se encontraron datos en ninguna hoja")
        
        # Concatenar todas las hojas
        df_final = pd.concat(dfs, ignore_index=True)
        
        elapsed = time.perf_counter() - start_time
        if progress_callback:
            progress_callback(0.50, f"Excel leído en {elapsed:.2f}s - {len(df_final):,} registros")
        
        # DEBUG: Mostrar información de columnas
        debug_print(f"Total de columnas: {len(df_final.columns)}", 1, True)
        debug_print(f"Primeras 5 columnas: {df_final.columns[:5].tolist()}", 1, True)
        
        # Verificar que existe la columna G (posición 6)
        if len(df_final.columns) > COLUMNA_SERIE_EXCEL:
            col_g = df_final.columns[COLUMNA_SERIE_EXCEL]
            debug_print(f"Columna G (posición {COLUMNA_SERIE_EXCEL}): '{col_g}'", 1, True)
            
            # Mostrar muestra de valores de la columna G
            muestra = df_final[col_g].head(10).tolist()
            debug_print(f"Muestra de valores en columna G: {muestra}", 1, True)
        else:
            debug_print(f"Menos de 7 columnas. Total: {len(df_final.columns)}", 1, True)
        
        return df_final
        
    except Exception as e:
        raise Exception(f"Error al leer el archivo Excel: {str(e)}")


# ============================================================================
# FUNCIONES DE SEPARACIÓN
# ============================================================================

def separar_por_tipo(df: pd.DataFrame, tipo_archivo: str = "txt") -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Separa el DataFrame en boletas (serie B) y otros (facturas, etc.)
    - TXT: usa columna en posición 7 (índice 7)
    - Excel: SIEMPRE usa columna G (posición 6)
    """
    debug_print("=" * 60, 0, True)
    debug_print("INICIANDO SEPARACIÓN POR TIPO", 0, True)
    debug_print(f"Tipo de archivo: {tipo_archivo}", 1, True)
    debug_print(f"Total de registros: {len(df):,}", 1, True)
    debug_print(f"Total de columnas: {len(df.columns)}", 1, True)
    debug_print("=" * 60, 0, True)
    
    if df.empty:
        debug_print("⚠️ DataFrame vacío", 1, True)
        return pd.DataFrame(), pd.DataFrame()
    
    # Buscar la columna de Serie según el tipo
    col_serie = None
    
    if tipo_archivo == "txt":
        # TXT: usar columna en posición 7 (índice 7)
        if len(df.columns) > COLUMNA_SERIE_TXT:
            col_serie = df.columns[COLUMNA_SERIE_TXT]
            debug_print(f"Usando columna por posición {COLUMNA_SERIE_TXT}: '{col_serie}'", 2, True)
        else:
            debug_print(f"❌ No hay suficientes columnas. Solo {len(df.columns)}", 1, True)
            return pd.DataFrame(), df
    else:
        # Excel: SIEMPRE usar columna G (posición 6)
        if len(df.columns) > COLUMNA_SERIE_EXCEL:
            col_serie = df.columns[COLUMNA_SERIE_EXCEL]
            debug_print(f"Usando columna G (posición {COLUMNA_SERIE_EXCEL}): '{col_serie}'", 2, True)
        elif len(df.columns) > 0:
            # Si hay menos de 7 columnas, usar la primera como fallback
            col_serie = df.columns[0]
            debug_print(f"⚠️ Menos de 7 columnas, usando primera columna: '{col_serie}'", 2, True)
        else:
            debug_print("❌ No se encontraron columnas", 1, True)
            return pd.DataFrame(), df
    
    if col_serie is None:
        debug_print("❌ No se encontró columna de serie", 1, True)
        return pd.DataFrame(), df
    
    debug_print(f"Columna de serie seleccionada: '{col_serie}'", 1, True)
    
    # Mostrar muestra de valores
    sample = df[col_serie].head(10).tolist()
    debug_print(f"Muestra de valores: {sample}", 1, True)
    
    # Normalizar serie (convertir a string y limpiar)
    df['Serie_Norm'] = df[col_serie].astype(str).str.strip().str.upper()
    
    # Contar tipos de serie
    series_unicas = df['Serie_Norm'].value_counts().head(10)
    debug_print("Tipos de serie más comunes:", 1, True)
    for serie, count in series_unicas.items():
        debug_print(f"  {serie}: {count:,} registros", 2, True)
    
    # Filtrar boletas (serie comienza con 'B')
    mask_boleta = df['Serie_Norm'].str.startswith('B', na=False)
    count_boletas = mask_boleta.sum()
    count_otros = (~mask_boleta).sum()
    
    debug_print(f"Boletas (serie B): {count_boletas:,}", 1, True)
    debug_print(f"Otros (facturas, etc.): {count_otros:,}", 1, True)
    
    df_boletas = df[mask_boleta].copy()
    df_otros = df[~mask_boleta].copy()
    
    # Eliminar columna temporal
    if 'Serie_Norm' in df_boletas.columns:
        df_boletas.drop('Serie_Norm', axis=1, inplace=True)
    if 'Serie_Norm' in df_otros.columns:
        df_otros.drop('Serie_Norm', axis=1, inplace=True)
    
    debug_print("=" * 60, 0, True)
    debug_print(f"✅ Separación completada: {len(df_boletas):,} boletas, {len(df_otros):,} otros", 0, True)
    debug_print("=" * 60, 0, True)
    
    return df_boletas, df_otros


# ============================================================================
# FUNCIONES DE GENERACIÓN DE EXCEL
# ============================================================================

def generar_excel_separado(df_boletas: pd.DataFrame,
                           df_otros: pd.DataFrame,
                           ruta_salida_boletas: str,
                           ruta_salida_otros: str,
                           progress_callback=None) -> Dict:
    """Genera los archivos Excel separados."""
    LIMITE_FILAS = 1_048_000
    
    resultado = {
        'boletas': {'success': False, 'hojas': 0, 'registros': 0, 'archivo': ruta_salida_boletas},
        'otros': {'success': False, 'hojas': 0, 'registros': 0, 'archivo': ruta_salida_otros}
    }
    
    # Función auxiliar para guardar DataFrame
    def guardar_dataframe(df, ruta, nombre_base):
        if df.empty:
            # Crear archivo con mensaje
            os.makedirs(os.path.dirname(ruta) or '.', exist_ok=True)
            pd.DataFrame({'Mensaje': [f'No se encontraron {nombre_base.lower()} en el archivo']}).to_excel(
                ruta, index=False
            )
            return {'success': True, 'hojas': 1, 'registros': 0, 'archivo': ruta}
        
        total_filas = len(df)
        num_hojas = (total_filas // LIMITE_FILAS) + 1 if total_filas > LIMITE_FILAS else 1
        
        # Limpiar datos
        df_clean = df.replace({np.nan: '', None: ''})
        
        try:
            os.makedirs(os.path.dirname(ruta) or '.', exist_ok=True)
            
            with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
                if num_hojas == 1:
                    nombre_hoja = nombre_base[:31]
                    df_clean.to_excel(writer, sheet_name=nombre_hoja, index=False)
                else:
                    for parte in range(num_hojas):
                        inicio = parte * LIMITE_FILAS
                        fin = min((parte + 1) * LIMITE_FILAS, total_filas)
                        df_parte = df_clean.iloc[inicio:fin]
                        nombre_hoja = f"{nombre_base}_{parte + 1}"[:31]
                        df_parte.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            return {
                'success': True,
                'hojas': num_hojas,
                'registros': total_filas,
                'archivo': ruta
            }
        except Exception as e:
            return {
                'success': False,
                'hojas': 0,
                'registros': 0,
                'error': str(e)
            }
    
    # Generar Excel de Boletas
    if progress_callback:
        progress_callback(0.50, f"Generando Excel de Boletas ({len(df_boletas):,} registros)...")
    
    resultado['boletas'] = guardar_dataframe(df_boletas, ruta_salida_boletas, 'Boletas')
    
    # Generar Excel de Otros
    if progress_callback:
        progress_callback(0.75, f"Generando Excel de Facturas y otros ({len(df_otros):,} registros)...")
    
    resultado['otros'] = guardar_dataframe(df_otros, ruta_salida_otros, 'Facturas_Otros')
    
    return resultado


# ============================================================================
# FUNCIÓN PRINCIPAL - ORQUESTADOR
# ============================================================================

def procesar_archivo_separador(file_path: str,
                               fila_inicio: int = 1,
                               es_excel: bool = False,
                               progress_callback=None) -> Dict:
    """
    Orquestador principal del proceso de separación.
    """
    start_time = time.perf_counter()
    
    debug_print("=" * 70, 0, True)
    debug_print("🚀 INICIANDO PROCESO DE SEPARACIÓN", 0, True)
    debug_print(f"Archivo: {file_path}", 1, True)
    debug_print(f"Es Excel: {es_excel}", 1, True)
    debug_print(f"Fila de inicio: {fila_inicio}", 1, True)
    debug_print("=" * 70, 0, True)
    
    try:
        if progress_callback:
            progress_callback(0.01, "Iniciando procesamiento...")
        
        # 1. Leer archivo
        debug_print("PASO 1: Leyendo archivo", 1, True)
        
        if es_excel:
            df = leer_excel_masivo(file_path, fila_inicio, progress_callback)
        else:
            df = leer_txt_masivo(file_path, fila_inicio, progress_callback=progress_callback)
        
        if df.empty:
            debug_print("❌ ERROR: No se encontraron datos", 1, True)
            return {
                'success': False,
                'error': 'No se encontraron datos en el archivo',
                'elapsed_seconds': time.perf_counter() - start_time
            }
        
        debug_print(f"✅ Total de registros: {len(df):,}", 1, True)
        debug_print(f"Total de columnas: {len(df.columns)}", 1, True)
        
        # 2. Separar por tipo
        if progress_callback:
            progress_callback(0.45, "Separando boletas de facturas...")
        
        debug_print("PASO 2: Separando por tipo", 1, True)
        tipo_archivo = "excel" if es_excel else "txt"
        df_boletas, df_otros = separar_por_tipo(df, tipo_archivo)
        
        debug_print(f"  Boletas: {len(df_boletas):,}", 2, True)
        debug_print(f"  Facturas y otros: {len(df_otros):,}", 2, True)
        
        # 3. Generar nombres de salida
        debug_print("PASO 3: Generando archivos de salida", 1, True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"separado_{timestamp}"
        
        # Crear directorio reportes si no existe
        os.makedirs("reportes", exist_ok=True)
        
        ruta_boletas = f"reportes/{base_name}_boletas.xlsx"
        ruta_otros = f"reportes/{base_name}_facturas_otros.xlsx"
        
        # 4. Generar Exceles
        debug_print("PASO 4: Generando Exceles", 1, True)
        resultado_excel = generar_excel_separado(
            df_boletas, df_otros,
            ruta_boletas, ruta_otros,
            progress_callback
        )
        
        elapsed = time.perf_counter() - start_time
        
        debug_print("=" * 70, 0, True)
        debug_print("✅ PROCESO COMPLETADO EXITOSAMENTE", 0, True)
        debug_print(f"Tiempo total: {elapsed:.2f} segundos", 1, True)
        debug_print(f"Archivo de boletas: {ruta_boletas}", 1, True)
        debug_print(f"Archivo de facturas: {ruta_otros}", 1, True)
        debug_print("=" * 70, 0, True)
        
        return {
            'success': True,
            'total_registros': len(df),
            'boletas': {
                'registros': len(df_boletas),
                'hojas': resultado_excel['boletas']['hojas'],
                'archivo': ruta_boletas,
                'success': resultado_excel['boletas']['success']
            },
            'otros': {
                'registros': len(df_otros),
                'hojas': resultado_excel['otros']['hojas'],
                'archivo': ruta_otros,
                'success': resultado_excel['otros']['success']
            },
            'elapsed_seconds': elapsed,
            'timestamp': timestamp,
            'tipo_archivo': 'Excel' if es_excel else 'TXT'
        }
        
    except Exception as e:
        elapsed = time.perf_counter() - start_time
        debug_print(f"❌ ERROR: {str(e)}", 1, True)
        import traceback
        debug_print(traceback.format_exc(), 1, True)
        return {
            'success': False,
            'error': str(e),
            'elapsed_seconds': elapsed
        }


# ============================================================================
# EJEMPLO DE USO
# ============================================================================

if __name__ == "__main__":
    # Ejemplo de uso para TXT
    resultado = procesar_archivo_separador(
        file_path="uploads/mi_archivo.txt",
        fila_inicio=1,
        es_excel=False
    )
    print(resultado)
    
    # Ejemplo de uso para Excel
    # resultado = procesar_archivo_separador(
    #     file_path="uploads/mi_archivo.xlsx",
    #     fila_inicio=1,
    #     es_excel=True
    # )
    # print(resultado)
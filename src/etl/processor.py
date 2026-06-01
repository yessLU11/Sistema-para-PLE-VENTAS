import pandas as pd
import io

# --- REGLAS DE NEGOCIO ---
PREFIX_RULES = {
    "01": "F",
    "03": "B"
}

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

def procesar_excel(file_stream):
    """Orquestador principal: Extrae, Normaliza y Detecta Duplicados."""
    try:
        excel_file = pd.ExcelFile(file_stream)
        hojas = excel_file.sheet_names
        
        df_limpio = pd.DataFrame()
        lista_errores = []
        
        total_filas = 0
        duplicados_encontrados = 0

        for hoja in hojas:
            df = pd.read_excel(excel_file, sheet_name=hoja)
            
            if df.empty:
                lista_errores.append(f"Hoja '{hoja}' está vacía.")
                continue

            # Ajustar índices de fila para auditoría (Excel inicia en 1, cabecera es 1)
            df['_ExcelRow'] = df.index + 2 

            # Normalización de Tipo de Documento
            if 'TipoDoc' in df.columns:
                df['TipoDoc_Norm'] = df['TipoDoc'].apply(normalizar_tipo_doc)
            else:
                lista_errores.append(f"Hoja '{hoja}': Falta columna 'TipoDoc'")
                continue

            # Normalización de Código de Establecimiento
            if 'CodigoEstablecimiento' in df.columns:
                df['CodigoEstablecimiento_Norm'] = df.apply(
                    lambda x: normalizar_codigo(x['TipoDoc_Norm'], x['CodigoEstablecimiento']), axis=1
                )
            else:
                lista_errores.append(f"Hoja '{hoja}': Falta columna 'CodigoEstablecimiento'")

            df['Hoja_Nombre'] = hoja
            df_limpio = pd.concat([df_limpio, df], ignore_index=True)

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

        return {
            "success": True,
            "hojas_procesadas": hojas,
            "total_registros": total_filas,
            "duplicados_count": duplicados_encontrados,
            "df_duplicados": df_trazabilidad,
            "df_limpio": df_limpio, 
            "errores": lista_errores
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"Error fatal en el procesamiento: {str(e)}",
            "errores": []
        }
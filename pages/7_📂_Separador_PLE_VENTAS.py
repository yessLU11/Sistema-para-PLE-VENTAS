# pages/7_📂_Separador_PLE_VENTAS.py
"""
Separador de PLE VENTAS
- Procesa archivos TXT o Excel (500MB+)
- Separa boletas (serie B) de facturas y otros
- Genera Exceles separados optimizados
- Descarga ambos archivos sin reprocesar
"""

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from src.separador_ple import procesar_archivo_separador
from config import OUTPUT_DIR

st.set_page_config(
    page_title="Separador PLE VENTAS",
    page_icon="📂",
    layout="wide"
)

st.title("Separador de Comprobantes PLE Ventas")
st.markdown("""
Procesa archivos TXT o Excel y separa automáticamente:
- **Boletas** (Series que comienzan con 'B')
- **Facturas y otros** (Series que comienzan con 'F' u otros)
- Soporte para **TXT** y **Excel** con múltiples hojas
- Para Excel: **SIEMPRE usa la columna G** para identificar la serie
- Optimizado para archivos **>500 MB**
""")

os.makedirs("reportes", exist_ok=True)
os.makedirs("uploads", exist_ok=True)

# ============================================================================
# INICIALIZAR ESTADO DE SESIÓN
# ============================================================================

if "resultado_separador" not in st.session_state:
    st.session_state.resultado_separador = None
if "archivo_procesado" not in st.session_state:
    st.session_state.archivo_procesado = None
if "ruta_boletas" not in st.session_state:
    st.session_state.ruta_boletas = None
if "ruta_otros" not in st.session_state:
    st.session_state.ruta_otros = None
if "df_boletas" not in st.session_state:
    st.session_state.df_boletas = None
if "df_otros" not in st.session_state:
    st.session_state.df_otros = None
if "procesando" not in st.session_state:
    st.session_state.procesando = False
if "ultimo_archivo" not in st.session_state:
    st.session_state.ultimo_archivo = None

# ============================================================================
# SELECCIÓN DE TIPO DE ARCHIVO
# ============================================================================

st.markdown("## 📤 Seleccionar archivo")

tipo_archivo = st.radio(
    "Tipo de archivo:",
    ["TXT (PLE)", "Excel (xlsx)"],
    horizontal=True,
    help="Selecciona el formato de tu archivo",
    key="tipo_archivo_radio"
)

# ============================================================================
# CARGA DE ARCHIVO
# ============================================================================

col1, col2 = st.columns([2, 1])

with col1:
    extensiones = ["txt"] if tipo_archivo == "TXT (PLE)" else ["xlsx", "xls"]
    archivo = st.file_uploader(
        f"Seleccionar archivo {tipo_archivo}",
        type=extensiones,
        key="file_upload",
        help=f"Archivo en formato {tipo_archivo}"
    )

with col2:
    fila_inicio = st.number_input(
        "Fila de inicio",
        min_value=1,
        value=1,
        step=1,
        help="Fila donde empiezan los datos (1-index)",
        key="fila_inicio_input"
    )
    
    # Mostrar requisitos según el tipo
    if tipo_archivo == "TXT (PLE)":
        st.info("**TXT:** Asegurate que la Serie esté en la posición 7 (índice 7)")
    else:
        st.info("**Excel:** Asegurate que la Serie esté en la columna G")
    st.caption("Límite: 500MB+")

# ============================================================================
# INFORMACIÓN DEL ARCHIVO
# ============================================================================

if archivo:
    tamano_mb = archivo.size / (1024 * 1024)
    st.info(f"""
    **Archivo:** {archivo.name}  
    **Tamaño:** {tamano_mb:.2f} MB ({archivo.size:,} bytes)  
    **Tipo:** {tipo_archivo}
    **Fila de inicio:** {fila_inicio}
    """)

# ============================================================================
# BOTÓN DE PROCESAMIENTO
# ============================================================================

st.markdown("---")

# Verificar si es un archivo nuevo para resetear resultados
if archivo and st.session_state.ultimo_archivo != archivo.name:
    st.session_state.resultado_separador = None
    st.session_state.ultimo_archivo = archivo.name

# Botón de procesar
if st.button("🚀 Separar Boletas y Facturas", use_container_width=True, type="primary"):
    if not archivo:
        st.error("❌ Por favor, sube un archivo.")
    elif st.session_state.procesando:
        st.warning("⏳ Ya se está procesando un archivo. Espera por favor.")
    else:
        st.session_state.procesando = True
        
        with st.spinner("Procesando archivo (puede tomar varios minutos)..."):
            try:
                temp_path = f"uploads/separador_{archivo.name}"
                with open(temp_path, "wb") as f:
                    f.write(archivo.getbuffer())

                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def actualizar_progreso(valor, mensaje):
                    progress_bar.progress(valor)
                    status_text.text(mensaje)
                
                es_excel = tipo_archivo == "Excel (xlsx)"
                resultado = procesar_archivo_separador(
                    temp_path,
                    fila_inicio=fila_inicio,
                    es_excel=es_excel,
                    progress_callback=actualizar_progreso
                )
                
                # Guardar todo en session_state
                st.session_state.resultado_separador = resultado
                st.session_state.archivo_procesado = archivo.name
                
                if resultado['success']:
                    # Guardar rutas de los archivos generados
                    st.session_state.ruta_boletas = resultado['boletas']['archivo']
                    st.session_state.ruta_otros = resultado['otros']['archivo']
                    
                    # Cargar DataFrames para previsualización
                    if resultado['boletas']['registros'] > 0 and os.path.exists(resultado['boletas']['archivo']):
                        st.session_state.df_boletas = pd.read_excel(
                            resultado['boletas']['archivo'], 
                            engine='openpyxl'
                        )
                    else:
                        st.session_state.df_boletas = pd.DataFrame()
                    
                    if resultado['otros']['registros'] > 0 and os.path.exists(resultado['otros']['archivo']):
                        st.session_state.df_otros = pd.read_excel(
                            resultado['otros']['archivo'], 
                            engine='openpyxl'
                        )
                    else:
                        st.session_state.df_otros = pd.DataFrame()
                
                # Limpiar archivo temporal
                try:
                    os.remove(temp_path)
                except:
                    pass
                
                # Marcar que ya no estamos procesando
                st.session_state.procesando = False
                
                if resultado['success']:
                    st.success("✅ Procesamiento completado exitosamente")
                    progress_bar.progress(1.0)
                    status_text.text("Proceso completado")
                    st.balloons()
                else:
                    st.error(f"❌ Error: {resultado.get('error', 'Error desconocido')}")
                    
            except MemoryError as e:
                st.error(f"❌ Error de memoria: {str(e)}")
                st.info("💡 Intenta con un archivo más pequeño o aumenta la memoria disponible.")
                st.session_state.procesando = False
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                with st.expander("Ver detalles técnicos"):
                    st.code(traceback.format_exc())
                st.session_state.procesando = False

# ============================================================================
# MOSTRAR RESULTADOS (SIEMPRE QUE ESTÉ PROCESADO)
# ============================================================================

if st.session_state.resultado_separador and st.session_state.resultado_separador['success']:
    
    resultado = st.session_state.resultado_separador
    
    st.markdown("### Resultados")
    
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.metric("Total registros", f"{resultado['total_registros']:,}")
    with col_b:
        st.metric("Boletas", f"{resultado['boletas']['registros']:,}")
    with col_c:
        st.metric("Facturas y otros", f"{resultado['otros']['registros']:,}")
    with col_d:
        st.metric("Tiempo", f"{resultado['elapsed_seconds']:.1f}s")
    
    st.markdown("---")
    st.markdown("### 📥 Descargar archivos generados")
    
    # Botones de descarga en dos columnas
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📄 Boletas")
        if resultado['boletas']['registros'] > 0:
            ruta_boletas = st.session_state.ruta_boletas
            if ruta_boletas and os.path.exists(ruta_boletas):
                with open(ruta_boletas, "rb") as f:
                    st.download_button(
                        label="📥 **Descargar Boletas**",
                        data=f,
                        file_name=f"boletas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                        key="btn_boletas"
                    )
                st.caption(f"{resultado['boletas']['registros']:,} registros en {resultado['boletas']['hojas']} hoja(s)")
            else:
                st.warning("⚠️ Archivo de boletas no disponible")
        else:
            st.info("📭 No se encontraron boletas en el archivo")
    
    with col2:
        st.markdown("#### 📄 Facturas y otros")
        if resultado['otros']['registros'] > 0:
            ruta_otros = st.session_state.ruta_otros
            if ruta_otros and os.path.exists(ruta_otros):
                with open(ruta_otros, "rb") as f:
                    st.download_button(
                        label="📥 **Descargar Facturas y otros**",
                        data=f,
                        file_name=f"facturas_&_otros.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="secondary",
                        key="btn_facturas"
                    )
                st.caption(f"{resultado['otros']['registros']:,} registros en {resultado['otros']['hojas']} hoja(s)")
            else:
                st.warning("⚠️ Archivo de facturas no disponible")
        else:
            st.info("📭 No se encontraron facturas u otros documentos")
    
    # ========================================================================
    # PREVISUALIZACIÓN DE DATOS
    # ========================================================================
    
    st.markdown("---")
    st.markdown("### Previsualización de Datos")
    
    tab1, tab2 = st.tabs(["1. Boletas", "2. Facturas y otros"])
    
    with tab1:
        if st.session_state.df_boletas is not None and not st.session_state.df_boletas.empty:
            st.dataframe(
                st.session_state.df_boletas.head(10),
                use_container_width=True,
                height=300
            )
            st.caption(f"Mostrando 10 de {len(st.session_state.df_boletas):,} registros")
        else:
            st.info("No hay boletas para mostrar")
    
    with tab2:
        if st.session_state.df_otros is not None and not st.session_state.df_otros.empty:
            st.dataframe(
                st.session_state.df_otros.head(10),
                use_container_width=True,
                height=300
            )
            st.caption(f"Mostrando 10 de {len(st.session_state.df_otros):,} registros")
        else:
            st.info("No hay facturas u otros para mostrar")
    
    # ========================================================================
    # BOTÓN PARA REINICIAR
    # ========================================================================
    
    if st.button("🔄 Procesar otro archivo", use_container_width=True):
        # Limpiar el estado pero mantener las variables inicializadas
        st.session_state.resultado_separador = None
        st.session_state.archivo_procesado = None
        st.session_state.ruta_boletas = None
        st.session_state.ruta_otros = None
        st.session_state.df_boletas = None
        st.session_state.df_otros = None
        st.session_state.procesando = False
        st.session_state.ultimo_archivo = None
        st.rerun()

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("### Notas importantes")
st.markdown("""
| Característica | Especificación |
|----------------|----------------|
| **Formatos soportados** | TXT (PLE) y Excel (xlsx) |
| **Identificación de boletas** | Serie que comienza con **'B'** |
| **Columna para Excel** | siempre en la **Columna G**  |
| **Posición para TXT** | **Posición 7** (índice 7) |
| **Límite de filas por hoja** | 1,048,000 (múltiples hojas automáticas) |
| **Archivos grandes** | Optimizado para >500 MB |
""")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #9e9e9e; padding: 20px;'>
    <p> <b>Área de Tributación - BN - Soporte Tecnico: Yessly Poma de la cruz</b></p>
    <p style='font-size: 12px;'>Separador PLE Ventas v2.1</p>
</div>
""", unsafe_allow_html=True)
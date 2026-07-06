# pages/7_📂_Separador_PLE_VENTAS.py
"""
Separador de PLE VENTAS
- Procesa archivos TXT o Excel (500MB+)
- Separa boletas (serie B) de facturas y otros
- Genera Exceles separados optimizados
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

st.title("📂 Separador de Comprobantes PLE Ventas")
st.markdown("""
Procesa archivos TXT o Excel y separa automáticamente:
- ✅ **Boletas** (Series que comienzan con 'B')
- ✅ **Facturas y otros** (Series que no comienzan con 'B')
- ✅ Soporte para **TXT** y **Excel** con múltiples hojas
- ✅ Para Excel: **SIEMPRE usa la columna G** para identificar la serie
- ✅ Optimizado para archivos **>500 MB**
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

# ============================================================================
# SELECCIÓN DE TIPO DE ARCHIVO
# ============================================================================

st.markdown("## 📤 Seleccionar archivo")

tipo_archivo = st.radio(
    "Tipo de archivo:",
    ["TXT (PLE)", "Excel (xlsx)"],
    horizontal=True,
    help="Selecciona el formato de tu archivo"
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
        help="Fila donde empiezan los datos (1-index)"
    )

# ============================================================================
# INFORMACIÓN DEL ARCHIVO
# ============================================================================

if archivo:
    tamano_mb = archivo.size / (1024 * 1024)
    st.info(f"""
    **📄 Archivo:** {archivo.name}  
    **📊 Tamaño:** {tamano_mb:.2f} MB ({archivo.size:,} bytes)  
    **📋 Tipo:** {tipo_archivo}
    **📋 Fila de inicio:** {fila_inicio}
    """)

# ============================================================================
# BOTÓN DE PROCESAMIENTO
# ============================================================================

st.markdown("---")

if st.button("🚀 Separar Boletas y Facturas", use_container_width=True, type="primary"):
    if not archivo:
        st.error("❌ Por favor, sube un archivo.")
    else:
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
                
                st.session_state.resultado_separador = resultado
                st.session_state.archivo_procesado = archivo.name
                
                try:
                    os.remove(temp_path)
                except:
                    pass
                
                if resultado['success']:
                    st.success("✅ Procesamiento completado exitosamente")
                    
                    st.markdown("### 📊 Resultados del Separador")
                    
                    col_a, col_b, col_c, col_d = st.columns(4)
                    with col_a:
                        st.metric("📊 Total registros", f"{resultado['total_registros']:,}")
                    with col_b:
                        st.metric("📋 Boletas", f"{resultado['boletas']['registros']:,}")
                    with col_c:
                        st.metric("📋 Facturas y otros", f"{resultado['otros']['registros']:,}")
                    with col_d:
                        st.metric("📁 Tipo", resultado.get('tipo_archivo', 'N/A'))
                    
                    st.markdown("---")
                    st.markdown("### 📥 Descargar archivos generados")
                    
                    # Botones de descarga en dos columnas
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### 📄 Boletas")
                        if resultado['boletas']['success'] and os.path.exists(resultado['boletas']['archivo']):
                            with open(resultado['boletas']['archivo'], "rb") as f:
                                st.download_button(
                                    label="📥 **Descargar Boletas**",
                                    data=f,
                                    file_name=os.path.basename(resultado['boletas']['archivo']),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    type="primary"
                                )
                            st.caption(f"✅ {resultado['boletas']['registros']:,} registros en {resultado['boletas']['hojas']} hoja(s)")
                        else:
                            st.info("No se encontraron boletas en el archivo")
                    
                    with col2:
                        st.markdown("#### 📄 Facturas y otros")
                        if resultado['otros']['success'] and os.path.exists(resultado['otros']['archivo']):
                            with open(resultado['otros']['archivo'], "rb") as f:
                                st.download_button(
                                    label="📥 **Descargar Facturas y otros**",
                                    data=f,
                                    file_name=os.path.basename(resultado['otros']['archivo']),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    type="primary"
                                )
                            st.caption(f"✅ {resultado['otros']['registros']:,} registros en {resultado['otros']['hojas']} hoja(s)")
                        else:
                            st.info("No se encontraron facturas u otros documentos")
                    
                    minutos = int(resultado['elapsed_seconds'] // 60)
                    segundos = int(resultado['elapsed_seconds'] % 60)
                    st.caption(f"⏱️ Tiempo de procesamiento: {minutos}m {segundos}s")
                    
                    progress_bar.progress(1.0)
                    status_text.text("✅ Proceso completado")
                    
                else:
                    st.error(f"❌ Error: {resultado.get('error', 'Error desconocido')}")
                    
            except MemoryError as e:
                st.error(f"❌ Error de memoria: {str(e)}")
                st.info("💡 Intenta con un archivo más pequeño o aumenta la memoria disponible.")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                with st.expander("Ver detalles técnicos"):
                    st.code(traceback.format_exc())

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("### 📖 Notas importantes")
st.markdown("""
| Característica | Especificación |
|----------------|----------------|
| **Formatos soportados** | TXT (PLE) y Excel (xlsx) |
| **Identificación de boletas** | Serie que comienza con **'B'** |
| **Columna para Excel** | **Columna G** (siempre) |
| **Límite de filas por hoja** | 1,048,000 (múltiples hojas automáticas) |
| **Archivos grandes** | Optimizado para >500 MB |
""")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #9e9e9e; padding: 20px;'>
    <p>🏦 <b>Área de Tributación - BN - Soporte Tecnico: Yessly Poma de la cruz</b></p>
    <p style='font-size: 12px;'>Separador PLE Ventas v2.1</p>
</div>
""", unsafe_allow_html=True)
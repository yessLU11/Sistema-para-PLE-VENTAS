# pages/1_🔄_Convertidor_SIRE.py
"""
CONVERTIDOR SIRE/PLE - SUNAT
Version 2.1 - Doble Formato + Archivos Grandes (hasta 2000MB)
"""
import streamlit as st
import pandas as pd
import os
import logging
import gc
import sys 
import time
from datetime import datetime
from pathlib import Path
import traceback

# 🚀 Importamos todo lo que está en el Core masivo
from src.sire_core import (
    conv_logger,
    SIREValidator,
    TXTProcessorConEncabezado,
    TXTProcessorSinEncabezado,
    ExcelGenerator
)

# ============================================================================
# CONFIGURACIÓN INICIAL
# ============================================================================

# Generar timestamp para nombres de archivos
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

# ============================================================================
# INICIALIZAR ESTADO DE SESIÓN (IMPORTANTE)
# ============================================================================

# Inicializar todas las variables de estado al inicio
if 'formato_seleccionado' not in st.session_state:
    st.session_state.formato_seleccionado = None
if 'mostrar_paso2' not in st.session_state:
    st.session_state.mostrar_paso2 = False
if 'archivo_cargado' not in st.session_state:
    st.session_state.archivo_cargado = None

# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.set_page_config(
    page_title="Convertidor SIRE/PLE",
    page_icon="🔄",
    layout="wide"
)

st.markdown('<p class="header-title">🏦 Convertidor SIRE/PLE - SUNAT</p>', unsafe_allow_html=True)
st.markdown('<p class="header-subtitle">Área de Tributación - Banco de la Nación</p>', unsafe_allow_html=True)
st.markdown("---")

# Información en sidebar
with st.sidebar:
    st.markdown("### 📋 INFORMACIÓN")
    st.info("""
**Versión:** 2.1

**Características:**
- ✅ Formato CON encabezados
- ✅ Formato SIN encabezados  
- ✅ Archivos hasta 2000MB
- ✅ Múltiples hojas Excel
- ✅ Optimizado 1M+ filas
    """)
    st.markdown("---")
    st.markdown("### 📊 ESTADÍSTICAS")
    st.write("📁 Archivos procesados: ", len(list(Path("output_files").glob("*.xlsx"))))

# ============================================================================
# PASO 1: SELECCIÓN DE FORMATO (Siempre visible)
# ============================================================================

st.markdown("### 📂 **Paso 1: Selecciona el Formato de tu Archivo TXT**")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="opcion-card">
        <div class="opcion-title">📋 Formato 1: CON Encabezados</div>
        <div class="opcion-desc">Primera línea contiene nombres de columnas</div>
        <div class="opcion-desc" style="margin-top:10px;color:#666;">Ej: Ruc|Razón Social|Total CP|...</div>
    </div>
    """, unsafe_allow_html=True)
    
    # ✅ ELIMINAR st.rerun() - Solo actualizar estado
    if st.button("📋 CON ENCABEZADOS", use_container_width=True, key="btn_formato1"):
        st.session_state.formato_seleccionado = "CON_ENCABEZADOS"
        st.session_state.mostrar_paso2 = True
        st.session_state.archivo_cargado = None  # Resetear archivo

with col2:
    st.markdown("""
    <div class="opcion-card">
        <div class="opcion-title">📋 Formato 2: SIN Encabezados</div>
        <div class="opcion-desc">Archivo inicia directamente con datos</div>
        <div class="opcion-desc" style="margin-top:10px;color:#666;">Ej: 1|20260500|651-18264653-14|...</div>
    </div>
    """, unsafe_allow_html=True)
    
    # ✅ ELIMINAR st.rerun() - Solo actualizar estado
    if st.button("📋 SIN ENCABEZADOS", use_container_width=True, key="btn_formato2"):
        st.session_state.formato_seleccionado = "SIN_ENCABEZADOS"
        st.session_state.mostrar_paso2 = True
        st.session_state.archivo_cargado = None  # Resetear archivo

# ============================================================================
# PASO 2: CARGA DE ARCHIVO (Se muestra automáticamente)
# ============================================================================

# ✅ MOSTRAR PASO 2 cuando se ha seleccionado un formato
if st.session_state.mostrar_paso2 and st.session_state.formato_seleccionado:
    
    formato = st.session_state.formato_seleccionado
    
    st.markdown("---")
    st.markdown(f"### 📂 **Paso 2: Cargar archivo ({formato})**")
    
    # Botón para cambiar de formato (volver al paso 1)
    if st.button("🔄 Cambiar formato", use_container_width=False):
        st.session_state.formato_seleccionado = None
        st.session_state.mostrar_paso2 = False
        st.session_state.archivo_cargado = None
        st.rerun()
    
    col1, col2 = st.columns([3, 1])
    with col1:
        uploaded_file = st.file_uploader(
            f"Selecciona el archivo TXT ({formato})",
            type=["txt"],
            help=f"Archivo en formato SIRE - {formato}",
            key="file_uploader_main"  # ✅ Key única para evitar conflictos
        )

    # ========================================================================
    # PASO 3: PROCESAMIENTO (Cuando se carga un archivo)
    # ========================================================================
    
    if uploaded_file is not None:
        # Guardar archivo en session_state
        st.session_state.archivo_cargado = uploaded_file
        
        # Info del archivo
        tamano_mb = uploaded_file.size / (1024 * 1024)
        st.markdown(f"""
        <div class="info-box">
            <b>📄 Archivo:</b> {uploaded_file.name}<br>
            <b>📊 Tamaño:</b> {tamano_mb:.2f} MB ({uploaded_file.size:,} bytes)<br>
            <b>📋 Formato:</b> {formato}
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### ⚙️ **Paso 3: Procesar y Convertir**")
        
        if st.button("🚀 CONVERTIR A EXCEL", use_container_width=True, type="primary"):
            start_time = time.perf_counter()
            progress_bar = st.progress(0.0)
            status_text = st.empty()

            def update_progress(value, message=''):
                try:
                    progress_bar.progress(value)
                except Exception:
                    pass
                if message:
                    status_text.info(message)

            with st.spinner('⏳ Procesando datos, por favor espere...\n\n⚠️ Archivos grandes pueden tomar varios minutos.'):
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    update_progress(0.05, 'Guardando archivo temporal...')
                    
                    # Guardar temporalmente
                    temp_path = os.path.join("input_files", "temp_upload.txt")
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    conv_logger.info(f"Iniciando conversión: {uploaded_file.name}")
                    conv_logger.info(f"Formato seleccionado: {formato}")
                    
                    # Elegir procesador según formato
                    if formato == "CON_ENCABEZADOS":
                        processor = TXTProcessorConEncabezado(max_size_mb=2000)
                    else:
                        processor = TXTProcessorSinEncabezado(max_size_mb=2000)
                    
                    update_progress(0.20, 'Leyendo el archivo TXT...')
                    df, error = processor.read_txt(temp_path)
                    
                    if error:
                        st.error(f"❌ Error al leer: {error}")
                    else:
                        update_progress(0.45, 'Validando datos...')
                        validator = SIREValidator()
                        validation = validator.validate_all(df)
                        
                        # Mostrar advertencias si hay
                        if validation['warnings']:
                            for warn in validation['warnings']:
                                st.warning(f"⚠️ {warn}")
                        
                        if not validation['is_valid']:
                            st.error("❌ Validación fallida:")
                            for err in validation['errors']:
                                st.error(f"  • {err}")
                        else:
                            # Generar Excel
                            nombre_original = uploaded_file.name
                            if nombre_original.lower().endswith('.txt'):
                                nombre_original = nombre_original[:-4]

                            output_filename = f"output_files/{nombre_original}.xlsx"
                            generator = ExcelGenerator()
                            success, message = generator.create_excel(
                                df,
                                output_filename,
                                progress_callback=update_progress
                            )
                            
                            elapsed = time.perf_counter() - start_time
                            update_progress(1.0, f'Conversión completada en {elapsed:.2f} segundos')
                            
                            if success:
                                # Limpiar temporal
                                if os.path.exists(temp_path):
                                    os.remove(temp_path)
                                
                                # Mostrar estadísticas
                                st.markdown("---")
                                st.markdown("### 📊 **Estadísticas de Conversión**")
                                
                                col1, col2, col3, col4 = st.columns(4)
                                with col1:
                                    st.metric("📝 Registros", f"{len(df):,}")
                                with col2:
                                    st.metric("📊 Columnas", len(df.columns))
                                with col3:
                                    st.metric("📄 Hojas Excel", generator.sheets_created)
                                with col4:
                                    st.metric("⏱️ Tiempo total", f"{elapsed:.2f} s")
                                
                                # Buscar columna de monto
                                monto_col = None
                                for col in ['MontoTotal', 'Total CP', 'Total']:
                                    if col in df.columns:
                                        monto_col = col
                                        break
                                if monto_col:
                                    try:
                                        total = pd.to_numeric(df[monto_col], errors='coerce').sum()
                                        st.metric("💰 Total", f"{total:,.2f}")
                                    except:
                                        st.metric("💰 Total", "N/A")
                                else:
                                    st.metric("💰 Total", "N/A")
                                
                                # Vista previa
                                st.markdown("---")
                                st.markdown("### 👀 **Vista Previa de Datos**")
                                st.dataframe(df.head(10), use_container_width=True, height=300)
                                
                                # Éxito
                                st.success("✅ ¡Conversión completada exitosamente! Archivo listo para descargar.")
                                
                                # Botón descarga
                                with open(output_filename, "rb") as f:
                                    st.download_button(
                                        label="📥 DESCARGAR EXCEL GENERADO",
                                        data=f,
                                        file_name=f"PLE_VENTAS_BN.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True,
                                        type="primary"
                                    )
                            else:
                                st.error(f"❌ {message}")
                                
                except Exception as e:
                    conv_logger.error(f"Excepción: {traceback.format_exc()}")
                    st.error(f"❌ Error inesperado: {str(e)}")
                    with st.expander("Ver detalles técnicos"):
                        st.code(traceback.format_exc())
else:
    # Mostrar mensaje si no se ha seleccionado formato
    if not st.session_state.mostrar_paso2:
        st.info("👆 Selecciona un formato arriba para continuar con la carga del archivo.")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #9e9e9e; padding: 20px;'>
    <p>🏦 <b>Área de Tributación - BN - Soporte Tecnico: Yessly Poma de la cruz</b></p>
    <p style='font-size: 12px;'>Convertidor SIRE/PLE v2.1 | Automatización de Declaraciones Tributarias</p>
    <p style='font-size: 11px; margin-top: 10px;'>© 2026 Todos los derechos reservados</p>
</div>
""", unsafe_allow_html=True)
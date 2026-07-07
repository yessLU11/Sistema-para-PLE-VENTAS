# pages/3_📅_Ordenar_Boletas.py
"""
Módulo: Ordenar Boletas por Día
Agrupa boletas de tipo 03 por fecha y establecimiento, genera un resumen.
"""

import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
from src.etl.processor import ordenar_boletas
from config import OUTPUT_DIR

# Configuración de la página
st.set_page_config(
    page_title="Ordenar Boletas",
    page_icon="📅",
    layout="wide"
)

st.title("📅 Ordenar Boletas por Día")
st.markdown("---")
st.markdown("""
Esta herramienta procesa un archivo Excel de boletas (PLE Ventas) y genera un resumen:
- Filtra solo documentos tipo **03**.
- Agrupa por **Fecha de Emisión** y **Código de Establecimiento** (B001, B007, ...).
- Para cada grupo:
  - Completa la columna **campo_10** (J) con el último número correlativo.
  - Suma los montos de **MontoOtrosConceptos** (T) y los coloca en T y **campo_25** (Z).
  - Genera un **IDComprobante** y **Serie** correlativos.
  - Asigna valor 1 a la columna **campo_35** (AJ).
- Genera un archivo Excel con una hoja por cada establecimiento.
""")
st.info("⚠️ El archivo debe contener las columnas: FechaEmision, TipoDoc, CodigoEstablecimiento, NumeroCorrelativo, MontoOtrosConceptos, IDComprobante, Serie y las demás columnas del PLE.")

# Área de carga
archivo_subido = st.file_uploader("⬆️📁 Sube tu archivo Excel (.xlsx)", type=['xlsx'])

if archivo_subido:
    st.markdown("---")
    st.markdown("##### ⚙️ Procesando archivo...")
    
    if st.button("🚀 Ordenar Boletas", use_container_width=True):
        start_time = time.perf_counter()
        progress_bar = st.progress(0.0)
        status_text = st.empty()

        def progress_callback(value, message=''):
            try:
                progress_bar.progress(value)
            except Exception:
                pass
            if message:
                status_text.info(message)

        with st.spinner("Procesando, por favor espera..."):
            resultado = ordenar_boletas(archivo_subido, progress_callback=progress_callback)
        
        elapsed = time.perf_counter() - start_time
        progress_callback(1.0, f'Finalizado en {elapsed:.2f} segundos')

        if resultado['success']:
            st.success("✅ Procesamiento completado exitosamente")
            
            # Mostrar estadísticas
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 Hojas generadas", len(resultado['sheets']))
            with col2:
                st.metric("📝 Filas totales", resultado['total_rows'])
            with col3:
                st.metric("⏱️ Tiempo total", f"{elapsed:.2f} s")
            
            st.markdown("**Establecimientos procesados:**")
            st.write(", ".join(resultado['sheets']))
            
            # Generar nombre de salida
            nombre_base = archivo_subido.name
            if nombre_base.lower().endswith('.xlsx'):
                nombre_base = nombre_base[:-5]
            output_filename = f"Ordenado_{nombre_base}.xlsx"
            
            # Botón de descarga
            st.markdown("---")
            st.download_button(
                label="📥 Descargar Excel ordenado",
                data=resultado['buffer'],
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Opcional: guardar también en output_files
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            with open(output_path, "wb") as f:
                f.write(resultado['buffer'].getbuffer())
            st.info(f"📁 Archivo guardado también en `{output_path}`")
            
        else:
            st.error(f"❌ Error: {resultado.get('error', 'Error desconocido')}")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #9e9e9e; padding: 20px;'>
    <p>🏦 <b>Área de Tributación - BN - Soporte Tecnico: Yessly Poma de la cruz</b></p>
    <p style='font-size: 12px;'>Ordenar Boletas v1.0</p>
</div>
""", unsafe_allow_html=True)
# pages/2_📊_Validador_Documental.py
"""
Módulo: Validador Documental
Detección de duplicados en documentos tipo 01 (Facturas) y 03 (Boletas)
"""

import streamlit as st
import pandas as pd
import os
import logging
import datetime
import time
from config import APP_TITLE, APP_ICON, INPUT_DIR, OUTPUT_DIR, LOG_DIR
import gc
import io
from src.etl.processor import procesar_excel

# Configuración de la página
st.set_page_config(
    page_title="Validador Documental",
    page_icon="📊",
    layout="wide"
)

# ============================================================================
# TÍTULO Y DESCRIPCIÓN AMIGABLE
# ============================================================================

st.title("📊 Validador Documental")
st.markdown("---")

# Explicación clara para el usuario
st.markdown ( """
### 🔍 ¿Qué hace esta herramienta?

Esta herramienta **revisa tu archivo Excel** y encuentra **registros duplicados** (comprobantes repetidos) en:
- **Facturas** (TipoDoc = 01)
- **Boletas** (TipoDoc = 03)

---

### 📋 ¿Qué archivo debes subir?

**Archivo Excel (.xlsx)** que contenga estas **3 columnas obligatorias**:

| Columna | Descripción | Ejemplo |
|---------|-------------|---------|
| **TipoDoc** | Tipo de documento | `01` (Factura), `03` (Boleta) |
| **CodigoEstablecimiento** | Código del establecimiento | `B001`, `F001` |
| **NumeroCorrelativo** | Número del comprobante | `42419003`, `00001067` |

---

### 🎯 ¿Dónde están las boletas y facturas?

| Tipo | TipoDoc | Código | Ejemplo |
|------|---------|--------|---------|
| 🧾 **Boleta** | `03` | Empieza con **`B`** | `B001`, `B025`, `B029` |
| 📄 **Factura** | `01` | Empieza con **`F`** | `F001`, `F029` |
| 📋 **Otros** | `07`, `08`, etc. | Otros códigos | `0128`, `0129` |

---

### 📝 Ejemplo de Excel válido

| TipoDoc | CodigoEstablecimiento | NumeroCorrelativo |
|---------|----------------------|-------------------|
| 03 | B001 | 42419003 |
| 03 | B001 | 42419004 |
| 01 | F001 | 00001067 |
| 01 | F001 | 00001068 |

"""
)
st.info("💡 **Importante:** Tu Excel puede tener más columnas, pero estas 3 son **OBLIGATORIAS** para que la herramienta funcione.")

st.markdown("---")

# ============================================================================
# ÁREA DE CARGA DE ARCHIVO
# ============================================================================

st.subheader("📤 Cargar archivo Excel")

# Mostrar ejemplo visual de columnas requeridas
with st.expander("👁️ Ver ejemplo de estructura requerida"):
    ejemplo_df = pd.DataFrame({
        'TipoDoc': ['03', '03', '01', '01'],
        'CodigoEstablecimiento': ['B001', 'B001', 'F001', 'F001'],
        'NumeroCorrelativo': ['42419003', '42419004', '00001067', '00001068'],
        'FechaEmision': ['16/06/2026', '16/06/2026', '26/06/2026', '26/06/2026'],
        'RazonSocialCliente': ['ARIAS OCHOA CESAR', 'VENEGAS CHAVEZ', 'CAJA MUNICIPAL', 'FINANCIERA']
    })
    st.dataframe(ejemplo_df, use_container_width=True)
    st.caption("✅ Las columnas en **negrita** son obligatorias")

archivo_subido = st.file_uploader(
    "Selecciona tu archivo Excel (.xlsx)",
    type=['xlsx'],
    help="Sube tu archivo Excel con los datos del PLE Ventas"
)

# Mostrar información del archivo si está cargado
if archivo_subido:
    tamano_mb = archivo_subido.size / (1024 * 1024)
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.metric("📁 Archivo", archivo_subido.name)
    with col_info2:
        st.metric("📊 Tamaño", f"{tamano_mb:.2f} MB")
    with col_info3:
        st.metric("📌 Formato", "Excel (.xlsx)")

# ============================================================================
# PROCESAMIENTO
# ============================================================================

if archivo_subido:
    st.markdown("---")
    st.subheader("🚀 Ejecutar Validación")
    
    # Botón para ejecutar
    if st.button("🔍 Buscar Duplicados", use_container_width=True, type="primary"):
        
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

        with st.spinner('🔍 Revisando tu archivo en busca de duplicados...'):
            resultado = procesar_excel(archivo_subido, progress_callback=progress_callback)

        elapsed = time.perf_counter() - start_time
        progress_callback(1.0, f'✅ Finalizado en {elapsed:.2f} segundos')

        # ====================================================================
        # MOSTRAR RESULTADOS
        # ====================================================================
        
        if resultado['success']:
            st.success("✅ ¡Procesamiento completado exitosamente!")
            
            # Dashboard de métricas
            st.subheader("📊 Resumen del Procesamiento")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📄 Total de registros", f"{resultado['total_registros']:,}")
            with col2:
                st.metric("📑 Hojas procesadas", len(resultado['hojas_procesadas']))
            with col3:
                st.metric("⚠️ Duplicados encontrados", resultado['duplicados_count'], 
                          delta="Revisar" if resultado['duplicados_count'] > 0 else "✅ OK",
                          delta_color="inverse" if resultado['duplicados_count'] > 0 else "normal")
            with col4:
                st.metric("⏱️ Tiempo total", f"{elapsed:.2f} s")

            # Mostrar advertencias si existen
            if resultado['errores']:
                with st.expander("⚠️ Advertencias del sistema"):
                    for err in resultado['errores']:
                        st.warning(err)

            # ================================================================
            # MOSTRAR DUPLICADOS ENCONTRADOS
            # ================================================================
            
            if resultado['duplicados_count'] > 0:
                st.subheader("🔍 Registros Duplicados Encontrados")
                
                st.markdown("""
                **Estos son los comprobantes que aparecen repetidos en tu archivo.**
                
                **Columnas:**
                - **Hoja_Origen**: En qué hoja está el registro
                - **_ExcelRow**: En qué fila de Excel está (la fila 1 es el encabezado)
                - **TipoDoc_Norm**: Tipo normalizado (F = Factura, B = Boleta)
                - **CodigoEstablecimiento_Norm**: Código del establecimiento
                - **NumeroCorrelativo**: Número del comprobante
                - **Clave_Unica**: Clave que identifica cada registro
                """)
                
                # Mostrar tabla de duplicados
                st.dataframe(
                    resultado['df_duplicados'],
                    use_container_width=True,
                    height=400
                )
                
                st.caption(f"Total de registros duplicados: {resultado['duplicados_count']:,}")
                
                # ============================================================
                # DESCARGA DE DUPLICADOS
                # ============================================================
                
                st.markdown("---")
                st.subheader("📥 Descargar Reporte de Duplicados")
                
                buffer_duplicados = io.BytesIO()
                with pd.ExcelWriter(buffer_duplicados, engine='xlsxwriter') as writer:
                    resultado['df_duplicados'].to_excel(writer, index=False, sheet_name='Duplicados')
                
                buffer_duplicados.seek(0)

                st.download_button(
                    label="📥 Descargar Archivo de Duplicados",
                    data=buffer_duplicados,
                    file_name=f"duplicados_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                st.info("💡 **Recomendación:** Revisa los duplicados y elimina los registros repetidos antes de continuar con el proceso.")
                
            else:
                # ============================================================
                # NO SE ENCONTRARON DUPLICADOS
                # ============================================================
                
                st.success("✅ ¡Excelente! No se detectaron duplicados en tus documentos 01 y 03.")
                
                st.markdown("""
                ### 🎯 ¿Qué significa esto?
                
                - Todos tus comprobantes tipo **Factura (01)** y **Boleta (03)** tienen números únicos
                - No hay registros repetidos que puedan causar problemas
                - Puedes continuar con el siguiente paso del proceso
                """)

                st.balloons()

        else:
            st.error(f"❌ Error: {resultado.get('message', 'Error desconocido')}")
            
            st.markdown("""
            ### 🔧 Posibles causas del error:
            
            1. **El archivo no tiene las columnas necesarias**
               - Verifica que tenga: `TipoDoc`, `CodigoEstablecimiento`, `NumeroCorrelativo`
            
            2. **El archivo está vacío o corrupto**
               - Intenta abrirlo en Excel para verificar que tenga datos
            
            3. **Formato incorrecto**
               - Asegúrate de que sea un archivo Excel (.xlsx)
            
            ### 💡 ¿Necesitas ayuda?
            Contacta al soporte técnico con el mensaje de error.
            """)

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #9e9e9e; padding: 20px;'>
    <p>🏦 <b>Área de Tributación - BN - Soporte Técnico: Yessly Poma de la Cruz</b></p>
    <p style='font-size: 12px;'>Validador Documental v1.0 | Detección de Duplicados</p>
</div>
""", unsafe_allow_html=True)
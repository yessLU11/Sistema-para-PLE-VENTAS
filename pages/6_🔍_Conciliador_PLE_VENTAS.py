# pages/6_🔍_Conciliador_PLE_VENTAS.py
"""
Conciliador para PLE VENTAS
- Optimizado para archivos masivos
- Soporte para múltiples hojas
- Reportes detallados con formato profesional
"""

import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime
from src.conciliador_compras import (
    leer_todas_hojas_conciliacion_compras,
    conciliar_archivos_compras,
    generar_reporte_presentes_no_presentes_sire_bn,
    generar_reporte_presentes_no_presentes_sire_sunat
)
from config import OUTPUT_DIR

st.set_page_config(
    page_title="Conciliador PLE VENTAS",
    page_icon="🔄",
    layout="wide"
)

st.title("Conciliador PLE VENTAS")
st.markdown("""
Compara **todos los datos** de dos archivos Excel (múltiples hojas) para **PLE VENTAS**.
- Optimizado para archivos masivos (500k+ registros)
- Usa columnas fijas **I (Serie)** y **J (Número)**
- Normaliza números a **8 dígitos**
- Reportes con formato profesional
""")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs("reportes", exist_ok=True)
os.makedirs("uploads", exist_ok=True)

# ============================================================================
# INICIALIZAR ESTADO DE SESIÓN
# ============================================================================

if "df_sire_sunat" not in st.session_state:
    st.session_state.df_sire_sunat = None
    st.session_state.df_sire_bn = None
    st.session_state.nombre_sire_sunat = None
    st.session_state.nombre_sire_bn = None
    st.session_state.fila_sire_sunat = 2
    st.session_state.fila_sire_bn = 2
    st.session_state.resultado_conciliacion = None

# ============================================================================
# CARGA DE ARCHIVOS
# ============================================================================

st.markdown("## Conciliación de PLE VENTAS")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Archivo 1 (SIRE_SUNAT)")
    archivo1 = st.file_uploader(
        "Seleccionar Excel - SIRE_SUNAT", 
        type=["xlsx", "xls"], 
        key="sire_sunat",
        help="Archivo con los datos de SUNAT"
    )
    fila1 = st.number_input(
        "Fila de inicio (Archivo 1)", 
        min_value=1, 
        value=2, 
        step=1, 
        key="fila_sunat",
        help="Indica la fila donde empiezan los DATOS (no los encabezados). Por defecto 2."
    )

with col2:
    st.subheader("📂 Archivo 2 (SIRE_BN)")
    archivo2 = st.file_uploader(
        "Seleccionar Excel - SIRE_BN", 
        type=["xlsx", "xls"], 
        key="sire_bn",
        help="Archivo con los datos del Banco de la Nación"
    )
    fila2 = st.number_input(
        "Fila de inicio (Archivo 2)", 
        min_value=1, 
        value=2, 
        step=1, 
        key="fila_bn",
        help="Indica la fila donde empiezan los DATOS (no los encabezados). Por defecto 2."
    )

# ============================================================================
# DETECTAR CAMBIOS EN ARCHIVOS O FILAS
# ============================================================================

if archivo1 and archivo2:
    if (st.session_state.nombre_sire_sunat != archivo1.name or 
        st.session_state.nombre_sire_bn != archivo2.name or
        st.session_state.fila_sire_sunat != fila1 or
        st.session_state.fila_sire_bn != fila2):
        st.session_state.df_sire_sunat = None
        st.session_state.df_sire_bn = None
        st.session_state.nombre_sire_sunat = archivo1.name
        st.session_state.nombre_sire_bn = archivo2.name
        st.session_state.fila_sire_sunat = fila1
        st.session_state.fila_sire_bn = fila2
        st.session_state.resultado_conciliacion = None

# ============================================================================
# BOTÓN 1: CONCILIACIÓN GENERAL (carga los DataFrames)
# ============================================================================

st.markdown("---")

if st.button("Cargar y Conciliar Archivos", use_container_width=True, type="primary"):
    if not archivo1 or not archivo2:
        st.error("❌ Por favor, sube ambos archivos.")
    else:
        with st.spinner("Procesando archivos..."):
            try:
                # Guardar archivos temporalmente
                path1 = f"uploads/conc_sunat_{archivo1.name}"
                path2 = f"uploads/conc_bn_{archivo2.name}"
                
                with open(path1, "wb") as f:
                    f.write(archivo1.getbuffer())
                with open(path2, "wb") as f:
                    f.write(archivo2.getbuffer())

                # Leer DataFrames
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def actualizar_progreso(valor, mensaje):
                    progress_bar.progress(valor)
                    status_text.text(mensaje)
                
                actualizar_progreso(0.1, "Leyendo SIRE_SUNAT...")
                df1 = leer_todas_hojas_conciliacion_compras(
                    path1, fila_inicio=fila1,
                    progress_callback=lambda v, m: actualizar_progreso(0.1 + 0.35 * v, f"SIRE_SUNAT: {m}")
                )
                
                actualizar_progreso(0.5, "Leyendo SIRE_BN...")
                df2 = leer_todas_hojas_conciliacion_compras(
                    path2, fila_inicio=fila2,
                    progress_callback=lambda v, m: actualizar_progreso(0.5 + 0.35 * v, f"SIRE_BN: {m}")
                )

                # Guardar en session_state
                st.session_state.df_sire_sunat = df1
                st.session_state.df_sire_bn = df2
                st.session_state.nombre_sire_sunat = archivo1.name
                st.session_state.nombre_sire_bn = archivo2.name

                # Conciliar
                actualizar_progreso(0.85, "Conciliando...")
                resultado = conciliar_archivos_compras(
                    df1, df2,
                    nombre1=archivo1.name,
                    nombre2=archivo2.name
                )
                st.session_state.resultado_conciliacion = resultado

                # Mostrar resumen
                actualizar_progreso(0.95, "Generando resumen...")
                resumen = resultado['resumen']
                
                st.success("✅ Archivos cargados y conciliados correctamente")
                st.markdown("### Resumen de Conciliación")
                
                col_a, col_b, col_c, col_d = st.columns(4)
                with col_a:
                    st.metric(f"📊 {archivo1.name}", f"{resumen['archivo1']['total_registros']:,}")
                with col_b:
                    st.metric(f"📊 {archivo2.name}", f"{resumen['archivo2']['total_registros']:,}")
                with col_c:
                    st.metric("⚖️ Diferencia", f"{resumen['archivo1']['total_registros'] - resumen['archivo2']['total_registros']:,}")
                with col_d:
                    st.metric("🔑 Coincidencias", f"{resumen['coincidencias']:,}")
                
                actualizar_progreso(1.0, "✅ Proceso completado")
                status_text.empty()
                progress_bar.empty()

                # Limpiar archivos temporales
                try:
                    os.remove(path1)
                    os.remove(path2)
                except:
                    pass

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                import traceback
                with st.expander("Ver detalles técnicos"):
                    st.code(traceback.format_exc())

# ============================================================================
# BOTÓN 2: REPORTE SIRE_BN
# ============================================================================

st.markdown("---")
st.markdown("### Reportes de Presencia")

if st.session_state.df_sire_sunat is not None and st.session_state.df_sire_bn is not None:
    
    # Reporte SIRE_BN
    with st.container():
        col_btn1, col_desc1 = st.columns([1, 3])
        
        with col_btn1:
            if st.button("📋 Reporte SIRE_BN", use_container_width=True):
                with st.spinner("Generando reporte SIRE_BN..."):
                    try:
                        df1 = st.session_state.df_sire_sunat
                        df2 = st.session_state.df_sire_bn
                        nombre1 = st.session_state.nombre_sire_sunat
                        nombre2 = st.session_state.nombre_sire_bn

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        nombre_salida = f"reportes/presentes_no_presentes_sire_bn_{timestamp}.xlsx"
                        
                        generar_reporte_presentes_no_presentes_sire_bn(
                            df_sire_sunat=df1,
                            df_sire_bn=df2,
                            nombre_sire_sunat=nombre1,
                            nombre_sire_bn=nombre2,
                            ruta_salida=nombre_salida
                        )

                        with open(nombre_salida, "rb") as f:
                            st.download_button(
                                label="📥 Descargar Reporte SIRE_BN",
                                data=f,
                                file_name=os.path.basename(nombre_salida),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        st.success("✅ Reporte SIRE_BN generado correctamente")
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
                        import traceback
                        with st.expander("Ver detalles técnicos"):
                            st.code(traceback.format_exc())
        
        with col_desc1:
            st.markdown("""
            **📋 Reporte SIRE_BN**  
            Muestra qué registros del **SIRE_BN** están **presentes** o **no presentes** en **SIRE_SUNAT**.  
            *Útil para identificar qué comprobantes declarados en PLE no están registrados en SIRE SUNAT.*
            """)
    
    # Reporte SIRE_SUNAT
    with st.container():
        col_btn2, col_desc2 = st.columns([1, 3])
        
        with col_btn2:
            if st.button("📋 Reporte SIRE_SUNAT", use_container_width=True):
                with st.spinner("Generando reporte SIRE_SUNAT..."):
                    try:
                        df1 = st.session_state.df_sire_sunat
                        df2 = st.session_state.df_sire_bn
                        nombre1 = st.session_state.nombre_sire_sunat
                        nombre2 = st.session_state.nombre_sire_bn

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        nombre_salida = f"reportes/presentes_no_presentes_sire_sunat_{timestamp}.xlsx"
                        
                        generar_reporte_presentes_no_presentes_sire_sunat(
                            df_sire_sunat=df1,
                            df_sire_bn=df2,
                            nombre_sire_sunat=nombre1,
                            nombre_sire_bn=nombre2,
                            ruta_salida=nombre_salida
                        )

                        with open(nombre_salida, "rb") as f:
                            st.download_button(
                                label="📥 Descargar Reporte SIRE_SUNAT",
                                data=f,
                                file_name=os.path.basename(nombre_salida),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        st.success("✅ Reporte SIRE_SUNAT generado correctamente")
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
                        import traceback
                        with st.expander("Ver detalles técnicos"):
                            st.code(traceback.format_exc())
        
        with col_desc2:
            st.markdown("""
            **📋 Reporte SIRE_SUNAT**  
            Muestra qué registros del **SIRE_SUNAT** están **presentes** o **no presentes** en **SIRE_BN**.  
            *Útil para identificar qué comprobantes registrados en SIRE SUNAT NO han sido declarados en el PLE.*
            """)

else:
    st.info("ℹ️ Primero carga y concilia los archivos usando el botón **'Cargar y Conciliar Archivos'** para poder acceder a los reportes SIRE SUNAT Y SIRE BN.")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("### 📖 Notas importantes para PLE VENTAS")
st.markdown("""
| Característica | Especificación |
|----------------|----------------|
| **Tipo de comprobante** | VENTAS (Facturas, Boletas, Notas) |
| **Columnas fijas** | **I = Serie**, **J = Número** |
| **Tipos de documento** | 01, 03, 07, 08, 02, 04, 05, 06, 09, 12, 14, 16, 30, 42, 46, 53, 87 |
| **Normalización de número** | 8 dígitos con ceros a la izquierda |
| **ID de conciliación** | `SERIE + NÚMERO_NORMALIZADO` |
""")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #9e9e9e; padding: 20px;'>
    <p>🏦 <b>Área de Tributación - BN - Soporte Tecnico: Yessly Poma de la cruz</b></p>
    <p style='font-size: 12px;'>Conciliador | SIRE BN | SIRE SUNAT</p>
    <p style='font-size: 11px; margin-top: 10px;'>© 2026 Todos los derechos reservados</p>
</div>
""", unsafe_allow_html=True)
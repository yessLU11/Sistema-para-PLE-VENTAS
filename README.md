# 🏦 Sistema de Procesamiento de Archivos para el PLE de Ventas

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)]()
[![License](https://img.shields.io/badge/license-MIT-green)]()
[![Status](https://img.shields.io/badge/status-Active-success)]()


# PLE VENTAS - Sistema de Procesamiento Tributario

## 📋 Descripción General

Sistema integral para el procesamiento de archivos pertenecientes a PLE Ventas, diseñado para el área de tributación del Banco de la Nación.

Diseñado por Poma De la cruz Yessly Josselyn **Área de Tributación del Banco de la Nación del Perú**, este sistema optimiza el procesamiento de grandes volúmenes de datos fiscales, permitiendo manejar archivos de hasta 500MB y más de 1 millón de filas.

## 🏗️ Arquitectura del Sistema
PLE_VENTAS/
├── app.py # Página principal
├── config.py # Configuración global
├── requirements.txt # Dependencias
├── pages/
│ ├── 1_🔄_Convertidor_SIRE.py # TXT → Excel
│ ├── 2_📊_Validador_Documental.py # Detección de duplicados
│ ├── 3_📅_Ordenar_Boletas.py # Ordenar boletas por día
│ ├── 4_📂_Separador_PLE_VENTAS.py # Separar B vs F
│ └── 5_🔍_Conciliador_PLE_VENTAS.py # PLE vs SUNAT
├── src/ # Lógica de negocio
│ ├── sire_core.py # Core de conversión
│ ├── separador_ple.py # Lógica de separación
│ ├── conciliador_compras.py # Lógica de conciliación
│ └── etl/
│ └── processor.py # Procesamiento ETL
├── input_files/ # Archivos de entrada
├── output_files/ # Archivos de salida
├── reportes/ # Reportes generados
├── logs/ # Logs del sistema
└── uploads/ # Uploads temporales

---

## 🔄 Flujo de Trabajo

1. **Convertir** - TXT → Excel (Módulo 1)
2. **Validar** - Detectar duplicados (Módulo 2)
3. **Ordenar** - Boletas por día (Módulo 3)
4. **Separar** - Boletas vs Facturas (Módulo 4)
5. **Conciliar** - PLE vs SUNAT (Módulo 5)

## 📦 Módulos

### 1️⃣ Convertidor SIRE → Excel
Convierte archivos TXT del SIRE a Excel estructurado.
- Soporte para archivos >500MB
- Formato CON/SIN encabezados
- Múltiples hojas automáticas

### 2️⃣ Validador Documental
Detecta duplicados en documentos tipo 01 y 03.
- Normaliza códigos automáticamente
- Reporte detallado de duplicados

### 3️⃣ Ordenar Boletas por Día
Agrupa boletas tipo 03 por fecha y establecimiento.
- Genera IDs correlativos
- Resumen por establecimiento

### 4️⃣ Separador PLE VENTAS
Separa boletas (B) de facturas y otros.
- Soporte TXT/Excel
- Descarga independiente

### 5️⃣ Conciliador PLE VENTAS
Compara PLE vs SUNAT.
- Reportes SIRE_BN y SIRE_SUNAT
- Formato profesional

## 🚀 Instalación

```bash
# Clonar repositorio
git clone [url-del-repo]

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar aplicación
streamlit run app.py

---

## 🛠️ Requisitos del Sistema

### Dependencias
- **Python 3.10+**
- **Streamlit 1.35.0** 
- **Archivos** hasta 2000MB
- **Memoria RAM** 8GB+ recomendado

### Especificaciones Recomendadas
- RAM: Mínimo 4GB (recomendado 8GB+)
- Espacio en disco: 1GB disponible
- Conexión a internet: No requerida (aplicación local)
```

---

```

La aplicación se abrirá en tu navegador en `http://localhost:8501`

### Flujo de Trabajo

#### **Opción 1: Convertidor SIRE/PLE**
1. Selecciona el formato de tu archivo TXT:
   - **CON encabezados**: Si tu archivo tiene una fila de títulos
   - **SIN encabezados**: Si es un archivo de datos puros
2. Sube tu archivo TXT (máximo 500MB)
3. Haz clic en "Procesar archivo"
4. Descarga tu archivo Excel generado

#### **Opción 2: Validador Documental**
1. Sube tu archivo Excel (.xlsx)
2. Haz clic en "Ejecutar Validación"
3. Revisa las métricas:
   - Total de filas leídas
   - Hojas procesadas
   - Duplicados detectados
4. Descarga el reporte de duplicados (si aplica)

---

## 📁 Estructura del Proyecto

```
codigo_PLE_VENTAS/
├── app.py                          # Aplicación principal (homepage)
├── config.py                       # Configuración global
├── requirements.txt                # Dependencias Python
├── pages/
│   ├── 1_🔄_Convertidor_SIRE.py  # Interfaz convertidor SIRE
│   └── 2_📊_Validador_Documental.py  # Interfaz validador
├── src/
│   ├── sire_core.py               # Motor core de conversión
│   └── etl/
│       └── processor.py           # Procesador ETL de validación
├── input_files/                   # Directorio para archivos de entrada
├── output_files/                  # Directorio para archivos procesados
├── logs/                          # Registros de operaciones
└── __pycache__/                   # Caché Python

```

---

## 🔧 Configuración

### Archivo `config.py`
Ajusta los parámetros según tus necesidades:

```python
# Información de la aplicación
APP_TITLE = "Convertidor SIRE/PLE - SUNAT"
APP_VERSION = "2.1"

# Límites de procesamiento
MAX_FILE_SIZE_MB = 500           # Tamaño máximo de archivo
MAX_ROWS_PER_SHEET = 1_000_000   # Máximo de filas por hoja
CHUNK_SIZE = 50000               # Tamaño de chunks para procesamiento

# Formato de archivo
ENCODING = 'latin-1'             # Codificación de entrada
SEPARATOR = '|'                  # Separador de campos TXT
```

---

## 📊 Módulos Principales

### `sire_core.py` - Motor de Conversión
Maneja el procesamiento completo de archivos SIRE:
- **ConversionLogger**: Sistema de logging especializado
- **SIREValidator**: Validación de formato SIRE
- **TXTProcessorConEncabezado**: Procesa TXT con headers
- **TXTProcessorSinEncabezado**: Procesa TXT sin headers
- **ExcelGenerator**: Genera archivos Excel optimizados

### `processor.py` - Procesador ETL
Valida y normaliza documentos Excel:
- Normalización de códigos
- Detección de duplicados
- Generación de estadísticas
- Exportación de reportes

---

## 💡 Ejemplos de Uso

### Conversión Básica
```python
from src.sire_core import SIREValidator, TXTProcessorConEncabezado, ExcelGenerator

# Validar archivo
validator = SIREValidator('input_files/archivo.txt')
if validator.validar():
    # Procesar
    processor = TXTProcessorConEncabezado('input_files/archivo.txt')
    df = processor.process()
    
    # Generar Excel
    generator = ExcelGenerator(df, 'output_files/resultado.xlsx')
    generator.generate()
```

---

## 📈 Rendimiento

### Capacidades de Procesamiento
| Métrica | Valor |
|---------|-------|
| Tamaño máximo de archivo | 500 MB |
| Máximo de filas por hoja | 1,000,000 |
| Tamaño de chunk de procesamiento | 50,000 filas |
| Encoding soportado | Latin-1, UTF-8 |
| Formato de separador | Pipe (\|) |

### Benchmarks (Aproximados)
- Archivo de 100MB (~500k filas): **2-5 minutos**
- Archivo de 250MB (~1M filas): **5-10 minutos**
- Validación de 10k registros: **1-2 segundos**

---

## 🐛 Solución de Problemas

### Error: "ModuleNotFoundError: No module named 'streamlit'"
```bash
pip install -r requirements.txt
```

### Error: "Permission denied" en input_files
```bash
# Verifica permisos de la carpeta
chmod 755 input_files output_files logs
```

### Error: "MemoryError" en archivos grandes
- Aumenta la RAM disponible
- Reduce el tamaño del chunk en `config.py`
- Divide el archivo en partes más pequeñas

### La conversión es muy lenta
- Verifica que no haya otros procesos consumiendo recursos
- Considera reducir el `MAX_ROWS_PER_SHEET` temporalmente

---

## 📝 Notas de Compatibilidad

### Formatos Soportados
| Formato | Entrada | Salida |
|---------|---------|--------|
| TXT (SIRE) | ✅ | - |
| XLSX (Excel) | ✅ | ✅ |
| CSV | - | - |

### Sistemas Operativos
- ✅ Windows 10/11
- ✅ Linux (Ubuntu 18.04+)
- ✅ macOS (10.13+)

---

## 🔐 Seguridad

### Medidas Implementadas
- ✅ Validación de formatos de archivo
- ✅ Limpieza de datos sensibles en logs
- ✅ Gestión de memoria para evitar desbordamientos
- ✅ Codificación de caracteres especiales
- ⚠️ Los archivos se almacenan localmente; no se envían a servidores externos

---

## 📞 Contacto y Soporte

**Desarrollado por:**Poma de la cruz Yessly Josselyn © 2026 Todos los derechos reservados Área de Tributación - Banco de la Nación del Perú

Para reportar bugs o sugerencias:
- 📧 Email: tributacion@bn.com.pe
- 📋 Issues: [GitHub Issues](https://github.com/yessLU11/Sistema-de-procesamiento-de-archivos-para-el-PLE-de-Ventas/issues)

---

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.

---

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Crea una rama para tu característica (`git checkout -b feature/mi-caracteristica`)
2. Realiza tus cambios y confirma (`git commit -m 'Agrega mi característica'`)
3. Empuja a la rama (`git push origin feature/mi-caracteristica`)
4. Abre un Pull Request

---

## 📚 Recursos Adicionales

- [Documentación de Streamlit](https://docs.streamlit.io/)
- [Documentación de Pandas](https://pandas.pydata.org/docs/)
- [SUNAT - PLE de Ventas](https://www.sunat.gob.pe/)
- [Especificaciones SIRE](https://www.sunat.gob.pe/orientacionaduanera/renta/)


---

**Última actualización:** Julio 2026  
**Versión:** 2.1  
**Mantenedor:** Banco de la Nación del Perú

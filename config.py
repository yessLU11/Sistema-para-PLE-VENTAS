import os

# Información de la aplicación
APP_TITLE = "Convertidor SIRE/PLE - SUNAT"
APP_ICON = "🏦"
APP_DESCRIPTION = "Convertidor profesional de archivos SIRE a Excel"
APP_VERSION = "2.1"
APP_AUTHOR = "Área de Tributación - Banco de la Nación del Perú"

# Directorios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input_files")
OUTPUT_DIR = os.path.join(BASE_DIR, "output_files")
LOG_DIR = os.path.join(BASE_DIR, "logs")

# Crear directorios si no existen
for dir_path in [INPUT_DIR, OUTPUT_DIR, LOG_DIR]:
    os.makedirs(dir_path, exist_ok=True)

# Límites y formatos
MAX_FILE_SIZE_MB = 500
MAX_ROWS_PER_SHEET = 1_000_000
CHUNK_SIZE = 50000
ENCODING = 'latin-1'
SEPARATOR = '|'
LAYOUT = "wide"
SIDEBAR_STATE = "expanded"
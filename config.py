# config.py
import os

APP_TITLE = "Convertidor SIRE/PLE - SUNAT"
APP_ICON = "🏦"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input_files")
OUTPUT_DIR = os.path.join(BASE_DIR, "output_files")
LOG_DIR = os.path.join(BASE_DIR, "logs")

for dir_path in [INPUT_DIR, OUTPUT_DIR, LOG_DIR]:
    os.makedirs(dir_path, exist_ok=True)

MAX_FILE_SIZE_MB = 2000
MAX_ROWS_PER_SHEET = 1_000_000
ENCODING = 'latin-1'
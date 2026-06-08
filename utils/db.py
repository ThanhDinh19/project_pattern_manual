from __future__ import annotations

import os
from pathlib import Path

import mysql.connector
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

BASE_PATH = os.getenv("BASE_PATH", "/PatternManual")

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

EXCEL_IMAGE_DIR = UPLOAD_DIR / "excel_images"
EXCEL_IMAGE_DIR.mkdir(parents=True, exist_ok=True)

FOLDER_IMPORT_DIR = UPLOAD_DIR / "folder_imports"
FOLDER_IMPORT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xlsm"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
PDF_EXTENSIONS = {".pdf"}

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "127.0.0.1"),
    "user": os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASSWORD", ""),
    "database": os.getenv("DB_NAME", "manual_db"),
    "charset": os.getenv("DB_CHARSET", "utf8mb4"),
}


def get_db_connection():
    return mysql.connector.connect(**DB_CONFIG)
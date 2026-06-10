from __future__ import annotations

import os
from pathlib import Path

import pyodbc
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

DB_DRIVER = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = os.getenv("DB_PORT", "").strip()
DB_NAME = os.getenv("DB_NAME", "manual_db")
DB_USER = os.getenv("DB_USER", "sa")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_TRUST_SERVER_CERTIFICATE = os.getenv("DB_TRUST_SERVER_CERTIFICATE", "yes")


def get_db_connection():
    server = DB_HOST if not DB_PORT else f"{DB_HOST},{DB_PORT}"

    conn_str = (
        f"DRIVER={{{DB_DRIVER}}};"
        f"SERVER={server};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASSWORD};"
        f"TrustServerCertificate={DB_TRUST_SERVER_CERTIFICATE};"
    )
    return pyodbc.connect(conn_str)
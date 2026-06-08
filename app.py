from __future__ import annotations

from flask import Flask, redirect, jsonify

from routes.admin_routes import register_admin_routes
from routes.auth_routes import register_auth_routes
from routes.excel_routes import register_excel_routes
from routes.folder_routes import register_folder_routes
from routes.reset_routes import register_reset_routes
from routes.state_routes import register_state_routes
from services.excel_service import restore_state_from_mysql
from services.folder_service import restore_folder_state_from_disk
from utils.db import BASE_PATH
import os
from dotenv import load_dotenv
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key")
    app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024 * 1024  # 5 GB

    @app.errorhandler(413)
    def request_entity_too_large(error):
        return jsonify({"success": False, "message": "Thư mục tải lên quá lớn (tối đa 5GB)"}), 413

    @app.get("/")
    def root():
        return redirect(f"{BASE_PATH}/login")

    register_auth_routes(app)
    register_admin_routes(app)
    register_excel_routes(app)
    register_folder_routes(app)
    register_reset_routes(app)
    register_state_routes(app)

    return app


app = create_app()

if __name__ == "__main__":
    try:
        restore_state_from_mysql()
        print("[OK] MySQL state restored successfully")
    except Exception as e:
        print(f"[WARN] Failed to restore MySQL state: {e}")
    
    try:
        restore_folder_state_from_disk()
        print("[OK] Folder state restored successfully")
    except Exception as e:
        print(f"[WARN] Failed to restore folder state: {e}")
    
    app.run(host="0.0.0.0", port=8000, debug=True, use_reloader=True)

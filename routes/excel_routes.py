from __future__ import annotations

from pathlib import Path

from flask import jsonify, request
from werkzeug.utils import secure_filename

from services.audit_service import write_audit_log
from services.excel_service import allowed_file, parse_excel_file, save_workbook_to_sqlserver
from utils.db import BASE_PATH, UPLOAD_DIR
from utils.helpers import STATE, get_public_imports_state


def register_excel_routes(app):
    @app.post(f"{BASE_PATH}/api/excel/upload")
    def upload_excel():
        uploaded_file = request.files.get("file")
        if not uploaded_file or uploaded_file.filename == "":
            return jsonify({"success": False, "message": "Bạn chưa chọn file Excel"}), 400
        if not allowed_file(uploaded_file.filename):
            return jsonify({"success": False, "message": "Chỉ hỗ trợ file .xlsx hoặc .xlsm"}), 400

        safe_name = secure_filename(uploaded_file.filename)
        file_path = UPLOAD_DIR / safe_name
        uploaded_file.save(file_path)

        try:
            result = parse_excel_file(Path(file_path))
            full_workbook = next((item for item in STATE["imports"] if item["id"] == result["id"]), None)
            if full_workbook:
                save_workbook_to_sqlserver(full_workbook)

            write_audit_log(
                action_type="IMPORT_EXCEL",
                action_label="Import Excel",
                target_type="excel_file",
                target_value=uploaded_file.filename,
                details={
                    "workbook_id": result.get("id"),
                    "file_name": result.get("fileName"),
                    "sheet_count": result.get("sheetCount"),
                    "total_rows": result.get("totalRows"),
                    "image_index_count": result.get("imageIndexCount"),
                },
            )

            return jsonify(
                {
                    "success": True,
                    "message": "Import Excel thành công, dữ liệu mới đã được cộng dồn",
                    "data": {
                        "workbook": result,
                        "imports": get_public_imports_state(),
                    },
                }
            )
        except Exception as error:
            import traceback

            traceback.print_exc()
            return jsonify({"success": False, "message": f"Không đọc được file Excel: {error}"}), 500
        finally:
            if file_path.exists():
                file_path.unlink(missing_ok=True)

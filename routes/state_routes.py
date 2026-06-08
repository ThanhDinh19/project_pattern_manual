from __future__ import annotations

from flask import jsonify, request

from services.excel_service import restore_state_from_mysql, search_by_uploaded_image
from services.folder_service import restore_folder_state_from_disk
from utils.db import BASE_PATH
from utils.helpers import STATE, get_public_imports_state, get_public_workbook_state, get_reset_options


def register_state_routes(app):
    @app.post(f"{BASE_PATH}/api/search-by-image")
    def search_by_image():
        pasted_image = request.files.get("image")
        try:
            result = search_by_uploaded_image(pasted_image)
            status = result.pop("status", 200)
            return jsonify(result), status
        except Exception as error:
            import traceback

            traceback.print_exc()
            return jsonify({"success": False, "message": f"Lỗi xử lý ảnh: {error}"}), 500

    @app.get(f"{BASE_PATH}/api/state")
    def get_state():
        if not STATE["imports"]:
            restore_state_from_mysql()
        if STATE["imports"] and not STATE["folder_index"]:
            restore_folder_state_from_disk()

        imports = get_public_imports_state()
        active_workbook = get_public_workbook_state()
        return jsonify(
            {
                "success": True,
                "data": {
                    "imports": imports,
                    "workbook": active_workbook,
                    "hasWorkbook": len(imports) > 0,
                    "importCount": len(imports),
                    "totalImageIndexCount": len(STATE.get("search_index", [])),
                    "resetOptions": get_reset_options(),
                },
            }
        )

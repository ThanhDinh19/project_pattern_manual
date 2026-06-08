from __future__ import annotations

import zipfile
from io import BytesIO

from flask import abort, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from services.audit_service import write_audit_log
from services.folder_service import (
    build_folder_index,
    map_folders_to_rows_by_style_no,
    restore_folder_state_from_disk,
    save_uploaded_folder_files,
)
from utils.db import BASE_PATH, UPLOAD_DIR
from utils.helpers import STATE, get_active_workbook, get_public_imports_state, get_public_workbook_state, natural_sort_key, safe_resolve_file, slugify_filename


def register_folder_routes(app):
    @app.get(f"{BASE_PATH}/media/<path:rel_path>")
    def media_file(rel_path: str):
        try:
            file_path = safe_resolve_file(UPLOAD_DIR, rel_path)
            return send_file(file_path, as_attachment=False, download_name=file_path.name)
        except Exception:
            abort(404)

    @app.get(f"{BASE_PATH}/folder/<folder_id>")
    def folder_detail(folder_id: str):
        folder = STATE["folder_index"].get(folder_id)
        if not folder:
            abort(404)
        return render_template("folder_detail.html", folder=folder, base_path=BASE_PATH)

    @app.get(f"{BASE_PATH}/api/folder/<folder_id>")
    def folder_detail_api(folder_id: str):
        folder = STATE["folder_index"].get(folder_id)
        if not folder:
            return jsonify({"success": False, "message": "Không tìm thấy folder"}), 404
        return jsonify({
            "success": True,
            "data": {
                "id": folder["id"],
                "name": folder["name"],
                "fileCount": folder["fileCount"],
                "files": [
                    {
                        "name": f["name"],
                        "relPath": f["relPath"],
                        "ext": f["ext"],
                        "sizeKb": f["sizeKb"],
                        "isImage": f["isImage"],
                        "isPdf": f["isPdf"],
                        "viewUrl": f["viewUrl"],
                        "downloadUrl": f["downloadUrl"],
                    }
                    for f in folder.get("files", [])
                ],
                "downloadAllUrl": f"{BASE_PATH}/folder-download/{folder_id}",
            },
        })

    @app.get(f"{BASE_PATH}/folder-file/<folder_id>/<path:rel_path>")
    def folder_file(folder_id: str, rel_path: str):
        folder = STATE["folder_index"].get(folder_id)
        if not folder:
            abort(404)
        try:
            file_path = safe_resolve_file(folder["path"], rel_path)
            return send_file(file_path, as_attachment=False, download_name=file_path.name)
        except Exception:
            abort(404)

    @app.get(f"{BASE_PATH}/folder-download/<folder_id>")
    def folder_download(folder_id: str):
        folder = STATE["folder_index"].get(folder_id)
        if not folder:
            abort(404)

        memory_file = BytesIO()
        with zipfile.ZipFile(memory_file, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(folder["path"].rglob("*"), key=lambda x: natural_sort_key(x.as_posix())):
                if not file_path.is_file():
                    continue
                arcname = f"{folder['name']}/{file_path.relative_to(folder['path']).as_posix()}"
                zf.write(file_path, arcname)

        memory_file.seek(0)
        zip_name = slugify_filename(folder["name"]) or "folder"
        return send_file(
            memory_file,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"{zip_name}.zip",
        )

    @app.get(f"{BASE_PATH}/folder-download-file/<folder_id>/<path:rel_path>")
    def folder_download_file(folder_id: str, rel_path: str):
        folder = STATE["folder_index"].get(folder_id)
        if not folder:
            abort(404)
        try:
            file_path = safe_resolve_file(folder["path"], rel_path)
            return send_file(file_path, as_attachment=True, download_name=file_path.name)
        except Exception:
            abort(404)

    @app.post(f"{BASE_PATH}/api/folder/import")
    def import_local_folder():
        workbook = get_active_workbook()
        if not workbook:
            return jsonify({"success": False, "message": "Bạn cần import Excel trước"}), 400

        files = request.files.getlist("files")
        relative_paths = request.form.getlist("relativePaths")
        root_folder_name = str(request.form.get("rootFolderName", "")).strip()

        if not files:
            return jsonify({"success": False, "message": "Bạn chưa chọn folder"}), 400
        if len(files) != len(relative_paths):
            return jsonify({"success": False, "message": "Dữ liệu folder upload không khớp"}), 400

        try:
            saved_root = save_uploaded_folder_files(files, relative_paths, root_folder_name)
            new_folder_index = build_folder_index(saved_root)

            existing_roots = STATE.get("folder_root_dir") or []
            if not isinstance(existing_roots, list):
                existing_roots = [existing_roots]
            existing_roots.append(saved_root)
            STATE["folder_root_dir"] = existing_roots

            merged_folder_index = dict(STATE.get("folder_index", {}))
            merged_folder_index.update(new_folder_index)
            STATE["folder_index"] = merged_folder_index

            mapped_count = 0
            for workbook_item in STATE.get("imports", []):
                mapped_count += map_folders_to_rows_by_style_no(workbook_item, STATE["folder_index"])
                workbook_item["folderImportName"] = root_folder_name or saved_root.name

            write_audit_log(
                action_type="IMPORT_FOLDER",
                action_label="Import Folder",
                target_type="folder_root",
                target_value=root_folder_name or saved_root.name,
                details={
                    "mapped_count": mapped_count,
                    "folder_count": len(new_folder_index),
                    "uploaded_file_count": len(files),
                },
            )

            return jsonify(
                {
                    "success": True,
                    "message": "Import folder thành công",
                    "data": {
                        "folderCount": len(STATE["folder_index"]),
                        "addedFolderCount": len(new_folder_index),
                        "mappedCount": mapped_count,
                        "workbook": get_public_workbook_state(),
                        "imports": get_public_imports_state(),
                    },
                }
            )
        except Exception as error:
            import traceback

            traceback.print_exc()
            return jsonify({"success": False, "message": f"Không import được folder: {error}"}), 500

from __future__ import annotations

from flask import jsonify, render_template, request
from werkzeug.security import generate_password_hash

from utils.auth import (
    login_required_api,
    login_required_page,
    permission_required_api,
    permission_required_page,
    row_to_dict,
)
from utils.db import BASE_PATH, get_db_connection
from utils.helpers import normalize_email


def register_admin_routes(app):
    @app.get(f"{BASE_PATH}/admin")
    @login_required_page
    @permission_required_page("can_manage_users")
    def admin_page():
        return render_template("admin.html", base_path=BASE_PATH)

    @app.get(f"{BASE_PATH}/api/admin/users")
    @login_required_api
    @permission_required_api("can_manage_users")
    def admin_list_users():
        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                """
                SELECT
                    id, email, full_name, is_active, is_admin,
                    can_import_excel, can_import_folder, can_search_image,
                    can_view_data, can_reset_data, can_manage_users,
                    can_view_audit_logs, created_at
                FROM users
                ORDER BY created_at DESC, id DESC
                """
            )
            db_rows = cursor.fetchall() or []
            users = [row_to_dict(cursor, r) for r in db_rows]
            return jsonify(
                {
                    "success": True,
                    "data": {
                        "users": [
                            {
                                "id": user["id"],
                                "email": user["email"],
                                "fullName": user.get("full_name") or "",
                                "isActive": bool(user["is_active"]),
                                "isAdmin": bool(user["is_admin"]),
                                "permissions": {
                                    "canImportExcel": bool(user["can_import_excel"]),
                                    "canImportFolder": bool(user["can_import_folder"]),
                                    "canSearchImage": bool(user["can_search_image"]),
                                    "canViewData": bool(user["can_view_data"]),
                                    "canResetData": bool(user["can_reset_data"]),
                                    "canManageUsers": bool(user["can_manage_users"]),
                                    "canViewAuditLogs": bool(user["can_view_audit_logs"]),
                                },
                                "createdAt": str(user["created_at"]),
                            }
                            for user in users
                        ]
                    },
                }
            )
        finally:
            cursor.close()
            conn.close()

    @app.post(f"{BASE_PATH}/api/admin/users")
    @login_required_api
    @permission_required_api("can_manage_users")
    def admin_create_user():
        payload = request.get_json(silent=True) or request.form
        email = normalize_email(payload.get("email"))
        password = str(payload.get("password", "")).strip()
        full_name = str(payload.get("full_name", "")).strip()

        if not email or not password:
            return jsonify({"success": False, "message": "Thiếu email hoặc mật khẩu"}), 400

        password_hash = generate_password_hash(password)
        can_import_excel = int(bool(payload.get("can_import_excel", True)))
        can_import_folder = int(bool(payload.get("can_import_folder", True)))
        can_search_image = int(bool(payload.get("can_search_image", True)))
        can_view_data = int(bool(payload.get("can_view_data", True)))
        can_reset_data = int(bool(payload.get("can_reset_data", False)))
        can_manage_users = int(bool(payload.get("can_manage_users", False)))
        can_view_audit_logs = int(bool(payload.get("can_view_audit_logs", False)))
        is_admin = int(bool(payload.get("is_admin", False)))
        is_active = int(bool(payload.get("is_active", True)))

        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                """
                INSERT INTO users (
                    email, password_hash, full_name, is_active, is_admin,
                    can_import_excel, can_import_folder, can_search_image,
                    can_view_data, can_reset_data, can_manage_users,
                    can_view_audit_logs
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    email,
                    password_hash,
                    full_name,
                    is_active,
                    is_admin,
                    can_import_excel,
                    can_import_folder,
                    can_search_image,
                    can_view_data,
                    can_reset_data,
                    can_manage_users,
                    can_view_audit_logs,
                ),
            )
            conn.commit()
            return jsonify({"success": True, "message": "Tạo tài khoản thành công"})
        except Exception as error:
            conn.rollback()
            return jsonify({"success": False, "message": f"Không tạo được tài khoản: {error}"}), 500
        finally:
            cursor.close()
            conn.close()

    @app.put(f"{BASE_PATH}/api/admin/users/<int:user_id>/permissions")
    @login_required_api
    @permission_required_api("can_manage_users")
    def admin_update_user_permissions(user_id: int):
        payload = request.get_json(silent=True) or {}
        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                """
                UPDATE users
                SET
                    full_name = ?,
                    is_active = ?,
                    is_admin = ?,
                    can_import_excel = ?,
                    can_import_folder = ?,
                    can_search_image = ?,
                    can_view_data = ?,
                    can_reset_data = ?,
                    can_manage_users = ?,
                    can_view_audit_logs = ?
                WHERE id = ?
                """,
                (
                    str(payload.get("full_name", "")).strip(),
                    int(bool(payload.get("is_active", True))),
                    int(bool(payload.get("is_admin", False))),
                    int(bool(payload.get("can_import_excel", True))),
                    int(bool(payload.get("can_import_folder", True))),
                    int(bool(payload.get("can_search_image", True))),
                    int(bool(payload.get("can_view_data", True))),
                    int(bool(payload.get("can_reset_data", False))),
                    int(bool(payload.get("can_manage_users", False))),
                    int(bool(payload.get("can_view_audit_logs", False))),
                    user_id,
                ),
            )
            conn.commit()
            return jsonify({"success": True, "message": "Cập nhật quyền thành công"})
        finally:
            cursor.close()
            conn.close()

    @app.put(f"{BASE_PATH}/api/admin/users/<int:user_id>/password")
    @login_required_api
    @permission_required_api("can_manage_users")
    def admin_reset_user_password(user_id: int):
        payload = request.get_json(silent=True) or {}
        new_password = str(payload.get("password", "")).strip()
        if not new_password:
            return jsonify({"success": False, "message": "Thiếu mật khẩu mới"}), 400

        password_hash = generate_password_hash(new_password)
        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "UPDATE users SET password_hash = ? WHERE id = ?",
                (password_hash, user_id),
            )
            conn.commit()
            return jsonify({"success": True, "message": "Đổi mật khẩu thành công"})
        finally:
            cursor.close()
            conn.close()

    @app.get(f"{BASE_PATH}/api/admin/audit-logs")
    @login_required_api
    @permission_required_api("can_view_audit_logs")
    def admin_audit_logs():
        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                """
                SELECT TOP 300
                    id, user_id, user_email,
                    action_type, action_label,
                    target_type, target_value,
                    details_json, ip_address,
                    created_at
                FROM audit_logs
                ORDER BY created_at DESC, id DESC
                """
            )
            db_rows = cursor.fetchall() or []
            rows = [row_to_dict(cursor, r) for r in db_rows]
            return jsonify({"success": True, "data": {"logs": rows}})
        finally:
            cursor.close()
            conn.close()

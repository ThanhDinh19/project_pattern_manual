from __future__ import annotations

from functools import wraps

from flask import jsonify, redirect, session, url_for

from utils.db import get_db_connection


def get_current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(
            "SELECT * FROM users WHERE id = %s AND is_active = 1",
            (user_id,),
        )
        return cursor.fetchone()
    finally:
        cursor.close()
        conn.close()



def user_to_public_payload(user: dict) -> dict:
    return {
        "id": user["id"],
        "email": user["email"],
        "fullName": user.get("full_name") or "",
        "isAdmin": bool(user.get("is_admin")),
        "permissions": {
            "canImportExcel": bool(user.get("can_import_excel")),
            "canImportFolder": bool(user.get("can_import_folder")),
            "canSearchImage": bool(user.get("can_search_image")),
            "canViewData": bool(user.get("can_view_data")),
            "canResetData": bool(user.get("can_reset_data")),
            "canManageUsers": bool(user.get("can_manage_users")),
            "canViewAuditLogs": bool(user.get("can_view_audit_logs")),
        },
    }



def login_required_page(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not get_current_user():
            return redirect(url_for("login_page"))
        return fn(*args, **kwargs)

    return wrapper



def login_required_api(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = get_current_user()
        if not user:
            return jsonify({"success": False, "message": "Bạn chưa đăng nhập"}), 401
        return fn(*args, **kwargs)

    return wrapper



def permission_required_api(permission_name: str):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            user = get_current_user()
            if not user:
                return jsonify({"success": False, "message": "Bạn chưa đăng nhập"}), 401

            if not bool(user.get(permission_name)):
                return jsonify(
                    {"success": False, "message": "Bạn không có quyền thực hiện thao tác này"}
                ), 403

            return fn(*args, **kwargs)

        return wrapper

    return decorator



def permission_required_page(permission_name: str):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            user = get_current_user()
            if not user:
                return redirect(url_for("login_page"))

            if not bool(user.get(permission_name)):
                return redirect(url_for("index"))

            return fn(*args, **kwargs)

        return wrapper

    return decorator

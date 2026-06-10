from __future__ import annotations

from flask import jsonify, render_template, request, session
from werkzeug.security import check_password_hash

from utils.auth import (
    get_current_user,
    login_required_api,
    login_required_page,
    user_to_public_payload,
    row_to_dict,
)
from utils.db import BASE_PATH, get_db_connection
from utils.helpers import normalize_email


def register_auth_routes(app):
    @app.get(f"{BASE_PATH}/login")
    def login_page():
        return render_template("login.html", base_path=BASE_PATH)

    @app.post(f"{BASE_PATH}/api/login")
    def api_login():
        payload = request.get_json(silent=True) or request.form
        email = normalize_email(payload.get("email"))
        password = str(payload.get("password", ""))

        if not email or not password:
            return jsonify({"success": False, "message": "Thiếu email hoặc mật khẩu"}), 400

        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "SELECT * FROM users WHERE email = ? AND is_active = 1",
                (email,),
            )
            row = cursor.fetchone()
            user = row_to_dict(cursor, row)
            if not user or not check_password_hash(user["password_hash"], password):
                return jsonify({"success": False, "message": "Sai email hoặc mật khẩu"}), 401

            session["user_id"] = user["id"]
            session["user_email"] = user["email"]
            return jsonify(
                {
                    "success": True,
                    "message": "Đăng nhập thành công",
                    "data": user_to_public_payload(user),
                }
            )
        finally:
            cursor.close()
            conn.close()

    @app.post(f"{BASE_PATH}/api/logout")
    def api_logout():
        session.clear()
        return jsonify({"success": True, "message": "Đã đăng xuất"})

    @app.get(f"{BASE_PATH}/")
    @login_required_page
    def index():
        return render_template("index.html", base_path=BASE_PATH)

    @app.get(f"{BASE_PATH}/api/me")
    @login_required_api
    def api_me():
        user = get_current_user()
        return jsonify({"success": True, "data": user_to_public_payload(user)})

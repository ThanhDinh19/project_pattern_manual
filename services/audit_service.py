from __future__ import annotations

import json

from flask import request

from utils.auth import get_current_user
from utils.db import get_db_connection


def write_audit_log(
    action_type: str,
    action_label: str,
    target_type: str | None = None,
    target_value: str | None = None,
    details: dict | None = None,
) -> None:
    user = get_current_user()
    user_id = user["id"] if user else None
    user_email = user["email"] if user else None
    ip_address = request.headers.get("X-Forwarded-For", request.remote_addr)

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            """
            INSERT INTO audit_logs (
                user_id, user_email,
                action_type, action_label,
                target_type, target_value,
                details_json, ip_address
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (
                user_id,
                user_email,
                action_type,
                action_label,
                target_type,
                target_value,
                json.dumps(details, ensure_ascii=False) if details else None,
                ip_address,
            ),
        )
        conn.commit()
    finally:
        cursor.close()
        conn.close()

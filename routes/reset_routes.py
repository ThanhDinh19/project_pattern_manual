from __future__ import annotations

from flask import jsonify, request

from services.audit_service import write_audit_log
from services.reset_service import hard_reset_all_data, hard_reset_by_customer, hard_reset_by_customer_and_season
from utils.db import BASE_PATH
from utils.helpers import get_public_imports_state, get_reset_options


def register_reset_routes(app):
    @app.post(f"{BASE_PATH}/api/reset")
    def reset_state():
        try:
            hard_reset_all_data()
            write_audit_log(
                action_type="RESET_ALL",
                action_label="Reset toàn bộ dữ liệu",
                target_type="system",
                target_value="all_data",
            )
            return jsonify({"success": True, "message": "Đã reset toàn bộ dữ liệu"})
        except Exception as error:
            import traceback

            traceback.print_exc()
            return jsonify({"success": False, "message": f"Không reset được dữ liệu: {error}"}), 500

    @app.post(f"{BASE_PATH}/api/reset/customer")
    def reset_customer_data():
        payload = request.get_json(silent=True) or request.form
        customer = str(payload.get("customer", "")).strip()
        if not customer:
            return jsonify({"success": False, "message": "Bạn chưa chọn customer"}), 400

        reset_result = hard_reset_by_customer(customer)
        if reset_result["removedRowCount"] == 0:
            return jsonify({"success": False, "message": "Không có dữ liệu để reset"}), 404

        write_audit_log(
            action_type="RESET_CUSTOMER",
            action_label="Reset theo customer",
            target_type="customer",
            target_value=customer,
            details={
                "customer": customer,
                "removed_row_count": reset_result["removedRowCount"],
                "affected_import_count": reset_result["affectedImportCount"],
                "removed_import_count": reset_result["removedImportCount"],
            },
        )
        return jsonify(
            {
                "success": True,
                "message": f"Đã reset {reset_result['removedRowCount']} dòng dữ liệu của customer {customer}",
                "data": {"imports": get_public_imports_state(), "resetOptions": get_reset_options()},
            }
        )

    @app.post(f"{BASE_PATH}/api/reset/season")
    def reset_customer_season_data():
        payload = request.get_json(silent=True) or request.form
        customer = str(payload.get("customer", "")).strip()
        season = str(payload.get("season", "")).strip()
        if not customer or not season:
            return jsonify({"success": False, "message": "Bạn cần chọn customer và season"}), 400

        reset_result = hard_reset_by_customer_and_season(customer, season)
        if reset_result["removedRowCount"] == 0:
            return jsonify({"success": False, "message": "Không có dữ liệu để reset"}), 404

        write_audit_log(
            action_type="RESET_SEASON",
            action_label="Reset theo customer / season",
            target_type="customer_season",
            target_value=f"{customer} / {season}",
            details={
                "customer": customer,
                "season": season,
                "removed_row_count": reset_result["removedRowCount"],
                "affected_import_count": reset_result["affectedImportCount"],
                "removed_import_count": reset_result["removedImportCount"],
            },
        )
        return jsonify(
            {
                "success": True,
                "message": f"Đã reset {reset_result['removedRowCount']} dòng dữ liệu của {customer} / {season}",
                "data": {"imports": get_public_imports_state(), "resetOptions": get_reset_options()},
            }
        )

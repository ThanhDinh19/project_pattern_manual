from __future__ import annotations

import os
import stat
import shutil
from pathlib import Path
from typing import Any

from services.excel_service import restore_state_from_mysql
from services.folder_service import restore_folder_state_from_disk
from utils.db import EXCEL_IMAGE_DIR, FOLDER_IMPORT_DIR, get_db_connection
from utils.helpers import STATE, clear_all_state

def _force_delete_dir(path_obj: Path) -> None:
    if not path_obj.exists():
        return
    def remove_readonly(func, path, excinfo):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception:
            pass
    try:
        shutil.rmtree(path_obj, onerror=remove_readonly)
    except Exception:
        try:
            shutil.rmtree(path_obj, ignore_errors=True)
        except Exception:
            pass


def _collect_folder_ids_from_state(customer: str, season: str | None = None) -> list[str]:
    folder_ids = []
    cust_lower = customer.strip().lower()
    seas_lower = season.strip().lower() if season else None

    for workbook in STATE.get("imports", []):
        for sheet in workbook.get("sheets", []):
            for row in sheet.get("rows", []):
                row_cust = str(row.get("Customer") or "").strip().lower()
                row_seas = str(row.get("Season") or "").strip().lower()

                match = (row_cust == cust_lower)
                if match and seas_lower:
                    match = (row_seas == seas_lower)

                if match:
                    fid = row.get("__folderId")
                    if fid:
                        folder_ids.append(fid)
    return list(set(folder_ids))


def hard_reset_all_data() -> None:
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SET FOREIGN_KEY_CHECKS = 0")
        cursor.execute("DELETE FROM pattern_row_images")
        cursor.execute("DELETE FROM pattern_rows")
        cursor.execute("DELETE FROM excel_sheets")
        cursor.execute("DELETE FROM excel_imports")
        cursor.execute("SET FOREIGN_KEY_CHECKS = 1")
        conn.commit()
    except Exception:
        conn.rollback()
        try:
            cursor.execute("SET FOREIGN_KEY_CHECKS = 1")
        except Exception:
            pass
        raise
    finally:
        cursor.close()
        conn.close()

    clear_all_state()
    for path_obj in [FOLDER_IMPORT_DIR, EXCEL_IMAGE_DIR]:
        try:
            if path_obj.exists():
                for child in path_obj.iterdir():
                    try:
                        if child.is_dir():
                            _force_delete_dir(child)
                        else:
                            os.chmod(child, stat.S_IWRITE)
                            child.unlink(missing_ok=True)
                    except Exception as child_err:
                        print(f"[ERROR] Failed to delete child {child}: {child_err}")
        except Exception as e:
            print(f"[ERROR] Failed to delete contents of {path_obj}: {e}")



def _reload_shared_state_after_partial_reset() -> None:
    STATE["imports"] = []
    STATE["search_index"] = []
    STATE["folder_index"] = {}
    STATE["folder_root_dir"] = []
    restore_state_from_mysql()
    restore_folder_state_from_disk()



def _hard_reset_rows(
    row_where_sql: str,
    row_where_params: tuple[Any, ...],
    joined_where_sql: str,
    joined_where_params: tuple[Any, ...],
    folder_ids: list[str] = [],
) -> dict[str, int]:
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(
            f"""
            SELECT COUNT(*) AS row_count
            FROM pattern_rows
            WHERE {row_where_sql}
            """,
            row_where_params,
        )
        removed_row_count = int((cursor.fetchone() or {}).get("row_count") or 0)
        if removed_row_count == 0:
            return {"removedRowCount": 0, "affectedImportCount": 0, "removedImportCount": 0}

        cursor.execute(
            f"""
            SELECT DISTINCT import_id
            FROM pattern_rows
            WHERE {row_where_sql}
            """,
            row_where_params,
        )
        affected_import_ids = [
            str(row["import_id"]) for row in (cursor.fetchall() or []) if row.get("import_id")
        ]

        cursor.execute("SET FOREIGN_KEY_CHECKS = 0")
        cursor.execute(
            f"""
            DELETE pri
            FROM pattern_row_images pri
            INNER JOIN pattern_rows pr ON pr.id = pri.pattern_row_id
            WHERE {joined_where_sql}
            """,
            joined_where_params,
        )
        cursor.execute(
            f"""
            DELETE FROM pattern_rows
            WHERE {row_where_sql}
            """,
            row_where_params,
        )

        removed_import_ids: list[str] = []
        if affected_import_ids:
            placeholders = ", ".join(["%s"] * len(affected_import_ids))
            cursor.execute(
                f"""
                DELETE es
                FROM excel_sheets es
                LEFT JOIN pattern_rows pr ON pr.sheet_id = es.id
                WHERE es.import_id IN ({placeholders})
                  AND pr.id IS NULL
                """,
                tuple(affected_import_ids),
            )

            for import_id in affected_import_ids:
                cursor.execute(
                    "SELECT COUNT(*) AS total_rows FROM pattern_rows WHERE import_id = %s",
                    (import_id,),
                )
                total_rows = int((cursor.fetchone() or {}).get("total_rows") or 0)

                cursor.execute(
                    "SELECT COUNT(*) AS sheet_count FROM excel_sheets WHERE import_id = %s",
                    (import_id,),
                )
                sheet_count = int((cursor.fetchone() or {}).get("sheet_count") or 0)

                cursor.execute(
                    """
                    SELECT COUNT(*) AS image_index_count
                    FROM pattern_row_images pri
                    INNER JOIN pattern_rows pr ON pr.id = pri.pattern_row_id
                    WHERE pr.import_id = %s
                    """,
                    (import_id,),
                )
                image_index_count = int((cursor.fetchone() or {}).get("image_index_count") or 0)

                if total_rows <= 0:
                    removed_import_ids.append(import_id)
                else:
                    cursor.execute(
                        """
                        UPDATE excel_imports
                        SET total_rows = %s, sheet_count = %s, image_index_count = %s
                        WHERE id = %s
                        """,
                        (total_rows, sheet_count, image_index_count, import_id),
                    )

            if removed_import_ids:
                placeholders_removed = ", ".join(["%s"] * len(removed_import_ids))
                cursor.execute(
                    f"DELETE FROM excel_imports WHERE id IN ({placeholders_removed})",
                    tuple(removed_import_ids),
                )

        cursor.execute("SET FOREIGN_KEY_CHECKS = 1")
        conn.commit()
    except Exception:
        conn.rollback()
        try:
            cursor.execute("SET FOREIGN_KEY_CHECKS = 1")
        except Exception:
            pass
        raise
    finally:
        cursor.close()
        conn.close()

    for import_id in removed_import_ids:
        try:
            image_dir = EXCEL_IMAGE_DIR / import_id
            _force_delete_dir(image_dir)
        except Exception:
            pass

    for fid in folder_ids:
        folder = STATE.get("folder_index", {}).get(fid)
        if folder and "path" in folder:
            try:
                folder_path = Path(folder["path"])
                relative = folder_path.relative_to(FOLDER_IMPORT_DIR)
                root_dir = FOLDER_IMPORT_DIR / relative.parts[0]
                _force_delete_dir(root_dir)
            except Exception:
                pass

    _reload_shared_state_after_partial_reset()
    return {
        "removedRowCount": removed_row_count,
        "affectedImportCount": len(affected_import_ids),
        "removedImportCount": len(removed_import_ids),
    }



def hard_reset_by_customer(customer: str) -> dict[str, int]:
    customer = str(customer or "").strip()
    if not customer:
        return {"removedRowCount": 0, "affectedImportCount": 0, "removedImportCount": 0}
    folder_ids = _collect_folder_ids_from_state(customer)
    return _hard_reset_rows(
        row_where_sql="customer = %s",
        row_where_params=(customer,),
        joined_where_sql="pr.customer = %s",
        joined_where_params=(customer,),
        folder_ids=folder_ids,
    )



def hard_reset_by_customer_and_season(customer: str, season: str) -> dict[str, int]:
    customer = str(customer or "").strip()
    season = str(season or "").strip()
    if not customer or not season:
        return {"removedRowCount": 0, "affectedImportCount": 0, "removedImportCount": 0}
    folder_ids = _collect_folder_ids_from_state(customer, season)
    return _hard_reset_rows(
        row_where_sql="customer = %s AND season = %s",
        row_where_params=(customer, season),
        joined_where_sql="pr.customer = %s AND pr.season = %s",
        joined_where_params=(customer, season),
        folder_ids=folder_ids,
    )

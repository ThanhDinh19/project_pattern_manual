from __future__ import annotations

import json
import re
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any


STATE: dict[str, Any] = {
    "imports": [],
    "search_index": [],
    "folder_index": {},
    "folder_root_dir": [],
}

DB_FIELD_MAPPING: dict[str, str] = {
    "No": "row_no",
    "Customer": "customer",
    "Season": "season",
    "Staff": "staff",
    "Style No": "style_no",
    "Style Name": "style_name",
    "Product": "product",
    "Categories": "categories",
    "Gender": "gender",
}

EXCEL_HEADER_ALIASES: dict[str, str] = {
    "no": "No",
    "stt": "No",
    "rowno": "No",
    "number": "No",
    "customer": "Customer",
    "customername": "Customer",
    "customercode": "Customer",
    "season": "Season",
    "staff": "Staff",
    "styleno": "Style No",
    "stylenumber": "Style No",
    "stylecode": "Style No",
    "stylename": "Style Name",
    "product": "Product",
    "category": "Categories",
    "categories": "Categories",
    "gender": "Gender",
    "gendermenwomen": "Gender",
    "genderwomenmen": "Gender",
    "gendermw": "Gender",
    "sketchdesign": "Sketch Design",
    "sketch": "Sketch Design",
}

STANDARD_HEADERS: set[str] = set(DB_FIELD_MAPPING.keys()) | {"Sketch Design"}


def normalize_email(email: str) -> str:
    return str(email or "").strip().lower()



def normalize_cell_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value



def normalize_header_key(header: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(header or "").strip().lower())



def normalize_excel_header(header: str) -> str:
    raw = str(header or "").strip()
    key = normalize_header_key(raw)
    return EXCEL_HEADER_ALIASES.get(key, raw)



def merge_extra_data_into_row(row_data: dict[str, Any], extra_data: dict[str, Any]) -> None:
    for raw_key, value in (extra_data or {}).items():
        canonical_key = normalize_excel_header(raw_key)
        if canonical_key in DB_FIELD_MAPPING:
            if row_data.get(canonical_key) in (None, "", []):
                row_data[canonical_key] = value
            continue
        if canonical_key == "Sketch Design":
            continue
        row_data[canonical_key] = value



def row_has_meaningful_data(row_data: dict[str, Any], headers: list[str]) -> bool:
    ignored_headers = {"no", "stt"}
    for header in headers:
        header_key = normalize_header_key(header)
        if header_key in ignored_headers:
            continue

        value = row_data.get(header)
        if isinstance(value, str):
            if value.strip() != "":
                return True
        elif value not in (None, "", []):
            return True
    return False



def detect_header_row(ws, scan_rows: int = 20, scan_cols: int = 50) -> int:
    best_row = 1
    best_count = -1
    max_row = min(ws.max_row, scan_rows)
    max_col = min(ws.max_column, scan_cols)

    for row_idx in range(1, max_row + 1):
        count = 0
        for col_idx in range(1, max_col + 1):
            value = ws.cell(row_idx, col_idx).value
            if value not in (None, ""):
                count += 1
        if count > best_count:
            best_count = count
            best_row = row_idx

    return best_row



def natural_sort_key(value: str) -> list[Any]:
    parts = re.split(r"(\d+)", str(value))
    result: list[Any] = []
    for part in parts:
        if part.isdigit():
            result.append(int(part))
        else:
            result.append(part.lower())
    return result



def normalize_match_value(value: Any) -> str:
    return str(value or "").strip()



def infer_partition_from_sheets(all_sheets: list[dict[str, Any]]) -> dict[str, Any]:
    customer_values: dict[str, int] = {}
    season_values: dict[str, int] = {}

    def add_count(counter: dict[str, int], value: Any) -> None:
        text = str(value or "").strip()
        if not text:
            return
        counter[text] = counter.get(text, 0) + 1

    for sheet in all_sheets:
        for row in sheet.get("rows", []):
            add_count(customer_values, row.get("Customer"))
            add_count(season_values, row.get("Season"))

    customers = sorted(customer_values.keys())
    seasons = sorted(season_values.keys())
    dominant_customer = max(customer_values, key=customer_values.get) if customer_values else None
    dominant_season = max(season_values, key=season_values.get) if season_values else None

    return {
        "customer": dominant_customer,
        "season": dominant_season,
        "customerValues": customers,
        "seasonValues": seasons,
        "isAmbiguous": len(customers) > 1 or len(seasons) > 1,
    }



def rebuild_global_search_index() -> None:
    merged_index: list[dict[str, Any]] = []
    for workbook in STATE.get("imports", []):
        merged_index.extend(workbook.get("search_index", []))
    STATE["search_index"] = merged_index



def get_active_workbook() -> dict[str, Any] | None:
    imports = STATE.get("imports", [])
    return imports[-1] if imports else None



def get_public_imports_state() -> list[dict[str, Any]]:
    result = []
    for workbook in STATE.get("imports", []):
        result.append({k: v for k, v in workbook.items() if k != "search_index"})
    return result



def get_public_workbook_state() -> dict[str, Any] | None:
    workbook = get_active_workbook()
    if not workbook:
        return None
    return {k: v for k, v in workbook.items() if k != "search_index"}



def extract_folder_match_keys(folder_name: str) -> list[str]:
    raw_name = normalize_match_value(folder_name)
    keys: list[str] = []

    def add_key(value: str) -> None:
        clean = normalize_match_value(value)
        if clean and clean not in keys:
            keys.append(clean)

    add_key(raw_name)
    main_part = re.split(r"\s*\(", raw_name, maxsplit=1)[0].strip()
    add_key(main_part)
    bracket_parts = re.findall(r"\(([^()]+)\)", raw_name)
    for part in bracket_parts:
        add_key(part)
    return keys



def safe_resolve_file(base_dir: Path, rel_path: str) -> Path:
    target = (base_dir / rel_path).resolve()
    base_resolved = base_dir.resolve()

    if base_resolved not in target.parents and target != base_resolved:
        raise FileNotFoundError("Đường dẫn không hợp lệ")
    if not target.exists() or not target.is_file():
        raise FileNotFoundError("Không tìm thấy file")
    return target



def clear_folder_state() -> None:
    folder_roots = STATE.get("folder_root_dir") or []
    if not isinstance(folder_roots, list):
        folder_roots = [folder_roots]

    for folder_root_dir in folder_roots:
        if folder_root_dir and Path(folder_root_dir).exists():
            shutil.rmtree(folder_root_dir, ignore_errors=True)

    STATE["folder_root_dir"] = []
    STATE["folder_index"] = {}

    for workbook in STATE.get("imports", []):
        for sheet in workbook.get("sheets", []):
            for row in sheet.get("rows", []):
                row["__folderId"] = None
                row["__folderName"] = None
                row["__detailUrl"] = None
        workbook["mappedFolderCount"] = 0
        workbook["folderImportName"] = None



def clear_all_state() -> None:
    clear_folder_state()
    STATE["imports"] = []
    STATE["search_index"] = []



def get_reset_options() -> dict[str, Any]:
    customer_map: dict[str, set[str]] = {}
    for workbook in STATE.get("imports", []):
        customer = str(workbook.get("customer") or "").strip() or "Chưa xác định"
        customer_map.setdefault(customer, set())
        season_values = workbook.get("seasonValues") or []
        if season_values:
            for season in season_values:
                season_text = str(season or "").strip() or "Chưa xác định"
                customer_map[customer].add(season_text)
        else:
            season = str(workbook.get("season") or "").strip() or "Chưa xác định"
            customer_map[customer].add(season)

    customers = sorted(customer_map.keys(), key=natural_sort_key)
    return {
        "customers": customers,
        "seasonsByCustomer": {
            customer: sorted(list(seasons), key=natural_sort_key)
            for customer, seasons in customer_map.items()
        },
    }



def normalize_relative_parts(relative_path: str) -> list[str]:
    parts = []
    for part in Path(relative_path).parts:
        clean = str(part).strip()
        if clean in ("", ".", ".."):
            continue
        parts.append(clean)
    return parts



def slugify_filename(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r"[^\w\-\.]+", "_", value, flags=re.UNICODE)
    return value.strip("_") or "sheet"



def json_loads_safe(value: Any) -> dict[str, Any]:
    if not value:
        return {}
    if isinstance(value, dict):
        return value
    if isinstance(value, str):
        try:
            return json.loads(value)
        except Exception:
            return {}
    return {}

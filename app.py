from __future__ import annotations

import base64
import re
import shutil
import uuid
from collections import defaultdict
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, render_template, request, send_file
from openpyxl import load_workbook
from PIL import Image, ImageChops, ImageOps, ImageStat
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

FOLDER_IMPORT_DIR = UPLOAD_DIR / "folder_imports"
FOLDER_IMPORT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xlsm"}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
PDF_EXTENSIONS = {".pdf"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB

STATE: dict[str, Any] = {
    "workbook": None,
    "search_index": [],
    "folder_index": {},
    "folder_root_dir": None,
}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_cell_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


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


def pil_to_data_url(image: Image.Image, max_width: int = 240) -> str:
    img = image.copy()

    if img.width > max_width:
        ratio = max_width / float(img.width)
        new_size = (max_width, max(1, int(img.height * ratio)))
        img = img.resize(new_size, Image.Resampling.LANCZOS)

    output = BytesIO()
    img.save(output, format="PNG", optimize=True)
    base64_string = base64.b64encode(output.getvalue()).decode("utf-8")
    return f"data:image/png;base64,{base64_string}"


def normalize_image(image: Image.Image) -> Image.Image:
    image = ImageOps.exif_transpose(image)

    if image.mode in ("RGBA", "LA"):
        background = Image.new("RGBA", image.size, (255, 255, 255, 255))
        image = Image.alpha_composite(background, image.convert("RGBA")).convert("RGB")
    else:
        image = image.convert("RGB")

    bg = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()

    if bbox:
        image = image.crop(bbox)

    return image


def load_normalized_image_from_bytes(image_bytes: bytes) -> Image.Image:
    with Image.open(BytesIO(image_bytes)) as img:
        normalized = normalize_image(img)
        return normalized.copy()


def build_compare_gray_png(image: Image.Image, size: tuple[int, int] = (128, 128)) -> bytes:
    compare_img = image.convert("L").resize(size, Image.Resampling.LANCZOS)
    output = BytesIO()
    compare_img.save(output, format="PNG", optimize=True)
    return output.getvalue()


def compute_dhash(image: Image.Image, hash_size: int = 16) -> int:
    gray = image.convert("L").resize((hash_size + 1, hash_size), Image.Resampling.LANCZOS)
    pixels = list(gray.getdata())

    value = 0
    for row in range(hash_size):
        row_offset = row * (hash_size + 1)
        for col in range(hash_size):
            left_px = pixels[row_offset + col]
            right_px = pixels[row_offset + col + 1]
            value = (value << 1) | (1 if left_px > right_px else 0)

    return value


def hamming_distance(left: int, right: int) -> int:
    return (left ^ right).bit_count()


def pixel_similarity(compare_a_png: bytes, compare_b_png: bytes) -> float:
    with Image.open(BytesIO(compare_a_png)) as img_a, Image.open(BytesIO(compare_b_png)) as img_b:
        diff = ImageChops.difference(img_a, img_b)
        mean = ImageStat.Stat(diff).mean[0] / 255.0
        similarity = 1.0 - mean
        return max(0.0, min(1.0, similarity))


def build_image_signature(image_bytes: bytes) -> dict[str, Any]:
    normalized = load_normalized_image_from_bytes(image_bytes)
    return {
        "hash": compute_dhash(normalized),
        "compare_png": build_compare_gray_png(normalized),
    }


def image_bytes_to_record(image_bytes: bytes) -> dict[str, Any]:
    normalized = load_normalized_image_from_bytes(image_bytes)
    return {
        "src": pil_to_data_url(normalized),
        "hash": compute_dhash(normalized),
        "compare_png": build_compare_gray_png(normalized),
    }


def extract_images_by_row(ws) -> dict[int, list[dict[str, Any]]]:
    images_by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)

    for image in getattr(ws, "_images", []):
        try:
            row_index = image.anchor._from.row + 1
            image_bytes = image._data()
            images_by_row[row_index].append(image_bytes_to_record(image_bytes))
        except Exception:
            continue

    return images_by_row


def clean_row_for_search(row_data: dict[str, Any]) -> dict[str, Any]:
    clean = {}
    for key, value in row_data.items():
        if key.startswith("__"):
            continue
        clean[key] = value
    return clean


def clear_folder_state() -> None:
    folder_root_dir = STATE.get("folder_root_dir")
    if folder_root_dir and Path(folder_root_dir).exists():
        shutil.rmtree(folder_root_dir, ignore_errors=True)

    STATE["folder_root_dir"] = None
    STATE["folder_index"] = {}

    workbook = STATE.get("workbook")
    if workbook:
        for sheet in workbook.get("sheets", []):
            for row in sheet.get("rows", []):
                row["__folderId"] = None
                row["__folderName"] = None
                row["__detailUrl"] = None
        workbook["mappedFolderCount"] = 0
        workbook["folderImportName"] = None

def clear_all_state() -> None:
    clear_folder_state()
    STATE["workbook"] = None
    STATE["search_index"] = []


def get_public_workbook_state() -> dict[str, Any] | None:
    workbook = STATE.get("workbook")
    if not workbook:
        return None
    return workbook

def parse_excel_file(file_path: Path) -> dict[str, Any]:
    clear_folder_state()

    workbook = load_workbook(file_path, data_only=True)
    all_sheets: list[dict[str, Any]] = []
    total_rows = 0
    search_index: list[dict[str, Any]] = []

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]

        header_row = detect_header_row(ws)
        headers = []
        column_indexes = []

        for col_idx in range(1, ws.max_column + 1):
            header_value = ws.cell(header_row, col_idx).value
            if header_value not in (None, ""):
                headers.append(str(header_value).strip())
                column_indexes.append(col_idx)

        if not headers:
            continue

        images_by_row = extract_images_by_row(ws)
        rows: list[dict[str, Any]] = []

        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_data: dict[str, Any] = {}
            has_value = False

            for header, col_idx in zip(headers, column_indexes):
                cell_value = normalize_cell_value(ws.cell(row_idx, col_idx).value)
                row_data[header] = cell_value
                if cell_value not in (None, ""):
                    has_value = True

            row_images = images_by_row.get(row_idx, [])
            row_data["__images"] = [item["src"] for item in row_images]
            row_data["__excelRow"] = row_idx
            row_data["__folderId"] = None
            row_data["__folderName"] = None
            row_data["__detailUrl"] = None

            if row_images:
                has_value = True

            if has_value:
                rows.append(row_data)

                for image_item in row_images:
                    search_index.append(
                        {
                            "sheetName": sheet_name,
                            "excelRow": row_idx,
                            "rowRef": row_data,
                            "matchedImage": image_item["src"],
                            "hash": image_item["hash"],
                            "compare_png": image_item["compare_png"],
                        }
                    )

        total_rows += len(rows)
        all_sheets.append(
            {
                "sheetName": sheet_name,
                "headerRow": header_row,
                "headers": headers,
                "rowCount": len(rows),
                "rows": rows,
            }
        )

    workbook_data = {
        "fileName": file_path.name,
        "sheetCount": len(all_sheets),
        "totalRows": total_rows,
        "imageIndexCount": len(search_index),
        "mappedFolderCount": 0,
        "folderImportName": None,
        "sheets": all_sheets,
    }

    STATE["workbook"] = workbook_data
    STATE["search_index"] = search_index

    return workbook_data


def compare_signature(query_signature: dict[str, Any], target_item: dict[str, Any]) -> float:
    hash_size = 16 * 16
    hash_similarity = 1.0 - (
        hamming_distance(query_signature["hash"], target_item["hash"]) / hash_size
    )
    img_similarity = pixel_similarity(query_signature["compare_png"], target_item["compare_png"])
    score = (hash_similarity * 0.6) + (img_similarity * 0.4)
    return max(0.0, min(1.0, score))


def normalize_relative_parts(relative_path: str) -> list[str]:
    parts = []
    for part in Path(relative_path).parts:
        clean = str(part).strip()
        if clean in ("", ".", ".."):
            continue
        parts.append(clean)
    return parts


def save_uploaded_folder_files(files, relative_paths: list[str], root_folder_name: str) -> Path:
    import_id = uuid.uuid4().hex
    save_root = FOLDER_IMPORT_DIR / import_id
    save_root.mkdir(parents=True, exist_ok=True)

    for uploaded_file, rel_path in zip(files, relative_paths):
        parts = normalize_relative_parts(rel_path)

        if root_folder_name and parts and parts[0] == root_folder_name:
            parts = parts[1:]

        if len(parts) < 2:
            continue

        target_path = save_root.joinpath(*parts)
        target_path.parent.mkdir(parents=True, exist_ok=True)
        uploaded_file.save(target_path)

    return save_root


def build_folder_index(saved_root: Path) -> dict[str, dict[str, Any]]:
    folder_index: dict[str, dict[str, Any]] = {}

    for child in sorted(saved_root.iterdir(), key=lambda x: natural_sort_key(x.name)):
        if not child.is_dir():
            continue

        folder_id = uuid.uuid4().hex
        file_items = []

        for file_path in sorted(child.rglob("*"), key=lambda x: natural_sort_key(x.as_posix())):
            if not file_path.is_file():
                continue

            rel_path = file_path.relative_to(child).as_posix()
            ext = file_path.suffix.lower()

            file_items.append(
                {
                    "name": file_path.name,
                    "relPath": rel_path,
                    "ext": ext,
                    "sizeKb": round(file_path.stat().st_size / 1024, 1),
                    "isImage": ext in IMAGE_EXTENSIONS,
                    "isPdf": ext in PDF_EXTENSIONS,
                }
            )

        folder_index[folder_id] = {
            "id": folder_id,
            "name": child.name,
            "path": child,
            "fileCount": len(file_items),
            "files": file_items,
        }

    return folder_index


def map_folders_to_rows_by_style_no(workbook_data: dict[str, Any], folder_index: dict[str, dict[str, Any]]) -> int:
    folder_by_name = {}
    for folder in folder_index.values():
        folder_by_name[str(folder["name"]).strip()] = folder

    mapped_count = 0

    for sheet in workbook_data.get("sheets", []):
        for row in sheet.get("rows", []):
            row["__folderId"] = None
            row["__folderName"] = None
            row["__detailUrl"] = None

            style_no = str(row.get("Style No", "") or "").strip()
            if not style_no:
                continue

            folder = folder_by_name.get(style_no)
            if not folder:
                continue

            row["__folderId"] = folder["id"]
            row["__folderName"] = folder["name"]
            row["__detailUrl"] = f"/folder/{folder['id']}"
            mapped_count += 1

    workbook_data["mappedFolderCount"] = mapped_count
    return mapped_count


def safe_resolve_file(base_dir: Path, rel_path: str) -> Path:
    target = (base_dir / rel_path).resolve()
    base_resolved = base_dir.resolve()

    if base_resolved not in target.parents and target != base_resolved:
        raise FileNotFoundError("Đường dẫn không hợp lệ")

    if not target.exists() or not target.is_file():
        raise FileNotFoundError("Không tìm thấy file")

    return target


@app.route("/")
def home():
    return render_template("index.html")


@app.get("/folder/<folder_id>")
def folder_detail(folder_id: str):
    folder = STATE["folder_index"].get(folder_id)
    if not folder:
        abort(404)

    return render_template("folder_detail.html", folder=folder)


@app.get("/folder-file/<folder_id>/<path:rel_path>")
def folder_file(folder_id: str, rel_path: str):
    folder = STATE["folder_index"].get(folder_id)
    if not folder:
        abort(404)

    try:
        file_path = safe_resolve_file(folder["path"], rel_path)
        return send_file(file_path, as_attachment=False, download_name=file_path.name)
    except Exception:
        abort(404)


@app.post("/api/excel/upload")
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
        result = parse_excel_file(file_path)
        return jsonify({"success": True, "message": "Import Excel thành công", "data": result})
    except Exception as error:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "message": f"Không đọc được file Excel: {error}"}), 500
    finally:
        if file_path.exists():
            file_path.unlink(missing_ok=True)


@app.post("/api/folder/import")
def import_local_folder():
    workbook = STATE.get("workbook")
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
        clear_folder_state()
        saved_root = save_uploaded_folder_files(files, relative_paths, root_folder_name)
        folder_index = build_folder_index(saved_root)

        STATE["folder_root_dir"] = saved_root
        STATE["folder_index"] = folder_index

        mapped_count = map_folders_to_rows_by_style_no(workbook, folder_index)
        workbook["folderImportName"] = root_folder_name or saved_root.name

        return jsonify(
            {
                "success": True,
                "message": "Import folder thành công",
                "data": {
                    "folderCount": len(folder_index),
                    "mappedCount": mapped_count,
                    "workbook": workbook,
                },
            }
        )
    except Exception as error:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "message": f"Không import được folder: {error}"}), 500


@app.post("/api/search-by-image")
def search_by_image():
    if not STATE["search_index"]:
        return jsonify(
            {
                "success": False,
                "message": "Bạn cần import Excel trước khi dán ảnh để tìm kiếm",
            }
        ), 400

    pasted_image = request.files.get("image")
    if not pasted_image:
        return jsonify({"success": False, "message": "Không nhận được ảnh từ clipboard"}), 400

    try:
        image_bytes = pasted_image.read()
        query_signature = build_image_signature(image_bytes)

        results = []
        for item in STATE["search_index"]:
            score = compare_signature(query_signature, item)
            row_ref = item["rowRef"]

            results.append(
                {
                    "sheetName": item["sheetName"],
                    "excelRow": item["excelRow"],
                    "score": round(score, 4),
                    "scorePercent": int(round(score * 100)),
                    "matchedImage": item["matchedImage"],
                    "detailUrl": row_ref.get("__detailUrl"),
                    "folderName": row_ref.get("__folderName"),
                    "row": clean_row_for_search(row_ref),
                }
            )

        results.sort(key=lambda x: x["score"], reverse=True)

        return jsonify(
            {
                "success": True,
                "message": "So khớp ảnh thành công",
                "data": {
                    "results": results[:10],
                },
            }
        )
    except Exception as error:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "message": f"Lỗi xử lý ảnh: {error}"}), 500



@app.get("/api/state")
def get_state():
    workbook = get_public_workbook_state()

    return jsonify(
        {
            "success": True,
            "data": {
                "workbook": workbook,
                "hasWorkbook": workbook is not None,
            },
        }
    )


@app.post("/api/reset")
def reset_state():
    try:
        clear_all_state()
        return jsonify(
            {
                "success": True,
                "message": "Đã reset toàn bộ dữ liệu",
            }
        )
    except Exception as error:
        import traceback
        traceback.print_exc()
        return jsonify(
            {"success": False, "message": f"Không reset được dữ liệu: {error}"}
        ), 500


# if __name__ == "__main__":
#     app.run(debug=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)

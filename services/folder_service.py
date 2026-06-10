from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any

from utils.db import BASE_PATH, FOLDER_IMPORT_DIR, IMAGE_EXTENSIONS, PDF_EXTENSIONS
from utils.helpers import (
    STATE,
    extract_folder_match_keys,
    natural_sort_key,
    normalize_match_value,
    normalize_relative_parts,
)


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

    folder_paths = [p for p in saved_root.rglob("*") if p.is_dir()]

    for folder_path in sorted(
        folder_paths,
        key=lambda x: natural_sort_key(x.relative_to(saved_root).as_posix()),
    ):
        all_files = [p for p in folder_path.rglob("*") if p.is_file()]
        if not all_files:
            continue

        file_items = []
        for file_path in sorted(
            all_files,
            key=lambda x: natural_sort_key(x.relative_to(folder_path).as_posix()),
        ):
            ext = file_path.suffix.lower()
            file_items.append(
                {
                    "name": file_path.name,
                    "relPath": file_path.relative_to(folder_path).as_posix(),
                    "ext": ext,
                    "sizeKb": round(file_path.stat().st_size / 1024, 1),
                    "isImage": ext in IMAGE_EXTENSIONS,
                    "isPdf": ext in PDF_EXTENSIONS,
                }
            )

        folder_id = uuid.uuid4().hex

        import urllib.parse

        for f in file_items:
            quoted_path = urllib.parse.quote(f["relPath"], safe="/")
            f["viewUrl"] = f"{BASE_PATH}/folder-file/{folder_id}/{quoted_path}"
            f["downloadUrl"] = f"{BASE_PATH}/folder-download-file/{folder_id}/{quoted_path}"

        folder_index[folder_id] = {
            "id": folder_id,
            "name": folder_path.relative_to(saved_root).as_posix(),
            "matchName": folder_path.name,
            "path": folder_path,
            "fileCount": len(file_items),
            "files": file_items,
        }

    return folder_index


def map_folders_to_rows_by_style_no(
    workbook_data: dict[str, Any],
    folder_index: dict[str, dict[str, Any]],
) -> int:
    folder_by_key: dict[str, dict[str, Any]] = {}

    for folder in folder_index.values():
        for key in extract_folder_match_keys(folder.get("matchName") or folder["name"]):
            folder_by_key[key] = folder

    mapped_count = 0

    for sheet in workbook_data.get("sheets", []):
        for row in sheet.get("rows", []):
            style_no = normalize_match_value(row.get("Style No"))
            if not style_no:
                continue

            folder = folder_by_key.get(style_no)
            if not folder:
                continue

            row["__folderId"] = folder["id"]
            row["__folderName"] = folder["name"]
            row["__detailUrl"] = f"{BASE_PATH}/folder/{folder['id']}"
            mapped_count += 1

    workbook_data["mappedFolderCount"] = mapped_count
    return mapped_count


def restore_folder_state_from_disk() -> dict[str, Any]:
    STATE["folder_index"] = {}
    STATE["folder_root_dir"] = []

    for workbook in STATE.get("imports", []):
        for sheet in workbook.get("sheets", []):
            for row in sheet.get("rows", []):
                row["__folderId"] = None
                row["__folderName"] = None
                row["__detailUrl"] = None

        workbook["mappedFolderCount"] = 0
        workbook["folderImportName"] = None

    if not FOLDER_IMPORT_DIR.exists():
        return {
            "rootCount": 0,
            "folderCount": 0,
            "mappedCount": 0,
        }

    saved_roots = [p for p in FOLDER_IMPORT_DIR.iterdir() if p.is_dir()]
    saved_roots.sort(key=lambda p: (p.stat().st_mtime, natural_sort_key(p.name)))

    merged_folder_index: dict[str, dict[str, Any]] = {}
    restored_roots: list[Path] = []

    for saved_root in saved_roots:
        try:
            folder_index = build_folder_index(saved_root)
        except Exception:
            continue

        if not folder_index:
            continue

        restored_roots.append(saved_root)
        merged_folder_index.update(folder_index)

    STATE["folder_root_dir"] = restored_roots
    STATE["folder_index"] = merged_folder_index

    total_mapped = 0
    for workbook in STATE.get("imports", []):
        mapped_count = map_folders_to_rows_by_style_no(workbook, STATE["folder_index"])
        total_mapped += mapped_count

        if restored_roots:
            workbook["folderImportName"] = restored_roots[-1].name

    return {
        "rootCount": len(restored_roots),
        "folderCount": len(STATE["folder_index"]),
        "mappedCount": total_mapped,
    }
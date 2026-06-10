from __future__ import annotations

import base64
from collections import defaultdict
from io import BytesIO
from typing import Any

from PIL import Image, ImageChops, ImageOps, ImageStat

from utils.db import BASE_PATH, EXCEL_IMAGE_DIR, UPLOAD_DIR
from utils.helpers import slugify_filename


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


def save_excel_image_file(
    image_bytes: bytes,
    import_id: str,
    sheet_name: str,
    row_idx: int,
    image_index: int,
) -> dict[str, Any]:
    normalized = load_normalized_image_from_bytes(image_bytes)
    sheet_safe = slugify_filename(sheet_name)

    target_dir = EXCEL_IMAGE_DIR / import_id / sheet_safe / str(row_idx)
    target_dir.mkdir(parents=True, exist_ok=True)

    file_name = f"image_{image_index + 1}.png"
    file_path = target_dir / file_name
    normalized.save(file_path, format="PNG", optimize=True)

    relative_path = file_path.relative_to(UPLOAD_DIR).as_posix()
    return {
        "src": f"{BASE_PATH}/media/{relative_path}",
        "relative_path": relative_path,
        "hash": compute_dhash(normalized),
        "compare_png": build_compare_gray_png(normalized),
    }


def extract_images_by_row(ws, import_id: str, sheet_name: str) -> dict[int, list[dict[str, Any]]]:
    images_by_row: dict[int, list[dict[str, Any]]] = defaultdict(list)

    for image_index, image_obj in enumerate(getattr(ws, "_images", [])):
        try:
            row_index = image_obj.anchor._from.row + 1
            image_bytes = image_obj._data()
            record = save_excel_image_file(
                image_bytes=image_bytes,
                import_id=import_id,
                sheet_name=sheet_name,
                row_idx=row_index,
                image_index=len(images_by_row[row_index]),
            )
            images_by_row[row_index].append(record)
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


def compare_signature(query_signature: dict[str, Any], target_item: dict[str, Any]) -> float:
    hash_size = 16 * 16
    hash_similarity = 1.0 - (
        hamming_distance(query_signature["hash"], target_item["hash"]) / hash_size
    )
    img_similarity = pixel_similarity(query_signature["compare_png"], target_item["compare_png"])
    score = (hash_similarity * 0.6) + (img_similarity * 0.4)
    return max(0.0, min(1.0, score))
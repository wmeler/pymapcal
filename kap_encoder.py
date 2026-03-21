from __future__ import annotations

import math
import struct
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from PySide6.QtGui import QImage


class KapEncodeCancelled(Exception):
    pass


@dataclass(frozen=True)
class IndexedColor:
    r: int
    g: int
    b: int
    weight: int
    old_index: int


@dataclass(frozen=True)
class IndexedImage:
    width: int
    height: int
    bits_per_pixel: int
    palette: list[tuple[int, int, int]]
    rows: list[bytes]


def write_kap_file(
    image_path: Path,
    output_path: Path,
    header_text: str,
    log_cb: Optional[Callable[[str], None]] = None,
    cancel_requested_cb: Optional[Callable[[], bool]] = None,
) -> None:
    indexed = load_indexed_image(
        image_path=image_path,
        max_colors=128,
        log_cb=log_cb,
        cancel_requested_cb=cancel_requested_cb,
    )
    kap_bytes = build_kap_bytes(
        header_text=header_text,
        indexed=indexed,
        log_cb=log_cb,
        cancel_requested_cb=cancel_requested_cb,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    try:
        temp_path.write_bytes(kap_bytes)
        temp_path.replace(output_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def load_indexed_image(
    image_path: Path,
    max_colors: int = 128,
    log_cb: Optional[Callable[[str], None]] = None,
    cancel_requested_cb: Optional[Callable[[], bool]] = None,
) -> IndexedImage:
    _check_cancel(cancel_requested_cb)
    image = QImage(str(image_path))
    if image.isNull():
        raise RuntimeError(f"cannot_load_image:{image_path}")
    indexed = image.convertToFormat(QImage.Format.Format_Indexed8)
    if indexed.isNull():
        raise RuntimeError(f"cannot_quantize_image:{image_path}")

    width = indexed.width()
    height = indexed.height()
    color_table = indexed.colorTable()
    original_rows: list[bytes] = []
    counts: list[int] = [0] * max(1, len(color_table))

    for y in range(height):
        if y % 64 == 0:
            _check_cancel(cancel_requested_cb)
        row = _scanline_bytes(indexed, y, width)
        original_rows.append(row)
        for index in row:
            counts[index] += 1

    used_entries = [
        IndexedColor(
            r=(color_table[index] >> 16) & 0xFF,
            g=(color_table[index] >> 8) & 0xFF,
            b=color_table[index] & 0xFF,
            weight=counts[index],
            old_index=index,
        )
        for index in range(len(color_table))
        if counts[index] > 0
    ]
    if not used_entries:
        raise RuntimeError(f"image_has_no_pixels:{image_path}")

    if len(used_entries) <= max_colors:
        old_to_new = {entry.old_index: new_index for new_index, entry in enumerate(used_entries)}
        palette = [(entry.r, entry.g, entry.b) for entry in used_entries]
    else:
        old_to_new, palette = _reduce_palette(used_entries, max_colors)

    rows = [bytes(old_to_new[index] for index in row) for row in original_rows]
    bits_per_pixel = max(1, math.ceil(math.log2(max(2, len(palette)))))
    if bits_per_pixel > 7:
        raise RuntimeError(f"too_many_output_colors:{len(palette)}")

    _log(
        log_cb,
        (
            f"[kap-encoder] quantized image={image_path} size={width}x{height} "
            f"input_colors={len(used_entries)} output_colors={len(palette)} bits={bits_per_pixel}\n"
        ),
    )
    return IndexedImage(
        width=width,
        height=height,
        bits_per_pixel=bits_per_pixel,
        palette=palette,
        rows=rows,
    )


def build_kap_bytes(
    header_text: str,
    indexed: IndexedImage,
    log_cb: Optional[Callable[[str], None]] = None,
    cancel_requested_cb: Optional[Callable[[], bool]] = None,
) -> bytes:
    _check_cancel(cancel_requested_cb)
    header = _normalize_header(header_text)
    out = bytearray(header)
    out.extend(b"OST/1\r\n")
    out.extend(f"IFM/{indexed.bits_per_pixel}\r\n".encode("ascii"))
    for index, (r, g, b) in enumerate(indexed.palette):
        out.extend(f"RGB/{index},{r},{g},{b}\r\n".encode("ascii"))
    out.extend(b"\x1A\x00")
    out.append(indexed.bits_per_pixel)

    offsets: list[int] = []
    for line_no, row in enumerate(indexed.rows):
        if line_no % 64 == 0:
            _check_cancel(cancel_requested_cb)
        offsets.append(len(out))
        out.extend(_compress_row(row, indexed.bits_per_pixel, line_no))
    offsets.append(len(out))
    for offset in offsets:
        out.extend(struct.pack(">I", offset))

    _log(
        log_cb,
        (
            f"[kap-encoder] encoded raster rows={indexed.height} palette={len(indexed.palette)} "
            f"index_entries={len(offsets)} bytes={len(out)}\n"
        ),
    )
    return bytes(out)


def _normalize_header(header_text: str) -> bytes:
    stripped = header_text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    return stripped.replace("\n", "\r\n").encode("utf-8") + b"\r\n"


def _scanline_bytes(image: QImage, y: int, width: int) -> bytes:
    ptr = image.constScanLine(y)
    return bytes(memoryview(ptr)[:width])


def _reduce_palette(
    entries: list[IndexedColor],
    target_colors: int,
) -> tuple[dict[int, int], list[tuple[int, int, int]]]:
    boxes: list[list[IndexedColor]] = [entries[:]]
    while len(boxes) < target_colors:
        split_index = _best_split_box_index(boxes)
        if split_index is None:
            break
        left, right = _split_box(boxes.pop(split_index))
        boxes.append(left)
        boxes.append(right)

    old_to_new: dict[int, int] = {}
    palette: list[tuple[int, int, int]] = []
    for new_index, box in enumerate(boxes):
        total = sum(entry.weight for entry in box)
        r = round(sum(entry.r * entry.weight for entry in box) / total)
        g = round(sum(entry.g * entry.weight for entry in box) / total)
        b = round(sum(entry.b * entry.weight for entry in box) / total)
        palette.append((r, g, b))
        for entry in box:
            old_to_new[entry.old_index] = new_index
    return old_to_new, palette


def _best_split_box_index(boxes: list[list[IndexedColor]]) -> Optional[int]:
    best_index: Optional[int] = None
    best_score = -1
    for index, box in enumerate(boxes):
        if len(box) < 2:
            continue
        r_range = max(entry.r for entry in box) - min(entry.r for entry in box)
        g_range = max(entry.g for entry in box) - min(entry.g for entry in box)
        b_range = max(entry.b for entry in box) - min(entry.b for entry in box)
        span = max(r_range, g_range, b_range)
        if span <= 0:
            continue
        total = sum(entry.weight for entry in box)
        score = span * total
        if score > best_score:
            best_score = score
            best_index = index
    return best_index


def _split_box(box: list[IndexedColor]) -> tuple[list[IndexedColor], list[IndexedColor]]:
    ranges = [
        max(entry.r for entry in box) - min(entry.r for entry in box),
        max(entry.g for entry in box) - min(entry.g for entry in box),
        max(entry.b for entry in box) - min(entry.b for entry in box),
    ]
    axis = max(range(3), key=ranges.__getitem__)
    ordered = sorted(box, key=lambda entry: (entry.r, entry.g, entry.b)[axis])
    total = sum(entry.weight for entry in ordered)
    cumulative = 0
    split_at = 1
    for index, entry in enumerate(ordered[:-1], start=1):
        cumulative += entry.weight
        if cumulative * 2 >= total:
            split_at = index
            break
    return ordered[:split_at], ordered[split_at:]


def _compress_row(row: bytes, bits_per_pixel: int, line_no: int) -> bytes:
    lower_bits = 7 - bits_per_pixel
    max_run_value = (1 << lower_bits) - 1
    out = bytearray()
    out.extend(_encode_bsb_number(line_no, 0, 0x7F))

    x = 0
    width = len(row)
    while x < width:
        pixel = row[x]
        x += 1
        run_length = 1
        while x < width and row[x] == pixel:
            x += 1
            run_length += 1
        out.extend(_encode_bsb_number(run_length - 1, pixel << lower_bits, max_run_value))
    out.append(0)
    return bytes(out)


def _encode_bsb_number(value: int, pixel_prefix: int, max_value: int) -> bytes:
    if value > max_value:
        return _encode_bsb_number(value >> 7, pixel_prefix | 0x80, max_value) + bytes(
            [(value & 0x7F) | (pixel_prefix & 0x80)]
        )
    final_byte = pixel_prefix | value
    if final_byte == 0:
        return b"\x80\x00"
    return bytes([final_byte])


def _check_cancel(cancel_requested_cb: Optional[Callable[[], bool]]) -> None:
    if cancel_requested_cb is not None and cancel_requested_cb():
        raise KapEncodeCancelled()


def _log(log_cb: Optional[Callable[[str], None]], message: str) -> None:
    if log_cb is not None:
        log_cb(message)

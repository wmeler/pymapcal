from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional, Protocol, Sequence

from PySide6.QtGui import QImage

from geo_positioning import build_geo_transform
from kap_export import (
    KapExportJob,
    KapExportResult,
    KapPolygonPoint,
    KapReference,
    run_kap_export_jobs,
    sanitize_kap_stem,
)


class ExportPointLike(Protocol):
    x: float
    y: float
    lon: Optional[float]
    lat: Optional[float]
    is_corner: bool


class ExportSheetLike(Protocol):
    name: str
    scale: str
    points: Sequence[ExportPointLike]


class ExportScanLike(Protocol):
    name: str
    image_path: str
    sheets: Sequence[ExportSheetLike]


@dataclass
class KapExportIssue:
    code: str
    scan_name: str
    sheet_name: str = ""
    path: Optional[Path] = None
    scale_text: str = ""


@dataclass
class KapExportPlan:
    jobs: list[KapExportJob]
    issues: list[KapExportIssue]


def parse_scale_value(value: str) -> Optional[int]:
    cleaned = re.sub(r"[^0-9]", "", value or "")
    if not cleaned:
        return None
    out = int(cleaned)
    if out <= 0:
        return None
    return out


def unique_kap_path(out_dir: Path, used_names: set[str], stem: str) -> Path:
    base = sanitize_kap_stem(stem, fallback="sheet")
    candidate = base
    idx = 2
    while candidate.lower() in used_names or (out_dir / f"{candidate}.kap").exists():
        candidate = f"{base}_{idx}"
        idx += 1
    used_names.add(candidate.lower())
    return out_dir / f"{candidate}.kap"


def point_geo_or_transform(
    point: ExportPointLike,
    transform,
) -> Optional[tuple[float, float]]:
    if point.lon is not None and point.lat is not None:
        return float(point.lon), float(point.lat)
    if transform is None:
        return None
    return transform.pixel_to_geo(point.x, point.y)


def collect_kap_export_plan(
    scans: Sequence[ExportScanLike],
    out_dir: Path,
    sheet_default_name: str,
) -> KapExportPlan:
    jobs: list[KapExportJob] = []
    issues: list[KapExportIssue] = []
    used_stems: set[str] = set()

    for scan in scans:
        scan_name = scan.name.strip() or (Path(scan.image_path).stem if scan.image_path else "scan")
        if not scan.sheets:
            continue

        image_path = Path(scan.image_path).expanduser()
        if not image_path.exists():
            issues.append(KapExportIssue("export_image_missing", scan_name=scan_name, path=image_path))
            continue

        image = QImage(str(image_path))
        if image.isNull():
            issues.append(KapExportIssue("export_image_open_error", scan_name=scan_name, path=image_path))
            continue
        width = image.width()
        height = image.height()

        for idx, sheet in enumerate(scan.sheets, start=1):
            sheet_name = sheet.name.strip() or f"{sheet_default_name} {idx}"
            scale_value = parse_scale_value(sheet.scale)
            if scale_value is None:
                issues.append(
                    KapExportIssue(
                        "export_sheet_scale_invalid",
                        scan_name=scan_name,
                        sheet_name=sheet_name,
                        scale_text=sheet.scale if sheet.scale else "-",
                    )
                )
                continue

            corners = [point for point in sheet.points if point.is_corner]
            if len(corners) < 3:
                issues.append(
                    KapExportIssue(
                        "export_sheet_corners_missing",
                        scan_name=scan_name,
                        sheet_name=sheet_name,
                    )
                )
                continue

            transform = build_geo_transform(sheet.points)
            polygon: list[KapPolygonPoint] = []
            missing_corner_geo = False
            for point in corners:
                geo = point_geo_or_transform(point, transform)
                if geo is None:
                    missing_corner_geo = True
                    break
                polygon.append(
                    KapPolygonPoint(
                        pixel_x=point.x,
                        pixel_y=point.y,
                        lon=geo[0],
                        lat=geo[1],
                    )
                )
            if missing_corner_geo:
                issues.append(
                    KapExportIssue(
                        "export_sheet_geo_missing",
                        scan_name=scan_name,
                        sheet_name=sheet_name,
                    )
                )
                continue

            references: list[KapReference] = []
            for point in sheet.points:
                if point.is_corner:
                    continue
                geo = point_geo_or_transform(point, transform)
                if geo is None:
                    continue
                references.append(
                    KapReference(
                        pixel_x=point.x,
                        pixel_y=point.y,
                        lon=geo[0],
                        lat=geo[1],
                    )
                )

            out_path = unique_kap_path(out_dir, used_stems, f"{scan_name}_{sheet_name}")
            jobs.append(
                KapExportJob(
                    sheet_name=f"{scan_name}/{sheet_name}",
                    image_path=image_path,
                    output_path=out_path,
                    width=width,
                    height=height,
                    scale=scale_value,
                    polygon=polygon,
                    references=references,
                )
            )

    return KapExportPlan(jobs=jobs, issues=issues)


def execute_kap_export_plan(
    jobs: list[KapExportJob],
    imgkap_path: str,
    sounding_datum: str,
    temp_dir: Optional[Path] = None,
    log_path: Optional[Path] = None,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
    cancel_requested_cb: Optional[Callable[[], bool]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
) -> list[KapExportResult]:
    return run_kap_export_jobs(
        jobs=jobs,
        imgkap_path=imgkap_path,
        sounding_datum=sounding_datum,
        temp_dir=temp_dir,
        log_path=log_path,
        progress_cb=progress_cb,
        cancel_requested_cb=cancel_requested_cb,
        log_cb=log_cb,
    )

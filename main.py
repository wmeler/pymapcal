from __future__ import annotations

import json
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QEvent, QPointF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QAction, QColor, QMouseEvent, QPainter, QPen, QPixmap, QWheelEvent
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

TRANSLATIONS = {
    "pl": {
        "settings_json_object": "Plik .pymapcal musi być obiektem JSON.",
        "settings_load_error": "Nie udało się wczytać .pymapcal: {error}",
        "title": "MapCal Qt",
        "ready": "Gotowe",
        "settings_title": "Ustawienia",
        "sheet_default": "Arkusz",
        "limit_title": "Limit",
        "limit_cal_points": "Maksymalnie 9 punktów kalibracyjnych na arkusz.",
        "btn_new_sheet": "Nowy arkusz",
        "btn_close_sheet": "Zamknij arkusz",
        "btn_cancel_drawing": "Anuluj rysowanie",
        "btn_add_cal_point": "Dodaj punkt kalibracyjny",
        "btn_select_mode": "Tryb wyboru",
        "btn_delete_sheet": "Usuń zaznaczony arkusz",
        "btn_use_as_corner": "Dodaj punkt kalibracji do obrysu",
        "btn_zoom_in": "Zoom +",
        "btn_zoom_out": "Zoom -",
        "btn_zoom_reset": "Zoom 100%",
        "panel_sheets": "Arkusze",
        "panel_sheet_meta": "Metadane arkusza",
        "field_name": "Nazwa",
        "field_scale": "Skala",
        "panel_point_geo": "Punkt (lon/lat)",
        "menu_file": "Plik",
        "menu_open_scan": "Otwórz skan mapy",
        "menu_save": "Zapisz",
        "menu_save_as": "Zapisz jako...",
        "menu_load_project": "Wczytaj projekt",
        "menu_help": "Pomoc",
        "menu_readme": "README",
        "help_title": "Pomoc",
        "help_missing_readme": "Nie znaleziono pliku README.md.",
        "help_read_error": "Nie udało się odczytać README.md:\n{error}",
        "help_dialog_title": "Pomoc - README",
        "status_cursor_no_geo": "Px: ({x:.1f}, {y:.1f}) | Geo: - | Zoom: {zoom}%",
        "status_cursor_geo": "Px: ({x:.1f}, {y:.1f}) | Geo: lon={lon:.6f}, lat={lat:.6f} | Zoom: {zoom}%",
        "error_title": "Błąd",
        "error_sheet_min_points": "Arkusz musi mieć min. 3 punkty.",
        "link_point_title": "Powiązanie punktu",
        "link_point_question": "Dociągnąć obrys kadrowania do zaznaczonego punktu kalibracji?",
        "tree_crop_outline": "Obrys kadrowania",
        "tree_cal_points": "Punkty kalibracyjne",
        "tree_sheet_item": "{name} ({count} pkt)",
        "tree_corner_item": "Narożnik {n}: x={x:.1f}, y={y:.1f}",
        "tree_point_item": "Punkt {n}: x={x:.1f}, y={y:.1f}",
        "error_lon_format_title": "Błąd formatu lon",
        "error_lon_format_msg": "Nie udało się odczytać długości geograficznej. Obsługiwane: DD, DMM, DMS + półkula (E/W).",
        "error_lat_format_title": "Błąd formatu lat",
        "error_lat_format_msg": "Nie udało się odczytać szerokości geograficznej. Obsługiwane: DD, DMM, DMS + półkula (N/S).",
        "dialog_pick_scan": "Wybierz skan mapy",
        "dialog_image_filter": "Image (*.tif *.tiff *.bmp *.png *.jpg *.jpeg)",
        "error_open_image": "Nie udało się otworzyć obrazu.",
        "dialog_save_as": "Zapisz projekt jako",
        "dialog_project_filter": "MapCal Project (*.json)",
        "default_project_name": "projekt.json",
        "dialog_load_project": "Wczytaj projekt",
        "error_load_project": "Nie udało się wczytać projektu:\n{error}",
        "error_open_project_image": "Nie udało się otworzyć obrazu zapisanego w projekcie.",
        "warning_title": "Uwaga",
        "warning_missing_project_image": "Brak obrazu z projektu. Wczytano tylko dane arkuszy.",
    },
    "en": {
        "settings_json_object": ".pymapcal must be a JSON object.",
        "settings_load_error": "Failed to load .pymapcal: {error}",
        "title": "MapCal Qt",
        "ready": "Ready",
        "settings_title": "Settings",
        "sheet_default": "Sheet",
        "limit_title": "Limit",
        "limit_cal_points": "Maximum 9 calibration points per sheet.",
        "btn_new_sheet": "New sheet",
        "btn_close_sheet": "Close sheet",
        "btn_cancel_drawing": "Cancel drawing",
        "btn_add_cal_point": "Add calibration point",
        "btn_select_mode": "Select mode",
        "btn_delete_sheet": "Delete selected sheet",
        "btn_use_as_corner": "Use calibration point in crop outline",
        "btn_zoom_in": "Zoom +",
        "btn_zoom_out": "Zoom -",
        "btn_zoom_reset": "Zoom 100%",
        "panel_sheets": "Sheets",
        "panel_sheet_meta": "Sheet metadata",
        "field_name": "Name",
        "field_scale": "Scale",
        "panel_point_geo": "Point (lon/lat)",
        "menu_file": "File",
        "menu_open_scan": "Open map scan",
        "menu_save": "Save",
        "menu_save_as": "Save as...",
        "menu_load_project": "Load project",
        "menu_help": "Help",
        "menu_readme": "README",
        "help_title": "Help",
        "help_missing_readme": "README.md file not found.",
        "help_read_error": "Failed to read README.md:\n{error}",
        "help_dialog_title": "Help - README",
        "status_cursor_no_geo": "Px: ({x:.1f}, {y:.1f}) | Geo: - | Zoom: {zoom}%",
        "status_cursor_geo": "Px: ({x:.1f}, {y:.1f}) | Geo: lon={lon:.6f}, lat={lat:.6f} | Zoom: {zoom}%",
        "error_title": "Error",
        "error_sheet_min_points": "A sheet must have at least 3 points.",
        "link_point_title": "Point link",
        "link_point_question": "Snap crop outline to selected calibration point?",
        "tree_crop_outline": "Crop outline",
        "tree_cal_points": "Calibration points",
        "tree_sheet_item": "{name} ({count} pt)",
        "tree_corner_item": "Corner {n}: x={x:.1f}, y={y:.1f}",
        "tree_point_item": "Point {n}: x={x:.1f}, y={y:.1f}",
        "error_lon_format_title": "Lon format error",
        "error_lon_format_msg": "Failed to parse longitude. Supported: DD, DMM, DMS + hemisphere (E/W).",
        "error_lat_format_title": "Lat format error",
        "error_lat_format_msg": "Failed to parse latitude. Supported: DD, DMM, DMS + hemisphere (N/S).",
        "dialog_pick_scan": "Select map scan",
        "dialog_image_filter": "Image (*.tif *.tiff *.bmp *.png *.jpg *.jpeg)",
        "error_open_image": "Failed to open image.",
        "dialog_save_as": "Save project as",
        "dialog_project_filter": "MapCal Project (*.json)",
        "default_project_name": "project.json",
        "dialog_load_project": "Load project",
        "error_load_project": "Failed to load project:\n{error}",
        "error_open_project_image": "Failed to open image stored in project.",
        "warning_title": "Warning",
        "warning_missing_project_image": "Project image is missing. Loaded only sheet data.",
    },
}


def t(lang: str, key: str, **kwargs) -> str:
    lang_dict = TRANSLATIONS.get(lang) or TRANSLATIONS["pl"]
    text = lang_dict.get(key) or TRANSLATIONS["pl"].get(key, key)
    return text.format(**kwargs) if kwargs else text


@dataclass
class CalibrationPoint:
    x: float
    y: float
    lon: Optional[float] = None
    lat: Optional[float] = None
    lon_text: str = ""
    lat_text: str = ""
    is_corner: bool = False

    def to_json(self) -> dict:
        return {
            "x": self.x,
            "y": self.y,
            "lon": self.lon,
            "lat": self.lat,
            "lon_text": self.lon_text,
            "lat_text": self.lat_text,
            "is_corner": self.is_corner,
        }

    @staticmethod
    def from_json(data: dict) -> "CalibrationPoint":
        return CalibrationPoint(
            x=float(data["x"]),
            y=float(data["y"]),
            lon=data.get("lon"),
            lat=data.get("lat"),
            lon_text=str(data.get("lon_text", "")),
            lat_text=str(data.get("lat_text", "")),
            is_corner=bool(data.get("is_corner", False)),
        )


@dataclass
class Sheet:
    name: str
    scale: str = ""
    points: list[CalibrationPoint] = field(default_factory=list)

    def corners(self) -> list[CalibrationPoint]:
        return [p for p in self.points if p.is_corner]

    def to_json(self) -> dict:
        return {
            "name": self.name,
            "scale": self.scale,
            "points": [p.to_json() for p in self.points],
        }

    @staticmethod
    def from_json(data: dict) -> "Sheet":
        return Sheet(
            name=data.get("name", "Sheet"),
            scale=data.get("scale", ""),
            points=[CalibrationPoint.from_json(p) for p in data.get("points", [])],
        )


@dataclass
class DisplaySettings:
    outline_width: int = 1
    outline_selected_width: int = 2
    draft_outline_width: int = 2
    crosshair_arm_corner: int = 14
    crosshair_arm_cal: int = 12
    crosshair_ring_corner: int = 4
    crosshair_ring_cal: int = 2
    crosshair_selected_arm_bonus: int = 4
    crosshair_selected_ring_bonus: int = 2
    cursor_guide_width: int = 2
    cursor_guide_alpha: int = 200
    cursor_guide_dash: int = 10
    cursor_guide_gap: int = 6
    cursor_guide_color: str = "#FFD84D"

    @staticmethod
    def _int_value(data: dict, key: str, default: int, min_value: int, max_value: int) -> int:
        value = data.get(key, default)
        try:
            iv = int(value)
        except (TypeError, ValueError):
            return default
        return max(min_value, min(max_value, iv))

    @staticmethod
    def _str_value(data: dict, key: str, default: str) -> str:
        value = data.get(key, default)
        if isinstance(value, str) and value.strip():
            return value.strip()
        return default

    @classmethod
    def from_dict(cls, data: dict) -> "DisplaySettings":
        return cls(
            outline_width=cls._int_value(data, "outline_width", 1, 1, 12),
            outline_selected_width=cls._int_value(data, "outline_selected_width", 2, 1, 16),
            draft_outline_width=cls._int_value(data, "draft_outline_width", 2, 1, 16),
            crosshair_arm_corner=cls._int_value(data, "crosshair_arm_corner", 14, 4, 80),
            crosshair_arm_cal=cls._int_value(data, "crosshair_arm_cal", 12, 4, 80),
            crosshair_ring_corner=cls._int_value(data, "crosshair_ring_corner", 4, 1, 40),
            crosshair_ring_cal=cls._int_value(data, "crosshair_ring_cal", 2, 1, 40),
            crosshair_selected_arm_bonus=cls._int_value(data, "crosshair_selected_arm_bonus", 4, 0, 40),
            crosshair_selected_ring_bonus=cls._int_value(data, "crosshair_selected_ring_bonus", 2, 0, 20),
            cursor_guide_width=cls._int_value(data, "cursor_guide_width", 2, 1, 16),
            cursor_guide_alpha=cls._int_value(data, "cursor_guide_alpha", 200, 0, 255),
            cursor_guide_dash=cls._int_value(data, "cursor_guide_dash", 10, 1, 80),
            cursor_guide_gap=cls._int_value(data, "cursor_guide_gap", 6, 1, 80),
            cursor_guide_color=cls._str_value(data, "cursor_guide_color", "#FFD84D"),
        )


def load_display_settings() -> tuple[DisplaySettings, str, Optional[Path], Optional[str]]:
    candidates = [Path.cwd() / ".pymapcal", Path.home() / ".pymapcal"]
    for path in candidates:
        if not path.exists():
            continue
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(raw, dict):
                return DisplaySettings(), "pl", path, t("pl", "settings_json_object")
            lang = raw.get("language", "pl")
            if lang not in ("pl", "en"):
                lang = "pl"
            if isinstance(raw.get("display"), dict):
                source = raw["display"]
            else:
                source = raw
            return DisplaySettings.from_dict(source), lang, path, None
        except Exception as exc:
            return DisplaySettings(), "pl", path, t("pl", "settings_load_error", error=exc)
    return DisplaySettings(), "pl", None, None


def point_in_polygon(x: float, y: float, polygon: list[CalibrationPoint]) -> bool:
    if len(polygon) < 3:
        return False
    inside = False
    j = len(polygon) - 1
    for i in range(len(polygon)):
        xi, yi = polygon[i].x, polygon[i].y
        xj, yj = polygon[j].x, polygon[j].y
        intersects = ((yi > y) != (yj > y)) and (
            x < (xj - xi) * (y - yi) / ((yj - yi) + 1e-12) + xi
        )
        if intersects:
            inside = not inside
        j = i
    return inside


def solve_3x3(a: list[list[float]], b: list[float]) -> Optional[list[float]]:
    m = [row[:] + [rhs] for row, rhs in zip(a, b)]
    n = 3
    for col in range(n):
        pivot = max(range(col, n), key=lambda r: abs(m[r][col]))
        if abs(m[pivot][col]) < 1e-12:
            return None
        if pivot != col:
            m[col], m[pivot] = m[pivot], m[col]
        f = m[col][col]
        for k in range(col, n + 1):
            m[col][k] /= f
        for r in range(n):
            if r == col:
                continue
            factor = m[r][col]
            for k in range(col, n + 1):
                m[r][k] -= factor * m[col][k]
    return [m[i][n] for i in range(n)]


def affine_from_points(points: list[CalibrationPoint]) -> Optional[tuple[list[float], list[float]]]:
    known = [p for p in points if p.lon is not None and p.lat is not None]
    if len(known) < 3:
        return None
    samples = [(p.x, p.y, float(p.lon), float(p.lat)) for p in known]
    return affine_fit(samples)


def affine_fit(samples: list[tuple[float, float, float, float]]) -> Optional[tuple[list[float], list[float]]]:
    if len(samples) < 3:
        return None

    s_xx = s_xy = s_x = s_yy = s_y = n = 0.0
    b_dst_x = [0.0, 0.0, 0.0]
    b_dst_y = [0.0, 0.0, 0.0]

    for x, y, dst_x, dst_y in samples:
        n += 1.0
        s_xx += x * x
        s_xy += x * y
        s_x += x
        s_yy += y * y
        s_y += y
        b_dst_x[0] += x * dst_x
        b_dst_x[1] += y * dst_x
        b_dst_x[2] += dst_x
        b_dst_y[0] += x * dst_y
        b_dst_y[1] += y * dst_y
        b_dst_y[2] += dst_y

    ata = [
        [s_xx, s_xy, s_x],
        [s_xy, s_yy, s_y],
        [s_x, s_y, n],
    ]
    dst_x_coef = solve_3x3(ata, b_dst_x)
    dst_y_coef = solve_3x3(ata, b_dst_y)
    if not dst_x_coef or not dst_y_coef:
        return None
    return dst_x_coef, dst_y_coef


def apply_affine(coef_x: list[float], coef_y: list[float], x: float, y: float) -> tuple[float, float]:
    out_x = coef_x[0] * x + coef_x[1] * y + coef_x[2]
    out_y = coef_y[0] * x + coef_y[1] * y + coef_y[2]
    return out_x, out_y


def apply_similarity(
    src_a: tuple[float, float],
    src_b: tuple[float, float],
    dst_a: tuple[float, float],
    dst_b: tuple[float, float],
    x: float,
    y: float,
) -> Optional[tuple[float, float]]:
    src_dx = src_b[0] - src_a[0]
    src_dy = src_b[1] - src_a[1]
    dst_dx = dst_b[0] - dst_a[0]
    dst_dy = dst_b[1] - dst_a[1]
    src_len = math.hypot(src_dx, src_dy)
    dst_len = math.hypot(dst_dx, dst_dy)
    if src_len < 1e-9 or dst_len < 1e-9:
        return None

    src_ex = src_dx / src_len
    src_ey = src_dy / src_len
    dst_ex = dst_dx / dst_len
    dst_ey = dst_dy / dst_len
    src_perp_x, src_perp_y = -src_ey, src_ex
    dst_perp_x, dst_perp_y = -dst_ey, dst_ex

    rel_x = x - src_a[0]
    rel_y = y - src_a[1]
    u = rel_x * src_ex + rel_y * src_ey
    v = rel_x * src_perp_x + rel_y * src_perp_y
    scale = dst_len / src_len

    out_x = dst_a[0] + scale * (u * dst_ex + v * dst_perp_x)
    out_y = dst_a[1] + scale * (u * dst_ey + v * dst_perp_y)
    return out_x, out_y


class GeoTransform:
    def __init__(
        self,
        mode: str,
        affine_forward: Optional[tuple[list[float], list[float]]] = None,
        affine_reverse: Optional[tuple[list[float], list[float]]] = None,
        sim_pixels: Optional[tuple[tuple[float, float], tuple[float, float]]] = None,
        sim_geo: Optional[tuple[tuple[float, float], tuple[float, float]]] = None,
    ) -> None:
        self.mode = mode
        self.affine_forward = affine_forward
        self.affine_reverse = affine_reverse
        self.sim_pixels = sim_pixels
        self.sim_geo = sim_geo

    def pixel_to_geo(self, x: float, y: float) -> Optional[tuple[float, float]]:
        if self.mode == "affine" and self.affine_forward is not None:
            lon_coef, lat_coef = self.affine_forward
            return apply_affine(lon_coef, lat_coef, x, y)
        if self.mode == "similarity" and self.sim_pixels and self.sim_geo:
            return apply_similarity(self.sim_pixels[0], self.sim_pixels[1], self.sim_geo[0], self.sim_geo[1], x, y)
        return None

    def geo_to_pixel(self, lon: float, lat: float) -> Optional[tuple[float, float]]:
        if self.mode == "affine" and self.affine_reverse is not None:
            x_coef, y_coef = self.affine_reverse
            return apply_affine(x_coef, y_coef, lon, lat)
        if self.mode == "similarity" and self.sim_pixels and self.sim_geo:
            return apply_similarity(self.sim_geo[0], self.sim_geo[1], self.sim_pixels[0], self.sim_pixels[1], lon, lat)
        return None


def build_geo_transform(points: list[CalibrationPoint]) -> Optional[GeoTransform]:
    known = [p for p in points if p.lon is not None and p.lat is not None]
    if len(known) >= 3:
        forward = affine_fit([(p.x, p.y, float(p.lon), float(p.lat)) for p in known])
        reverse = affine_fit([(float(p.lon), float(p.lat), p.x, p.y) for p in known])
        if forward and reverse:
            return GeoTransform("affine", affine_forward=forward, affine_reverse=reverse)

    if len(known) >= 2:
        p1, p2 = known[0], known[1]
        sim_pixels = ((p1.x, p1.y), (p2.x, p2.y))
        sim_geo = ((float(p1.lon), float(p1.lat)), (float(p2.lon), float(p2.lat)))
        if (
            math.hypot(sim_pixels[1][0] - sim_pixels[0][0], sim_pixels[1][1] - sim_pixels[0][1]) > 1e-9
            and math.hypot(sim_geo[1][0] - sim_geo[0][0], sim_geo[1][1] - sim_geo[0][1]) > 1e-9
        ):
            return GeoTransform("similarity", sim_pixels=sim_pixels, sim_geo=sim_geo)
    return None


def format_dmm(value: float, kind: str) -> str:
    hemi = "N" if kind == "lat" else "E"
    if value < 0:
        hemi = "S" if kind == "lat" else "W"
    abs_value = abs(value)
    deg = int(abs_value)
    minutes = (abs_value - deg) * 60.0
    deg_width = 2 if kind == "lat" else 3
    return f"{deg:0{deg_width}d} {minutes:06.3f} {hemi}"


def parse_geo_coordinate(value: str, kind: str) -> Optional[float]:
    text = value.strip().upper()
    if not text:
        return None

    text = text.replace(",", ".")
    hemis = re.findall(r"[NSEW]", text)
    hemisphere = hemis[0] if hemis else None

    if hemisphere is not None:
        if kind == "lat" and hemisphere not in ("N", "S"):
            return None
        if kind == "lon" and hemisphere not in ("E", "W"):
            return None

    nums = [float(n) for n in re.findall(r"[-+]?\d+(?:\.\d+)?", text)]
    if not nums:
        return None

    deg_raw = nums[0]
    deg = abs(deg_raw)
    minutes = 0.0
    seconds = 0.0
    if len(nums) >= 2:
        minutes = nums[1]
    if len(nums) >= 3:
        seconds = nums[2]

    if minutes < 0 or seconds < 0 or minutes >= 60 or seconds >= 60:
        return None

    decimal = deg + minutes / 60.0 + seconds / 3600.0
    sign = -1.0 if deg_raw < 0 else 1.0
    if hemisphere in ("S", "W"):
        sign = -1.0
    elif hemisphere in ("N", "E"):
        sign = 1.0
    result = sign * decimal

    if kind == "lat" and abs(result) > 90:
        return None
    if kind == "lon" and abs(result) > 180:
        return None
    return result


class MapCanvas(QWidget):
    cursorMoved = Signal(float, float, object)
    selectionChanged = Signal(object, object)
    sheetsChanged = Signal()
    panBy = Signal(float, float)
    zoomRequest = Signal(float, object)

    MODE_SELECT = "select"
    MODE_NEW_SHEET = "new_sheet"
    MODE_ADD_CAL_POINT = "add_cal"

    def __init__(self) -> None:
        super().__init__()
        self.setMouseTracking(True)
        self.lang = "pl"
        self.image_path: Optional[Path] = None
        self.pixmap = QPixmap()
        self.sheets: list[Sheet] = []
        self.selected_sheet_idx: Optional[int] = None
        self.selected_point_idx: Optional[int] = None
        self.mode = self.MODE_SELECT
        self.new_sheet_temp_points: list[CalibrationPoint] = []
        self.dragging = False
        self.zoom_factor = 1.0
        self.is_panning = False
        self.pan_last_pos = QPointF()
        self.show_geo_grid = True
        self.display = DisplaySettings()
        self.cursor_x: Optional[float] = None
        self.cursor_y: Optional[float] = None

    def apply_display_settings(self, settings: DisplaySettings) -> None:
        self.display = settings
        self.update()

    def set_language(self, lang: str) -> None:
        self.lang = lang if lang in ("pl", "en") else "pl"

    @property
    def selected_sheet(self) -> Optional[Sheet]:
        if self.selected_sheet_idx is None:
            return None
        if self.selected_sheet_idx < 0 or self.selected_sheet_idx >= len(self.sheets):
            return None
        return self.sheets[self.selected_sheet_idx]

    @property
    def selected_point(self) -> Optional[CalibrationPoint]:
        s = self.selected_sheet
        if s is None or self.selected_point_idx is None:
            return None
        if self.selected_point_idx < 0 or self.selected_point_idx >= len(s.points):
            return None
        return s.points[self.selected_point_idx]

    def set_image(self, path: Path) -> bool:
        px = QPixmap(str(path))
        if px.isNull():
            return False
        self.pixmap = px
        self.image_path = path
        self.zoom_factor = 1.0
        self.updateGeometry()
        self.update()
        return True

    def sizeHint(self) -> QSize:
        if not self.pixmap.isNull():
            w = max(1, int(self.pixmap.width() * self.zoom_factor))
            h = max(1, int(self.pixmap.height() * self.zoom_factor))
            return QSize(w, h)
        return QSize(1000, 700)

    def minimumSizeHint(self) -> QSize:
        return QSize(400, 300)

    def set_zoom(self, factor: float) -> None:
        factor = max(0.1, min(8.0, factor))
        if abs(factor - self.zoom_factor) < 1e-9:
            return
        self.zoom_factor = factor
        self.updateGeometry()
        self.resize(self.sizeHint())
        self.update()

    def to_image_coords(self, pos: QPointF) -> tuple[float, float]:
        if self.zoom_factor <= 0:
            return pos.x(), pos.y()
        return pos.x() / self.zoom_factor, pos.y() / self.zoom_factor

    def clear_all(self) -> None:
        self.sheets = []
        self.selected_sheet_idx = None
        self.selected_point_idx = None
        self.new_sheet_temp_points = []
        self.mode = self.MODE_SELECT
        self.selectionChanged.emit(self.selected_sheet_idx, self.selected_point_idx)
        self.sheetsChanged.emit()
        self.update()

    def start_new_sheet(self) -> None:
        self.mode = self.MODE_NEW_SHEET
        self.new_sheet_temp_points = []
        self.selected_point_idx = None
        self.update()

    def close_new_sheet(self) -> bool:
        if self.mode != self.MODE_NEW_SHEET:
            return False
        if len(self.new_sheet_temp_points) < 3:
            return False
        sheet = Sheet(name=f"{t(self.lang, 'sheet_default')} {len(self.sheets) + 1}")
        for p in self.new_sheet_temp_points:
            sheet.points.append(CalibrationPoint(p.x, p.y, is_corner=True))
        self.sheets.append(sheet)
        self.selected_sheet_idx = len(self.sheets) - 1
        self.selected_point_idx = None
        self.new_sheet_temp_points = []
        self.mode = self.MODE_SELECT
        self.selectionChanged.emit(self.selected_sheet_idx, self.selected_point_idx)
        self.sheetsChanged.emit()
        self.update()
        return True

    def cancel_new_sheet(self) -> None:
        self.new_sheet_temp_points = []
        self.mode = self.MODE_SELECT
        self.update()

    def set_mode_select(self) -> None:
        self.mode = self.MODE_SELECT

    def set_mode_add_cal_point(self) -> None:
        self.mode = self.MODE_ADD_CAL_POINT

    def nearest_point(self, x: float, y: float, radius: float = 8.0) -> tuple[Optional[int], Optional[int]]:
        best = (None, None)
        best_d = radius
        for si, sheet in enumerate(self.sheets):
            for pi, p in enumerate(sheet.points):
                d = math.hypot(p.x - x, p.y - y)
                if d <= best_d:
                    best = (si, pi)
                    best_d = d
        return best

    def sheet_under(self, x: float, y: float) -> Optional[int]:
        for i in reversed(range(len(self.sheets))):
            if point_in_polygon(x, y, self.sheets[i].corners()):
                return i
        return None

    def geo_for_cursor(self, x: float, y: float) -> Optional[tuple[float, float]]:
        sheet = self.selected_sheet
        if sheet is None:
            idx = self.sheet_under(x, y)
            if idx is None:
                return None
            sheet = self.sheets[idx]
        transform = build_geo_transform(sheet.points)
        if transform is None:
            return None
        return transform.pixel_to_geo(x, y)

    def draw_geo_grid(self, p: QPainter) -> None:
        if self.pixmap.isNull() or not self.show_geo_grid:
            return

        sheet = self.selected_sheet or (self.sheets[0] if self.sheets else None)
        if sheet is None:
            return
        transform = build_geo_transform(sheet.points)
        if transform is None:
            return

        w = float(self.pixmap.width())
        h = float(self.pixmap.height())
        corner_geo = []
        for gx, gy in ((0.0, 0.0), (w, 0.0), (0.0, h), (w, h)):
            g = transform.pixel_to_geo(gx, gy)
            if g is not None:
                corner_geo.append(g)
        if len(corner_geo) < 2:
            return

        lon_vals = [g[0] for g in corner_geo]
        lat_vals = [g[1] for g in corner_geo]
        min_lon, max_lon = min(lon_vals), max(lon_vals)
        min_lat, max_lat = min(lat_vals), max(lat_vals)
        if abs(max_lon - min_lon) < 1e-12 or abs(max_lat - min_lat) < 1e-12:
            return

        def nice_step(span: float, target_lines: int = 8) -> float:
            raw = abs(span) / max(1, target_lines)
            if raw <= 0:
                return 1.0
            power = 10 ** math.floor(math.log10(raw))
            for m in (1.0, 2.0, 5.0, 10.0):
                step = m * power
                if step >= raw:
                    return step
            return 10.0 * power

        lon_step = nice_step(max_lon - min_lon)
        lat_step = nice_step(max_lat - min_lat)

        p.setPen(QPen(QColor(255, 255, 255, 90), 1))
        samples = 24

        lon = math.floor(min_lon / lon_step) * lon_step
        while lon <= max_lon + 1e-12:
            prev = None
            for i in range(samples + 1):
                lat = min_lat + (max_lat - min_lat) * (i / samples)
                xy = transform.geo_to_pixel(lon, lat)
                if xy is None:
                    prev = None
                    continue
                cur = QPointF(xy[0], xy[1])
                if prev is not None:
                    p.drawLine(prev, cur)
                prev = cur
            lon += lon_step

        lat = math.floor(min_lat / lat_step) * lat_step
        while lat <= max_lat + 1e-12:
            prev = None
            for i in range(samples + 1):
                lon = min_lon + (max_lon - min_lon) * (i / samples)
                xy = transform.geo_to_pixel(lon, lat)
                if xy is None:
                    prev = None
                    continue
                cur = QPointF(xy[0], xy[1])
                if prev is not None:
                    p.drawLine(prev, cur)
                prev = cur
            lat += lat_step

    def mouseMoveEvent(self, e: QMouseEvent) -> None:
        if self.is_panning:
            cur = e.position()
            delta = cur - self.pan_last_pos
            self.pan_last_pos = cur
            self.panBy.emit(delta.x(), delta.y())
            return

        x, y = self.to_image_coords(e.position())
        self.cursor_x = x
        self.cursor_y = y
        if self.dragging and self.selected_point is not None:
            self.selected_point.x = x
            self.selected_point.y = y
            self.update()
            self.sheetsChanged.emit()
        elif self.should_draw_cursor_guides():
            self.update()
        self.cursorMoved.emit(x, y, self.geo_for_cursor(x, y))

    def mousePressEvent(self, e: QMouseEvent) -> None:
        if e.button() == Qt.MiddleButton:
            self.is_panning = True
            self.pan_last_pos = e.position()
            self.setCursor(Qt.ClosedHandCursor)
            return

        if e.button() != Qt.LeftButton:
            return
        x, y = self.to_image_coords(e.position())
        if self.mode == self.MODE_NEW_SHEET:
            self.new_sheet_temp_points.append(CalibrationPoint(x=x, y=y, is_corner=True))
            self.update()
            return

        si, pi = self.nearest_point(x, y)
        if si is not None and pi is not None:
            self.selected_sheet_idx = si
            self.selected_point_idx = pi
            self.dragging = True
            self.selectionChanged.emit(si, pi)
            self.update()
            return

        if self.mode == self.MODE_ADD_CAL_POINT:
            if self.selected_sheet is None:
                if self.sheets:
                    self.selected_sheet_idx = 0
                else:
                    idx = self.sheet_under(x, y)
                    if idx is not None:
                        self.selected_sheet_idx = idx
            if self.selected_sheet is not None:
                cal_count = sum(1 for p in self.selected_sheet.points if not p.is_corner)
                if cal_count >= 9:
                    QMessageBox.information(self, t(self.lang, "limit_title"), t(self.lang, "limit_cal_points"))
                    return
                pred = self.geo_for_cursor(x, y)
                new_point = CalibrationPoint(x=x, y=y, is_corner=False)
                if pred is not None:
                    lon, lat = pred
                    new_point.lon = lon
                    new_point.lat = lat
                    new_point.lon_text = format_dmm(lon, "lon")
                    new_point.lat_text = format_dmm(lat, "lat")
                self.selected_sheet.points.append(new_point)
                self.selected_point_idx = len(self.selected_sheet.points) - 1
                self.selectionChanged.emit(self.selected_sheet_idx, self.selected_point_idx)
                self.sheetsChanged.emit()
                self.update()
            return

        idx = self.sheet_under(x, y)
        self.selected_sheet_idx = idx
        self.selected_point_idx = None
        self.selectionChanged.emit(self.selected_sheet_idx, self.selected_point_idx)
        self.update()

    def mouseReleaseEvent(self, e: QMouseEvent) -> None:
        if e.button() == Qt.MiddleButton:
            self.is_panning = False
            self.setCursor(Qt.ArrowCursor)
            return
        if e.button() == Qt.LeftButton:
            self.dragging = False
            self.update()

    def wheelEvent(self, e: QWheelEvent) -> None:
        if self.pixmap.isNull():
            return
        step = 1.15 if e.angleDelta().y() > 0 else 1 / 1.15
        self.zoomRequest.emit(self.zoom_factor * step, e.position())
        e.accept()

    def paintEvent(self, _event) -> None:
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#1e1e1e"))
        marker_items: list[tuple[float, float, QColor, bool, bool]] = []
        p.save()
        p.scale(self.zoom_factor, self.zoom_factor)
        if not self.pixmap.isNull():
            p.drawPixmap(0, 0, self.pixmap)
            self.draw_geo_grid(p)

        for si, sheet in enumerate(self.sheets):
            is_selected = si == self.selected_sheet_idx
            corners = sheet.corners()
            if len(corners) >= 2:
                pen = QPen(QColor("#00ff99") if is_selected else QColor("#00a2ff"))
                pen.setWidth(
                    self.display.outline_selected_width if is_selected else self.display.outline_width
                )
                pen.setCosmetic(True)
                p.setPen(pen)
                for i in range(len(corners)):
                    a = corners[i]
                    b = corners[(i + 1) % len(corners)]
                    p.drawLine(int(a.x), int(a.y), int(b.x), int(b.y))

            for pi, pt in enumerate(sheet.points):
                if pt.is_corner:
                    color = QColor("#ffb703")
                    is_corner = True
                else:
                    color = QColor("#8ecae6")
                    is_corner = False
                is_selected_point = is_selected and pi == self.selected_point_idx
                if is_selected_point:
                    color = QColor("#ff4d6d")
                marker_items.append((pt.x, pt.y, color, is_corner, is_selected_point))

        if self.mode == self.MODE_NEW_SHEET and self.new_sheet_temp_points:
            draft_pen = QPen(QColor("#f94144"), self.display.draft_outline_width, Qt.DashLine)
            draft_pen.setCosmetic(True)
            p.setPen(draft_pen)
            for i in range(1, len(self.new_sheet_temp_points)):
                a = self.new_sheet_temp_points[i - 1]
                b = self.new_sheet_temp_points[i]
                p.drawLine(int(a.x), int(a.y), int(b.x), int(b.y))
            for pt in self.new_sheet_temp_points:
                marker_items.append((pt.x, pt.y, QColor("#f94144"), True, False))
        p.restore()

        if self.should_draw_cursor_guides() and self.cursor_x is not None and self.cursor_y is not None:
            sx = self.cursor_x * self.zoom_factor
            sy = self.cursor_y * self.zoom_factor
            self.draw_cursor_guides(p, sx, sy)

        # Draw point symbols in screen space so their size stays constant regardless of zoom.
        for x, y, color, is_corner, is_selected_point in marker_items:
            sx = x * self.zoom_factor
            sy = y * self.zoom_factor
            arm = self.display.crosshair_arm_corner if is_corner else self.display.crosshair_arm_cal
            ring = self.display.crosshair_ring_corner if is_corner else self.display.crosshair_ring_cal
            if is_selected_point:
                arm += self.display.crosshair_selected_arm_bonus
                ring += self.display.crosshair_selected_ring_bonus
            self.draw_screen_target(p, sx, sy, color, arm, ring)

    @staticmethod
    def draw_screen_target(
        painter: QPainter,
        sx: float,
        sy: float,
        color: QColor,
        arm: int,
        ring: int,
    ) -> None:
        radius = max(5, arm - 2)
        gap = max(3, ring + 1)
        tick = max(4, ring + 2)

        def draw_cross(pen_color: QColor, width: int) -> None:
            painter.setPen(QPen(pen_color, width))
            # Main MAP cross with center gap.
            painter.drawLine(QPointF(sx - radius, sy), QPointF(sx - gap, sy))
            painter.drawLine(QPointF(sx + gap, sy), QPointF(sx + radius, sy))
            painter.drawLine(QPointF(sx, sy - radius), QPointF(sx, sy - gap))
            painter.drawLine(QPointF(sx, sy + gap), QPointF(sx, sy + radius))
            # Small ticks to form target grid impression.
            painter.drawLine(QPointF(sx - radius / 2, sy - tick / 2), QPointF(sx - radius / 2, sy + tick / 2))
            painter.drawLine(QPointF(sx + radius / 2, sy - tick / 2), QPointF(sx + radius / 2, sy + tick / 2))
            painter.drawLine(QPointF(sx - tick / 2, sy - radius / 2), QPointF(sx + tick / 2, sy - radius / 2))
            painter.drawLine(QPointF(sx - tick / 2, sy + radius / 2), QPointF(sx + tick / 2, sy + radius / 2))

        draw_cross(QColor(0, 0, 0, 180), 3)  # halo
        draw_cross(color, 1)
        painter.setPen(QPen(color, 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(QPointF(sx, sy), ring, ring)

    def should_draw_cursor_guides(self) -> bool:
        if self.mode == self.MODE_ADD_CAL_POINT:
            return True
        if self.dragging and self.selected_point is not None and not self.selected_point.is_corner:
            return True
        return False

    def draw_cursor_guides(self, painter: QPainter, sx: float, sy: float) -> None:
        base = QColor(self.display.cursor_guide_color)
        if not base.isValid():
            base = QColor("#FFD84D")
        base.setAlpha(self.display.cursor_guide_alpha)

        halo = QColor(0, 0, 0, min(255, self.display.cursor_guide_alpha))
        width = self.display.cursor_guide_width
        dash = self.display.cursor_guide_dash
        gap = self.display.cursor_guide_gap

        halo_pen = QPen(halo, width + 2, Qt.DashLine)
        halo_pen.setDashPattern([dash, gap])
        color_pen = QPen(base, width, Qt.DashLine)
        color_pen.setDashPattern([dash, gap])

        for pen in (halo_pen, color_pen):
            painter.setPen(pen)
            painter.drawLine(QPointF(0, sy), QPointF(float(self.width()), sy))
            painter.drawLine(QPointF(sx, 0), QPointF(sx, float(self.height())))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.display_settings, self.lang, self.display_settings_path, settings_error = load_display_settings()
        self.base_title = t(self.lang, "title")
        self.setWindowTitle(self.base_title)
        self.resize(1400, 900)
        self.project_path: Optional[Path] = None
        self.auto_fit_enabled = True

        self.canvas = MapCanvas()
        self.canvas.apply_display_settings(self.display_settings)
        self.canvas.set_language(self.lang)
        self.canvas.cursorMoved.connect(self.on_cursor_moved)
        self.canvas.selectionChanged.connect(self.on_canvas_selection_changed)
        self.canvas.sheetsChanged.connect(self.refresh_sheet_list)
        self.canvas.panBy.connect(self.on_canvas_pan_by)
        self.canvas.zoomRequest.connect(self.on_canvas_zoom_request)
        self.canvas_scroll = QScrollArea()
        self.canvas_scroll.setWidget(self.canvas)
        self.canvas_scroll.setWidgetResizable(False)
        self.canvas_scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.canvas_scroll.viewport().installEventFilter(self)

        root = QWidget()
        root_layout = QHBoxLayout(root)
        splitter = QSplitter()
        splitter.addWidget(self.canvas_scroll)
        splitter.addWidget(self.build_sidebar())
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)
        root_layout.addWidget(splitter)
        self.setCentralWidget(root)
        self.statusBar().showMessage(t(self.lang, "ready"))
        self.build_menu()
        self.update_window_title()
        if settings_error:
            QMessageBox.warning(self, t(self.lang, "settings_title"), settings_error)

    def tr(self, key: str, **kwargs) -> str:
        return t(self.lang, key, **kwargs)

    def update_window_title(self) -> None:
        if self.project_path is not None:
            self.setWindowTitle(f"{self.base_title} - {self.project_path.name}")
        else:
            self.setWindowTitle(self.base_title)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if self.auto_fit_enabled:
            self.fit_to_window()
        else:
            self.canvas.update()

    def eventFilter(self, watched, event) -> bool:
        if watched is self.canvas_scroll.viewport() and event.type() == QEvent.Resize:
            if self.auto_fit_enabled:
                self.fit_to_window()
            else:
                self.canvas.update()
        return super().eventFilter(watched, event)

    def build_sidebar(self) -> QWidget:
        panel = QFrame()
        panel.setFrameShape(QFrame.StyledPanel)
        layout = QVBoxLayout(panel)

        self.btn_new_sheet = QPushButton(self.tr("btn_new_sheet"))
        self.btn_close_sheet = QPushButton(self.tr("btn_close_sheet"))
        self.btn_cancel_sheet = QPushButton(self.tr("btn_cancel_drawing"))
        self.btn_add_cal = QPushButton(self.tr("btn_add_cal_point"))
        self.btn_select_mode = QPushButton(self.tr("btn_select_mode"))
        self.btn_delete_sheet = QPushButton(self.tr("btn_delete_sheet"))
        self.btn_use_as_corner = QPushButton(self.tr("btn_use_as_corner"))
        self.btn_zoom_in = QPushButton(self.tr("btn_zoom_in"))
        self.btn_zoom_out = QPushButton(self.tr("btn_zoom_out"))
        self.btn_zoom_reset = QPushButton(self.tr("btn_zoom_reset"))

        self.btn_new_sheet.clicked.connect(self.canvas.start_new_sheet)
        self.btn_close_sheet.clicked.connect(self.close_sheet_clicked)
        self.btn_cancel_sheet.clicked.connect(self.canvas.cancel_new_sheet)
        self.btn_add_cal.clicked.connect(self.canvas.set_mode_add_cal_point)
        self.btn_select_mode.clicked.connect(self.canvas.set_mode_select)
        self.btn_delete_sheet.clicked.connect(self.delete_sheet)
        self.btn_use_as_corner.clicked.connect(self.use_selected_point_as_corner)
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        self.btn_zoom_reset.clicked.connect(self.zoom_reset)

        layout.addWidget(self.btn_new_sheet)
        layout.addWidget(self.btn_close_sheet)
        layout.addWidget(self.btn_cancel_sheet)
        layout.addWidget(self.btn_add_cal)
        layout.addWidget(self.btn_select_mode)
        layout.addWidget(self.btn_delete_sheet)
        layout.addWidget(self.btn_use_as_corner)
        layout.addWidget(self.btn_zoom_in)
        layout.addWidget(self.btn_zoom_out)
        layout.addWidget(self.btn_zoom_reset)

        layout.addWidget(QLabel(self.tr("panel_sheets")))
        self.sheet_tree = QTreeWidget()
        self.sheet_tree.setHeaderHidden(True)
        self.sheet_tree.currentItemChanged.connect(self.on_tree_selection_changed)
        layout.addWidget(self.sheet_tree)

        meta_title = QLabel(self.tr("panel_sheet_meta"))
        layout.addWidget(meta_title)
        meta_form = QFormLayout()
        self.name_edit = QLineEdit()
        self.scale_edit = QLineEdit()
        self.name_edit.textEdited.connect(self.on_sheet_meta_changed)
        self.scale_edit.textEdited.connect(self.on_sheet_meta_changed)
        meta_form.addRow(self.tr("field_name"), self.name_edit)
        meta_form.addRow(self.tr("field_scale"), self.scale_edit)
        layout.addLayout(meta_form)

        point_title = QLabel(self.tr("panel_point_geo"))
        layout.addWidget(point_title)
        point_form = QFormLayout()
        self.lon_edit = QLineEdit()
        self.lat_edit = QLineEdit()
        self.lon_edit.editingFinished.connect(self.on_point_geo_changed)
        self.lat_edit.editingFinished.connect(self.on_point_geo_changed)
        point_form.addRow("Lon", self.lon_edit)
        point_form.addRow("Lat", self.lat_edit)
        layout.addLayout(point_form)

        layout.addStretch(1)
        return panel

    def build_menu(self) -> None:
        file_menu = self.menuBar().addMenu(self.tr("menu_file"))
        open_img = QAction(self.tr("menu_open_scan"), self)
        open_img.triggered.connect(self.open_image)
        file_menu.addAction(open_img)

        save_proj = QAction(self.tr("menu_save"), self)
        save_proj.triggered.connect(self.save_project)
        file_menu.addAction(save_proj)

        save_as_proj = QAction(self.tr("menu_save_as"), self)
        save_as_proj.triggered.connect(self.save_project_as)
        file_menu.addAction(save_as_proj)

        load_proj = QAction(self.tr("menu_load_project"), self)
        load_proj.triggered.connect(self.load_project)
        file_menu.addAction(load_proj)

        help_menu = self.menuBar().addMenu(self.tr("menu_help"))
        show_readme = QAction(self.tr("menu_readme"), self)
        show_readme.triggered.connect(self.show_readme_help)
        help_menu.addAction(show_readme)

    def show_readme_help(self) -> None:
        base = Path(__file__).resolve().parent
        if self.lang == "en":
            preferred = base / "README.en.md"
        else:
            preferred = base / "README.pl.md"
        readme_path = preferred if preferred.exists() else (base / "README.md")
        if not readme_path.exists():
            QMessageBox.warning(self, self.tr("help_title"), self.tr("help_missing_readme"))
            return
        try:
            content = readme_path.read_text(encoding="utf-8")
        except Exception as exc:
            QMessageBox.critical(self, self.tr("help_title"), self.tr("help_read_error", error=exc))
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(self.tr("help_dialog_title"))
        dlg.resize(900, 700)
        layout = QVBoxLayout(dlg)
        viewer = QTextBrowser(dlg)
        viewer.setMarkdown(content)
        viewer.setOpenExternalLinks(True)
        layout.addWidget(viewer)
        buttons = QDialogButtonBox(QDialogButtonBox.Close, parent=dlg)
        buttons.rejected.connect(dlg.reject)
        buttons.accepted.connect(dlg.accept)
        buttons.button(QDialogButtonBox.Close).clicked.connect(dlg.close)
        layout.addWidget(buttons)
        dlg.exec()

    def on_cursor_moved(self, x: float, y: float, geo: Optional[tuple[float, float]]) -> None:
        zoom_pct = int(round(self.canvas.zoom_factor * 100))
        if geo is None:
            msg = self.tr("status_cursor_no_geo", x=x, y=y, zoom=zoom_pct)
        else:
            lon, lat = geo
            msg = self.tr("status_cursor_geo", x=x, y=y, lon=lon, lat=lat, zoom=zoom_pct)
        self.statusBar().showMessage(msg)

    def on_canvas_pan_by(self, dx: float, dy: float) -> None:
        hbar = self.canvas_scroll.horizontalScrollBar()
        vbar = self.canvas_scroll.verticalScrollBar()
        hbar.setValue(hbar.value() - int(dx))
        vbar.setValue(vbar.value() - int(dy))

    def on_canvas_zoom_request(self, requested_zoom: float, anchor_widget_pos: QPointF) -> None:
        self.auto_fit_enabled = False
        self.apply_zoom(requested_zoom, anchor_widget_pos)

    def apply_zoom(self, new_zoom: float, anchor_widget_pos: Optional[QPointF] = None) -> None:
        old_zoom = self.canvas.zoom_factor
        if self.canvas.pixmap.isNull():
            return
        hbar = self.canvas_scroll.horizontalScrollBar()
        vbar = self.canvas_scroll.verticalScrollBar()
        h_old = hbar.value()
        v_old = vbar.value()

        if anchor_widget_pos is None:
            vp = self.canvas_scroll.viewport()
            anchor_widget_pos = QPointF(
                h_old + vp.width() / 2.0,
                v_old + vp.height() / 2.0,
            )

        if old_zoom <= 0:
            old_zoom = 1.0
        img_x = anchor_widget_pos.x() / old_zoom
        img_y = anchor_widget_pos.y() / old_zoom
        view_x = anchor_widget_pos.x() - h_old
        view_y = anchor_widget_pos.y() - v_old

        self.canvas.set_zoom(new_zoom)
        new_zoom = self.canvas.zoom_factor
        new_h = int(img_x * new_zoom - view_x)
        new_v = int(img_y * new_zoom - view_y)
        hbar.setValue(new_h)
        vbar.setValue(new_v)

    def zoom_in(self) -> None:
        self.auto_fit_enabled = False
        self.apply_zoom(self.canvas.zoom_factor * 1.15)

    def zoom_out(self) -> None:
        self.auto_fit_enabled = False
        self.apply_zoom(self.canvas.zoom_factor / 1.15)

    def zoom_reset(self) -> None:
        self.auto_fit_enabled = False
        self.apply_zoom(1.0)

    def fit_to_window(self) -> None:
        if self.canvas.pixmap.isNull():
            return
        vp = self.canvas_scroll.viewport()
        vw = max(1, vp.width())
        vh = max(1, vp.height())
        iw = max(1, self.canvas.pixmap.width())
        ih = max(1, self.canvas.pixmap.height())
        target_zoom = min(vw / iw, vh / ih)
        self.apply_zoom(target_zoom)

    def close_sheet_clicked(self) -> None:
        if not self.canvas.close_new_sheet():
            QMessageBox.warning(self, self.tr("error_title"), self.tr("error_sheet_min_points"))

    def delete_sheet(self) -> None:
        idx = self.canvas.selected_sheet_idx
        if idx is None:
            return
        del self.canvas.sheets[idx]
        if self.canvas.sheets:
            self.canvas.selected_sheet_idx = max(0, idx - 1)
        else:
            self.canvas.selected_sheet_idx = None
        self.canvas.selected_point_idx = None
        self.refresh_sheet_list()
        self.canvas.update()

    def use_selected_point_as_corner(self) -> None:
        point = self.canvas.selected_point
        if point is None:
            return
        if point.is_corner:
            return
        answer = QMessageBox.question(
            self,
            self.tr("link_point_title"),
            self.tr("link_point_question"),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if answer != QMessageBox.Yes:
            return
        point.is_corner = True
        self.refresh_sheet_list()
        self.canvas.update()

    def refresh_sheet_list(self) -> None:
        self.sheet_tree.blockSignals(True)
        self.sheet_tree.clear()

        selected_item: Optional[QTreeWidgetItem] = None
        for si, s in enumerate(self.canvas.sheets):
            top = QTreeWidgetItem([self.tr("tree_sheet_item", name=s.name, count=len(s.points))])
            top.setData(0, Qt.UserRole, {"type": "sheet", "sheet_idx": si})
            self.sheet_tree.addTopLevelItem(top)

            corners_item = QTreeWidgetItem([self.tr("tree_crop_outline")])
            corners_item.setData(0, Qt.UserRole, {"type": "corner_group", "sheet_idx": si})
            top.addChild(corners_item)

            cal_item = QTreeWidgetItem([self.tr("tree_cal_points")])
            cal_item.setData(0, Qt.UserRole, {"type": "cal_group", "sheet_idx": si})
            top.addChild(cal_item)

            corner_no = 1
            cal_no = 1
            for pi, p in enumerate(s.points):
                if p.is_corner:
                    label = self.tr("tree_corner_item", n=corner_no, x=p.x, y=p.y)
                    parent = corners_item
                    corner_no += 1
                else:
                    label = self.tr("tree_point_item", n=cal_no, x=p.x, y=p.y)
                    parent = cal_item
                    cal_no += 1
                point_item = QTreeWidgetItem([label])
                point_item.setData(0, Qt.UserRole, {"type": "point", "sheet_idx": si, "point_idx": pi})
                parent.addChild(point_item)

                if si == self.canvas.selected_sheet_idx and pi == self.canvas.selected_point_idx:
                    selected_item = point_item

            top.setExpanded(True)
            corners_item.setExpanded(True)
            cal_item.setExpanded(True)

            if si == self.canvas.selected_sheet_idx and self.canvas.selected_point_idx is None:
                selected_item = top

        if selected_item is not None:
            self.sheet_tree.setCurrentItem(selected_item)
        self.sheet_tree.blockSignals(False)
        self.refresh_editors()

    def on_tree_selection_changed(self, current: Optional[QTreeWidgetItem], _previous: Optional[QTreeWidgetItem]) -> None:
        if current is None:
            return
        data = current.data(0, Qt.UserRole)
        if not isinstance(data, dict):
            return

        sheet_idx = data.get("sheet_idx")
        if not isinstance(sheet_idx, int) or sheet_idx < 0 or sheet_idx >= len(self.canvas.sheets):
            return

        self.canvas.selected_sheet_idx = sheet_idx
        item_type = data.get("type")
        if item_type == "point":
            point_idx = data.get("point_idx")
            if isinstance(point_idx, int):
                self.canvas.selected_point_idx = point_idx
            else:
                self.canvas.selected_point_idx = None
        else:
            self.canvas.selected_point_idx = None

        self.canvas.selectionChanged.emit(self.canvas.selected_sheet_idx, self.canvas.selected_point_idx)
        self.canvas.update()

    def on_canvas_selection_changed(self, _sheet_idx, _point_idx) -> None:
        self.refresh_sheet_list()
        self.refresh_editors()

    def refresh_editors(self) -> None:
        sheet = self.canvas.selected_sheet
        point = self.canvas.selected_point

        self.name_edit.blockSignals(True)
        self.scale_edit.blockSignals(True)
        self.lon_edit.blockSignals(True)
        self.lat_edit.blockSignals(True)

        if sheet is None:
            self.name_edit.setText("")
            self.scale_edit.setText("")
        else:
            self.name_edit.setText(sheet.name)
            self.scale_edit.setText(sheet.scale)

        if point is None:
            self.lon_edit.setText("")
            self.lat_edit.setText("")
        else:
            lon_display = point.lon_text if point.lon_text else ("" if point.lon is None else f"{point.lon:.8f}")
            lat_display = point.lat_text if point.lat_text else ("" if point.lat is None else f"{point.lat:.8f}")
            self.lon_edit.setText(lon_display)
            self.lat_edit.setText(lat_display)

        self.name_edit.blockSignals(False)
        self.scale_edit.blockSignals(False)
        self.lon_edit.blockSignals(False)
        self.lat_edit.blockSignals(False)

    def on_sheet_meta_changed(self) -> None:
        sheet = self.canvas.selected_sheet
        if sheet is None:
            return
        sheet.name = self.name_edit.text().strip() or self.tr("sheet_default")
        sheet.scale = self.scale_edit.text().strip()
        self.refresh_sheet_list()
        self.canvas.update()

    def on_point_geo_changed(self) -> None:
        point = self.canvas.selected_point
        if point is None:
            return
        lon_text = self.lon_edit.text().strip()
        lat_text = self.lat_edit.text().strip()
        lon = parse_geo_coordinate(lon_text, "lon")
        lat = parse_geo_coordinate(lat_text, "lat")

        if lon_text and lon is None:
            QMessageBox.warning(
                self,
                self.tr("error_lon_format_title"),
                self.tr("error_lon_format_msg"),
            )
        if lat_text and lat is None:
            QMessageBox.warning(
                self,
                self.tr("error_lat_format_title"),
                self.tr("error_lat_format_msg"),
            )

        point.lon_text = lon_text
        point.lat_text = lat_text
        point.lon = lon
        point.lat = lat
        self.refresh_editors()
        self.canvas.update()

    def open_image(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(
            self,
            self.tr("dialog_pick_scan"),
            "",
            self.tr("dialog_image_filter"),
        )
        if not path_str:
            return
        path = Path(path_str)
        if not self.canvas.set_image(path):
            QMessageBox.critical(self, self.tr("error_title"), self.tr("error_open_image"))
            return
        self.canvas.clear_all()
        self.canvas.sheets.append(Sheet(name=path.stem or f"{self.tr('sheet_default')} 1"))
        self.canvas.selected_sheet_idx = 0
        self.canvas.selected_point_idx = None
        self.canvas.sheetsChanged.emit()
        self.auto_fit_enabled = True
        QTimer.singleShot(0, self.fit_to_window)
        self.project_path = None
        self.update_window_title()

    def save_project(self) -> None:
        if self.project_path is not None:
            self.write_project_to_path(self.project_path)
            return

        self.save_project_as()

    def save_project_as(self) -> None:
        if self.canvas.image_path is not None:
            suggested_path = self.canvas.image_path.with_suffix(".json")
        elif self.project_path is not None:
            suggested_path = self.project_path
        else:
            suggested_path = Path.cwd() / self.tr("default_project_name")

        path_str, _ = QFileDialog.getSaveFileName(
            self,
            self.tr("dialog_save_as"),
            str(suggested_path),
            self.tr("dialog_project_filter"),
        )
        if not path_str:
            return
        path = Path(path_str)
        if path.suffix == "":
            path = path.with_suffix(".json")

        self.write_project_to_path(path)

    def write_project_to_path(self, path: Path) -> None:
        data = {
            "image_path": str(self.canvas.image_path) if self.canvas.image_path else "",
            "sheets": [s.to_json() for s in self.canvas.sheets],
        }
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.project_path = path
        self.update_window_title()

    def load_project(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(
            self,
            self.tr("dialog_load_project"),
            "",
            self.tr("dialog_project_filter"),
        )
        if not path_str:
            return
        self.load_project_path(Path(path_str))

    def load_project_path(self, path: Path) -> bool:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            QMessageBox.critical(self, self.tr("error_title"), self.tr("error_load_project", error=exc))
            return False

        image_path = Path(data.get("image_path", ""))
        if image_path and image_path.exists():
            if not self.canvas.set_image(image_path):
                QMessageBox.critical(self, self.tr("error_title"), self.tr("error_open_project_image"))
                return False
        else:
            QMessageBox.warning(self, self.tr("warning_title"), self.tr("warning_missing_project_image"))
            self.canvas.pixmap = QPixmap()
        self.canvas.sheets = [Sheet.from_json(s) for s in data.get("sheets", [])]
        self.canvas.selected_sheet_idx = 0 if self.canvas.sheets else None
        self.canvas.selected_point_idx = None
        self.canvas.update()
        self.refresh_sheet_list()
        self.auto_fit_enabled = True
        QTimer.singleShot(0, self.fit_to_window)
        self.project_path = path
        self.update_window_title()
        return True


def main() -> int:
    app = QApplication(sys.argv)
    win = MainWindow()
    if len(sys.argv) >= 2:
        arg_path = Path(sys.argv[1]).expanduser()
        if arg_path.suffix.lower() == ".json":
            win.load_project_path(arg_path)
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

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
    QComboBox,
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
    QProgressDialog,
    QPushButton,
    QScrollArea,
    QSplitter,
    QSpinBox,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from kap_export import (
    KapExportJob,
    KapPolygonPoint,
    KapReference,
    run_kap_export_jobs,
    sanitize_kap_stem,
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
        "btn_add_outline_point": "Dodaj punkt obrysu",
        "btn_select_mode": "Tryb wyboru",
        "btn_delete_sheet": "Usuń zaznaczony arkusz",
        "btn_delete_point": "Usuń zaznaczony punkt",
        "btn_use_as_corner": "Dodaj punkt kalibracji do obrysu",
        "btn_save_point": "Zapisz punkt",
        "btn_zoom_in": "Zoom +",
        "btn_zoom_out": "Zoom -",
        "btn_zoom_reset": "Zoom 100%",
        "panel_sheets": "Arkusze",
        "panel_sheet_meta": "Metadane arkusza",
        "field_name": "Nazwa",
        "field_scale": "Skala",
        "field_lat": "Lat",
        "field_lon": "Lon",
        "panel_point_geo": "Punkt (lat/lon)",
        "menu_file": "Plik",
        "menu_open_scan": "Dodaj skan mapy",
        "menu_save": "Zapisz",
        "menu_save_as": "Zapisz jako...",
        "menu_load_project": "Wczytaj album",
        "menu_import_map": "Importuj MAP...",
        "menu_export_kap": "Eksportuj KAP...",
        "menu_export_all_kap": "Eksportuj wszystkie arkusze do KAP...",
        "menu_settings": "Ustawienia",
        "menu_edit_settings": "Edytuj ustawienia...",
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
        "tree_scan_item": "Skan: {name}",
        "tree_sheet_item": "{name} ({count} pkt)",
        "tree_corner_item": "Narożnik {n}: x={x:.1f}, y={y:.1f}",
        "tree_point_item": "Punkt {n}: x={x:.1f}, y={y:.1f}",
        "error_lon_format_title": "Błąd formatu lon",
        "error_lon_format_msg": "Nie udało się odczytać długości geograficznej. Obsługiwane: DD, DMM, DMS + półkula (E/W).",
        "error_lat_format_title": "Błąd formatu lat",
        "error_lat_format_msg": "Nie udało się odczytać szerokości geograficznej. Obsługiwane: DD, DMM, DMS + półkula (N/S).",
        "error_point_not_selected": "Najpierw zaznacz punkt do edycji.",
        "error_point_save_title": "Błąd zapisu punktu",
        "error_point_save_msg": "Popraw błędne współrzędne i zapisz ponownie.",
        "unsaved_point_title": "Niezapisane zmiany punktu",
        "unsaved_point_msg": "Punkt ma niezapisane zmiany. Co zrobić?",
        "status_grid_hint": "Siatka pojawi się po zapisaniu min. 2 punktów z poprawnym Lat/Lon.",
        "dialog_pick_scan": "Wybierz skan mapy",
        "dialog_image_filter": "Image (*.tif *.tiff *.bmp *.png *.jpg *.jpeg)",
        "error_open_image": "Nie udało się otworzyć obrazu.",
        "dialog_save_as": "Zapisz projekt jako",
        "dialog_project_filter": "MapCal Project (*.json)",
        "default_project_name": "projekt.json",
        "dialog_load_project": "Wczytaj album",
        "dialog_import_map": "Importuj pliki MAP",
        "dialog_map_filter": "Ozi Map (*.map)",
        "dialog_export_kap": "Wybierz katalog docelowy KAP",
        "error_load_project": "Nie udało się wczytać projektu:\n{error}",
        "error_open_project_image": "Nie udało się otworzyć obrazu zapisanego w projekcie.",
        "warning_title": "Uwaga",
        "confirm_title": "Potwierdzenie",
        "warning_missing_project_image": "Brak obrazu z projektu. Wczytano tylko dane arkuszy.",
        "warning_missing_scan_image": "Brak pliku skanu: {path}",
        "confirm_delete_sheet": "Usunąć zaznaczony arkusz?",
        "confirm_delete_point": "Usunąć zaznaczony punkt?",
        "confirm_delete_sheet_named": "Usunąć arkusz \"{name}\"?",
        "confirm_delete_point_named": "Usunąć punkt {kind} #{index}?",
        "point_kind_outline": "obrysu",
        "point_kind_calibration": "kalibracyjny",
        "import_summary_title": "Import MAP",
        "import_summary": "Zaimportowano: {ok}\nBłędy: {fail}",
        "export_title": "Eksport KAP",
        "export_no_scan": "Brak aktywnego skanu do eksportu.",
        "export_no_scans": "Album nie zawiera skanów do eksportu.",
        "export_no_sheets": "Aktywny skan nie zawiera arkuszy.",
        "export_no_sheets_all": "Album nie zawiera arkuszy do eksportu.",
        "export_image_missing": "Nie znaleziono pliku obrazu skanu: {path}",
        "export_image_open_error": "Nie udało się odczytać rozmiaru obrazu skanu.",
        "export_sheet_scale_invalid": "Arkusz \"{name}\": niepoprawna skala \"{scale}\".",
        "export_sheet_corners_missing": "Arkusz \"{name}\": obrys musi mieć co najmniej 3 punkty.",
        "export_sheet_geo_missing": "Arkusz \"{name}\": brak współrzędnych geo dla punktów obrysu.",
        "export_sheet_imgkap_failed": "Arkusz \"{name}\": imgkap zwrócił błąd.",
        "export_sheet_imgkap_missing": "Nie znaleziono binarki imgkap: {path}",
        "export_summary": "Eksport zakończony. OK: {ok}, błędy: {fail}\nKatalog: {out}",
        "export_progress_title": "Eksport KAP",
        "export_progress_label": "Generowanie KAP ({current}/{total}): {name}",
        "export_progress_cancel": "Anuluj",
        "export_cancelled": "Eksport przerwany przez użytkownika ({done}/{total}).\nKatalog: {out}",
        "settings_dialog_title": "Ustawienia programu",
        "settings_field_language": "Język",
        "settings_field_imgkap_path": "Ścieżka do binarki imgkap",
        "settings_field_sounding_datum": "KAP soundingDatum",
        "settings_field_imgkap_work_dir": "Katalog debug imgkap (tmp+log)",
        "settings_lang_pl": "Polski",
        "settings_lang_en": "English",
        "settings_field_outline_width": "Grubość obrysu",
        "settings_field_outline_selected_width": "Grubość obrysu (zaznaczony)",
        "settings_field_draft_outline_width": "Grubość obrysu roboczego",
        "settings_field_crosshair_arm_corner": "Celownik obrysu: długość ramienia",
        "settings_field_crosshair_arm_cal": "Celownik kalibracji: długość ramienia",
        "settings_field_crosshair_ring_corner": "Celownik obrysu: promień środka",
        "settings_field_crosshair_ring_cal": "Celownik kalibracji: promień środka",
        "settings_field_crosshair_selected_arm_bonus": "Celownik zaznaczony: +ramię",
        "settings_field_crosshair_selected_ring_bonus": "Celownik zaznaczony: +promień",
        "settings_field_cursor_guide_width": "Linie kursora: grubość",
        "settings_field_cursor_guide_alpha": "Linie kursora: przezroczystość",
        "settings_field_cursor_guide_dash": "Linie kursora: długość kreski",
        "settings_field_cursor_guide_gap": "Linie kursora: odstęp kresek",
        "settings_field_cursor_guide_color": "Linie kursora: kolor (#RRGGBB)",
        "settings_save_ok": "Ustawienia zapisane do: {path}",
        "settings_save_error": "Nie udało się zapisać ustawień:\n{error}",
        "settings_invalid_color": "Niepoprawny kolor. Użyj formatu #RRGGBB.",
        "settings_lang_restart_hint": "Zmiana języka będzie w pełni widoczna po ponownym uruchomieniu aplikacji.",
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
        "btn_add_outline_point": "Add outline point",
        "btn_select_mode": "Select mode",
        "btn_delete_sheet": "Delete selected sheet",
        "btn_delete_point": "Delete selected point",
        "btn_use_as_corner": "Use calibration point in crop outline",
        "btn_save_point": "Save point",
        "btn_zoom_in": "Zoom +",
        "btn_zoom_out": "Zoom -",
        "btn_zoom_reset": "Zoom 100%",
        "panel_sheets": "Sheets",
        "panel_sheet_meta": "Sheet metadata",
        "field_name": "Name",
        "field_scale": "Scale",
        "field_lat": "Lat",
        "field_lon": "Lon",
        "panel_point_geo": "Point (lat/lon)",
        "menu_file": "File",
        "menu_open_scan": "Add map scan",
        "menu_save": "Save",
        "menu_save_as": "Save as...",
        "menu_load_project": "Load album",
        "menu_import_map": "Import MAP...",
        "menu_export_kap": "Export KAP...",
        "menu_export_all_kap": "Export all sheets to KAP...",
        "menu_settings": "Settings",
        "menu_edit_settings": "Edit settings...",
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
        "tree_scan_item": "Scan: {name}",
        "tree_sheet_item": "{name} ({count} pt)",
        "tree_corner_item": "Corner {n}: x={x:.1f}, y={y:.1f}",
        "tree_point_item": "Point {n}: x={x:.1f}, y={y:.1f}",
        "error_lon_format_title": "Lon format error",
        "error_lon_format_msg": "Failed to parse longitude. Supported: DD, DMM, DMS + hemisphere (E/W).",
        "error_lat_format_title": "Lat format error",
        "error_lat_format_msg": "Failed to parse latitude. Supported: DD, DMM, DMS + hemisphere (N/S).",
        "error_point_not_selected": "Select a point first.",
        "error_point_save_title": "Point save error",
        "error_point_save_msg": "Fix invalid coordinates and save again.",
        "unsaved_point_title": "Unsaved point changes",
        "unsaved_point_msg": "Point has unsaved changes. What do you want to do?",
        "status_grid_hint": "Grid appears after saving at least 2 points with valid Lat/Lon.",
        "dialog_pick_scan": "Select map scan",
        "dialog_image_filter": "Image (*.tif *.tiff *.bmp *.png *.jpg *.jpeg)",
        "error_open_image": "Failed to open image.",
        "dialog_save_as": "Save project as",
        "dialog_project_filter": "MapCal Project (*.json)",
        "default_project_name": "project.json",
        "dialog_load_project": "Load album",
        "dialog_import_map": "Import MAP files",
        "dialog_map_filter": "Ozi Map (*.map)",
        "dialog_export_kap": "Select output directory for KAP",
        "error_load_project": "Failed to load project:\n{error}",
        "error_open_project_image": "Failed to open image stored in project.",
        "warning_title": "Warning",
        "confirm_title": "Confirmation",
        "warning_missing_project_image": "Project image is missing. Loaded only sheet data.",
        "warning_missing_scan_image": "Missing scan file: {path}",
        "confirm_delete_sheet": "Delete selected sheet?",
        "confirm_delete_point": "Delete selected point?",
        "confirm_delete_sheet_named": "Delete sheet \"{name}\"?",
        "confirm_delete_point_named": "Delete {kind} point #{index}?",
        "point_kind_outline": "outline",
        "point_kind_calibration": "calibration",
        "import_summary_title": "MAP import",
        "import_summary": "Imported: {ok}\nFailed: {fail}",
        "export_title": "KAP export",
        "export_no_scan": "No active scan to export.",
        "export_no_scans": "Album does not contain scans to export.",
        "export_no_sheets": "Current scan has no sheets.",
        "export_no_sheets_all": "Album does not contain sheets to export.",
        "export_image_missing": "Scan image file not found: {path}",
        "export_image_open_error": "Failed to read scan image dimensions.",
        "export_sheet_scale_invalid": "Sheet \"{name}\": invalid scale \"{scale}\".",
        "export_sheet_corners_missing": "Sheet \"{name}\": crop outline must have at least 3 points.",
        "export_sheet_geo_missing": "Sheet \"{name}\": missing geo coordinates for outline points.",
        "export_sheet_imgkap_failed": "Sheet \"{name}\": imgkap returned an error.",
        "export_sheet_imgkap_missing": "imgkap executable not found: {path}",
        "export_summary": "Export finished. OK: {ok}, errors: {fail}\nDirectory: {out}",
        "export_progress_title": "KAP export",
        "export_progress_label": "Generating KAP ({current}/{total}): {name}",
        "export_progress_cancel": "Cancel",
        "export_cancelled": "Export canceled by user ({done}/{total}).\nDirectory: {out}",
        "settings_dialog_title": "Application settings",
        "settings_field_language": "Language",
        "settings_field_imgkap_path": "Path to imgkap executable",
        "settings_field_sounding_datum": "KAP soundingDatum",
        "settings_field_imgkap_work_dir": "imgkap debug directory (tmp+log)",
        "settings_lang_pl": "Polish",
        "settings_lang_en": "English",
        "settings_field_outline_width": "Outline width",
        "settings_field_outline_selected_width": "Outline width (selected)",
        "settings_field_draft_outline_width": "Draft outline width",
        "settings_field_crosshair_arm_corner": "Outline crosshair: arm length",
        "settings_field_crosshair_arm_cal": "Calibration crosshair: arm length",
        "settings_field_crosshair_ring_corner": "Outline crosshair: center radius",
        "settings_field_crosshair_ring_cal": "Calibration crosshair: center radius",
        "settings_field_crosshair_selected_arm_bonus": "Selected crosshair: +arm",
        "settings_field_crosshair_selected_ring_bonus": "Selected crosshair: +radius",
        "settings_field_cursor_guide_width": "Cursor guides: width",
        "settings_field_cursor_guide_alpha": "Cursor guides: alpha",
        "settings_field_cursor_guide_dash": "Cursor guides: dash length",
        "settings_field_cursor_guide_gap": "Cursor guides: dash gap",
        "settings_field_cursor_guide_color": "Cursor guides: color (#RRGGBB)",
        "settings_save_ok": "Settings saved to: {path}",
        "settings_save_error": "Failed to save settings:\n{error}",
        "settings_invalid_color": "Invalid color. Use #RRGGBB format.",
        "settings_lang_restart_hint": "Language change is fully applied after restarting the app.",
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
class Scan:
    name: str
    image_path: str
    sheets: list[Sheet] = field(default_factory=list)

    def to_json(self) -> dict:
        return {
            "name": self.name,
            "image_path": self.image_path,
            "sheets": [s.to_json() for s in self.sheets],
        }

    @staticmethod
    def from_json(data: dict) -> "Scan":
        return Scan(
            name=data.get("name", ""),
            image_path=str(data.get("image_path", "")),
            sheets=[Sheet.from_json(s) for s in data.get("sheets", [])],
        )


@dataclass
class MapImportEntry:
    scan_name: str
    image_path: str
    sheet: Sheet


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

    def to_dict(self) -> dict:
        return {
            "outline_width": self.outline_width,
            "outline_selected_width": self.outline_selected_width,
            "draft_outline_width": self.draft_outline_width,
            "crosshair_arm_corner": self.crosshair_arm_corner,
            "crosshair_arm_cal": self.crosshair_arm_cal,
            "crosshair_ring_corner": self.crosshair_ring_corner,
            "crosshair_ring_cal": self.crosshair_ring_cal,
            "crosshair_selected_arm_bonus": self.crosshair_selected_arm_bonus,
            "crosshair_selected_ring_bonus": self.crosshair_selected_ring_bonus,
            "cursor_guide_width": self.cursor_guide_width,
            "cursor_guide_alpha": self.cursor_guide_alpha,
            "cursor_guide_dash": self.cursor_guide_dash,
            "cursor_guide_gap": self.cursor_guide_gap,
            "cursor_guide_color": self.cursor_guide_color,
        }


def load_display_settings() -> tuple[DisplaySettings, str, Optional[Path], str, str, str, Optional[str]]:
    candidates = [Path.cwd() / ".pymapcal", Path.home() / ".pymapcal"]
    for path in candidates:
        if not path.exists():
            continue
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(raw, dict):
                return DisplaySettings(), "pl", path, "imgkap", "UNKNOWN", "", t("pl", "settings_json_object")
            lang = raw.get("language", "pl")
            if lang not in ("pl", "en"):
                lang = "pl"
            imgkap_path = raw.get("imgkap_path", "imgkap")
            if not isinstance(imgkap_path, str) or not imgkap_path.strip():
                imgkap_path = "imgkap"
            sounding_datum = raw.get("kap_sounding_datum", "UNKNOWN")
            if not isinstance(sounding_datum, str) or not sounding_datum.strip():
                sounding_datum = "UNKNOWN"
            imgkap_work_dir = raw.get("imgkap_work_dir", "")
            if not isinstance(imgkap_work_dir, str):
                imgkap_work_dir = ""
            if isinstance(raw.get("display"), dict):
                source = raw["display"]
            else:
                source = raw
            return (
                DisplaySettings.from_dict(source),
                lang,
                path,
                imgkap_path.strip(),
                sounding_datum.strip(),
                imgkap_work_dir.strip(),
                None,
            )
        except Exception as exc:
            return DisplaySettings(), "pl", path, "imgkap", "UNKNOWN", "", t("pl", "settings_load_error", error=exc)
    return DisplaySettings(), "pl", None, "imgkap", "UNKNOWN", "", None


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


def _hemisphere_sign(h: str) -> float:
    up = h.strip().upper()
    if up in ("S", "W"):
        return -1.0
    return 1.0


def _parse_dmm_fields(deg_text: str, min_text: str, hemi_text: str) -> Optional[float]:
    try:
        deg = abs(float(deg_text.strip()))
        minutes = float(min_text.strip())
    except ValueError:
        return None
    if minutes < 0 or minutes >= 60:
        return None
    return _hemisphere_sign(hemi_text) * (deg + minutes / 60.0)


def _split_map_line(line: str) -> list[str]:
    return [part.strip() for part in line.strip().split(",")]


def _resolve_scan_path_from_map(map_path: Path, image_line: str, path_line: str) -> Path:
    candidates: list[Path] = []
    if image_line:
        candidates.append(map_path.parent / image_line)
    if path_line:
        normalized = path_line.replace("\\", "/").strip()
        candidates.append(Path(normalized))
        candidates.append(map_path.parent / Path(normalized).name)
    for c in candidates:
        if c.exists():
            return c
    if candidates:
        return candidates[0]
    return map_path.with_suffix(".bmp")


def parse_ozi_map_file(map_path: Path) -> Optional[MapImportEntry]:
    try:
        lines = map_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return None
    if len(lines) < 3:
        return None
    if not lines[0].startswith("OziExplorer Map Data File"):
        return None

    image_line = lines[1].strip()
    path_line = lines[2].strip()
    scan_path = _resolve_scan_path_from_map(map_path, image_line, path_line)
    scan_name = map_path.stem

    sheet = Sheet(name=map_path.stem or "Sheet")
    boundary_xy: dict[int, tuple[float, float]] = {}
    boundary_geo: dict[int, tuple[float, float]] = {}

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("Point") and ",xy," in line:
            parts = _split_map_line(line)
            if len(parts) < 12:
                continue
            x_text, y_text = parts[2], parts[3]
            if not x_text or not y_text:
                continue
            try:
                x = float(x_text)
                y = float(y_text)
            except ValueError:
                continue
            lat = _parse_dmm_fields(parts[6], parts[7], parts[8])
            lon = _parse_dmm_fields(parts[9], parts[10], parts[11])
            p = CalibrationPoint(x=x, y=y, is_corner=False)
            if lon is not None and lat is not None:
                p.lon = lon
                p.lat = lat
                p.lon_text = format_dmm(lon, "lon")
                p.lat_text = format_dmm(lat, "lat")
            sheet.points.append(p)
            continue

        if line.startswith("MMPXY"):
            parts = _split_map_line(line)
            if len(parts) >= 4:
                try:
                    idx = int(parts[1])
                    boundary_xy[idx] = (float(parts[2]), float(parts[3]))
                except ValueError:
                    pass
            continue

        if line.startswith("MMPLL"):
            parts = _split_map_line(line)
            if len(parts) >= 4:
                try:
                    idx = int(parts[1])
                    lon = float(parts[2])
                    lat = float(parts[3])
                    boundary_geo[idx] = (lon, lat)
                except ValueError:
                    pass
            continue

    for idx in sorted(set(boundary_xy.keys()) & set(boundary_geo.keys())):
        x, y = boundary_xy[idx]
        lon, lat = boundary_geo[idx]
        sheet.points.append(
            CalibrationPoint(
                x=x,
                y=y,
                lon=lon,
                lat=lat,
                lon_text=format_dmm(lon, "lon"),
                lat_text=format_dmm(lat, "lat"),
                is_corner=True,
            )
        )

    return MapImportEntry(scan_name=scan_name, image_path=str(scan_path), sheet=sheet)


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


def parse_scale_value(value: str) -> Optional[int]:
    text = value.strip()
    if not text:
        return None
    compact = text.replace(" ", "").replace("\u00A0", "")
    for sep in (":", "/"):
        if sep in compact:
            compact = compact.split(sep)[-1]
    digits = "".join(ch for ch in compact if ch.isdigit())
    if not digits:
        return None
    try:
        out = int(digits)
    except ValueError:
        return None
    if out <= 0:
        return None
    return out


class MapCanvas(QWidget):
    cursorMoved = Signal(float, float, object)
    selectionChanged = Signal(object, object)
    sheetsChanged = Signal()
    panBy = Signal(float, float)
    zoomRequest = Signal(float, object)

    MODE_SELECT = "select"
    MODE_NEW_SHEET = "new_sheet"
    MODE_ADD_CAL_POINT = "add_cal"
    MODE_ADD_OUTLINE_POINT = "add_outline"

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

    def set_mode_add_outline_point(self) -> None:
        self.mode = self.MODE_ADD_OUTLINE_POINT

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
                    msg = QMessageBox(self)
                    msg.setIcon(QMessageBox.Information)
                    msg.setWindowTitle(t(self.lang, "limit_title"))
                    msg.setText(t(self.lang, "limit_cal_points"))
                    msg.setWindowModality(Qt.ApplicationModal)
                    app = QApplication.instance()
                    if app is not None and app.applicationState() == Qt.ApplicationState.ApplicationActive:
                        msg.setWindowFlag(Qt.WindowStaysOnTopHint, True)
                    msg.exec()
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

        if self.mode == self.MODE_ADD_OUTLINE_POINT:
            if self.selected_sheet is None:
                if self.sheets:
                    self.selected_sheet_idx = 0
                else:
                    idx = self.sheet_under(x, y)
                    if idx is not None:
                        self.selected_sheet_idx = idx
            if self.selected_sheet is not None:
                self.selected_sheet.points.append(CalibrationPoint(x=x, y=y, is_corner=True))
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
        (
            self.display_settings,
            self.lang,
            self.display_settings_path,
            self.imgkap_path,
            self.kap_sounding_datum,
            self.imgkap_work_dir,
            settings_error,
        ) = load_display_settings()
        self.base_title = t(self.lang, "title")
        self.setWindowTitle(self.base_title)
        self.resize(1400, 900)
        self.project_path: Optional[Path] = None
        self.auto_fit_enabled = True
        self.scans: list[Scan] = []
        self.current_scan_idx: Optional[int] = None
        self._suppress_selection_handlers = False
        self.edit_point_key: Optional[tuple[int, int, int]] = None
        self._editor_baseline_lat = ""
        self._editor_baseline_lon = ""

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
        self.refresh_editors()
        self.update_window_title()
        if settings_error:
            self.show_warning(t(self.lang, "settings_title"), settings_error)

    def tr(self, key: str, **kwargs) -> str:
        return t(self.lang, key, **kwargs)

    def prepare_dialog(self, dialog) -> None:
        dialog.setWindowModality(Qt.ApplicationModal)
        app = QApplication.instance()
        if app is not None and app.applicationState() == Qt.ApplicationState.ApplicationActive:
            dialog.setWindowFlag(Qt.WindowStaysOnTopHint, True)

    def show_warning(self, title: str, text: str) -> None:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle(title)
        msg.setText(text)
        self.prepare_dialog(msg)
        msg.exec()

    def show_information(self, title: str, text: str) -> None:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle(title)
        msg.setText(text)
        self.prepare_dialog(msg)
        msg.exec()

    def show_critical(self, title: str, text: str) -> None:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle(title)
        msg.setText(text)
        self.prepare_dialog(msg)
        msg.exec()

    def ask_yes_no(self, title: str, text: str, default_button=QMessageBox.No) -> bool:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(default_button)
        self.prepare_dialog(msg)
        return msg.exec() == QMessageBox.Yes

    def get_open_file_name(self, title: str, start_dir: str, file_filter: str) -> tuple[str, str]:
        dlg = QFileDialog(self, title, start_dir, file_filter)
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setAcceptMode(QFileDialog.AcceptOpen)
        dlg.setNameFilter(file_filter)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self.prepare_dialog(dlg)
        if dlg.exec() == QDialog.Accepted and dlg.selectedFiles():
            return dlg.selectedFiles()[0], dlg.selectedNameFilter()
        return "", file_filter

    def get_open_file_names(self, title: str, start_dir: str, file_filter: str) -> tuple[list[str], str]:
        dlg = QFileDialog(self, title, start_dir, file_filter)
        dlg.setFileMode(QFileDialog.ExistingFiles)
        dlg.setAcceptMode(QFileDialog.AcceptOpen)
        dlg.setNameFilter(file_filter)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self.prepare_dialog(dlg)
        if dlg.exec() == QDialog.Accepted:
            return dlg.selectedFiles(), dlg.selectedNameFilter()
        return [], file_filter

    def get_save_file_name(self, title: str, start_path: str, file_filter: str) -> tuple[str, str]:
        dlg = QFileDialog(self, title, start_path, file_filter)
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilter(file_filter)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self.prepare_dialog(dlg)
        if dlg.exec() == QDialog.Accepted and dlg.selectedFiles():
            return dlg.selectedFiles()[0], dlg.selectedNameFilter()
        return "", file_filter

    def get_existing_directory(self, title: str, start_dir: str) -> str:
        dlg = QFileDialog(self, title, start_dir)
        dlg.setFileMode(QFileDialog.Directory)
        dlg.setOption(QFileDialog.ShowDirsOnly, True)
        dlg.setAcceptMode(QFileDialog.AcceptOpen)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self.prepare_dialog(dlg)
        if dlg.exec() == QDialog.Accepted and dlg.selectedFiles():
            return dlg.selectedFiles()[0]
        return ""

    @staticmethod
    def normalize_scan_path(path_text: str) -> str:
        return str(Path(path_text).expanduser().resolve(strict=False))

    def find_scan_index_by_image_path(self, image_path: str) -> Optional[int]:
        key = self.normalize_scan_path(image_path)
        for idx, scan in enumerate(self.scans):
            if self.normalize_scan_path(scan.image_path) == key:
                return idx
        return None

    @staticmethod
    def tree_key_from_data(data: object) -> Optional[tuple]:
        if not isinstance(data, dict):
            return None
        return (
            data.get("type"),
            data.get("scan_idx"),
            data.get("sheet_idx"),
            data.get("point_idx"),
        )

    def find_tree_item_by_key(self, key: tuple) -> Optional[QTreeWidgetItem]:
        def walk(item: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
            if self.tree_key_from_data(item.data(0, Qt.UserRole)) == key:
                return item
            for i in range(item.childCount()):
                found = walk(item.child(i))
                if found is not None:
                    return found
            return None

        for i in range(self.sheet_tree.topLevelItemCount()):
            found = walk(self.sheet_tree.topLevelItem(i))
            if found is not None:
                return found
        return None

    def current_selected_point_key(self) -> Optional[tuple[int, int, int]]:
        if self.current_scan_idx is None:
            return None
        if self.canvas.selected_sheet_idx is None or self.canvas.selected_point_idx is None:
            return None
        return (self.current_scan_idx, self.canvas.selected_sheet_idx, self.canvas.selected_point_idx)

    def get_point_by_key(self, key: tuple[int, int, int]) -> Optional[CalibrationPoint]:
        scan_idx, sheet_idx, point_idx = key
        if not (0 <= scan_idx < len(self.scans)):
            return None
        sheets = self.scans[scan_idx].sheets
        if not (0 <= sheet_idx < len(sheets)):
            return None
        points = sheets[sheet_idx].points
        if not (0 <= point_idx < len(points)):
            return None
        return points[point_idx]

    def point_editor_dirty(self) -> bool:
        if self.edit_point_key is None:
            return False
        return (
            self.lat_edit.text().strip() != self._editor_baseline_lat
            or self.lon_edit.text().strip() != self._editor_baseline_lon
        )

    def confirm_leave_point_editor(self, new_point_key: Optional[tuple[int, int, int]]) -> bool:
        if self.edit_point_key is None or self.edit_point_key == new_point_key:
            return True
        if not self.point_editor_dirty():
            return True

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle(self.tr("unsaved_point_title"))
        msg.setText(self.tr("unsaved_point_msg"))
        save_btn = msg.addButton(QMessageBox.Save)
        discard_btn = msg.addButton(QMessageBox.Discard)
        cancel_btn = msg.addButton(QMessageBox.Cancel)
        msg.setDefaultButton(save_btn)
        self.prepare_dialog(msg)
        msg.exec()
        clicked = msg.clickedButton()

        if clicked == save_btn:
            point = self.get_point_by_key(self.edit_point_key)
            if point is None:
                return True
            return self.save_editor_values_to_point(point, show_errors=True, show_grid_hint=False)
        if clicked == discard_btn:
            return True
        if clicked == cancel_btn:
            return False
        return False

    @property
    def current_scan(self) -> Optional[Scan]:
        if self.current_scan_idx is None:
            return None
        if self.current_scan_idx < 0 or self.current_scan_idx >= len(self.scans):
            return None
        return self.scans[self.current_scan_idx]

    def set_current_scan(self, idx: Optional[int]) -> None:
        if idx is None or idx < 0 or idx >= len(self.scans):
            self.current_scan_idx = None
            self.canvas.pixmap = QPixmap()
            self.canvas.image_path = None
            self.canvas.sheets = []
            self.canvas.selected_sheet_idx = None
            self.canvas.selected_point_idx = None
            self.canvas.updateGeometry()
            self.canvas.update()
            self.refresh_sheet_list()
            return

        scan = self.scans[idx]
        self.current_scan_idx = idx
        scan_path = Path(scan.image_path) if scan.image_path else None
        if scan_path and scan_path.exists():
            if not self.canvas.set_image(scan_path):
                self.show_warning(self.tr("warning_title"), self.tr("warning_missing_scan_image", path=str(scan_path)))
                self.canvas.pixmap = QPixmap()
                self.canvas.image_path = scan_path
                self.canvas.updateGeometry()
        elif scan_path is not None:
            missing = str(scan_path) if scan_path else "-"
            self.show_warning(self.tr("warning_title"), self.tr("warning_missing_scan_image", path=missing))
            self.canvas.pixmap = QPixmap()
            self.canvas.image_path = scan_path
            self.canvas.updateGeometry()
        else:
            self.canvas.pixmap = QPixmap()
            self.canvas.image_path = None
            self.canvas.updateGeometry()

        self.canvas.sheets = scan.sheets
        self.canvas.selected_sheet_idx = 0 if self.canvas.sheets else None
        self.canvas.selected_point_idx = None
        self.canvas.new_sheet_temp_points = []
        self.canvas.mode = self.canvas.MODE_SELECT
        self.canvas.dragging = False
        self.canvas.update()
        self.refresh_sheet_list()
        self.auto_fit_enabled = True
        QTimer.singleShot(0, self.fit_to_window)

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
        self.btn_add_outline = QPushButton(self.tr("btn_add_outline_point"))
        self.btn_select_mode = QPushButton(self.tr("btn_select_mode"))
        self.btn_delete_sheet = QPushButton(self.tr("btn_delete_sheet"))
        self.btn_delete_point = QPushButton(self.tr("btn_delete_point"))
        self.btn_use_as_corner = QPushButton(self.tr("btn_use_as_corner"))
        self.btn_zoom_in = QPushButton(self.tr("btn_zoom_in"))
        self.btn_zoom_out = QPushButton(self.tr("btn_zoom_out"))
        self.btn_zoom_reset = QPushButton(self.tr("btn_zoom_reset"))

        self.btn_new_sheet.clicked.connect(self.canvas.start_new_sheet)
        self.btn_close_sheet.clicked.connect(self.close_sheet_clicked)
        self.btn_cancel_sheet.clicked.connect(self.canvas.cancel_new_sheet)
        self.btn_add_cal.clicked.connect(self.canvas.set_mode_add_cal_point)
        self.btn_add_outline.clicked.connect(self.canvas.set_mode_add_outline_point)
        self.btn_select_mode.clicked.connect(self.canvas.set_mode_select)
        self.btn_delete_sheet.clicked.connect(self.delete_sheet)
        self.btn_delete_point.clicked.connect(self.delete_selected_point)
        self.btn_use_as_corner.clicked.connect(self.use_selected_point_as_corner)
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        self.btn_zoom_reset.clicked.connect(self.zoom_reset)

        layout.addWidget(self.btn_new_sheet)
        layout.addWidget(self.btn_close_sheet)
        layout.addWidget(self.btn_cancel_sheet)
        layout.addWidget(self.btn_add_cal)
        layout.addWidget(self.btn_add_outline)
        layout.addWidget(self.btn_select_mode)
        layout.addWidget(self.btn_delete_sheet)
        layout.addWidget(self.btn_delete_point)
        layout.addWidget(self.btn_use_as_corner)
        layout.addWidget(self.btn_zoom_in)
        layout.addWidget(self.btn_zoom_out)
        layout.addWidget(self.btn_zoom_reset)

        self.lbl_sheets = QLabel(self.tr("panel_sheets"))
        layout.addWidget(self.lbl_sheets)
        self.sheet_tree = QTreeWidget()
        self.sheet_tree.setHeaderHidden(True)
        self.sheet_tree.currentItemChanged.connect(self.on_tree_selection_changed)
        layout.addWidget(self.sheet_tree)

        self.lbl_meta_title = QLabel(self.tr("panel_sheet_meta"))
        layout.addWidget(self.lbl_meta_title)
        self.meta_form = QFormLayout()
        self.name_edit = QLineEdit()
        self.scale_edit = QLineEdit()
        self.name_edit.textChanged.connect(self.on_sheet_meta_changed)
        self.scale_edit.textChanged.connect(self.on_sheet_meta_changed)
        self.meta_form.addRow(self.tr("field_name"), self.name_edit)
        self.meta_form.addRow(self.tr("field_scale"), self.scale_edit)
        layout.addLayout(self.meta_form)

        self.lbl_point_title = QLabel(self.tr("panel_point_geo"))
        layout.addWidget(self.lbl_point_title)
        self.point_form = QFormLayout()
        self.lat_edit = QLineEdit()
        self.lon_edit = QLineEdit()
        self.point_form.addRow(self.tr("field_lat"), self.lat_edit)
        self.point_form.addRow(self.tr("field_lon"), self.lon_edit)
        layout.addLayout(self.point_form)
        self.btn_save_point = QPushButton(self.tr("btn_save_point"))
        self.btn_save_point.clicked.connect(self.save_selected_point_geo)
        layout.addWidget(self.btn_save_point)

        layout.addStretch(1)
        return panel

    def build_menu(self) -> None:
        self.menu_file = self.menuBar().addMenu(self.tr("menu_file"))
        self.action_open_img = QAction(self.tr("menu_open_scan"), self)
        self.action_open_img.triggered.connect(self.open_image)
        self.menu_file.addAction(self.action_open_img)

        self.action_import_map = QAction(self.tr("menu_import_map"), self)
        self.action_import_map.triggered.connect(self.import_map_files)
        self.menu_file.addAction(self.action_import_map)

        self.action_export_kap = QAction(self.tr("menu_export_kap"), self)
        self.action_export_kap.triggered.connect(self.export_kap_current_scan)
        self.menu_file.addAction(self.action_export_kap)

        self.action_export_all_kap = QAction(self.tr("menu_export_all_kap"), self)
        self.action_export_all_kap.triggered.connect(self.export_kap_all_scans)
        self.menu_file.addAction(self.action_export_all_kap)

        self.action_save_proj = QAction(self.tr("menu_save"), self)
        self.action_save_proj.triggered.connect(self.save_project)
        self.menu_file.addAction(self.action_save_proj)

        self.action_save_as_proj = QAction(self.tr("menu_save_as"), self)
        self.action_save_as_proj.triggered.connect(self.save_project_as)
        self.menu_file.addAction(self.action_save_as_proj)

        self.action_load_proj = QAction(self.tr("menu_load_project"), self)
        self.action_load_proj.triggered.connect(self.load_project)
        self.menu_file.addAction(self.action_load_proj)

        self.menu_settings = self.menuBar().addMenu(self.tr("menu_settings"))
        self.action_edit_settings = QAction(self.tr("menu_edit_settings"), self)
        self.action_edit_settings.triggered.connect(self.edit_settings)
        self.menu_settings.addAction(self.action_edit_settings)

        self.menu_help = self.menuBar().addMenu(self.tr("menu_help"))
        self.action_show_readme = QAction(self.tr("menu_readme"), self)
        self.action_show_readme.triggered.connect(self.show_readme_help)
        self.menu_help.addAction(self.action_show_readme)

    def apply_language_to_ui(self) -> None:
        self.base_title = t(self.lang, "title")
        self.update_window_title()
        self.canvas.set_language(self.lang)

        self.btn_new_sheet.setText(self.tr("btn_new_sheet"))
        self.btn_close_sheet.setText(self.tr("btn_close_sheet"))
        self.btn_cancel_sheet.setText(self.tr("btn_cancel_drawing"))
        self.btn_add_cal.setText(self.tr("btn_add_cal_point"))
        self.btn_add_outline.setText(self.tr("btn_add_outline_point"))
        self.btn_select_mode.setText(self.tr("btn_select_mode"))
        self.btn_delete_sheet.setText(self.tr("btn_delete_sheet"))
        self.btn_delete_point.setText(self.tr("btn_delete_point"))
        self.btn_use_as_corner.setText(self.tr("btn_use_as_corner"))
        self.btn_zoom_in.setText(self.tr("btn_zoom_in"))
        self.btn_zoom_out.setText(self.tr("btn_zoom_out"))
        self.btn_zoom_reset.setText(self.tr("btn_zoom_reset"))
        self.btn_save_point.setText(self.tr("btn_save_point"))

        self.lbl_sheets.setText(self.tr("panel_sheets"))
        self.lbl_meta_title.setText(self.tr("panel_sheet_meta"))
        self.lbl_point_title.setText(self.tr("panel_point_geo"))

        name_label = self.meta_form.labelForField(self.name_edit)
        if name_label is not None:
            name_label.setText(self.tr("field_name"))
        scale_label = self.meta_form.labelForField(self.scale_edit)
        if scale_label is not None:
            scale_label.setText(self.tr("field_scale"))
        lat_label = self.point_form.labelForField(self.lat_edit)
        if lat_label is not None:
            lat_label.setText(self.tr("field_lat"))
        lon_label = self.point_form.labelForField(self.lon_edit)
        if lon_label is not None:
            lon_label.setText(self.tr("field_lon"))

        self.menu_file.setTitle(self.tr("menu_file"))
        self.action_open_img.setText(self.tr("menu_open_scan"))
        self.action_import_map.setText(self.tr("menu_import_map"))
        self.action_export_kap.setText(self.tr("menu_export_kap"))
        self.action_export_all_kap.setText(self.tr("menu_export_all_kap"))
        self.action_save_proj.setText(self.tr("menu_save"))
        self.action_save_as_proj.setText(self.tr("menu_save_as"))
        self.action_load_proj.setText(self.tr("menu_load_project"))
        self.menu_settings.setTitle(self.tr("menu_settings"))
        self.action_edit_settings.setText(self.tr("menu_edit_settings"))
        self.menu_help.setTitle(self.tr("menu_help"))
        self.action_show_readme.setText(self.tr("menu_readme"))

        point_dirty = self.point_editor_dirty()
        dirty_lat = self.lat_edit.text()
        dirty_lon = self.lon_edit.text()
        dirty_lat_style = self.lat_edit.styleSheet()
        dirty_lon_style = self.lon_edit.styleSheet()
        dirty_key = self.edit_point_key

        self.refresh_sheet_list()

        if point_dirty and dirty_key is not None and self.current_selected_point_key() == dirty_key:
            self.lat_edit.setText(dirty_lat)
            self.lon_edit.setText(dirty_lon)
            self.lat_edit.setStyleSheet(dirty_lat_style)
            self.lon_edit.setStyleSheet(dirty_lon_style)

        if self.canvas.cursor_x is not None and self.canvas.cursor_y is not None:
            geo = self.canvas.geo_for_cursor(self.canvas.cursor_x, self.canvas.cursor_y)
            self.on_cursor_moved(self.canvas.cursor_x, self.canvas.cursor_y, geo)
        else:
            self.statusBar().showMessage(self.tr("ready"))

    @staticmethod
    def _settings_path_for_write(existing_path: Optional[Path]) -> Path:
        if existing_path is not None:
            return existing_path
        return Path.cwd() / ".pymapcal"

    def edit_settings(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle(self.tr("settings_dialog_title"))
        dlg.resize(700, 700)
        self.prepare_dialog(dlg)
        layout = QVBoxLayout(dlg)
        form = QFormLayout()

        lang_combo = QComboBox(dlg)
        lang_combo.addItem(self.tr("settings_lang_pl"), "pl")
        lang_combo.addItem(self.tr("settings_lang_en"), "en")
        lang_combo.setCurrentIndex(0 if self.lang == "pl" else 1)
        form.addRow(self.tr("settings_field_language"), lang_combo)
        imgkap_path_edit = QLineEdit(self.imgkap_path, dlg)
        form.addRow(self.tr("settings_field_imgkap_path"), imgkap_path_edit)
        imgkap_work_dir_edit = QLineEdit(self.imgkap_work_dir, dlg)
        form.addRow(self.tr("settings_field_imgkap_work_dir"), imgkap_work_dir_edit)
        sounding_datum_edit = QLineEdit(self.kap_sounding_datum, dlg)
        form.addRow(self.tr("settings_field_sounding_datum"), sounding_datum_edit)

        def make_spin(value: int, min_v: int, max_v: int) -> QSpinBox:
            s = QSpinBox(dlg)
            s.setRange(min_v, max_v)
            s.setValue(value)
            return s

        outline_width = make_spin(self.display_settings.outline_width, 1, 12)
        outline_sel_width = make_spin(self.display_settings.outline_selected_width, 1, 16)
        draft_width = make_spin(self.display_settings.draft_outline_width, 1, 16)
        arm_corner = make_spin(self.display_settings.crosshair_arm_corner, 4, 80)
        arm_cal = make_spin(self.display_settings.crosshair_arm_cal, 4, 80)
        ring_corner = make_spin(self.display_settings.crosshair_ring_corner, 1, 40)
        ring_cal = make_spin(self.display_settings.crosshair_ring_cal, 1, 40)
        sel_arm = make_spin(self.display_settings.crosshair_selected_arm_bonus, 0, 40)
        sel_ring = make_spin(self.display_settings.crosshair_selected_ring_bonus, 0, 20)
        guide_width = make_spin(self.display_settings.cursor_guide_width, 1, 16)
        guide_alpha = make_spin(self.display_settings.cursor_guide_alpha, 0, 255)
        guide_dash = make_spin(self.display_settings.cursor_guide_dash, 1, 80)
        guide_gap = make_spin(self.display_settings.cursor_guide_gap, 1, 80)
        guide_color = QLineEdit(self.display_settings.cursor_guide_color, dlg)

        form.addRow(self.tr("settings_field_outline_width"), outline_width)
        form.addRow(self.tr("settings_field_outline_selected_width"), outline_sel_width)
        form.addRow(self.tr("settings_field_draft_outline_width"), draft_width)
        form.addRow(self.tr("settings_field_crosshair_arm_corner"), arm_corner)
        form.addRow(self.tr("settings_field_crosshair_arm_cal"), arm_cal)
        form.addRow(self.tr("settings_field_crosshair_ring_corner"), ring_corner)
        form.addRow(self.tr("settings_field_crosshair_ring_cal"), ring_cal)
        form.addRow(self.tr("settings_field_crosshair_selected_arm_bonus"), sel_arm)
        form.addRow(self.tr("settings_field_crosshair_selected_ring_bonus"), sel_ring)
        form.addRow(self.tr("settings_field_cursor_guide_width"), guide_width)
        form.addRow(self.tr("settings_field_cursor_guide_alpha"), guide_alpha)
        form.addRow(self.tr("settings_field_cursor_guide_dash"), guide_dash)
        form.addRow(self.tr("settings_field_cursor_guide_gap"), guide_gap)
        form.addRow(self.tr("settings_field_cursor_guide_color"), guide_color)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel, parent=dlg)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec() != QDialog.Accepted:
            return

        color_text = guide_color.text().strip()
        if not QColor(color_text).isValid():
            self.show_warning(self.tr("error_title"), self.tr("settings_invalid_color"))
            return

        new_lang = lang_combo.currentData()
        new_imgkap_path = imgkap_path_edit.text().strip() or "imgkap"
        new_imgkap_work_dir = imgkap_work_dir_edit.text().strip()
        new_sounding_datum = sounding_datum_edit.text().strip() or "UNKNOWN"
        new_settings = DisplaySettings(
            outline_width=outline_width.value(),
            outline_selected_width=outline_sel_width.value(),
            draft_outline_width=draft_width.value(),
            crosshair_arm_corner=arm_corner.value(),
            crosshair_arm_cal=arm_cal.value(),
            crosshair_ring_corner=ring_corner.value(),
            crosshair_ring_cal=ring_cal.value(),
            crosshair_selected_arm_bonus=sel_arm.value(),
            crosshair_selected_ring_bonus=sel_ring.value(),
            cursor_guide_width=guide_width.value(),
            cursor_guide_alpha=guide_alpha.value(),
            cursor_guide_dash=guide_dash.value(),
            cursor_guide_gap=guide_gap.value(),
            cursor_guide_color=color_text,
        )

        settings_path = self._settings_path_for_write(self.display_settings_path)
        payload = {
            "language": new_lang,
            "imgkap_path": new_imgkap_path,
            "imgkap_work_dir": new_imgkap_work_dir,
            "kap_sounding_datum": new_sounding_datum,
            "display": new_settings.to_dict(),
        }
        try:
            settings_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception as exc:
            self.show_critical(self.tr("error_title"), self.tr("settings_save_error", error=exc))
            return

        self.display_settings = new_settings
        self.display_settings_path = settings_path
        self.imgkap_path = new_imgkap_path
        self.imgkap_work_dir = new_imgkap_work_dir
        self.kap_sounding_datum = new_sounding_datum
        self.canvas.apply_display_settings(new_settings)

        lang_changed = new_lang != self.lang
        self.lang = new_lang
        if lang_changed:
            self.apply_language_to_ui()
            self.statusBar().showMessage(self.tr("settings_save_ok", path=str(settings_path)))
        else:
            self.statusBar().showMessage(self.tr("settings_save_ok", path=str(settings_path)))

    def show_readme_help(self) -> None:
        base = Path(__file__).resolve().parent
        if self.lang == "en":
            preferred = base / "README.en.md"
        else:
            preferred = base / "README.pl.md"
        readme_path = preferred if preferred.exists() else (base / "README.md")
        if not readme_path.exists():
            self.show_warning(self.tr("help_title"), self.tr("help_missing_readme"))
            return
        try:
            content = readme_path.read_text(encoding="utf-8")
        except Exception as exc:
            self.show_critical(self.tr("help_title"), self.tr("help_read_error", error=exc))
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(self.tr("help_dialog_title"))
        dlg.resize(900, 700)
        self.prepare_dialog(dlg)
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
            self.show_warning(self.tr("error_title"), self.tr("error_sheet_min_points"))

    def delete_sheet(self) -> None:
        idx = self.canvas.selected_sheet_idx
        sheet = self.canvas.selected_sheet
        if idx is None or sheet is None:
            return
        if not self.ask_yes_no(
            self.tr("confirm_title"),
            self.tr("confirm_delete_sheet_named", name=sheet.name),
            default_button=QMessageBox.No,
        ):
            return
        del self.canvas.sheets[idx]
        if self.canvas.sheets:
            self.canvas.selected_sheet_idx = max(0, idx - 1)
        else:
            self.canvas.selected_sheet_idx = None
        self.canvas.selected_point_idx = None
        self.refresh_sheet_list()
        self.canvas.update()

    def delete_selected_point(self) -> None:
        sheet = self.canvas.selected_sheet
        idx = self.canvas.selected_point_idx
        point = self.canvas.selected_point
        if sheet is None or idx is None or point is None:
            return
        if idx < 0 or idx >= len(sheet.points):
            return
        kind = self.tr("point_kind_outline") if point.is_corner else self.tr("point_kind_calibration")
        kind_index = 0
        for i, p in enumerate(sheet.points):
            if p.is_corner == point.is_corner:
                kind_index += 1
            if i == idx:
                break

        if not self.ask_yes_no(
            self.tr("confirm_title"),
            self.tr("confirm_delete_point_named", kind=kind, index=kind_index),
            default_button=QMessageBox.No,
        ):
            return
        del sheet.points[idx]
        self.canvas.selected_point_idx = None
        self.refresh_sheet_list()
        self.canvas.update()

    def use_selected_point_as_corner(self) -> None:
        point = self.canvas.selected_point
        if point is None:
            return
        if point.is_corner:
            return
        if not self.ask_yes_no(
            self.tr("link_point_title"),
            self.tr("link_point_question"),
            default_button=QMessageBox.Yes,
        ):
            return
        point.is_corner = True
        self.refresh_sheet_list()
        self.canvas.update()

    def refresh_sheet_list(self) -> None:
        vbar = self.sheet_tree.verticalScrollBar()
        hbar = self.sheet_tree.horizontalScrollBar()
        vpos = vbar.value()
        hpos = hbar.value()

        expanded_keys: set[tuple] = set()
        selected_key: Optional[tuple] = None

        key_from_data = self.tree_key_from_data

        def collect_state(item: QTreeWidgetItem) -> None:
            nonlocal selected_key
            data = item.data(0, Qt.UserRole)
            key = key_from_data(data)
            if key is not None and item.isExpanded():
                expanded_keys.add(key)
            if item is self.sheet_tree.currentItem() and key is not None:
                selected_key = key
            for i in range(item.childCount()):
                collect_state(item.child(i))

        for i in range(self.sheet_tree.topLevelItemCount()):
            collect_state(self.sheet_tree.topLevelItem(i))

        self.sheet_tree.blockSignals(True)
        self.sheet_tree.clear()

        selected_item: Optional[QTreeWidgetItem] = None
        key_to_item: dict[tuple, QTreeWidgetItem] = {}

        def register_item(item: QTreeWidgetItem, data: dict) -> None:
            key = key_from_data(data)
            if key is not None:
                key_to_item[key] = item

        for scan_idx, scan in enumerate(self.scans):
            scan_name = scan.name if scan.name else Path(scan.image_path).stem
            scan_item = QTreeWidgetItem([self.tr("tree_scan_item", name=scan_name)])
            scan_data = {"type": "scan", "scan_idx": scan_idx}
            scan_item.setData(0, Qt.UserRole, scan_data)
            register_item(scan_item, scan_data)
            self.sheet_tree.addTopLevelItem(scan_item)

            for sheet_idx, sheet in enumerate(scan.sheets):
                sheet_item = QTreeWidgetItem([self.tr("tree_sheet_item", name=sheet.name, count=len(sheet.points))])
                sheet_data = {"type": "sheet", "scan_idx": scan_idx, "sheet_idx": sheet_idx}
                sheet_item.setData(0, Qt.UserRole, sheet_data)
                register_item(sheet_item, sheet_data)
                scan_item.addChild(sheet_item)

                corners_item = QTreeWidgetItem([self.tr("tree_crop_outline")])
                corners_data = {"type": "corner_group", "scan_idx": scan_idx, "sheet_idx": sheet_idx}
                corners_item.setData(0, Qt.UserRole, corners_data)
                register_item(corners_item, corners_data)
                sheet_item.addChild(corners_item)

                cal_item = QTreeWidgetItem([self.tr("tree_cal_points")])
                cal_data = {"type": "cal_group", "scan_idx": scan_idx, "sheet_idx": sheet_idx}
                cal_item.setData(0, Qt.UserRole, cal_data)
                register_item(cal_item, cal_data)
                sheet_item.addChild(cal_item)

                corner_no = 1
                cal_no = 1
                for point_idx, point in enumerate(sheet.points):
                    if point.is_corner:
                        label = self.tr("tree_corner_item", n=corner_no, x=point.x, y=point.y)
                        parent = corners_item
                        corner_no += 1
                    else:
                        label = self.tr("tree_point_item", n=cal_no, x=point.x, y=point.y)
                        parent = cal_item
                        cal_no += 1
                    point_item = QTreeWidgetItem([label])
                    point_data = {
                        "type": "point",
                        "scan_idx": scan_idx,
                        "sheet_idx": sheet_idx,
                        "point_idx": point_idx,
                    }
                    point_item.setData(0, Qt.UserRole, point_data)
                    register_item(point_item, point_data)
                    parent.addChild(point_item)

                    if (
                        scan_idx == self.current_scan_idx
                        and sheet_idx == self.canvas.selected_sheet_idx
                        and point_idx == self.canvas.selected_point_idx
                    ):
                        selected_item = point_item

                sheet_key = key_from_data(sheet_data)
                corners_key = key_from_data(corners_data)
                cal_key = key_from_data(cal_data)
                sheet_item.setExpanded(sheet_key in expanded_keys if sheet_key is not None else True)
                corners_item.setExpanded(corners_key in expanded_keys if corners_key is not None else True)
                cal_item.setExpanded(cal_key in expanded_keys if cal_key is not None else True)

                if (
                    scan_idx == self.current_scan_idx
                    and sheet_idx == self.canvas.selected_sheet_idx
                    and self.canvas.selected_point_idx is None
                ):
                    selected_item = sheet_item

            scan_key = key_from_data(scan_data)
            scan_item.setExpanded(scan_key in expanded_keys if scan_key is not None else True)
            if scan_idx == self.current_scan_idx and self.canvas.selected_sheet_idx is None:
                selected_item = scan_item

        if selected_key is not None and selected_key in key_to_item:
            selected_item = key_to_item[selected_key]
        if selected_item is not None:
            self.sheet_tree.setCurrentItem(selected_item)
        self.sheet_tree.blockSignals(False)
        vbar.setValue(vpos)
        hbar.setValue(hpos)
        self.refresh_editors()

    def on_tree_selection_changed(self, current: Optional[QTreeWidgetItem], _previous: Optional[QTreeWidgetItem]) -> None:
        if self._suppress_selection_handlers:
            return
        if current is None:
            return
        data = current.data(0, Qt.UserRole)
        if not isinstance(data, dict):
            return

        scan_idx = data.get("scan_idx")
        if not isinstance(scan_idx, int) or scan_idx < 0 or scan_idx >= len(self.scans):
            return

        new_point_key: Optional[tuple[int, int, int]] = None
        if (
            data.get("type") == "point"
            and isinstance(data.get("sheet_idx"), int)
            and isinstance(data.get("point_idx"), int)
        ):
            new_point_key = (scan_idx, data["sheet_idx"], data["point_idx"])

        if not self.confirm_leave_point_editor(new_point_key):
            if self.edit_point_key is not None:
                old_key = ("point", self.edit_point_key[0], self.edit_point_key[1], self.edit_point_key[2])
                old_item = self.find_tree_item_by_key(old_key)
                if old_item is not None:
                    self._suppress_selection_handlers = True
                    self.sheet_tree.setCurrentItem(old_item)
                    self._suppress_selection_handlers = False
            return

        if scan_idx != self.current_scan_idx:
            self.set_current_scan(scan_idx)

        item_type = data.get("type")
        sheet_idx = data.get("sheet_idx")
        if item_type == "scan":
            self.canvas.selected_sheet_idx = None
            self.canvas.selected_point_idx = None
        elif isinstance(sheet_idx, int) and 0 <= sheet_idx < len(self.canvas.sheets):
            self.canvas.selected_sheet_idx = sheet_idx
            if item_type == "point":
                point_idx = data.get("point_idx")
                if isinstance(point_idx, int):
                    self.canvas.selected_point_idx = point_idx
                else:
                    self.canvas.selected_point_idx = None
            else:
                self.canvas.selected_point_idx = None
        else:
            return

        self.canvas.selectionChanged.emit(self.canvas.selected_sheet_idx, self.canvas.selected_point_idx)
        self.canvas.update()

    def on_canvas_selection_changed(self, _sheet_idx, _point_idx) -> None:
        if self._suppress_selection_handlers:
            return
        new_point_key = self.current_selected_point_key()
        if not self.confirm_leave_point_editor(new_point_key):
            if self.edit_point_key is not None:
                scan_idx, sheet_idx, point_idx = self.edit_point_key
                self._suppress_selection_handlers = True
                if scan_idx != self.current_scan_idx:
                    self.set_current_scan(scan_idx)
                self.canvas.selected_sheet_idx = sheet_idx
                self.canvas.selected_point_idx = point_idx
                self.canvas.update()
                self.refresh_sheet_list()
                self.refresh_editors()
                self._suppress_selection_handlers = False
            return
        self.refresh_sheet_list()
        self.refresh_editors()

    def refresh_editors(self) -> None:
        sheet = self.canvas.selected_sheet
        point = self.canvas.selected_point
        has_scan = self.current_scan is not None
        has_sheet = sheet is not None
        has_point = point is not None
        point_is_corner = bool(point.is_corner) if point is not None else False

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
        self.lon_edit.setStyleSheet("")
        self.lat_edit.setStyleSheet("")

        # Context-sensitive controls.
        self.name_edit.setEnabled(has_sheet)
        self.scale_edit.setEnabled(has_sheet)
        self.lon_edit.setEnabled(has_point)
        self.lat_edit.setEnabled(has_point)
        self.btn_save_point.setEnabled(has_point)

        self.btn_delete_sheet.setEnabled(has_sheet)
        self.btn_delete_point.setEnabled(has_point)
        self.btn_use_as_corner.setEnabled(has_point and (not point_is_corner))

        self.btn_new_sheet.setEnabled(has_scan)
        self.btn_add_cal.setEnabled(has_scan and has_sheet)
        self.btn_add_outline.setEnabled(has_scan and has_sheet)
        self.btn_select_mode.setEnabled(has_scan)
        self.btn_zoom_in.setEnabled(has_scan)
        self.btn_zoom_out.setEnabled(has_scan)
        self.btn_zoom_reset.setEnabled(has_scan)

        self.edit_point_key = self.current_selected_point_key()
        self._editor_baseline_lat = self.lat_edit.text().strip()
        self._editor_baseline_lon = self.lon_edit.text().strip()

    def flush_current_sheet_meta(self) -> None:
        sheet = self.canvas.selected_sheet
        if sheet is None:
            return
        sheet.name = self.name_edit.text().strip() or self.tr("sheet_default")
        sheet.scale = self.scale_edit.text().strip()

    def on_sheet_meta_changed(self, *_args) -> None:
        self.flush_current_sheet_meta()
        self.refresh_sheet_list()
        self.canvas.update()

    def save_editor_values_to_point(
        self,
        point: CalibrationPoint,
        show_errors: bool = True,
        show_grid_hint: bool = True,
    ) -> bool:
        lon_text = self.lon_edit.text().strip()
        lat_text = self.lat_edit.text().strip()
        lon = parse_geo_coordinate(lon_text, "lon")
        lat = parse_geo_coordinate(lat_text, "lat")
        has_error = False

        self.lon_edit.setStyleSheet("")
        self.lat_edit.setStyleSheet("")

        if lon_text and lon is None:
            has_error = True
            self.lon_edit.setStyleSheet("border: 2px solid #ff4d6d;")
        if lat_text and lat is None:
            has_error = True
            self.lat_edit.setStyleSheet("border: 2px solid #ff4d6d;")

        if has_error:
            if show_errors:
                details: list[str] = []
                if lon_text and lon is None:
                    details.append(self.tr("error_lon_format_msg"))
                if lat_text and lat is None:
                    details.append(self.tr("error_lat_format_msg"))
                self.show_warning(
                    self.tr("error_point_save_title"),
                    self.tr("error_point_save_msg") + "\n\n" + "\n".join(details),
                )
            return False

        point.lon_text = lon_text
        point.lat_text = lat_text
        point.lon = lon
        point.lat = lat
        if show_grid_hint:
            known = [p for p in self.canvas.selected_sheet.points if p.lon is not None and p.lat is not None] if self.canvas.selected_sheet else []
            if len(known) < 2:
                self.statusBar().showMessage(self.tr("status_grid_hint"))
        return True

    def save_selected_point_geo(self) -> None:
        point = self.canvas.selected_point
        if point is None:
            self.show_warning(self.tr("error_title"), self.tr("error_point_not_selected"))
            return
        if not self.save_editor_values_to_point(point, show_errors=True, show_grid_hint=True):
            return
        self.refresh_editors()
        self.canvas.update()

    def open_image(self) -> None:
        path_str, _ = self.get_open_file_name(
            self.tr("dialog_pick_scan"),
            "",
            self.tr("dialog_image_filter"),
        )
        if not path_str:
            return
        path = Path(path_str)
        test_pix = QPixmap(str(path))
        if test_pix.isNull():
            self.show_critical(self.tr("error_title"), self.tr("error_open_image"))
            return
        new_scan = Scan(
            name=path.stem,
            image_path=str(path),
            sheets=[Sheet(name=f"{self.tr('sheet_default')} 1")],
        )
        self.scans.append(new_scan)
        self.set_current_scan(len(self.scans) - 1)
        self.auto_fit_enabled = True
        QTimer.singleShot(0, self.fit_to_window)
        self.project_path = None
        self.update_window_title()

    def import_map_files(self) -> None:
        file_list, _ = self.get_open_file_names(
            self.tr("dialog_import_map"),
            "",
            self.tr("dialog_map_filter"),
        )
        if not file_list:
            return

        ok = 0
        fail = 0
        last_scan_idx: Optional[int] = None
        for path_str in file_list:
            map_path = Path(path_str)
            entry = parse_ozi_map_file(map_path)
            if entry is None:
                fail += 1
                continue

            existing_idx = self.find_scan_index_by_image_path(entry.image_path)
            if existing_idx is None:
                self.scans.append(
                    Scan(
                        name=entry.scan_name,
                        image_path=entry.image_path,
                        sheets=[entry.sheet],
                    )
                )
                last_scan_idx = len(self.scans) - 1
            else:
                self.scans[existing_idx].sheets.append(entry.sheet)
                last_scan_idx = existing_idx
            ok += 1

        if last_scan_idx is not None:
            self.set_current_scan(last_scan_idx)
            self.project_path = None
            self.update_window_title()
        self.show_information(
            self.tr("import_summary_title"),
            self.tr("import_summary", ok=ok, fail=fail),
        )

    @staticmethod
    def unique_kap_path(out_dir: Path, used_names: set[str], stem: str) -> Path:
        base = sanitize_kap_stem(stem, fallback="sheet")
        candidate = base
        idx = 2
        while candidate.lower() in used_names or (out_dir / f"{candidate}.kap").exists():
            candidate = f"{base}_{idx}"
            idx += 1
        used_names.add(candidate.lower())
        return out_dir / f"{candidate}.kap"

    @staticmethod
    def point_geo_or_transform(
        point: CalibrationPoint,
        transform: Optional[GeoTransform],
    ) -> Optional[tuple[float, float]]:
        if point.lon is not None and point.lat is not None:
            return float(point.lon), float(point.lat)
        if transform is None:
            return None
        return transform.pixel_to_geo(point.x, point.y)

    def default_export_dir(self) -> Path:
        scan = self.current_scan
        if scan is not None and scan.image_path:
            p = Path(scan.image_path).expanduser()
            if p.parent.exists():
                return p.parent
        return Path.cwd()

    def collect_kap_jobs_for_scans(self, scans: list[Scan], out_dir: Path) -> tuple[list[KapExportJob], list[str]]:
        jobs: list[KapExportJob] = []
        details: list[str] = []
        used_stems: set[str] = set()

        for scan in scans:
            scan_name = scan.name.strip() or (Path(scan.image_path).stem if scan.image_path else "scan")
            if not scan.sheets:
                continue
            image_path = Path(scan.image_path).expanduser()
            if not image_path.exists():
                details.append(f"[{scan_name}] " + self.tr("export_image_missing", path=str(image_path)))
                continue
            px = QPixmap(str(image_path))
            if px.isNull():
                details.append(f"[{scan_name}] " + self.tr("export_image_open_error"))
                continue
            width = px.width()
            height = px.height()

            for idx, sheet in enumerate(scan.sheets, start=1):
                sheet_name = sheet.name.strip() or f"{self.tr('sheet_default')} {idx}"
                sheet_label = f"{scan_name}/{sheet_name}"
                scale_value = parse_scale_value(sheet.scale)
                if scale_value is None:
                    details.append(
                        self.tr(
                            "export_sheet_scale_invalid",
                            name=sheet_label,
                            scale=sheet.scale if sheet.scale else "-",
                        )
                    )
                    continue

                corners = [p for p in sheet.points if p.is_corner]
                if len(corners) < 3:
                    details.append(self.tr("export_sheet_corners_missing", name=sheet_label))
                    continue

                transform = build_geo_transform(sheet.points)
                polygon: list[KapPolygonPoint] = []
                missing_corner_geo = False
                for p in corners:
                    geo = self.point_geo_or_transform(p, transform)
                    if geo is None:
                        missing_corner_geo = True
                        break
                    polygon.append(
                        KapPolygonPoint(
                            pixel_x=p.x,
                            pixel_y=p.y,
                            lon=geo[0],
                            lat=geo[1],
                        )
                    )
                if missing_corner_geo:
                    details.append(self.tr("export_sheet_geo_missing", name=sheet_label))
                    continue

                refs: list[KapReference] = []
                for p in sheet.points:
                    if p.is_corner:
                        continue
                    geo = self.point_geo_or_transform(p, transform)
                    if geo is None:
                        continue
                    refs.append(
                        KapReference(
                            pixel_x=p.x,
                            pixel_y=p.y,
                            lon=geo[0],
                            lat=geo[1],
                        )
                    )

                out_path = self.unique_kap_path(out_dir, used_stems, f"{scan_name}_{sheet_name}")
                jobs.append(
                    KapExportJob(
                        sheet_name=sheet_label,
                        image_path=image_path,
                        output_path=out_path,
                        width=width,
                        height=height,
                        scale=scale_value,
                        polygon=polygon,
                        references=refs,
                    )
                )

        return jobs, details

    def run_kap_export_and_show_summary(self, jobs: list[KapExportJob], details: list[str], out_dir: Path) -> None:
        if not jobs:
            self.show_warning(
                self.tr("export_title"),
                self.tr("export_summary", ok=0, fail=len(details), out=str(out_dir))
                + ("\n\n" + "\n".join(details) if details else ""),
            )
            return

        pre_fail = len(details)
        debug_temp_dir = Path(self.imgkap_work_dir).expanduser() if self.imgkap_work_dir else None
        progress = QProgressDialog(self)
        progress.setWindowTitle(self.tr("export_progress_title"))
        progress.setLabelText(self.tr("export_progress_label", current=0, total=len(jobs), name="-"))
        progress.setCancelButtonText(self.tr("export_progress_cancel"))
        progress.setMinimum(0)
        progress.setMaximum(len(jobs))
        progress.setValue(0)
        self.prepare_dialog(progress)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.show()
        QApplication.processEvents()
        cancel_requested = False

        def on_cancel() -> None:
            nonlocal cancel_requested
            cancel_requested = True

        progress.canceled.connect(on_cancel)

        def on_progress(done: int, total: int, name: str) -> None:
            progress.setLabelText(self.tr("export_progress_label", current=done, total=total, name=name))
            progress.setValue(done)
            QApplication.processEvents()

        def is_cancel_requested() -> bool:
            QApplication.processEvents()
            return cancel_requested or progress.wasCanceled()

        try:
            run_results = run_kap_export_jobs(
                jobs=jobs,
                imgkap_path=self.imgkap_path,
                sounding_datum=self.kap_sounding_datum,
                temp_dir=debug_temp_dir,
                progress_cb=on_progress,
                cancel_requested_cb=is_cancel_requested,
            )
        finally:
            progress.close()
        ok = 0
        run_fail = 0
        cancelled = False
        imgkap_missing_reported = False
        for result in run_results:
            if result.success:
                ok += 1
                continue
            if result.error == "cancelled":
                cancelled = True
                break
            run_fail += 1
            if result.error == "imgkap_not_found":
                if not imgkap_missing_reported:
                    details.append(self.tr("export_sheet_imgkap_missing", path=self.imgkap_path))
                    imgkap_missing_reported = True
                continue
            details.append(self.tr("export_sheet_imgkap_failed", name=result.sheet_name))
            if result.stderr:
                details.append(f"  {result.stderr.splitlines()[0]}")
            elif result.stdout:
                details.append(f"  {result.stdout.splitlines()[0]}")
        if not cancelled and (cancel_requested or progress.wasCanceled()) and len(run_results) < len(jobs):
            cancelled = True

        total_fail = pre_fail + run_fail
        if cancelled:
            summary = self.tr("export_cancelled", done=len(run_results), total=len(jobs), out=str(out_dir))
        else:
            summary = self.tr("export_summary", ok=ok, fail=total_fail, out=str(out_dir))
        details_text = "\n".join(details[:40])
        if cancelled or total_fail > 0 or details:
            self.show_warning(
                self.tr("export_title"),
                summary + ("\n\n" + details_text if details_text else ""),
            )
        else:
            self.show_information(self.tr("export_title"), summary)

    def export_kap_current_scan(self) -> None:
        self.flush_current_sheet_meta()
        scan = self.current_scan
        if scan is None:
            self.show_warning(self.tr("export_title"), self.tr("export_no_scan"))
            return
        if not scan.sheets:
            self.show_warning(self.tr("export_title"), self.tr("export_no_sheets"))
            return

        out_dir_str = self.get_existing_directory(
            self.tr("dialog_export_kap"),
            str(self.default_export_dir()),
        )
        if not out_dir_str:
            return
        out_dir = Path(out_dir_str)
        jobs, details = self.collect_kap_jobs_for_scans([scan], out_dir)
        self.run_kap_export_and_show_summary(jobs, details, out_dir)

    def export_kap_all_scans(self) -> None:
        self.flush_current_sheet_meta()
        if not self.scans:
            self.show_warning(self.tr("export_title"), self.tr("export_no_scans"))
            return
        if not any(scan.sheets for scan in self.scans):
            self.show_warning(self.tr("export_title"), self.tr("export_no_sheets_all"))
            return

        out_dir_str = self.get_existing_directory(
            self.tr("dialog_export_kap"),
            str(self.default_export_dir()),
        )
        if not out_dir_str:
            return
        out_dir = Path(out_dir_str)
        jobs, details = self.collect_kap_jobs_for_scans(self.scans, out_dir)
        self.run_kap_export_and_show_summary(jobs, details, out_dir)

    def save_project(self) -> None:
        self.flush_current_sheet_meta()
        if self.project_path is not None:
            self.write_project_to_path(self.project_path)
            return

        self.save_project_as()

    def save_project_as(self) -> None:
        self.flush_current_sheet_meta()
        current_scan = self.current_scan
        if current_scan and current_scan.image_path:
            suggested_path = Path(current_scan.image_path).with_suffix(".json")
        elif self.project_path is not None:
            suggested_path = self.project_path
        else:
            suggested_path = Path.cwd() / self.tr("default_project_name")

        path_str, _ = self.get_save_file_name(
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
            "scans": [scan.to_json() for scan in self.scans],
        }
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.project_path = path
        self.update_window_title()

    def load_project(self) -> None:
        path_str, _ = self.get_open_file_name(
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
            self.show_critical(self.tr("error_title"), self.tr("error_load_project", error=exc))
            return False

        loaded_scans: list[Scan] = []
        if isinstance(data.get("scans"), list):
            loaded_scans = [Scan.from_json(s) for s in data.get("scans", []) if isinstance(s, dict)]
        else:
            # Backward compatibility: single scan format.
            image_path = str(data.get("image_path", ""))
            legacy_sheets = [Sheet.from_json(s) for s in data.get("sheets", [])]
            if image_path or legacy_sheets:
                loaded_scans = [
                    Scan(
                        name=Path(image_path).stem if image_path else self.tr("sheet_default"),
                        image_path=image_path,
                        sheets=legacy_sheets,
                    )
                ]

        self.scans = loaded_scans
        if self.scans:
            self.set_current_scan(0)
        else:
            self.set_current_scan(None)

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

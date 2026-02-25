# pymapcal

Qt MVP app for calibrating map sheets on top of a scanned map.

## Requirements
- Python 3.10+
- PySide6 (`pip install -r requirements.txt`)

## Run
```bash
python3 main.py
```

## Usage
1. `File -> Open map scan` (`tif/tiff/bmp/png/jpg/jpeg`).
2. Click `New sheet` and add corner points by clicking on the map.
3. Click `Close sheet` to create the polygon.
4. `Add calibration point` adds a calibration point (max 9 per sheet).
   - after clicking, coordinates are auto-predicted from current calibration,
   - you can edit them manually,
   - when adding/moving a calibration point, cursor guide lines are shown.
5. Drag points to move them.
6. In the side panel:
   - project tree: each sheet has `Crop outline` and `Calibration points`,
   - clicking a tree item selects the corresponding sheet/point on the map,
   - `Use calibration point in crop outline` allows reusing a calibration point as a crop point,
   - edit selected sheet name and scale,
   - edit `Lon` and `Lat` for selected point (supported formats below).
7. Status bar shows cursor position in pixels and approximate geo coordinates.
8. Pan and zoom:
   - mouse wheel: zoom to cursor,
   - middle mouse button + drag: pan,
   - `Zoom +`, `Zoom -`, `Zoom 100%` buttons in side panel.

## Notes
- Cursor geo position and grid are computed using:
  - 2 points: similarity transform (rotation + scale + translation),
  - 3+ points: affine least-squares fit.
- Outline lines have constant on-screen thickness (not scaled with zoom).
- `File -> Save` writes to current project file (no prompt).
- `File -> Save as...` writes to a new file/path.
- `File -> Load project` loads sheet data from JSON.

## `.pymapcal` Settings
The app loads settings from:
1. `./.pymapcal` (current directory),
2. `~/.pymapcal` (fallback if local file does not exist).

File format: JSON (either direct keys or nested under `display`).

i18n:
- `language: "pl"` or `language: "en"`

Example:
```json
{
  "language": "en",
  "display": {
    "outline_width": 2,
    "outline_selected_width": 3,
    "draft_outline_width": 2,
    "crosshair_arm_corner": 16,
    "crosshair_arm_cal": 14,
    "crosshair_ring_corner": 4,
    "crosshair_ring_cal": 2,
    "crosshair_selected_arm_bonus": 6,
    "crosshair_selected_ring_bonus": 2,
    "cursor_guide_width": 2,
    "cursor_guide_alpha": 200,
    "cursor_guide_dash": 10,
    "cursor_guide_gap": 6,
    "cursor_guide_color": "#FFD84D"
  }
}
```

## Coordinate formats
Supported for `Lon/Lat`:
- `DD` (decimal degrees), e.g. `18.654321`, `-54.1234`
- `DD + hemisphere`, e.g. `54.1234N`, `18.6543 E`
- `DMM + hemisphere`, e.g. `54 12.34 N`, `18° 39.26' E`
- `DMS + hemisphere`, e.g. `54 12 20.5 N`, `18°39'15.2"E`

Hemispheres:
- `N/S` for latitude (`Lat`)
- `E/W` for longitude (`Lon`)

## Input examples (ready to paste)
Same location in different formats:

- `Lat`:
  - `54.205694`
  - `54.205694N`
  - `54 12.3416 N`
  - `54°12.3416'N`
  - `54 12 20.5 N`
  - `54°12'20.5"N`

- `Lon`:
  - `18.652611`
  - `18.652611E`
  - `18 39.1567 E`
  - `18°39.1567'E`
  - `18 39 9.4 E`
  - `18°39'9.4"E`

Examples for west/south hemispheres:
- `Lat`: `33.9249S`, `33 55.494 S`, `33 55 29.6 S`
- `Lon`: `151.2093W`, `151 12.558 W`, `151 12 33.5 W`

## Start with a project file
You can open a project directly on startup:
```bash
python3 main.py /path/to/project.json
```

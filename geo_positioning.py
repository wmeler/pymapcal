from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Optional, Protocol, Sequence


class GeoPointLike(Protocol):
    x: float
    y: float
    lon: Optional[float]
    lat: Optional[float]
    is_corner: bool


WGS84_A = 6378137.0
WGS84_E = 0.08181919084262149


@dataclass
class GeoPositionResult:
    geo: Optional[tuple[float, float]]
    inside_outline: bool
    used_mercator: bool

    def to_dict(self) -> dict:
        return {
            "geo": None if self.geo is None else {"lon": self.geo[0], "lat": self.geo[1]},
            "inside_outline": self.inside_outline,
            "used_mercator": self.used_mercator,
        }


@dataclass
class CalibrationSnapResult:
    raw_x: float
    raw_y: float
    final_x: float
    final_y: float
    raw_geo: Optional[tuple[float, float]]
    final_geo: Optional[tuple[float, float]]
    inside_outline: bool
    used_mercator: bool
    did_snap: bool
    step_minutes: Optional[float]
    threshold_image_px: float

    def to_dict(self) -> dict:
        return {
            "raw_xy": {"x": self.raw_x, "y": self.raw_y},
            "final_xy": {"x": self.final_x, "y": self.final_y},
            "raw_geo": None if self.raw_geo is None else {"lon": self.raw_geo[0], "lat": self.raw_geo[1]},
            "final_geo": None if self.final_geo is None else {"lon": self.final_geo[0], "lat": self.final_geo[1]},
            "inside_outline": self.inside_outline,
            "used_mercator": self.used_mercator,
            "did_snap": self.did_snap,
            "step_minutes": self.step_minutes,
            "threshold_image_px": self.threshold_image_px,
        }


def _corners(points: Sequence[GeoPointLike]) -> list[GeoPointLike]:
    return [p for p in points if p.is_corner]


def point_in_polygon(x: float, y: float, polygon: Sequence[GeoPointLike]) -> bool:
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


def grid_step_minutes_for_scale(scale_value: Optional[int]) -> Optional[float]:
    if scale_value is None or scale_value <= 0:
        return None
    if scale_value <= 50_000:
        return 1.0
    if scale_value <= 100_000:
        return 2.0
    if scale_value <= 250_000:
        return 5.0
    if scale_value <= 500_000:
        return 10.0
    if scale_value <= 1_000_000:
        return 30.0
    if scale_value <= 2_000_000:
        return 60.0
    return 120.0


def mercator_wgs84_lat_to_y(lat_deg: float) -> float:
    lat = max(-89.999999, min(89.999999, lat_deg))
    phi = math.radians(lat)
    sin_phi = math.sin(phi)
    part = ((1.0 - WGS84_E * sin_phi) / (1.0 + WGS84_E * sin_phi)) ** (WGS84_E / 2.0)
    return WGS84_A * math.log(math.tan(math.pi / 4.0 + phi / 2.0) * part)


def mercator_wgs84_y_to_lat(y: float) -> float:
    ts = math.exp(-y / WGS84_A)
    phi = math.pi / 2.0 - 2.0 * math.atan(ts)
    for _ in range(8):
        sin_phi = math.sin(phi)
        part = ((1.0 - WGS84_E * sin_phi) / (1.0 + WGS84_E * sin_phi)) ** (WGS84_E / 2.0)
        phi = math.pi / 2.0 - 2.0 * math.atan(ts * part)
    return math.degrees(phi)


class GeoTransform:
    def __init__(
        self,
        mode: str,
        affine_forward: Optional[tuple[list[float], list[float]]] = None,
        affine_reverse: Optional[tuple[list[float], list[float]]] = None,
        sim_pixels: Optional[tuple[tuple[float, float], tuple[float, float]]] = None,
        sim_geo: Optional[tuple[tuple[float, float], tuple[float, float]]] = None,
        geo_space: str = "lonlat",
    ) -> None:
        self.mode = mode
        self.affine_forward = affine_forward
        self.affine_reverse = affine_reverse
        self.sim_pixels = sim_pixels
        self.sim_geo = sim_geo
        self.geo_space = geo_space

    def _project_geo(self, lon: float, lat: float) -> tuple[float, float]:
        if self.geo_space == "mercator":
            return lon, mercator_wgs84_lat_to_y(lat)
        return lon, lat

    def _unproject_geo(self, geo_x: float, geo_y: float) -> tuple[float, float]:
        if self.geo_space == "mercator":
            return geo_x, mercator_wgs84_y_to_lat(geo_y)
        return geo_x, geo_y

    def pixel_to_geo(self, x: float, y: float) -> Optional[tuple[float, float]]:
        if self.mode == "affine" and self.affine_forward is not None:
            geo_x_coef, geo_y_coef = self.affine_forward
            geo_x, geo_y = apply_affine(geo_x_coef, geo_y_coef, x, y)
            return self._unproject_geo(geo_x, geo_y)
        if self.mode == "similarity" and self.sim_pixels and self.sim_geo:
            projected = apply_similarity(self.sim_pixels[0], self.sim_pixels[1], self.sim_geo[0], self.sim_geo[1], x, y)
            if projected is None:
                return None
            return self._unproject_geo(projected[0], projected[1])
        return None

    def geo_to_pixel(self, lon: float, lat: float) -> Optional[tuple[float, float]]:
        geo_x, geo_y = self._project_geo(lon, lat)
        if self.mode == "affine" and self.affine_reverse is not None:
            x_coef, y_coef = self.affine_reverse
            return apply_affine(x_coef, y_coef, geo_x, geo_y)
        if self.mode == "similarity" and self.sim_pixels and self.sim_geo:
            return apply_similarity(self.sim_geo[0], self.sim_geo[1], self.sim_pixels[0], self.sim_pixels[1], geo_x, geo_y)
        return None


def _build_geo_transform_from_known_points(
    known: Sequence[GeoPointLike],
    use_mercator: bool,
) -> Optional[GeoTransform]:
    geo_samples = []
    for p in known:
        lon = float(p.lon)
        lat = float(p.lat)
        geo_y = mercator_wgs84_lat_to_y(lat) if use_mercator else lat
        geo_samples.append((p.x, p.y, lon, geo_y))

    if len(known) >= 3:
        forward = affine_fit(geo_samples)
        reverse = affine_fit([(lon, geo_y, x, y) for x, y, lon, geo_y in geo_samples])
        if forward and reverse:
            return GeoTransform(
                "affine",
                affine_forward=forward,
                affine_reverse=reverse,
                geo_space="mercator" if use_mercator else "lonlat",
            )

    if len(known) >= 2:
        x1, y1, lon1, geo_y1 = geo_samples[0]
        x2, y2, lon2, geo_y2 = geo_samples[1]
        sim_pixels = ((x1, y1), (x2, y2))
        sim_geo = ((lon1, geo_y1), (lon2, geo_y2))
        if (
            math.hypot(sim_pixels[1][0] - sim_pixels[0][0], sim_pixels[1][1] - sim_pixels[0][1]) > 1e-9
            and math.hypot(sim_geo[1][0] - sim_geo[0][0], sim_geo[1][1] - sim_geo[0][1]) > 1e-9
        ):
            return GeoTransform(
                "similarity",
                sim_pixels=sim_pixels,
                sim_geo=sim_geo,
                geo_space="mercator" if use_mercator else "lonlat",
            )
    return None


def build_geo_transform(points: Sequence[GeoPointLike], use_mercator: bool = False) -> Optional[GeoTransform]:
    known = [p for p in points if p.lon is not None and p.lat is not None]
    if not known:
        return None

    corner_known = [p for p in known if p.is_corner]
    if len(corner_known) < 3:
        return _build_geo_transform_from_known_points(known, use_mercator)

    transform = _build_geo_transform_from_known_points(corner_known, use_mercator)
    if transform is None:
        return _build_geo_transform_from_known_points(known, use_mercator)

    extra_known = [p for p in known if not p.is_corner]
    if not extra_known:
        return transform

    accepted_points = list(corner_known)
    for p in extra_known:
        predicted_xy = transform.geo_to_pixel(float(p.lon), float(p.lat))
        if predicted_xy is None:
            continue
        residual_px = math.hypot(predicted_xy[0] - p.x, predicted_xy[1] - p.y)
        # Keep only calibration points that are consistent with the sheet geometry.
        if residual_px <= 25.0:
            accepted_points.append(p)

    if len(accepted_points) == len(corner_known):
        return transform

    refined_transform = _build_geo_transform_from_known_points(accepted_points, use_mercator)
    return refined_transform or transform


def geo_position_from_image_xy(
    points: Sequence[GeoPointLike],
    x: float,
    y: float,
) -> GeoPositionResult:
    corners = _corners(points)
    inside_outline = len(corners) >= 3 and point_in_polygon(x, y, corners)
    transform = build_geo_transform(points, use_mercator=inside_outline)
    if transform is None:
        return GeoPositionResult(geo=None, inside_outline=inside_outline, used_mercator=inside_outline)
    return GeoPositionResult(
        geo=transform.pixel_to_geo(x, y),
        inside_outline=inside_outline,
        used_mercator=inside_outline,
    )


def _point_to_segment_distance(
    px: float,
    py: float,
    ax: float,
    ay: float,
    bx: float,
    by: float,
) -> tuple[float, tuple[float, float]]:
    dx = bx - ax
    dy = by - ay
    length_sq = dx * dx + dy * dy
    if length_sq <= 1e-12:
        return math.hypot(px - ax, py - ay), (ax, ay)

    t = ((px - ax) * dx + (py - ay) * dy) / length_sq
    t = max(0.0, min(1.0, t))
    qx = ax + t * dx
    qy = ay + t * dy
    return math.hypot(px - qx, py - qy), (qx, qy)


def _meridian_segment(
    transform: GeoTransform,
    lon: float,
    lat: float,
    lat_span_deg: float,
) -> Optional[tuple[tuple[float, float], tuple[float, float]]]:
    a = transform.geo_to_pixel(lon, lat - lat_span_deg)
    b = transform.geo_to_pixel(lon, lat + lat_span_deg)
    if a is None or b is None:
        return None
    return a, b


def _parallel_segment(
    transform: GeoTransform,
    lon: float,
    lat: float,
    lon_span_deg: float,
) -> Optional[tuple[tuple[float, float], tuple[float, float]]]:
    a = transform.geo_to_pixel(lon - lon_span_deg, lat)
    b = transform.geo_to_pixel(lon + lon_span_deg, lat)
    if a is None or b is None:
        return None
    return a, b


def _distance_to_geo_segment(
    x: float,
    y: float,
    segment: Optional[tuple[tuple[float, float], tuple[float, float]]],
) -> tuple[Optional[float], Optional[tuple[float, float]]]:
    if segment is None:
        return None, None
    (ax, ay), (bx, by) = segment
    return _point_to_segment_distance(x, y, ax, ay, bx, by)


def _grid_line_threshold_image_px(distances: list[float]) -> float:
    if not distances:
        return 0.0
    local_grid_spacing = min(distances)
    return max(6.0, min(local_grid_spacing * 0.2, 48.0))


def meridian_snap_threshold_image_px(
    transform: GeoTransform,
    lon: float,
    lat: float,
    step_deg: float,
    x: float,
    y: float,
) -> float:
    lat_span_deg = max(step_deg * 2.0, 1e-6)
    distances = []
    for lon_offset in (-step_deg, step_deg):
        distance, _ = _distance_to_geo_segment(
            x,
            y,
            _meridian_segment(transform, lon + lon_offset, lat, lat_span_deg),
        )
        if distance is not None:
            distances.append(distance)
    return _grid_line_threshold_image_px(distances)


def parallel_snap_threshold_image_px(
    transform: GeoTransform,
    lon: float,
    lat: float,
    step_deg: float,
    x: float,
    y: float,
) -> float:
    lon_span_deg = max(step_deg * 2.0, 1e-6)
    distances = []
    for lat_offset in (-step_deg, step_deg):
        distance, _ = _distance_to_geo_segment(
            x,
            y,
            _parallel_segment(transform, lon, lat + lat_offset, lon_span_deg),
        )
        if distance is not None:
            distances.append(distance)
    return _grid_line_threshold_image_px(distances)


def snap_calibration_position(
    points: Sequence[GeoPointLike],
    scale_value: Optional[int],
    x: float,
    y: float,
) -> CalibrationSnapResult:
    base = geo_position_from_image_xy(points, x, y)
    threshold = 0.0

    if not base.inside_outline or scale_value is None:
        return CalibrationSnapResult(
            raw_x=x,
            raw_y=y,
            final_x=x,
            final_y=y,
            raw_geo=base.geo,
            final_geo=base.geo,
            inside_outline=base.inside_outline,
            used_mercator=base.used_mercator,
            did_snap=False,
            step_minutes=None,
            threshold_image_px=threshold,
        )

    transform = build_geo_transform(points, use_mercator=True)
    step_minutes = grid_step_minutes_for_scale(scale_value)
    if transform is None or step_minutes is None or base.geo is None:
        return CalibrationSnapResult(
            raw_x=x,
            raw_y=y,
            final_x=x,
            final_y=y,
            raw_geo=base.geo,
            final_geo=base.geo,
            inside_outline=base.inside_outline,
            used_mercator=base.used_mercator,
            did_snap=False,
            step_minutes=step_minutes,
            threshold_image_px=threshold,
        )

    lon, lat = base.geo
    step_deg = step_minutes / 60.0
    snapped_lon = round(lon / step_deg) * step_deg
    snapped_lat = round(lat / step_deg) * step_deg

    final_x = x
    final_y = y
    final_lon = lon
    final_lat = lat
    did_snap = False
    snapped_meridian = False

    lon_threshold = meridian_snap_threshold_image_px(transform, lon, lat, step_deg, final_x, final_y)
    meridian_distance, meridian_xy = _distance_to_geo_segment(
        final_x,
        final_y,
        _meridian_segment(transform, snapped_lon, lat, max(step_deg * 2.0, 1e-6)),
    )
    if meridian_distance is not None and meridian_xy is not None and meridian_distance <= lon_threshold:
        final_x, final_y = meridian_xy
        projected_geo = transform.pixel_to_geo(final_x, final_y)
        if projected_geo is not None:
            final_lon = snapped_lon
            final_lat = projected_geo[1]
            did_snap = True
            snapped_meridian = True

    lat_threshold = parallel_snap_threshold_image_px(transform, final_lon, final_lat, step_deg, final_x, final_y)
    parallel_distance, parallel_xy = _distance_to_geo_segment(
        final_x,
        final_y,
        _parallel_segment(transform, final_lon, snapped_lat, max(step_deg * 2.0, 1e-6)),
    )
    if parallel_distance is not None and parallel_xy is not None and parallel_distance <= lat_threshold:
        final_x, final_y = parallel_xy
        projected_geo = transform.pixel_to_geo(final_x, final_y)
        if projected_geo is not None:
            if not snapped_meridian:
                final_lon = projected_geo[0]
            final_lat = snapped_lat
            did_snap = True

    threshold = max(lon_threshold, lat_threshold)
    final_geo = (final_lon, final_lat) if did_snap else base.geo

    return CalibrationSnapResult(
        raw_x=x,
        raw_y=y,
        final_x=final_x,
        final_y=final_y,
        raw_geo=base.geo,
        final_geo=final_geo,
        inside_outline=base.inside_outline,
        used_mercator=base.used_mercator,
        did_snap=did_snap,
        step_minutes=step_minutes,
        threshold_image_px=threshold,
    )

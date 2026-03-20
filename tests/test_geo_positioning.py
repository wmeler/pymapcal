import unittest
from types import SimpleNamespace

from geo_positioning import geo_position_from_image_xy, snap_calibration_position


GEO_TOLERANCE = 1e-2

SNAP_TOLERANCE = 1e-6

def point(x: float, y: float, lon: float, lat: float, is_corner: bool) -> SimpleNamespace:
    return SimpleNamespace(x=x, y=y, lon=lon, lat=lat, is_corner=is_corner)


# These three fixtures were distilled from:
# - click logs recorded on 2026-03-20 in geo_position_cases.jsonl
# - corrected saved point values from 1020_00.json (scan 1020_01 / Arkusz 1)
BASE_POINTS = [
    point(273.375, 296.75, 14.0, 55.75, True),
    point(6740.125, 291.25, 18.916666666666668, 55.75, True),
    point(6743.375, 4652.125, 18.916666666666668, 53.833333333333336, True),
    point(276.34298749999977, 4664.8182945312465, 14.0, 53.833333333333336, True),
    point(4224.75, 3165.0, 17.0, 55.5, False),
    point(1586.5, 3169.875, 15.0, 55.5, False),
]

CASE_FIXTURES = [
    {
        "name": "15E_55N",
        "points": BASE_POINTS,
        "raw_xy": (1585.7880357747003, 2030.9875238727775),
        "expected_geo": (15.0, 55.0),
    },
    {
        "name": "16E_55N",
        "points": BASE_POINTS
        + [
            point(1585.7880357747003, 2030.9875238727775, 15.0, 55.0, False),
        ],
        "raw_xy": (2905.4831592000764, 2029.1691034284333),
        "expected_geo": (16.0, 55.0),
    },
    {
        "name": "15E_55_5N",
        "points": BASE_POINTS
        + [
            point(1585.7880357747003, 2030.9875238727775, 15.0, 55.0, False),
            point(2905.4831592000764, 2029.1691034284333, 16.0, 55.0, False),
        ],
        "raw_xy": (1585.4608836707705, 875.1318780127544),
        "expected_geo": (15.0, 55.5),
    },
]


class GeoPositioning102001Tests(unittest.TestCase):
    def test_geo_position_from_image_xy_matches_corrected_points(self) -> None:
        for case in CASE_FIXTURES:
            with self.subTest(case=case["name"]):
                x, y = case["raw_xy"]
                expected_lon, expected_lat = case["expected_geo"]
                result = geo_position_from_image_xy(case["points"], x, y)

                self.assertIsNotNone(result.geo)
                self.assertTrue(result.inside_outline)
                self.assertTrue(result.used_mercator)
                assert result.geo is not None
                self.assertAlmostEqual(result.geo[0], expected_lon, delta=GEO_TOLERANCE)
                self.assertAlmostEqual(result.geo[1], expected_lat, delta=GEO_TOLERANCE)

    def test_snap_calibration_position_matches_corrected_points(self) -> None:
        for case in CASE_FIXTURES:
            with self.subTest(case=case["name"]):
                x, y = case["raw_xy"]
                expected_lon, expected_lat = case["expected_geo"]
                result = snap_calibration_position(
                    points=case["points"],
                    scale_value=500000,
                    x=x,
                    y=y,
                )

                self.assertIsNotNone(result.final_geo)
                self.assertTrue(result.inside_outline)
                self.assertTrue(result.did_snap)
                assert result.final_geo is not None
                self.assertAlmostEqual(result.final_geo[0], expected_lon, delta=SNAP_TOLERANCE)
                self.assertAlmostEqual(result.final_geo[1], expected_lat, delta=SNAP_TOLERANCE)


if __name__ == "__main__":
    unittest.main()

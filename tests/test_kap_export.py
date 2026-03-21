import unittest
from pathlib import Path

from kap_export import KapExportJob, KapPolygonPoint, KapReference, _build_header


class KapExportHeaderTests(unittest.TestCase):
    def test_header_uses_integer_image_coordinates_and_float_geo_coordinates(self) -> None:
        job = KapExportJob(
            sheet_name="Test Sheet",
            image_path=Path("/tmp/source.png"),
            output_path=Path("/tmp/out.kap"),
            width=1234,
            height=987,
            scale=50000,
            polygon=[
                KapPolygonPoint(pixel_x=10.6, pixel_y=20.4, lon=15.125, lat=54.5),
                KapPolygonPoint(pixel_x=110.2, pixel_y=20.9, lon=15.25, lat=54.5),
                KapPolygonPoint(pixel_x=109.7, pixel_y=220.8, lon=15.25, lat=54.25),
            ],
            references=[
                KapReference(pixel_x=50.7, pixel_y=100.3, lon=15.2, lat=54.4),
            ],
        )

        header = _build_header(
            job=job,
            ed_date="03/21/2026",
            sounding_unit="Meters",
            sounding_datum="WGS84",
        )

        self.assertIn("REF/1,11,20,54.5000000000,15.1250000000", header)
        self.assertIn("REF/2,110,21,54.5000000000,15.2500000000", header)
        self.assertIn("REF/4,51,100,54.4000000000,15.2000000000", header)


if __name__ == "__main__":
    unittest.main()

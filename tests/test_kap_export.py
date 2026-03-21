import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from PySide6.QtGui import QImage

from kap_export import KapExportJob, KapPolygonPoint, KapReference, _build_header, run_kap_export_jobs


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
            ed_date="21/03/2026",
            sounding_unit="Meters",
            sounding_datum="WGS84",
        )

        self.assertIn("CED/SE=,RE=,ED=21/03/2026", header)
        self.assertIn("REF/1,11,20,54.5000000000,15.1250000000", header)
        self.assertIn("REF/2,110,21,54.5000000000,15.2500000000", header)
        self.assertIn("REF/4,51,100,54.4000000000,15.2000000000", header)

    def test_run_kap_export_jobs_writes_kap_without_external_imgkap(self) -> None:
        with TemporaryDirectory() as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            image_path = tmp_dir / "source.png"
            output_path = tmp_dir / "out.kap"

            image = QImage(4, 4, QImage.Format.Format_RGB32)
            for y in range(4):
                for x in range(4):
                    image.setPixel(x, y, 0x00FF00 if (x + y) % 2 == 0 else 0xFF0000)
            self.assertTrue(image.save(str(image_path), "PNG"))

            job = KapExportJob(
                sheet_name="Test Sheet",
                image_path=image_path,
                output_path=output_path,
                width=4,
                height=4,
                scale=50000,
                polygon=[
                    KapPolygonPoint(pixel_x=0, pixel_y=0, lon=15.0, lat=54.5),
                    KapPolygonPoint(pixel_x=3, pixel_y=0, lon=15.1, lat=54.5),
                    KapPolygonPoint(pixel_x=3, pixel_y=3, lon=15.1, lat=54.4),
                    KapPolygonPoint(pixel_x=0, pixel_y=3, lon=15.0, lat=54.4),
                ],
                references=[
                    KapReference(pixel_x=1.4, pixel_y=1.6, lon=15.05, lat=54.45),
                ],
            )

            results = run_kap_export_jobs(
                jobs=[job],
                imgkap_path="python",
                ed_date="21/03/2026",
                sounding_datum="WGS84",
            )

            self.assertEqual(len(results), 1)
            self.assertTrue(results[0].success)
            self.assertTrue(output_path.exists())

            data = output_path.read_bytes()
            self.assertIn(b"BSB/NA=Test Sheet\r\n", data)
            self.assertIn(b"CED/SE=,RE=,ED=21/03/2026\r\n", data)
            self.assertIn(b"REF/1,0,0,54.5000000000,15.0000000000\r\n", data)
            self.assertIn(b"REF/5,1,2,54.4500000000,15.0500000000\r\n", data)
            self.assertIn(b"PLY/4,54.4000000000,15.0000000000\r\n", data)
            self.assertIn(b"IFM/1\r\n", data)
            self.assertIn(b"RGB/0,0,255,0\r\n", data)
            self.assertIn(b"RGB/1,255,0,0\r\n", data)

            index_offset = int.from_bytes(data[-4:], "big")
            self.assertGreater(index_offset, 0)
            self.assertEqual(len(data) - index_offset, (job.height + 1) * 4)


if __name__ == "__main__":
    unittest.main()

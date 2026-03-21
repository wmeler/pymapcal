import unittest

from wz_downloader import extract_wz_links, parse_total_size_from_content_range


class WzDownloaderTests(unittest.TestCase):
    def test_extract_wz_links_filters_wz_pdfs(self) -> None:
        html = """
        <html><body>
          <a href="/files/WZ_01_2026.pdf">ok</a>
          <a href="https://example.com/WZ_02_2026.PDF">ok2</a>
          <a href="/files/not_wz.pdf">no</a>
          <a href="/files/WZ_03_2026.doc">no2</a>
        </body></html>
        """

        links = extract_wz_links("https://bhmw.gov.pl/pl/news?strona=1", html)

        self.assertEqual(
            links,
            {
                "WZ_01_2026.pdf": "https://bhmw.gov.pl/files/WZ_01_2026.pdf",
                "WZ_02_2026.PDF": "https://example.com/WZ_02_2026.PDF",
            },
        )

    def test_parse_total_size_from_content_range(self) -> None:
        self.assertEqual(parse_total_size_from_content_range("bytes 0-0/12345"), 12345)
        self.assertIsNone(parse_total_size_from_content_range("bytes 0-0/*"))
        self.assertIsNone(parse_total_size_from_content_range(None))


if __name__ == "__main__":
    unittest.main()

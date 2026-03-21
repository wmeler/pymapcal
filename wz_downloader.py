from __future__ import annotations

import time
from dataclasses import dataclass, field
from http.cookiejar import CookieJar
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urljoin, urlparse
from urllib.request import HTTPCookieProcessor, Request, build_opener


NEWS_URL_TEMPLATE = "https://bhmw.gov.pl/pl/news?strona={page}"
PDF_NAME_PREFIX = "WZ"
PDF_SUFFIX = ".pdf"
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/145.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,"
        "image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"
    ),
}
MIN_REQUEST_INTERVAL_SECONDS = 1.0


class DownloadCancelled(Exception):
    pass


@dataclass
class WzDownloadResult:
    total_links: int
    downloaded: int
    skipped: int
    failed: int
    output_dir: Path
    cancelled: bool = False
    errors: list[str] = field(default_factory=list)


class LinkExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.hrefs: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        for name, value in attrs:
            if name.lower() == "href" and value:
                self.hrefs.append(value)
                return


class WzDownloader:
    def __init__(
        self,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[int, int, str], None]] = None,
        cancel_requested_cb: Optional[Callable[[], bool]] = None,
    ) -> None:
        self.log_cb = log_cb
        self.progress_cb = progress_cb
        self.cancel_requested_cb = cancel_requested_cb
        self.cookie_jar = CookieJar()
        self.http_opener = build_opener(HTTPCookieProcessor(self.cookie_jar))
        self.last_request_ts = 0.0

    def run(self, start_page: int, end_page: int, out_dir: Path) -> WzDownloadResult:
        if start_page < 1 or end_page < start_page:
            raise ValueError("Niepoprawny zakres stron. Użyj start >= 1 i end >= start.")

        out_dir.mkdir(parents=True, exist_ok=True)
        unique_links: dict[str, str] = {}
        errors: list[str] = []
        total_pages = end_page - start_page + 1

        for page_index, page_number in enumerate(range(start_page, end_page + 1), start=1):
            self._check_cancel()
            page_url = NEWS_URL_TEMPLATE.format(page=page_number)
            self._set_progress(page_index, total_pages, f"[SCAN] {page_url}")
            self._log(f"[SCAN] {page_url}\n")
            try:
                html = self.fetch_page(page_url)
            except Exception as exc:  # noqa: BLE001
                msg = f"Nie można pobrać strony {page_url}: {exc}"
                self._log(f"[ERR] {msg}\n")
                errors.append(msg)
                continue

            page_links = extract_wz_links(page_url, html)
            new_links = 0
            for filename, pdf_url in page_links.items():
                if filename not in unique_links:
                    unique_links[filename] = pdf_url
                    new_links += 1
            self._log(f"[SCAN] znalezione linki WZ*.pdf: {len(page_links)}, nowe: {new_links}\n")

        if not unique_links:
            return WzDownloadResult(
                total_links=0,
                downloaded=0,
                skipped=0,
                failed=len(errors),
                output_dir=out_dir,
                cancelled=False,
                errors=errors,
            )

        downloaded = 0
        skipped = 0
        failed = 0
        sorted_links = sorted(unique_links.items(), key=lambda item: (out_dir / item[0]).exists())
        total_units = max(1, len(sorted_links) * 1000)

        for file_index, (filename, pdf_url) in enumerate(sorted_links, start=1):
            self._check_cancel()
            destination = out_dir / filename
            base_units = (file_index - 1) * 1000
            label = f"{filename} ({file_index}/{len(sorted_links)})"
            self._set_progress(base_units, total_units, label)

            if not destination.exists():
                try:
                    saved_size = self.download_file(
                        pdf_url,
                        destination,
                        expected_size=None,
                        file_index=file_index,
                        total_files=len(sorted_links),
                    )
                    self._log(f"[OK] {filename} ({saved_size} B)\n")
                    downloaded += 1
                except DownloadCancelled:
                    return WzDownloadResult(
                        total_links=len(unique_links),
                        downloaded=downloaded,
                        skipped=skipped,
                        failed=failed,
                        output_dir=out_dir,
                        cancelled=True,
                        errors=errors,
                    )
                except Exception as exc:  # noqa: BLE001
                    msg = f"Nie udało się pobrać {pdf_url}: {exc}"
                    self._log(f"[ERR] {msg}\n")
                    errors.append(msg)
                    failed += 1
                continue

            remote_size = self.get_remote_size(pdf_url)
            if remote_size is not None:
                local_size = destination.stat().st_size
                if local_size == remote_size:
                    self._log(f"[SKIP] {filename} (rozmiar zgodny: {local_size} B)\n")
                    skipped += 1
                    self._set_progress(file_index * 1000, total_units, label)
                    continue

            if remote_size is None:
                self._log(
                    f"[SKIP] {filename} (plik istnieje, nie udało się odczytać rozmiaru z serwera)\n"
                )
                skipped += 1
                self._set_progress(file_index * 1000, total_units, label)
                continue

            try:
                saved_size = self.download_file(
                    pdf_url,
                    destination,
                    expected_size=remote_size,
                    file_index=file_index,
                    total_files=len(sorted_links),
                )
                self._log(f"[OK] {filename} ({saved_size} B)\n")
                downloaded += 1
            except DownloadCancelled:
                return WzDownloadResult(
                    total_links=len(unique_links),
                    downloaded=downloaded,
                    skipped=skipped,
                    failed=failed,
                    output_dir=out_dir,
                    cancelled=True,
                    errors=errors,
                )
            except Exception as exc:  # noqa: BLE001
                msg = f"Nie udało się pobrać {pdf_url}: {exc}"
                self._log(f"[ERR] {msg}\n")
                errors.append(msg)
                failed += 1

        return WzDownloadResult(
            total_links=len(unique_links),
            downloaded=downloaded,
            skipped=skipped,
            failed=failed,
            output_dir=out_dir,
            cancelled=False,
            errors=errors,
        )

    def fetch_page(self, url: str) -> str:
        with self.make_request(url, method="GET") as response:
            body = response.read()
            encoding = response.headers.get_content_charset() or "utf-8"
        return body.decode(encoding, errors="replace")

    def get_remote_size(self, url: str) -> int | None:
        self._check_cancel()
        try:
            with self.make_request(url, method="HEAD") as response:
                content_length = response.headers.get("Content-Length")
                if content_length and content_length.isdigit():
                    return int(content_length)
        except HTTPError as exc:
            if exc.code not in {403, 405, 501}:
                return None
        except URLError:
            return None

        headers = dict(DEFAULT_HEADERS)
        headers["Range"] = "bytes=0-0"
        try:
            with self.make_request(url, method="GET", headers=headers) as response:
                total_from_range = parse_total_size_from_content_range(
                    response.headers.get("Content-Range")
                )
                if total_from_range is not None:
                    return total_from_range
                content_length = response.headers.get("Content-Length")
                if content_length and content_length.isdigit():
                    return int(content_length)
        except (HTTPError, URLError):
            return None
        return None

    def download_file(
        self,
        url: str,
        destination: Path,
        expected_size: int | None,
        file_index: int,
        total_files: int,
    ) -> int:
        temp_file = destination.with_suffix(destination.suffix + ".part")
        total_units = max(1, total_files * 1000)
        base_units = (file_index - 1) * 1000
        label = f"{destination.name} ({file_index}/{total_files})"

        try:
            with self.make_request(url, method="GET") as response, temp_file.open("wb") as out_file:
                total_bytes = expected_size
                if total_bytes is None:
                    content_length = response.headers.get("Content-Length")
                    if content_length and content_length.isdigit():
                        total_bytes = int(content_length)

                downloaded = 0
                while True:
                    self._check_cancel()
                    chunk = response.read(128 * 1024)
                    if not chunk:
                        break
                    out_file.write(chunk)
                    downloaded += len(chunk)
                    if total_bytes and total_bytes > 0:
                        fraction = min(0.999, downloaded / total_bytes)
                        done_units = base_units + int(fraction * 1000)
                        self._set_progress(done_units, total_units, label)
            actual_size = temp_file.stat().st_size
            if expected_size is not None and actual_size != expected_size:
                temp_file.unlink(missing_ok=True)
                raise RuntimeError(
                    f"Niepoprawny rozmiar pliku: pobrano {actual_size} B, oczekiwano {expected_size} B"
                )
            temp_file.replace(destination)
            self._set_progress(file_index * 1000, total_units, label)
            return actual_size
        except DownloadCancelled:
            temp_file.unlink(missing_ok=True)
            raise
        except Exception:
            temp_file.unlink(missing_ok=True)
            raise

    def make_request(self, url: str, method: str = "GET", headers: dict[str, str] | None = None):
        req = Request(url, method=method, headers=headers or DEFAULT_HEADERS)
        self._log(f"[HTTP->] {req.get_method()} {req.full_url}\n")
        for key, value in req.header_items():
            self._log(f"[HTTP->] {key}: {value}\n")

        self._check_cancel()
        elapsed = time.monotonic() - self.last_request_ts
        wait_seconds = MIN_REQUEST_INTERVAL_SECONDS - elapsed
        if wait_seconds > 0:
            self._log(f"[RATE] czekam {wait_seconds:.2f}s przed kolejnym requestem\n")
            time.sleep(wait_seconds)
        self.last_request_ts = time.monotonic()

        try:
            response = self.http_opener.open(req, timeout=30)
        except HTTPError as exc:
            self._log(f"[HTTP<-] {exc.code} {exc.reason} {exc.url}\n")
            for key, value in exc.headers.items():
                self._log(f"[HTTP<-] {key}: {value}\n")
            raise

        status = getattr(response, "status", response.getcode())
        reason = getattr(response, "reason", "")
        self._log(f"[HTTP<-] {status} {reason} {response.geturl()}\n")
        for key, value in response.headers.items():
            self._log(f"[HTTP<-] {key}: {value}\n")
        return response

    def _set_progress(self, current: int, total: int, label: str) -> None:
        if self.progress_cb is not None:
            self.progress_cb(max(0, current), max(1, total), label)

    def _log(self, text: str) -> None:
        if self.log_cb is not None:
            self.log_cb(text)

    def _check_cancel(self) -> None:
        if self.cancel_requested_cb is not None and self.cancel_requested_cb():
            raise DownloadCancelled()


def extract_wz_links(page_url: str, html: str) -> dict[str, str]:
    parser = LinkExtractor()
    parser.feed(html)
    links: dict[str, str] = {}
    for href in parser.hrefs:
        full_url = urljoin(page_url, href)
        filename = Path(unquote(urlparse(full_url).path)).name
        upper = filename.upper()
        if filename and upper.startswith(PDF_NAME_PREFIX) and upper.endswith(PDF_SUFFIX.upper()):
            links.setdefault(filename, full_url)
    return links


def parse_total_size_from_content_range(content_range: str | None) -> int | None:
    if not content_range or "/" not in content_range:
        return None
    total = content_range.rsplit("/", 1)[-1].strip()
    if total.isdigit():
        return int(total)
    return None


def download_wz_messages(
    out_dir: Path,
    start_page: int = 1,
    end_page: int = 16,
    log_cb: Optional[Callable[[str], None]] = None,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
    cancel_requested_cb: Optional[Callable[[], bool]] = None,
) -> WzDownloadResult:
    downloader = WzDownloader(
        log_cb=log_cb,
        progress_cb=progress_cb,
        cancel_requested_cb=cancel_requested_cb,
    )
    return downloader.run(start_page=start_page, end_page=end_page, out_dir=out_dir)

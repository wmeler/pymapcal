#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from wz_downloader import DownloadCancelled, download_wz_messages


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Skanuj strony BHMW i pobieraj PDF-y o nazwach WZ*.pdf."
    )
    parser.add_argument("--start", type=int, default=1, help="Pierwszy numer strony (domyślnie: 1)")
    parser.add_argument("--end", type=int, default=16, help="Ostatni numer strony (domyślnie: 16)")
    parser.add_argument(
        "--out-dir",
        default="WZ",
        help="Katalog docelowy dla plików PDF (domyślnie: WZ)",
    )
    args = parser.parse_args()

    try:
        result = download_wz_messages(
            out_dir=Path(args.out_dir),
            start_page=args.start,
            end_page=args.end,
            log_cb=lambda text: print(text, end=""),
        )
    except ValueError as exc:
        parser.error(str(exc))
    except DownloadCancelled:
        print("[SUMMARY] pobieranie przerwane przez użytkownika", file=sys.stderr)
        return 1

    print(
        f"[SUMMARY] łącznie: {result.total_links}, pobrane: {result.downloaded}, "
        f"pominięte: {result.skipped}, błędy: {result.failed}"
    )
    return 1 if result.failed else 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""
run_headless.py — headless CLI runner for Gmail Media Extractor

Run this on a server/VM where there is no display.
Before first use, copy token.json from the desktop app to:
  ~/.config/gmail_media_extractor/token.json

Files needed in the same directory as this script:
  core.py, credentials.json

Usage examples:
  python3 run_headless.py
  python3 run_headless.py --output-dir ~/GmailMedia --after 2024/01/01
  python3 run_headless.py --categories "PDF Documents,Images" --no-sheets
  python3 run_headless.py --list-categories
  python3 run_headless.py --retry-report ~/GmailMedia/report_20260608_120000.json
"""
from __future__ import annotations

import argparse
import json
import queue
import signal
import sys
import threading
from pathlib import Path

HERE     = Path(__file__).parent.resolve()
DATA_DIR = Path.home() / ".config" / "gmail_media_extractor"

sys.path.insert(0, str(HERE))

from core import (
    ALL_CATEGORIES,
    DownloadManifest,
    ExtractionWorker,
    WorkerConfig,
)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Extract Gmail attachments headlessly (no GUI required).",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    p.add_argument(
        "--output-dir", "-o",
        default=str(Path.home() / "GmailMedia"),
        help="Where to save downloaded files  (default: ~/GmailMedia)",
    )
    p.add_argument(
        "--categories", "-c",
        default="",
        help=(
            "Comma-separated categories to download.  Default: all.\n"
            f"Available: {', '.join(sorted(ALL_CATEGORIES))}"
        ),
    )
    p.add_argument("--after",  default="", help="Only emails AFTER  YYYY/MM/DD")
    p.add_argument("--before", default="", help="Only emails BEFORE YYYY/MM/DD")
    p.add_argument(
        "--token-path",
        default=str(DATA_DIR / "token.json"),
        help=f"Google OAuth token path  (default: {DATA_DIR / 'token.json'})",
    )
    p.add_argument(
        "--creds-path",
        default=str(HERE / "credentials.json"),
        help=f"credentials.json path  (default: {HERE / 'credentials.json'})",
    )
    p.add_argument("--no-sheets", action="store_true", help="Disable Google Sheets logging")
    p.add_argument("--list-categories", action="store_true", help="Print available categories and exit")
    p.add_argument(
        "--retry-report",
        default="",
        metavar="REPORT_JSON",
        help=(
            "Path to a previous run's report JSON.\n"
            "Retries only the message IDs that failed in that run."
        ),
    )
    p.add_argument(
        "--direction",
        default="both",
        choices=["both", "received", "sent"],
        help=(
            "Which emails to scan for attachments:\n"
            "  both      — all emails (default)\n"
            "  received  — only emails others sent to you\n"
            "  sent      — only emails you sent"
        ),
    )
    p.add_argument(
        "--sort-by-domain",
        action="store_true",
        help=(
            "Organise folders by email domain instead of sender address.\n"
            "Structure: @domain.com/Category/file  (or @domain.com/file when only one type selected)"
        ),
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()

    if args.list_categories:
        print("Available categories:")
        for cat in sorted(ALL_CATEGORIES):
            print(f"  {cat}")
        return

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    output_dir    = Path(args.output_dir).expanduser().resolve()
    token_path    = Path(args.token_path).expanduser()
    creds_path    = Path(args.creds_path).expanduser()
    manifest_path = output_dir / "download_manifest.json"
    sheet_id_path = DATA_DIR / "sheets_id.txt"

    if not token_path.exists():
        print(f"ERROR: token.json not found at {token_path}")
        print()
        print("Copy token.json from the desktop app to this machine:")
        print("  Linux / Mac : <app folder>/.app_data/token.json")
        print("  Windows     : %APPDATA%\\GmailMediaExtractor\\token.json")
        print()
        print(f"Then place it at: {token_path}")
        sys.exit(1)

    if not creds_path.exists():
        print(f"ERROR: credentials.json not found at {creds_path}")
        print(f"Copy credentials.json from the project's files/ folder to: {creds_path}")
        sys.exit(1)

    if args.categories:
        categories = {c.strip() for c in args.categories.split(",") if c.strip()}
        unknown    = categories - set(ALL_CATEGORIES)
        if unknown:
            print(f"ERROR: Unknown categories: {', '.join(sorted(unknown))}")
            print("Run with --list-categories to see valid options.")
            sys.exit(1)
    else:
        categories = set(ALL_CATEGORIES)

    override_ids = None
    if args.retry_report:
        report_path = Path(args.retry_report).expanduser()
        if not report_path.exists():
            print(f"ERROR: Report not found: {report_path}", file=sys.stderr)
            sys.exit(1)
        report_data  = json.loads(report_path.read_text(encoding="utf-8"))
        override_ids = [
            f["msg_id"] for f in report_data.get("failures", []) if f.get("msg_id")
        ]
        if not override_ids:
            print("No failed message IDs found in that report — nothing to retry.")
            return

    print("=" * 52)
    print("  Gmail Media Extractor — headless mode")
    print("=" * 52)
    print(f"  Output dir  : {output_dir}")
    print(f"  Categories  : {', '.join(sorted(categories))}")
    if override_ids is not None:
        print(f"  Retry mode  : {len(override_ids)} failed IDs from {Path(args.retry_report).name}")
    else:
        if args.after:
            print(f"  After       : {args.after}")
        if args.before:
            print(f"  Before      : {args.before}")
    direction_label = {"both": "received & sent", "received": "received only", "sent": "sent only"}
    print(f"  Emails      : {direction_label[args.direction]}")
    print(f"  Sort mode   : {'domain (@domain.com folders)' if args.sort_by_domain else 'sender (per-address folders)'}")
    print(f"  Sheets log  : {'disabled' if args.no_sheets else 'enabled'}")
    print()

    output_dir.mkdir(parents=True, exist_ok=True)
    manifest  = DownloadManifest(manifest_path)
    msg_q: queue.Queue = queue.Queue()

    cfg = WorkerConfig(
        output_dir      = output_dir,
        categories      = categories,
        after           = args.after,
        before          = args.before,
        creds_path      = creds_path,
        token_path      = token_path,
        manifest        = manifest,
        msg_queue       = msg_q,
        storage_mode    = "local",
        enable_sheets   = not args.no_sheets,
        sheet_id_path   = sheet_id_path,
        scan_links      = True,
        override_ids    = override_ids,
        email_direction = args.direction,
        sort_by_domain  = args.sort_by_domain,
    )

    worker = ExtractionWorker(cfg)

    def _shutdown(_sig, _frame):
        print("\nShutting down gracefully…", flush=True)
        worker.stop()

    signal.signal(signal.SIGINT,  _shutdown)
    signal.signal(signal.SIGTERM, _shutdown)

    thread = threading.Thread(target=worker.run, daemon=True)
    thread.start()

    total = 0
    BAR   = 30

    while thread.is_alive() or not msg_q.empty():
        try:
            msg = msg_q.get(timeout=0.5)
        except queue.Empty:
            continue

        kind = msg["kind"]

        if kind == "status":
            print(f"\n{msg['text']}", flush=True)

        elif kind == "total":
            total = msg["count"]
            print(f"  Found {total:,} emails to scan.\n", flush=True)

        elif kind == "progress":
            current = msg["current"]
            pct     = current / total * 100 if total else 0
            filled  = int(BAR * current / total) if total else 0
            bar     = "#" * filled + "-" * (BAR - filled)
            print(f"\r  [{bar}] {current:,}/{total:,}  ({pct:.1f}%)", end="", flush=True)

        elif kind == "error":
            print(f"\n[ERROR] {msg['text']}", flush=True)

        elif kind == "sheet_url":
            print(f"  Sheets log  : {msg['url']}", flush=True)

        elif kind == "done":
            report_glob = sorted(Path(msg["output_dir"]).glob("report_*.json"))
            report_hint = (
                f"\n  Report     : {report_glob[-1]}"
                if report_glob else ""
            )
            errors_hint = (
                f"\n  Re-run     : python3 run_headless.py --retry-report {report_glob[-1]}"
                if report_glob and msg.get("errors", 0) > 0 else ""
            )
            print(
                f"\n\n{'=' * 52}"
                f"\n  Done!"
                f"\n  Downloaded : {msg['downloaded']:,}"
                f"\n  Skipped    : {msg['skipped']:,}"
                f"\n  Errors     : {msg['errors']:,}"
                f"\n  Files in   : {msg['output_dir']}"
                f"{report_hint}"
                f"{errors_hint}"
                f"\n{'=' * 52}",
                flush=True,
            )

    thread.join()


if __name__ == "__main__":
    main()

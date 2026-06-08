"""
core.py — Gmail extraction engine
Handles auth, message scanning, attachment download,
OneDrive link extraction, OneDrive upload, and Google Sheets logging.
"""
from __future__ import annotations

import base64
import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
from email.utils import parsedate_to_datetime
from typing import Optional
import queue

_thread_local    = threading.local()
_build_svc_lock  = threading.Lock()  # httplib2 init is not thread-safe on Python 3.14
_http_exec_lock  = threading.Lock()  # httplib2 SSL ops crash under concurrency on Python 3.14

# ── Google libs ────────────────────────────────────────────────────────────
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build as _build_service
    from googleapiclient.errors import HttpError
    GOOGLE_OK = True
except ImportError:
    GOOGLE_OK = False

# ── Microsoft / requests libs (optional) ──────────────────────────────────
try:
    import requests as _req
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

try:
    import msal
    MSAL_OK = True
except ImportError:
    MSAL_OK = False

# ── Retry helpers ──────────────────────────────────────────────────────────

_RETRYABLE_STATUS = frozenset({429, 500, 502, 503, 504})


def _gapi_retry(fn, *, attempts: int = 4):
    """Retry a Google API callable on transient HTTP or network errors."""
    for attempt in range(attempts):
        try:
            return fn()
        except HttpError as exc:
            if attempt < attempts - 1 and exc.resp.status in _RETRYABLE_STATUS:
                time.sleep(2 ** attempt)
                continue
            raise
        except OSError:
            if attempt < attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise


def _req_retry(fn, *, attempts: int = 4):
    """Retry a requests call on transient HTTP status codes or network errors."""
    for attempt in range(attempts):
        try:
            resp = fn()
            if resp.status_code in _RETRYABLE_STATUS and attempt < attempts - 1:
                time.sleep(2 ** attempt)
                continue
            return resp
        except Exception as exc:
            is_transient = isinstance(exc, OSError)
            if REQUESTS_OK:
                is_transient = is_transient or isinstance(
                    exc,
                    (_req.exceptions.ConnectionError,
                     _req.exceptions.Timeout,
                     _req.exceptions.ChunkedEncodingError),
                )
            if is_transient and attempt < attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise


# Google OAuth scopes: Gmail (read) + Sheets (log)
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ── Azure App Registration ─────────────────────────────────────────────────
# Create at portal.azure.com → App registrations → New registration
#   Platform : Mobile and desktop application  →  http://localhost
#   Permissions: Files.ReadWrite, User.Read (delegated)
# Replace the placeholder below with your Application (client) ID,
# OR add it as GitHub secret MS_CLIENT_ID and let build.yml patch it.
MICROSOFT_CLIENT_ID = "3d72e275-7e00-482e-b67a-9c1ee834b51a"

ONEDRIVE_CONFIGURED = MICROSOFT_CLIENT_ID != "YOUR_AZURE_APP_CLIENT_ID_HERE"

MAX_WORKERS = 4  # parallel email-processing threads (httplib2 serialized; 4 avoids lock contention)

# ── MIME registry ──────────────────────────────────────────────────────────
MIME_TO_CATEGORY: dict[str, str] = {
    "application/pdf":                                                              "PDF Documents",
    "application/msword":                                                           "Word Documents",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document":     "Word Documents",
    "application/vnd.oasis.opendocument.text":                                     "Word Documents",
    "application/vnd.ms-excel":                                                    "Spreadsheets",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":           "Spreadsheets",
    "application/vnd.oasis.opendocument.spreadsheet":                              "Spreadsheets",
    "application/vnd.ms-powerpoint":                                               "Presentations",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation":  "Presentations",
    "application/vnd.oasis.opendocument.presentation":                            "Presentations",
    "image/jpeg": "Images", "image/png": "Images", "image/gif": "Images",
    "image/webp": "Images", "image/bmp": "Images", "image/tiff": "Images",
    "image/svg+xml": "Images", "image/heic": "Images", "image/heif": "Images",
    "video/mp4": "Videos", "video/mpeg": "Videos", "video/quicktime": "Videos",
    "video/x-msvideo": "Videos", "video/x-matroska": "Videos", "video/webm": "Videos",
    "audio/mpeg": "Audio", "audio/wav": "Audio", "audio/ogg": "Audio",
    "audio/mp4": "Audio", "audio/flac": "Audio",
    "application/zip": "Archives", "application/x-rar-compressed": "Archives",
    "application/x-7z-compressed": "Archives", "application/x-tar": "Archives",
    "application/gzip": "Archives",
}

ALL_CATEGORIES: list[str] = sorted(set(MIME_TO_CATEGORY.values()))

CATEGORY_ICONS: dict[str, str] = {
    "PDF Documents": "📄", "Word Documents": "📝", "Spreadsheets": "📊",
    "Presentations": "📽", "Images": "🖼", "Videos": "🎬",
    "Audio": "🎵", "Archives": "🗜",
}

# Extension-to-category fallback for files downloaded via OneDrive links
_EXT_CATEGORY: dict[str, str] = {
    "pdf": "PDF Documents",
    "doc": "Word Documents", "docx": "Word Documents", "odt": "Word Documents",
    "xls": "Spreadsheets", "xlsx": "Spreadsheets", "ods": "Spreadsheets",
    "ppt": "Presentations", "pptx": "Presentations", "odp": "Presentations",
    "jpg": "Images", "jpeg": "Images", "png": "Images", "gif": "Images",
    "webp": "Images", "bmp": "Images", "tiff": "Images", "tif": "Images",
    "heic": "Images", "heif": "Images",
    "mp4": "Videos", "mov": "Videos", "avi": "Videos", "mkv": "Videos",
    "wmv": "Videos", "webm": "Videos", "mpeg": "Videos", "mpg": "Videos",
    "mp3": "Audio", "wav": "Audio", "ogg": "Audio", "flac": "Audio",
    "aac": "Audio", "m4a": "Audio",
    "zip": "Archives", "rar": "Archives", "7z": "Archives",
    "tar": "Archives", "gz": "Archives",
}


def category_from_filename(filename: str) -> Optional[str]:
    return _EXT_CATEGORY.get(Path(filename).suffix.lower().lstrip("."))


# ── Sheets column headers ──────────────────────────────────────────────────
SHEET_HEADERS = [
    "Timestamp", "Sender Name", "Sender Email", "Email Date",
    "Email Subject", "Filename", "Category", "Source", "Storage Location",
]


# ── Persistent manifest ────────────────────────────────────────────────────

class DownloadManifest:
    """
    Tracks every saved file to prevent re-downloads.
    Key format: "storage_mode/sender_dir/YYYY-MM-DD/Category/msg_id"
    Files with the same name from different emails are kept (different msg_id → different key).
    Old keys without msg_id are still recognised by count_for().
    """

    def __init__(self, manifest_path: Path):
        self._path          = manifest_path
        self._lock          = threading.Lock()
        self._data:          dict[str, list[str]] = {}
        self._pending_saves = 0
        self._load()

    def _load(self):
        if self._path.exists():
            try:
                raw = json.loads(self._path.read_text(encoding="utf-8"))
                data = {k: list(v) for k, v in raw.get("files_by_category", {}).items()}
                # Migrate old keys that have no storage-mode prefix → assume "local/"
                known = ("local/", "onedrive/")
                self._data = {
                    k if any(k.startswith(p) for p in known) else f"local/{k}": v
                    for k, v in data.items()
                }
            except Exception:
                self._data = {}

    _SAVE_INTERVAL = 1  # write on every modification so a crash loses at most one entry

    def _save(self):
        payload = {
            "schema_version": 2,
            "last_updated": datetime.now().isoformat(),
            "files_by_category": self._data,
        }
        content = json.dumps(payload, indent=2, ensure_ascii=False)
        # Atomic write: temp file in same dir then rename so the manifest is
        # never left in a half-written state if the process is killed mid-write.
        tmp = self._path.with_suffix(".tmp")
        tmp.write_text(content, encoding="utf-8")
        tmp.replace(self._path)
        self._pending_saves = 0

    def _save_if_due(self):
        self._pending_saves += 1
        if self._pending_saves >= self._SAVE_INTERVAL:
            self._save()

    def flush(self):
        """Force-write any pending changes to disk (call at end of extraction)."""
        with self._lock:
            if self._pending_saves > 0:
                self._save()

    def clear(self):
        """Erase all download history so every file is treated as new on the next run."""
        with self._lock:
            self._data = {}
            self._save()

    def already_downloaded(self, key: str, filename: str) -> bool:
        with self._lock:
            return filename in self._data.get(key, [])

    def mark_downloaded(self, key: str, filename: str):
        with self._lock:
            self._data.setdefault(key, [])
            if filename not in self._data[key]:
                self._data[key].append(filename)
            self._save_if_due()

    def remove(self, key: str, filename: str):
        """Remove a filename from the manifest (used when a file is found missing)."""
        with self._lock:
            bucket = self._data.get(key, [])
            if filename in bucket:
                bucket.remove(filename)
                self._save_if_due()

    def claim(self, key: str, filename: str) -> bool:
        """Atomically check-and-mark. Returns True if this call should download the file."""
        with self._lock:
            if filename in self._data.get(key, []):
                return False
            self._data.setdefault(key, [])
            self._data[key].append(filename)
            self._save_if_due()
            return True

    def total_count(self) -> int:
        with self._lock:
            return sum(len(v) for v in self._data.values())

    def count_for(self, category: str) -> int:
        """Count across all sender/date/msg combos for a given category."""
        with self._lock:
            return sum(
                len(files) for key, files in self._data.items()
                if key == category
                or key.endswith(f"/{category}")        # old schema: category at end
                or f"/{category}/" in key              # new schema: category before msg_id
            )


# ── Helpers ────────────────────────────────────────────────────────────────

def sanitise_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "unnamed_file"


def sanitise_dirname(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip()
    return cleaned[:80] or "unknown"


def unique_path(directory: Path, filename: str) -> Path:
    target = directory / filename
    if not target.exists():
        return target
    stem, suffix = Path(filename).stem, Path(filename).suffix
    for i in range(1, 99_999):
        c = directory / f"{stem}_{i}{suffix}"
        if not c.exists():
            return c
    return target


def walk_parts(parts: list[dict]):
    for part in parts:
        sub = part.get("parts")
        if sub:
            yield from walk_parts(sub)
        else:
            yield part


def parse_sender(from_header: str) -> tuple[str, str]:
    """Parse 'Name <email>' → (display_name, email_addr)."""
    m = re.match(r'"?([^"<]*)"?\s*<([^>]+)>', from_header.strip())
    if m:
        name = m.group(1).strip().strip('"') or m.group(2).strip()
        return name, m.group(2).strip().lower()
    addr = from_header.strip().lower()
    return addr, addr


def parse_email_date(date_header: str) -> str:
    """Parse RFC-2822 date → 'YYYY-MM-DD'."""
    try:
        return parsedate_to_datetime(date_header).strftime("%Y-%m-%d")
    except Exception:
        return "unknown-date"


def extract_body_text(msg: dict) -> str:
    """Return all decoded text (plain + html) from a message."""
    payload = msg.get("payload", {})
    parts = payload.get("parts") or [payload]
    chunks: list[str] = []
    for part in walk_parts(parts):
        if part.get("mimeType", "") in ("text/plain", "text/html"):
            data = part.get("body", {}).get("data", "")
            if data:
                try:
                    chunks.append(
                        base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="ignore")
                    )
                except Exception:
                    pass
    return " ".join(chunks)


_ONEDRIVE_RE = re.compile(
    r'https?://(?:'
    r'1drv\.ms/[^\s<>"\')\]]+|'
    r'onedrive\.live\.com/[^\s<>"\')\]]+|'
    r'[\w-]+-my\.sharepoint\.com[^\s<>"\')\]]+|'
    r'[\w-]+\.sharepoint\.com/[^\s<>"\')\]]+'
    r')',
    re.IGNORECASE,
)


def find_onedrive_links(text: str) -> list[str]:
    seen: dict[str, None] = {}
    for url in _ONEDRIVE_RE.findall(text):
        seen[url] = None
    return list(seen)


def download_onedrive_link(
    sharing_url: str,
    ms_token: Optional[str] = None,
) -> Optional[tuple[bytes, str]]:
    """
    Download a file from an OneDrive sharing URL.
    Uses auth token if provided (handles private shared-with-me links).
    Falls back to unauthenticated for fully public links.
    Returns (file_bytes, filename) or None.
    """
    if not REQUESTS_OK:
        return None

    encoded = base64.urlsafe_b64encode(sharing_url.encode()).decode().rstrip("=")
    share_token = f"u!{encoded}"

    headers = {}
    if ms_token:
        headers["Authorization"] = f"Bearer {ms_token}"

    resp = _req_retry(lambda: _req.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_token}/driveItem",
        headers=headers,
        timeout=30,
    ))
    if not resp.ok:
        resp = _req_retry(lambda: _req.get(
            f"https://api.onedrive.com/v1.0/shares/{share_token}/driveItem",
            timeout=30,
        ))
    if not resp.ok:
        return None

    item = resp.json()
    filename = item.get("name", "onedrive_file")
    dl_url = item.get("@microsoft.graph.downloadUrl") or item.get("@content.downloadUrl")
    if not dl_url:
        return None

    file_resp = _req_retry(lambda u=dl_url: _req.get(u, timeout=120))
    if file_resp.ok:
        return file_resp.content, filename
    return None


# ── OneDrive upload client ─────────────────────────────────────────────────

class OneDriveClient:
    """Microsoft Graph API client for upload and shared-link access."""

    GRAPH = "https://graph.microsoft.com/v1.0"
    MS_SCOPES = ["Files.ReadWrite", "User.Read"]

    def __init__(self, token_cache_path: Path):
        if not MSAL_OK:
            raise RuntimeError("msal package is not installed.")
        if not ONEDRIVE_CONFIGURED:
            raise RuntimeError(
                "Microsoft client ID is not configured.\n"
                "Replace MICROSOFT_CLIENT_ID in core.py with your Azure App Registration client ID."
            )
        self._cache_path = token_cache_path
        self._cache = msal.SerializableTokenCache()
        if token_cache_path.exists():
            self._cache.deserialize(token_cache_path.read_text(encoding="utf-8"))
        self._app = msal.PublicClientApplication(
            MICROSOFT_CLIENT_ID,
            authority="https://login.microsoftonline.com/aab598e8-253c-409a-86c2-e136f77ea34d",
            token_cache=self._cache,
        )
        self._token: Optional[str] = None
        self._token_lock = threading.Lock()

    def _save_cache(self):
        if self._cache.has_state_changed:
            self._cache_path.write_text(self._cache.serialize(), encoding="utf-8")

    @property
    def token(self) -> Optional[str]:
        return self._token

    def is_authenticated(self) -> bool:
        return self._token is not None

    def authenticate(self) -> bool:
        """Try silent auth first; fall back to interactive browser flow."""
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(self.MS_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._token = result["access_token"]
                self._save_cache()
                return True
        result = self._app.acquire_token_interactive(scopes=self.MS_SCOPES)
        if "access_token" in result:
            self._token = result["access_token"]
            self._save_cache()
            return True
        return False

    def sign_out(self):
        for acct in self._app.get_accounts():
            self._app.remove_account(acct)
        self._token = None
        if self._cache_path.exists():
            self._cache_path.unlink()

    def file_exists(self, drive_path: str) -> bool:
        """Return True if the file exists at drive_path on the user's OneDrive."""
        self._refresh_token()
        resp = _req_retry(lambda p=drive_path: _req.get(
            f"{self.GRAPH}/me/drive/root:/{p}",
            headers=self._headers(),
            timeout=15,
        ))
        return resp.status_code == 200

    def _refresh_token(self):
        """Silently refresh the access token using the cached refresh token (thread-safe)."""
        with self._token_lock:
            accounts = self._app.get_accounts()
            if not accounts:
                return
            result = self._app.acquire_token_silent(self.MS_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._token = result["access_token"]
                self._save_cache()

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._token}"}

    def upload(self, file_bytes: bytes, drive_path: str) -> str:
        """
        Upload bytes to /me/drive/root:/{drive_path}:/content.
        Uses simple PUT ≤ 4 MB, upload session for larger files.
        Returns webUrl on success; raises RuntimeError on failure.
        """
        self._refresh_token()
        url = f"{self.GRAPH}/me/drive/root:/{drive_path}:/content"
        if len(file_bytes) <= 4 * 1024 * 1024:
            resp = _req_retry(lambda: _req.put(
                url,
                headers={**self._headers(), "Content-Type": "application/octet-stream"},
                data=file_bytes,
                timeout=60,
            ))
            if not resp.ok:
                raise RuntimeError(
                    f"OneDrive upload failed ({resp.status_code}): {drive_path}"
                )
            return resp.json().get("webUrl", drive_path)

        # Large file: resumable upload session — restart the whole session on failure
        for _attempt in range(4):
            try:
                sess = _req_retry(lambda: _req.post(
                    f"{self.GRAPH}/me/drive/root:/{drive_path}:/createUploadSession",
                    headers=self._headers(),
                    json={"item": {"@microsoft.graph.conflictBehavior": "rename"}},
                    timeout=30,
                ))
                if not sess.ok:
                    raise RuntimeError(
                        f"OneDrive upload session failed ({sess.status_code}): {drive_path}"
                    )
                upload_url = sess.json()["uploadUrl"]
                chunk_size = 10 * 1024 * 1024
                total = len(file_bytes)
                resp = None
                for start in range(0, total, chunk_size):
                    chunk = file_bytes[start : start + chunk_size]
                    end = start + len(chunk) - 1
                    resp = _req_retry(lambda c=chunk, s=start, e=end, u=upload_url: _req.put(
                        u, data=c,
                        headers={"Content-Range": f"bytes {s}-{e}/{total}"},
                        timeout=120,
                    ))
                    if resp.status_code not in (200, 201, 202):
                        raise RuntimeError(
                            f"OneDrive chunk upload failed ({resp.status_code}): {drive_path}"
                        )
                return resp.json().get("webUrl", drive_path) if resp else drive_path
            except RuntimeError:
                if _attempt < 3:
                    time.sleep(2 ** _attempt)
                    self._refresh_token()
                    continue
                raise


# ── Google Sheets logger ───────────────────────────────────────────────────

class SheetsLogger:
    """Auto-creates a log spreadsheet and appends one row per saved file."""

    SHEET_TITLE = "Gmail Media Extractor Log"

    def __init__(self, creds: "Credentials", sheet_id_path: Path):
        self._svc = _build_service("sheets", "v4", credentials=creds)
        self._id_path = sheet_id_path
        self._sheet_id: Optional[str] = None
        self._pending: list[list] = []
        self._lock = threading.Lock()

    @property
    def sheet_url(self) -> str:
        if self._sheet_id:
            return f"https://docs.google.com/spreadsheets/d/{self._sheet_id}/edit"
        return ""

    def initialize(self):
        """Locate or create the log spreadsheet; write headers on first creation."""
        if self._id_path.exists():
            self._sheet_id = self._id_path.read_text(encoding="utf-8").strip()
            return
        ss = self._svc.spreadsheets().create(body={
            "properties": {"title": self.SHEET_TITLE},
            "sheets": [{"properties": {"title": "Log"}}],
        }).execute()
        self._sheet_id = ss["spreadsheetId"]
        self._svc.spreadsheets().values().append(
            spreadsheetId=self._sheet_id,
            range="Log!A1",
            valueInputOption="RAW",
            body={"values": [SHEET_HEADERS]},
        ).execute()
        self._id_path.write_text(self._sheet_id, encoding="utf-8")

    def log(
        self,
        sender_name: str,
        sender_email: str,
        email_date: str,
        subject: str,
        filename: str,
        category: str,
        source: str,
        storage: str,
    ):
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            sender_name, sender_email, email_date,
            subject, filename, category, source, storage,
        ]
        with self._lock:
            self._pending.append(row)

    def flush(self):
        """Write all buffered rows to the sheet in one API call."""
        with self._lock:
            if not self._pending or not self._sheet_id:
                return
            rows, self._pending = self._pending[:], []
        def _do():
            with _http_exec_lock:   # Sheets API also uses httplib2 — must serialize with Gmail calls
                self._svc.spreadsheets().values().append(
                    spreadsheetId=self._sheet_id,
                    range="Log!A1",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": rows},
                ).execute()
        try:
            _gapi_retry(_do)
        except Exception:
            pass  # Sheets errors must not abort extraction


# ── Worker config & engine ─────────────────────────────────────────────────

@dataclass
class WorkerConfig:
    output_dir:      Path           # local save root (also holds JSON report in OneDrive mode)
    categories:      set[str]
    after:           str            # YYYY/MM/DD or ""
    before:          str            # YYYY/MM/DD or ""
    creds_path:      Path
    token_path:      Path
    manifest:        DownloadManifest
    msg_queue:       queue.Queue
    storage_mode:    str = "local"  # "local" | "onedrive"
    onedrive_client: Optional[OneDriveClient] = None
    enable_sheets:   bool = True
    sheet_id_path:   Optional[Path] = None
    scan_links:      bool = True    # scan email body for OneDrive links
    override_ids:    Optional[list] = None  # if set, skip _list_ids and process these msg IDs only
    email_direction: str = "both"   # "both" | "received" | "sent"


class ExtractionWorker:
    """Runs entirely in a background thread; communicates with the UI via msg_queue."""

    def __init__(self, cfg: WorkerConfig):
        self.cfg   = cfg
        self._stop = threading.Event()

    def stop(self):
        self._stop.set()

    def _emit(self, kind: str, **kwargs):
        self.cfg.msg_queue.put({"kind": kind, **kwargs})

    # ── Entry point ────────────────────────────────────────────────────────

    def run(self):
        if not GOOGLE_OK:
            self._emit("error", text="Google client libraries are missing. Please reinstall.")
            return

        self._emit("status", text="🔐  Opening Google sign-in in your browser…")
        try:
            creds = self._authenticate()
        except FileNotFoundError as exc:
            self._emit("error", text=str(exc))
            return
        except Exception as exc:
            self._emit("error", text=f"Sign-in failed:\n{exc}")
            return

        gmail = _build_service("gmail", "v1", credentials=creds)

        # Set up Sheets logger
        sheets: Optional[SheetsLogger] = None
        if self.cfg.enable_sheets and self.cfg.sheet_id_path:
            try:
                sheets = SheetsLogger(creds, self.cfg.sheet_id_path)
                sheets.initialize()
                self._emit("sheet_url", url=sheets.sheet_url)
            except Exception as exc:
                self._emit("status", text=f"⚠️  Sheets logging unavailable: {exc}")

        self._emit("status", text="✅  Signed in! Counting emails…")

        qparts = ["has:attachment"]
        if self.cfg.after:
            qparts.append(f"after:{self.cfg.after}")
        if self.cfg.before:
            qparts.append(f"before:{self.cfg.before}")
        if self.cfg.email_direction == "received":
            qparts.append("-in:sent")
        elif self.cfg.email_direction == "sent":
            qparts.append("in:sent")

        if self.cfg.override_ids is not None:
            ids   = self.cfg.override_ids
            total = len(ids)
            self._emit("total", count=total)
            self._emit("status", text=f"🔁  Retrying {total:,} previously failed item(s)…")
        else:
            try:
                ids = self._list_ids(gmail, " ".join(qparts))
            except Exception as exc:
                self._emit("error", text=f"Failed to list emails:\n{exc}")
                return
            total = len(ids)
            self._emit("total", count=total)
            self._emit("status", text=f"📬  {total:,} emails to scan. Starting download…")

        active_mime  = {m for m, cat in MIME_TO_CATEGORY.items() if cat in self.cfg.categories}
        state_lock   = threading.Lock()
        state        = {"downloaded": 0, "skipped": 0, "errors": 0}
        records: list[dict] = []
        records_lock = threading.Lock()
        failures: list[dict] = []
        failures_lock = threading.Lock()

        def _submit(msg_id: str):
            self._process_one(msg_id, creds, active_mime, sheets,
                              state, state_lock, records, records_lock,
                              failures, failures_lock)

        done = 0
        try:
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = [executor.submit(_submit, mid) for mid in ids]
                for future in as_completed(futures):
                    done += 1
                    self._emit("progress", current=done, total=total)
                    try:
                        future.result()
                    except Exception:
                        with state_lock:
                            state["errors"] += 1
                    if self._stop.is_set():
                        self._emit("status", text="⏹  Cancelled.")
                        for f in futures:
                            f.cancel()
                        break
        finally:
            # Always persist the manifest — even on unexpected exceptions or
            # user-initiated cancellation — so the next run can resume cleanly.
            if sheets:
                try:
                    sheets.flush()
                except Exception:
                    pass
            self.cfg.manifest.flush()

        downloaded       = state["downloaded"]
        skipped_manifest = state["skipped"]
        errors           = state["errors"]

        report = {
            "generated_at": datetime.now().isoformat(),
            "stats": {
                "emails_scanned": total,
                "downloaded": downloaded,
                "skipped_already_downloaded": skipped_manifest,
                "errors": errors,
            },
            "failures": failures,
            "files": records,
        }
        self.cfg.output_dir.mkdir(parents=True, exist_ok=True)
        rp = self.cfg.output_dir / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        rp.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

        self._emit("done",
                   downloaded=downloaded,
                   skipped=skipped_manifest,
                   errors=errors,
                   output_dir=str(self.cfg.output_dir),
                   storage_mode=self.cfg.storage_mode)

    # ── Helpers ────────────────────────────────────────────────────────────

    def _get_gmail(self, creds):
        """Return a thread-local Gmail service (one instance per worker thread)."""
        if not hasattr(_thread_local, "gmail"):
            with _build_svc_lock:   # serialize httplib2 SSL init across threads
                _thread_local.gmail = _build_service("gmail", "v1", credentials=creds)
        return _thread_local.gmail

    def _process_one(
        self,
        msg_id: str,
        creds,
        active_mime: set,
        sheets: Optional[SheetsLogger],
        state: dict,
        state_lock: threading.Lock,
        records: list,
        records_lock: threading.Lock,
        failures: list,
        failures_lock: threading.Lock,
    ):
        """Fetch and process a single email. Runs inside a thread-pool worker."""
        svc = self._get_gmail(creds)

        def _fetch_msg():
            with _http_exec_lock:
                return svc.users().messages().get(
                    userId="me", id=msg_id, format="full"
                ).execute()
        try:
            msg = _gapi_retry(_fetch_msg)
        except Exception as exc:
            with failures_lock:
                failures.append({"msg_id": msg_id, "step": "fetch_message", "reason": str(exc)})
            self._emit("status", text=f"⚠️  Could not fetch email {msg_id}: {exc}")
            with state_lock:
                state["errors"] += 1
            return

        hdrs         = {h["name"]: h["value"]
                        for h in msg.get("payload", {}).get("headers", [])}
        sender_name, sender_email = parse_sender(hdrs.get("From", "unknown"))
        subject      = hdrs.get("Subject", "(no subject)")
        email_date   = parse_email_date(hdrs.get("Date", ""))
        sender_dir   = sanitise_dirname(sender_email or sender_name)

        # ── Direct attachments ─────────────────────────────────────────────
        payload = msg.get("payload", {})
        parts   = payload.get("parts", [])
        if not parts and payload.get("filename"):
            parts = [payload]

        for part in walk_parts(parts):
            if self._stop.is_set():
                return
            filename  = (part.get("filename") or "").strip()
            mime_type = part.get("mimeType", "").lower().split(";")[0].strip()
            att_id    = part.get("body", {}).get("attachmentId")

            if not att_id or mime_type not in active_mime:
                continue

            category = MIME_TO_CATEGORY[mime_type]
            if not filename:
                ext      = mime_type.split("/")[-1].split("+")[0]
                filename = f"attachment_{msg_id[:8]}.{ext}"
            safe_name    = sanitise_filename(filename)
            manifest_key = f"{self.cfg.storage_mode}/{sender_dir}/{email_date}/{category}/{msg_id}"

            if not self.cfg.manifest.claim(manifest_key, safe_name):
                if self.cfg.storage_mode == "onedrive" and self.cfg.onedrive_client:
                    od_path = f"GmailMedia/{sender_dir}/{email_date}/{category}/{safe_name}"
                    try:
                        exists = self.cfg.onedrive_client.file_exists(od_path)
                    except Exception:
                        exists = False  # can't confirm → re-download to be safe
                    if exists:
                        with state_lock:
                            state["skipped"] += 1
                        self._emit("skipped", name=safe_name, reason="already on OneDrive")
                        continue
                    # File is gone from OneDrive — remove from manifest and re-download
                    self.cfg.manifest.remove(manifest_key, safe_name)
                    self.cfg.manifest.claim(manifest_key, safe_name)
                else:
                    local_path = self.cfg.output_dir / sender_dir / email_date / category / safe_name
                    if local_path.exists():
                        with state_lock:
                            state["skipped"] += 1
                        self._emit("skipped", name=safe_name, reason="already downloaded")
                        continue
                    # File was deleted locally — remove from manifest and re-download
                    self.cfg.manifest.remove(manifest_key, safe_name)
                    self.cfg.manifest.claim(manifest_key, safe_name)

            def _fetch_att(aid=att_id):
                with _http_exec_lock:
                    return svc.users().messages().attachments().get(
                        userId="me", messageId=msg_id, id=aid
                    ).execute()
            try:
                att = _gapi_retry(_fetch_att)
            except Exception as exc:
                with failures_lock:
                    failures.append({
                        "msg_id": msg_id, "step": "fetch_attachment",
                        "filename": safe_name, "reason": str(exc),
                    })
                self._emit("status", text=f"⚠️  Could not fetch attachment {safe_name}: {exc}")
                with state_lock:
                    state["errors"] += 1
                continue
            raw = base64.urlsafe_b64decode(att.get("data", "") + "==")

            try:
                storage_loc = self._store(
                    raw, safe_name, sender_dir, email_date, category,
                    sender_name, sender_email, subject, "Direct Attachment", sheets,
                )
            except RuntimeError as exc:
                with failures_lock:
                    failures.append({
                        "msg_id": msg_id, "step": "store",
                        "filename": safe_name, "reason": str(exc),
                    })
                self._emit("status", text=f"⚠️  {exc}")
                with state_lock:
                    state["errors"] += 1
                continue
            with state_lock:
                state["downloaded"] += 1
                count = state["downloaded"]
            with records_lock:
                records.append({
                    "message_id": msg_id, "sender_email": sender_email,
                    "subject": subject, "date": email_date,
                    "filename": safe_name, "category": category,
                    "source": "attachment", "storage": storage_loc,
                })
            self._emit("file_done", name=safe_name, category=category, count=count)

        # ── OneDrive links in email body ───────────────────────────────────
        if self.cfg.scan_links and REQUESTS_OK:
            body_text = extract_body_text(msg)
            ms_token  = self.cfg.onedrive_client.token if self.cfg.onedrive_client else None

            for link_url in find_onedrive_links(body_text):
                if self._stop.is_set():
                    return
                try:
                    result = download_onedrive_link(link_url, ms_token)
                except Exception as exc:
                    with failures_lock:
                        failures.append({
                            "msg_id": msg_id, "step": "onedrive_link",
                            "url": link_url, "reason": str(exc),
                        })
                    self._emit("status", text=f"⚠️  OneDrive link failed after retries: {exc}")
                    with state_lock:
                        state["errors"] += 1
                    continue
                if result is None:
                    continue

                link_bytes, link_filename = result
                if not link_filename:
                    continue

                link_cat = category_from_filename(link_filename)
                if not link_cat or link_cat not in self.cfg.categories:
                    continue

                safe_name    = sanitise_filename(link_filename)
                manifest_key = f"{sender_dir}/{email_date}/{link_cat}/{msg_id}"

                if not self.cfg.manifest.claim(manifest_key, safe_name):
                    if self.cfg.storage_mode == "onedrive" and self.cfg.onedrive_client:
                        od_path = f"GmailMedia/{sender_dir}/{email_date}/{link_cat}/{safe_name}"
                        try:
                            exists = self.cfg.onedrive_client.file_exists(od_path)
                        except Exception:
                            exists = False  # can't confirm → re-download to be safe
                        if exists:
                            with state_lock:
                                state["skipped"] += 1
                            continue
                        self.cfg.manifest.remove(manifest_key, safe_name)
                        self.cfg.manifest.claim(manifest_key, safe_name)
                    else:
                        local_path = self.cfg.output_dir / sender_dir / email_date / link_cat / safe_name
                        if local_path.exists():
                            with state_lock:
                                state["skipped"] += 1
                            continue
                        self.cfg.manifest.remove(manifest_key, safe_name)
                        self.cfg.manifest.claim(manifest_key, safe_name)

                try:
                    storage_loc = self._store(
                        link_bytes, safe_name, sender_dir, email_date, link_cat,
                        sender_name, sender_email, subject, "OneDrive Link", sheets,
                    )
                except RuntimeError as exc:
                    with failures_lock:
                        failures.append({
                            "msg_id": msg_id, "step": "store_onedrive_link",
                            "filename": safe_name, "url": link_url, "reason": str(exc),
                        })
                    self._emit("status", text=f"⚠️  {exc}")
                    with state_lock:
                        state["errors"] += 1
                    continue
                with state_lock:
                    state["downloaded"] += 1
                    count = state["downloaded"]
                with records_lock:
                    records.append({
                        "message_id": msg_id, "sender_email": sender_email,
                        "subject": subject, "date": email_date,
                        "filename": safe_name, "category": link_cat,
                        "source": link_url, "storage": storage_loc,
                    })
                self._emit("file_done", name=safe_name, category=link_cat, count=count)

        if sheets:
            sheets.flush()

    def _store(
        self,
        data: bytes,
        filename: str,
        sender_dir: str,
        email_date: str,
        category: str,
        sender_name: str,
        sender_email: str,
        subject: str,
        source: str,
        sheets: Optional[SheetsLogger],
    ) -> str:
        """Save bytes to local disk or OneDrive. Returns storage location string."""
        if self.cfg.storage_mode == "onedrive" and self.cfg.onedrive_client:
            drive_path = f"GmailMedia/{sender_dir}/{email_date}/{category}/{filename}"
            web_url    = self.cfg.onedrive_client.upload(data, drive_path)
            if sheets:
                sheets.log(sender_name, sender_email, email_date,
                           subject, filename, category, source, web_url)
            return web_url

        target_dir = self.cfg.output_dir / sender_dir / email_date / category
        target_dir.mkdir(parents=True, exist_ok=True)
        out_path = unique_path(target_dir, filename)
        out_path.write_bytes(data)
        if sheets:
            sheets.log(sender_name, sender_email, email_date,
                       subject, filename, category, source, str(out_path))
        return str(out_path)

    def _authenticate(self) -> "Credentials":
        creds = None
        if self.cfg.token_path.exists():
            creds = Credentials.from_authorized_user_file(str(self.cfg.token_path), SCOPES)

        # Force re-auth if existing token is missing a required scope
        if creds and creds.valid and creds.scopes:
            if not set(SCOPES).issubset(creds.scopes):
                creds = None
                self.cfg.token_path.unlink(missing_ok=True)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    if creds.scopes and not set(SCOPES).issubset(creds.scopes):
                        creds = None
                        self.cfg.token_path.unlink(missing_ok=True)
                    else:
                        self.cfg.token_path.write_text(creds.to_json(), encoding="utf-8")
                        return creds
                except Exception:
                    creds = None

        if not creds or not creds.valid:
            if not self.cfg.creds_path.exists():
                raise FileNotFoundError(
                    "credentials.json not found.\n\n"
                    "Please place credentials.json in the same folder as this app\n"
                    "and restart. See Help → Setup Guide for instructions."
                )
            flow  = InstalledAppFlow.from_client_secrets_file(str(self.cfg.creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
            self.cfg.token_path.write_text(creds.to_json(), encoding="utf-8")

        return creds

    def _list_ids(self, service, query: str) -> list[str]:
        ids, page_token = [], None
        while True:
            kwargs = {"userId": "me", "q": query, "maxResults": 500}
            if page_token:
                kwargs["pageToken"] = page_token
            res = _gapi_retry(lambda kw=kwargs: service.users().messages().list(**kw).execute())
            ids       += [m["id"] for m in res.get("messages", [])]
            page_token = res.get("nextPageToken")
            if not page_token:
                break
        return ids

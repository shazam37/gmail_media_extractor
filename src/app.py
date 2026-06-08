"""
app.py — Gmail Media Extractor
Beautiful desktop UI built with customtkinter.
"""
from __future__ import annotations

import os
import sys
import queue
import datetime
import threading
import subprocess
import webbrowser
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from core import (
    ALL_CATEGORIES, CATEGORY_ICONS, ExtractionWorker,
    DownloadManifest, WorkerConfig,
    OneDriveClient, MSAL_OK, ONEDRIVE_CONFIGURED, REQUESTS_OK,
)

# Python 3.14 changed tkinter so event.widget can be a string name instead of
# a widget object, which breaks customtkinter's scroll handler.
import customtkinter.windows.widgets.ctk_scrollable_frame as _sf
_orig_check = _sf.CTkScrollableFrame.check_if_master_is_canvas
def _safe_check(self, widget):
    try:
        return _orig_check(self, widget)
    except AttributeError:
        return False
_sf.CTkScrollableFrame.check_if_master_is_canvas = _safe_check

# ── Theme ──────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Paths ──────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    APP_DIR    = Path(sys.executable).parent
    DATA_DIR   = Path(os.environ.get("APPDATA", Path.home())) / "GmailMediaExtractor"
    CREDS_PATH = Path(sys._MEIPASS) / "credentials.json"
else:
    APP_DIR    = Path(__file__).parent
    DATA_DIR   = APP_DIR / ".app_data"
    CREDS_PATH = APP_DIR / "credentials.json"

DATA_DIR.mkdir(parents=True, exist_ok=True)
TOKEN_PATH         = DATA_DIR / "token.json"
MANIFEST_PATH      = DATA_DIR / "download_manifest.json"
SHEET_ID_PATH      = DATA_DIR / "sheets_id.txt"
MS_TOKEN_CACHE     = DATA_DIR / "ms_token_cache.json"
IS_FROZEN          = getattr(sys, "frozen", False)

# ── Colour palette ─────────────────────────────────────────────────────────
C = {
    "bg":         "#0d0f1a",
    "surface":    "#141728",
    "surface2":   "#1c2035",
    "border":     "#252a45",
    "accent":     "#4f6ef7",
    "accent_h":   "#6b84ff",
    "accent_dim": "#2a3a8a",
    "green":      "#1e7a3e",
    "green_h":    "#25a050",
    "muted":      "#6b7280",
    "text":       "#e2e8f0",
    "text_dim":   "#94a3b8",
    "danger":     "#7f1d1d",
    "ms_blue":    "#0078d4",
    "ms_blue_h":  "#106ebe",
}


def open_folder(path: Path):
    if sys.platform == "win32":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


# ══════════════════════════════════════════════════════════════════════════════
#  Reusable UI widgets
# ══════════════════════════════════════════════════════════════════════════════

class SectionHeader(ctk.CTkLabel):
    def __init__(self, parent, text: str, **kwargs):
        super().__init__(
            parent, text=text,
            font=ctk.CTkFont(family="Helvetica", size=22, weight="bold"),
            text_color="#94a3b8",
            **kwargs,
        )


class Card(ctk.CTkFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            fg_color=C["surface"],
            corner_radius=12,
            border_width=1,
            border_color=C["border"],
            **kwargs,
        )


class PrimaryButton(ctk.CTkButton):
    def __init__(self, parent, **kwargs):
        kwargs.setdefault("height", 58)
        super().__init__(
            parent,
            fg_color=C["accent"],
            hover_color=C["accent_h"],
            corner_radius=10,
            font=ctk.CTkFont(size=20, weight="bold"),
            **kwargs,
        )


class GhostButton(ctk.CTkButton):
    def __init__(self, parent, **kwargs):
        kwargs.setdefault("height", 58)
        super().__init__(
            parent,
            fg_color=C["surface2"],
            hover_color=C["border"],
            corner_radius=10,
            font=ctk.CTkFont(size=19),
            **kwargs,
        )


class DateSelector(ctk.CTkFrame):
    """Three linked dropdowns (Year / Month / Day) that emit a YYYY/MM/DD string."""

    MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)

        dropdown_kw = dict(
            fg_color=C["surface2"],
            button_color=C["border"],
            button_hover_color=C["accent"],
            dropdown_fg_color=C["surface"],
            dropdown_hover_color=C["accent_dim"],
            text_color=C["text"],
            font=ctk.CTkFont(size=19),
            dropdown_font=ctk.CTkFont(size=18),
            corner_radius=8,
            height=54,
        )

        current_year = datetime.date.today().year
        years  = ["Year"]  + [str(y) for y in range(current_year, 2009, -1)]
        months = ["Month"] + self.MONTHS
        days   = ["Day"]   + [str(d) for d in range(1, 32)]

        self._year_var  = ctk.StringVar(value="Year")
        self._month_var = ctk.StringVar(value="Month")
        self._day_var   = ctk.StringVar(value="Day")

        ctk.CTkOptionMenu(self, variable=self._year_var,  values=years,
                          width=120, **dropdown_kw).pack(side="left", padx=(0, 8))
        ctk.CTkOptionMenu(self, variable=self._month_var, values=months,
                          width=128, **dropdown_kw).pack(side="left", padx=(0, 8))
        ctk.CTkOptionMenu(self, variable=self._day_var,   values=days,
                          width=105, **dropdown_kw).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            self, text="✕", width=48, height=54, corner_radius=8,
            fg_color=C["surface2"], hover_color=C["danger"],
            text_color=C["text_dim"], font=ctk.CTkFont(size=19),
            command=self._clear,
        ).pack(side="left")

    def _clear(self):
        self._year_var.set("Year")
        self._month_var.set("Month")
        self._day_var.set("Day")

    def get(self) -> str:
        """Returns YYYY/MM/DD or '' when no year is chosen."""
        y = self._year_var.get()
        if y == "Year":
            return ""
        m = self._month_var.get()
        d = self._day_var.get()
        month_num = (self.MONTHS.index(m) + 1) if m != "Month" else 1
        day_num   = int(d) if d != "Day" else 1
        return f"{y}/{month_num:02d}/{day_num:02d}"


# ══════════════════════════════════════════════════════════════════════════════
#  Main application
# ══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Gmail Media Extractor")
        self.geometry("860x1220")
        self.minsize(820, 1080)
        self.configure(fg_color=C["bg"])

        self._worker:      ExtractionWorker | None = None
        self._thread:      threading.Thread | None = None
        self._q:           queue.Queue = queue.Queue()
        self._total        = 0
        self._manifest     = DownloadManifest(MANIFEST_PATH)
        self._output_dir:  Path | None = None
        self._ms_client:   OneDriveClient | None = None
        self._sheets_url:  str | None = None

        self._build()
        self._poll()

    # ══════════════════════════════════════════════════════════════════════
    #  UI construction
    # ══════════════════════════════════════════════════════════════════════

    def _build(self):
        # ── Top bar ───────────────────────────────────────────────────────
        bar = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=96)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        ctk.CTkLabel(
            bar, text="✉   Gmail Media Extractor",
            font=ctk.CTkFont(family="Helvetica", size=32, weight="bold"),
            text_color=C["text"],
        ).pack(side="left", padx=28, pady=20)

        self._badge_var = ctk.StringVar(value=self._badge_text())
        self._badge = ctk.CTkLabel(
            bar, textvariable=self._badge_var,
            font=ctk.CTkFont(size=17),
            text_color=C["accent"],
            fg_color=C["accent_dim"],
            corner_radius=8,
        )
        self._badge.pack(side="right", padx=(4, 28), ipadx=14, ipady=8)

        ctk.CTkButton(
            bar, text="Clear history",
            font=ctk.CTkFont(size=15),
            height=40, corner_radius=8,
            fg_color=C["surface2"], hover_color=C["border"],
            text_color=C["text_dim"], border_width=0,
            command=self._clear_history,
        ).pack(side="right", padx=(0, 6), pady=16)

        # ── Scrollable body ───────────────────────────────────────────────
        body = ctk.CTkScrollableFrame(self, fg_color=C["bg"], corner_radius=0)
        body.pack(fill="both", expand=True)

        self._build_date_section(body)
        self._build_source_section(body)
        self._build_type_section(body)
        self._build_storage_section(body)
        self._build_options_section(body)
        self._build_hint_section(body)


        # ── Footer ────────────────────────────────────────────────────────
        self._build_footer()

    def _build_date_section(self, parent):
        SectionHeader(parent, "DATE RANGE").pack(anchor="w", padx=24, pady=(24, 10))

        card = Card(parent)
        card.pack(fill="x", padx=24, pady=(0, 4))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=20)

        label_kw = dict(font=ctk.CTkFont(size=19), text_color=C["text_dim"], width=66)

        from_row = ctk.CTkFrame(inner, fg_color="transparent")
        from_row.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(from_row, text="From", **label_kw).pack(side="left", padx=(0, 12))
        self._after_sel = DateSelector(from_row)
        self._after_sel.pack(side="left")

        to_row = ctk.CTkFrame(inner, fg_color="transparent")
        to_row.pack(fill="x")
        ctk.CTkLabel(to_row, text="To", **label_kw).pack(side="left", padx=(0, 12))
        self._before_sel = DateSelector(to_row)
        self._before_sel.pack(side="left")

    def _build_source_section(self, parent):
        SectionHeader(parent, "EMAIL SOURCE").pack(anchor="w", padx=24, pady=(24, 10))

        card = Card(parent)
        card.pack(fill="x", padx=24, pady=(0, 4))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=18)

        self._direction_var = ctk.StringVar(value="both")

        radio_kw = dict(
            variable=self._direction_var,
            font=ctk.CTkFont(size=19),
            text_color=C["text"],
            fg_color=C["accent"],
            hover_color=C["accent_h"],
            border_color=C["border"],
        )

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")

        ctk.CTkRadioButton(
            row, text="  📨  Both (received & sent)",
            value="both", **radio_kw,
        ).pack(side="left", padx=(0, 28))

        ctk.CTkRadioButton(
            row, text="  📥  Received only",
            value="received", **radio_kw,
        ).pack(side="left", padx=(0, 28))

        ctk.CTkRadioButton(
            row, text="  📤  Sent only",
            value="sent", **radio_kw,
        ).pack(side="left")

        ctk.CTkLabel(
            inner,
            text="Choose which emails to scan for attachments and shared files.",
            font=ctk.CTkFont(size=16), text_color=C["muted"],
        ).pack(anchor="w", pady=(12, 0))

    def _build_type_section(self, parent):
        SectionHeader(parent, "WHAT TO DOWNLOAD").pack(anchor="w", padx=24, pady=(24, 10))

        card = Card(parent)
        card.pack(fill="x", padx=24, pady=(0, 4))

        top_row = ctk.CTkFrame(card, fg_color="transparent")
        top_row.pack(fill="x", padx=18, pady=(18, 12))

        self._all_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            top_row, text="  Select All",
            variable=self._all_var,
            command=self._toggle_all,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=C["text"],
            fg_color=C["accent"], hover_color=C["accent_h"],
            checkmark_color="white", corner_radius=5,
        ).pack(side="left")

        ctk.CTkFrame(card, height=1, fg_color=C["border"]).pack(fill="x", padx=18)

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=18, pady=(12, 20))
        grid.columnconfigure((0, 1), weight=1)

        self._cat_vars: dict[str, ctk.BooleanVar] = {}
        for i, cat in enumerate(ALL_CATEGORIES):
            var = ctk.BooleanVar(value=True)
            self._cat_vars[cat] = var
            icon  = CATEGORY_ICONS.get(cat, "•")
            count = self._manifest.count_for(cat)
            label = f"  {icon}  {cat}" + (f"  ({count} saved)" if count else "")
            ctk.CTkCheckBox(
                grid, text=label,
                variable=var,
                command=self._sync_all,
                font=ctk.CTkFont(size=19),
                text_color=C["text_dim"],
                fg_color=C["accent"], hover_color=C["accent_h"],
                checkmark_color="white", corner_radius=5,
            ).grid(row=i // 2, column=i % 2, sticky="w", pady=10, padx=12)

    def _build_storage_section(self, parent):
        SectionHeader(parent, "WHERE TO SAVE").pack(anchor="w", padx=24, pady=(24, 10))

        card = Card(parent)
        card.pack(fill="x", padx=24, pady=(0, 4))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=(18, 6))

        # Radio buttons
        self._storage_var = ctk.StringVar(value="local")
        radio_row = ctk.CTkFrame(inner, fg_color="transparent")
        radio_row.pack(fill="x", pady=(0, 16))

        radio_kw = dict(
            variable=self._storage_var,
            command=self._on_storage_change,
            font=ctk.CTkFont(size=19),
            text_color=C["text"],
            fg_color=C["accent"],
            hover_color=C["accent_h"],
            border_color=C["border"],
        )
        ctk.CTkRadioButton(radio_row, text="  💾  On this computer",  value="local",    **radio_kw).pack(side="left", padx=(0, 36))
        ctk.CTkRadioButton(radio_row, text="  ☁️  My Microsoft OneDrive", value="onedrive", **radio_kw).pack(side="left")

        ctk.CTkFrame(inner, height=1, fg_color=C["border"]).pack(fill="x", pady=(0, 16))

        # ── Local panel ───────────────────────────────────────────────────
        self._local_panel = ctk.CTkFrame(inner, fg_color="transparent")
        self._local_panel.pack(fill="x", pady=(0, 8))

        default = str(Path.home() / "Downloads" / "GmailMedia")
        self._out_var = ctk.StringVar(value=default)
        ctk.CTkEntry(
            self._local_panel, textvariable=self._out_var,
            height=54, corner_radius=8,
            fg_color=C["surface2"], border_color=C["border"],
            font=ctk.CTkFont(size=17), text_color=C["text_dim"],
        ).pack(side="left", fill="x", expand=True, padx=(0, 12))
        GhostButton(
            self._local_panel, text="📂  Browse",
            width=140, height=54,
            command=self._browse,
        ).pack(side="right")

        # ── OneDrive panel ────────────────────────────────────────────────
        self._od_panel = ctk.CTkFrame(inner, fg_color="transparent")
        # (not packed initially — shown when OneDrive radio is selected)

        od_top = ctk.CTkFrame(self._od_panel, fg_color="transparent")
        od_top.pack(fill="x")

        self._od_status_var = ctk.StringVar(value="You need to sign in to use OneDrive")
        ctk.CTkLabel(
            od_top, textvariable=self._od_status_var,
            font=ctk.CTkFont(size=18), text_color=C["text_dim"],
        ).pack(side="left")

        self._od_signin_btn = ctk.CTkButton(
            od_top,
            text="Sign in to Microsoft →",
            height=46, corner_radius=8,
            fg_color=C["ms_blue"], hover_color=C["ms_blue_h"],
            font=ctk.CTkFont(size=18, weight="bold"),
            command=self._ms_signin,
        )
        self._od_signin_btn.pack(side="right")

        self._od_signout_btn = ctk.CTkButton(
            od_top,
            text="Sign out →",
            height=46, corner_radius=8,
            fg_color=C["surface2"], hover_color=C["danger"],
            font=ctk.CTkFont(size=17),
            command=self._ms_signout,
        )
        # Shown only when signed in

        ctk.CTkLabel(
            self._od_panel,
            text="Your files will be saved to your OneDrive, organised by sender and date.",
            font=ctk.CTkFont(size=16), text_color=C["muted"],
        ).pack(anchor="w", pady=(8, 8))

        # Try silent MS auth on startup if cached token exists
        if MS_TOKEN_CACHE.exists() and MSAL_OK and ONEDRIVE_CONFIGURED:
            threading.Thread(target=self._ms_silent_auth, daemon=True).start()

    def _build_options_section(self, parent):
        SectionHeader(parent, "EXTRA FEATURES").pack(anchor="w", padx=24, pady=(24, 10))

        card = Card(parent)
        card.pack(fill="x", padx=24, pady=(0, 4))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=18)

        cb_kw = dict(
            fg_color=C["accent"], hover_color=C["accent_h"],
            checkmark_color="white", corner_radius=5,
        )

        # OneDrive link scanning
        self._scan_links_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            inner,
            text="  🔗  Also grab files that were shared as links in emails",
            variable=self._scan_links_var,
            font=ctk.CTkFont(size=19), text_color=C["text"],
            **cb_kw,
        ).pack(anchor="w", pady=(0, 6))

        if not REQUESTS_OK:
            ctk.CTkLabel(
                inner,
                text="      ⚠️  Cannot download linked files — a required component is missing. Please reinstall.",
                font=ctk.CTkFont(size=16), text_color="#ef4444",
            ).pack(anchor="w", pady=(0, 10))
        else:
            ctk.CTkLabel(
                inner,
                text="      Works automatically for most links. Signing in to OneDrive above also unlocks privately shared files.",
                font=ctk.CTkFont(size=16), text_color=C["muted"],
            ).pack(anchor="w", pady=(0, 14))

        ctk.CTkFrame(inner, height=1, fg_color=C["border"]).pack(fill="x", pady=(0, 14))

        # Google Sheets logging
        self._sheets_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            inner,
            text="  📊  Save a record of all downloads to Google Sheets",
            variable=self._sheets_var,
            font=ctk.CTkFont(size=19), text_color=C["text"],
            **cb_kw,
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            inner,
            text="      A spreadsheet is automatically created in your Google Drive listing every file downloaded.",
            font=ctk.CTkFont(size=16), text_color=C["muted"],
        ).pack(anchor="w")

    def _build_hint_section(self, parent):
        card = ctk.CTkFrame(parent, fg_color="#0f1e38", corner_radius=10,
                            border_width=1, border_color="#1e3a6e")
        card.pack(fill="x", padx=24, pady=(20, 12))

        if not IS_FROZEN:
            ctk.CTkLabel(
                card,
                text="ℹ️  Dev mode — place credentials.json in this folder before running. "
                     "End-user builds have it bundled automatically.",
                font=ctk.CTkFont(size=17), text_color="#6ea8fe",
                wraplength=680, justify="left",
            ).pack(anchor="w", padx=18, pady=(16, 6))

            ctk.CTkButton(
                card,
                text="Open setup guide →",
                fg_color="transparent", hover_color="#1e3a6e",
                text_color="#4d8aff",
                font=ctk.CTkFont(size=17, underline=True),
                height=38, anchor="w",
                command=lambda: webbrowser.open(
                    "https://developers.google.com/gmail/api/quickstart/python"
                    "#authorize_credentials_for_a_desktop_application"
                ),
            ).pack(anchor="w", padx=12, pady=(0, 4))

        ctk.CTkLabel(
            card,
            text="ℹ️  First time only — your browser will open so you can sign in to Google. "
                 "Just sign in and click Allow. The app will continue on its own after that.",
            font=ctk.CTkFont(size=17), text_color="#6ea8fe",
            wraplength=680, justify="left",
        ).pack(anchor="w", padx=18, pady=(16, 6))

        self._signout_btn = ctk.CTkButton(
            card,
            text="Sign out / switch Google account →",
            fg_color="transparent", hover_color="#1e3a6e",
            text_color="#ef4444",
            font=ctk.CTkFont(size=17, underline=True),
            height=38, anchor="w",
            command=self._sign_out,
        )
        if TOKEN_PATH.exists():
            self._signout_btn.pack(anchor="w", padx=12, pady=(0, 14))

        ctk.CTkFrame(parent, height=12, fg_color="transparent").pack()

    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=230)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        self._status_var = ctk.StringVar(value="Ready — set your options above and click Start Download.")
        ctk.CTkLabel(
            footer, textvariable=self._status_var,
            font=ctk.CTkFont(size=18), text_color=C["text_dim"],
            wraplength=760,
        ).pack(pady=(18, 10))

        self._progress = ctk.CTkProgressBar(
            footer, width=760, height=14, corner_radius=7,
            fg_color=C["border"], progress_color=C["accent"],
        )
        self._progress.set(0)
        self._progress.pack(pady=(0, 8))

        self._prog_label = ctk.CTkLabel(
            footer, text="",
            font=ctk.CTkFont(size=16), text_color=C["muted"],
        )
        self._prog_label.pack()

        btn_row = ctk.CTkFrame(footer, fg_color="transparent")
        btn_row.pack(pady=(14, 0))

        self._start_btn = PrimaryButton(
            btn_row, text="▶   Start Download",
            width=250, command=self._start,
        )
        self._start_btn.pack(side="left", padx=8)

        self._cancel_btn = GhostButton(
            btn_row, text="⏹  Cancel",
            width=140, state="disabled", command=self._cancel,
        )
        self._cancel_btn.pack(side="left", padx=8)

        self._open_btn = ctk.CTkButton(
            btn_row, text="📂  Open Folder",
            width=180, height=58, corner_radius=10,
            fg_color=C["green"], hover_color=C["green_h"],
            font=ctk.CTkFont(size=19), state="disabled",
            command=self._open_output,
        )
        self._open_btn.pack(side="left", padx=8)

        self._sheet_btn = ctk.CTkButton(
            btn_row, text="📊  View Spreadsheet",
            width=180, height=58, corner_radius=10,
            fg_color="#1e4d1e", hover_color="#2d6b2d",
            font=ctk.CTkFont(size=19), state="disabled",
            command=self._open_sheet,
        )
        self._sheet_btn.pack(side="left", padx=8)

    # ══════════════════════════════════════════════════════════════════════
    #  Storage section logic
    # ══════════════════════════════════════════════════════════════════════

    def _on_storage_change(self, *_):
        if self._storage_var.get() == "local":
            self._od_panel.pack_forget()
            self._local_panel.pack(fill="x", pady=(0, 8))
        else:
            self._local_panel.pack_forget()
            self._od_panel.pack(fill="x", pady=(0, 8))

    def _ms_silent_auth(self):
        """Try to restore Microsoft session from cached token (no browser)."""
        try:
            client = OneDriveClient(MS_TOKEN_CACHE)
            accounts = client._app.get_accounts()
            if accounts:
                result = client._app.acquire_token_silent(
                    OneDriveClient.MS_SCOPES, account=accounts[0]
                )
                if result and "access_token" in result:
                    client._token = result["access_token"]
                    client._save_cache()
                    self._ms_client = client
                    self.after(0, self._update_ms_ui_signed_in)
        except Exception:
            pass

    def _ms_signin(self):
        self._od_signin_btn.configure(state="disabled", text="Signing in…")
        threading.Thread(target=self._ms_do_signin, daemon=True).start()

    def _ms_do_signin(self):
        try:
            if not MSAL_OK:
                self.after(0, lambda: messagebox.showerror(
                    "Missing package",
                    "The 'msal' package is not installed.\n\n"
                    "Run:  pip install msal",
                    parent=self,
                ))
                self.after(0, self._update_ms_ui_signed_out)
                return
            if not ONEDRIVE_CONFIGURED:
                self.after(0, lambda: messagebox.showerror(
                    "Not configured",
                    "Microsoft client ID is not set.\n\n"
                    "Create an Azure App Registration and replace\n"
                    "MICROSOFT_CLIENT_ID in core.py before building.",
                    parent=self,
                ))
                self.after(0, self._update_ms_ui_signed_out)
                return
            client = OneDriveClient(MS_TOKEN_CACHE)
            ok = client.authenticate()
            if ok:
                self._ms_client = client
                self.after(0, self._update_ms_ui_signed_in)
            else:
                self.after(0, lambda: messagebox.showerror(
                    "Sign-in failed", "Could not sign in to Microsoft.", parent=self
                ))
                self.after(0, self._update_ms_ui_signed_out)
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror(
                "Sign-in error", str(exc), parent=self
            ))
            self.after(0, self._update_ms_ui_signed_out)

    def _ms_signout(self):
        if self._ms_client:
            self._ms_client.sign_out()
            self._ms_client = None
        self._update_ms_ui_signed_out()

    def _update_ms_ui_signed_in(self):
        self._od_status_var.set("✅  Signed in — files will go to your OneDrive")
        self._od_signin_btn.pack_forget()
        self._od_signout_btn.pack(side="right")

    def _update_ms_ui_signed_out(self):
        self._od_status_var.set("You need to sign in to use OneDrive")
        self._od_signout_btn.pack_forget()
        self._od_signin_btn.configure(state="normal", text="Sign in to Microsoft →")
        self._od_signin_btn.pack(side="right")

    # ══════════════════════════════════════════════════════════════════════
    #  UI helpers
    # ══════════════════════════════════════════════════════════════════════

    def _badge_text(self) -> str:
        n = self._manifest.total_count() if hasattr(self, "_manifest") else 0
        return f"  {n} files saved  "

    def _clear_history(self):
        if not messagebox.askyesno(
            "Clear download history?",
            "This removes the record of every previously downloaded file.\n\n"
            "The next run will re-download everything from scratch.\n"
            "(Your actual files on disk / OneDrive are NOT deleted.)",
            parent=self,
        ):
            return
        self._manifest.clear()
        # self._badge_var.set(self._badge_text())

    def _toggle_all(self):
        v = self._all_var.get()
        for var in self._cat_vars.values():
            var.set(v)

    def _sync_all(self):
        self._all_var.set(all(v.get() for v in self._cat_vars.values()))

    def _browse(self):
        d = filedialog.askdirectory(title="Choose save folder", initialdir=self._out_var.get())
        if d:
            self._out_var.set(d)

    def _open_output(self):
        if self._output_dir and self._output_dir.exists():
            open_folder(self._output_dir)

    def _open_sheet(self):
        if self._sheets_url:
            webbrowser.open(self._sheets_url)

    def _sign_out(self):
        if messagebox.askyesno(
            "Sign out",
            "This will sign you out of Google.\n"
            "You'll be asked to sign in again next time you start the app.\n\nContinue?",
            parent=self,
        ):
            TOKEN_PATH.unlink(missing_ok=True)
            SHEET_ID_PATH.unlink(missing_ok=True)
            self._status_var.set("Signed out. You'll be asked to sign in again next time.")
            self._signout_btn.pack_forget()

    # ── Validation ─────────────────────────────────────────────────────────

    def _validate(self) -> tuple[bool, str]:
        if not any(v.get() for v in self._cat_vars.values()):
            return False, "Please select at least one file type."
        mode = self._storage_var.get()
        if mode == "local" and not self._out_var.get().strip():
            return False, "Please choose a save folder."
        if mode == "onedrive" and not (self._ms_client and self._ms_client.is_authenticated()):
            return False, "Please sign in to Microsoft before using OneDrive storage."
        return True, ""

    # ── Start / cancel ─────────────────────────────────────────────────────

    def _start(self):
        ok, err = self._validate()
        if not ok:
            messagebox.showerror("Check your settings", err, parent=self)
            return

        mode = self._storage_var.get()
        if mode == "local":
            self._output_dir = Path(self._out_var.get()).expanduser().resolve()
        else:
            # OneDrive: store JSON report in DATA_DIR/reports
            self._output_dir = DATA_DIR / "reports"

        self._output_dir.mkdir(parents=True, exist_ok=True)

        self._progress.set(0)
        self._prog_label.configure(text="")
        self._open_btn.configure(state="disabled")
        self._sheet_btn.configure(state="disabled")
        self._start_btn.configure(state="disabled")
        self._cancel_btn.configure(state="normal")

        cats = {c for c, v in self._cat_vars.items() if v.get()}

        cfg = WorkerConfig(
            output_dir       = self._output_dir,
            categories       = cats,
            after            = self._after_sel.get(),
            before           = self._before_sel.get(),
            creds_path       = CREDS_PATH,
            token_path       = TOKEN_PATH,
            manifest         = self._manifest,
            msg_queue        = self._q,
            storage_mode     = mode,
            onedrive_client  = self._ms_client if mode == "onedrive" else None,
            enable_sheets    = self._sheets_var.get(),
            sheet_id_path    = SHEET_ID_PATH if self._sheets_var.get() else None,
            scan_links       = self._scan_links_var.get(),
            email_direction  = self._direction_var.get(),
        )
        self._worker = ExtractionWorker(cfg)
        self._thread = threading.Thread(target=self._worker.run, daemon=True)
        self._thread.start()

    def _cancel(self):
        if self._worker:
            self._worker.stop()
        self._cancel_btn.configure(state="disabled")
        self._start_btn.configure(state="normal")

    # ── Queue polling ──────────────────────────────────────────────────────

    def _poll(self):
        try:
            while True:
                msg  = self._q.get_nowait()
                kind = msg["kind"]

                if kind == "status":
                    self._status_var.set(msg["text"])

                elif kind == "total":
                    self._total = msg["count"]

                elif kind == "progress":
                    cur, tot = msg["current"], msg["total"]
                    if tot:
                        self._progress.set(cur / tot)
                        self._prog_label.configure(
                            text=f"Checking email {cur:,} of {tot:,}…"
                        )

                elif kind == "file_done":
                    self._status_var.set(
                        f"💾  Saved: {msg['name']}  "
                        f"({msg['count']} files downloaded so far)"
                    )
                    self._badge_var.set(self._badge_text())

                elif kind == "skipped":
                    self._status_var.set(f"⏭  Already have this file: {msg['name']}")

                elif kind == "sheet_url":
                    self._sheets_url = msg["url"]

                elif kind == "done":
                    dl, sk, er = msg["downloaded"], msg["skipped"], msg["errors"]
                    self._progress.set(1)
                    self._prog_label.configure(text="")
                    self._start_btn.configure(state="normal")
                    self._cancel_btn.configure(state="disabled")
                    self._badge_var.set(self._badge_text())

                    # Enable relevant action buttons
                    if self._storage_var.get() == "local":
                        self._open_btn.configure(state="normal")
                    if self._sheets_url:
                        self._sheet_btn.configure(state="normal")

                    self._status_var.set(
                        f"✅  Done!  {dl} new files saved  ·  "
                        f"{sk} already had (skipped)  "
                        + (f"·  {er} errors" if er else "")
                    )
                    if msg.get("storage_mode") == "onedrive":
                        saved_line = "\nFiles saved to your Microsoft OneDrive\n(organised by sender → date → file type)"
                    else:
                        saved_line = f"\nSaved to:\n{msg['output_dir']}"
                    extra = ""
                    if self._sheets_url:
                        extra = "\n\nClick \"View Spreadsheet\" to see the full record."
                    messagebox.showinfo(
                        "All done!",
                        f"✅  Download complete!\n\n"
                        f"  New files saved   :  {dl}\n"
                        f"  Already had       :  {sk}\n"
                        + (f"  Errors            :  {er}\n" if er else "")
                        + saved_line
                        + extra,
                        parent=self,
                    )

                elif kind == "error":
                    self._status_var.set(f"❌  {msg['text']}")
                    self._start_btn.configure(state="normal")
                    self._cancel_btn.configure(state="disabled")
                    messagebox.showerror("Something went wrong", msg["text"], parent=self)

        except Exception:
            pass

        self.after(120, self._poll)


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    App().mainloop()

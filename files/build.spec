# build.spec — PyInstaller build configuration
# Usage:  pyinstaller build.spec
#
# Output:
#   Windows → dist/GmailMediaExtractor.exe   (single file)
#   macOS   → dist/GmailMediaExtractor.app   (app bundle)
#   Linux   → dist/GmailMediaExtractor       (single binary)

import sys
import subprocess
from pathlib import Path

block_cipher = None

# Resolve customtkinter asset path without importing it at spec-parse time
# (importing it here would require tkinter to be present during the build,
#  which is not guaranteed on Linux).
_ctk_path_raw = subprocess.check_output(
    [sys.executable, "-c",
     "import importlib.util; s=importlib.util.find_spec('customtkinter'); print(s.origin)"],
    text=True,
).strip()
CTK_DIR = Path(_ctk_path_raw).parent

a = Analysis(
    ["app.py"],
    pathex=[SPECPATH],
    binaries=[],
    datas=[
        # customtkinter themes + fonts
        (str(CTK_DIR / "assets"), "customtkinter/assets"),
        # OAuth client credentials — bundled so end-users need no setup files
        ("credentials.json", "."),
    ],
    hiddenimports=[
        "customtkinter",
        "PIL._tkinter_finder",
        "google.auth",
        "google.auth.transport",
        "google.auth.transport.requests",
        "google.oauth2",
        "google.oauth2.credentials",
        "google_auth_oauthlib",
        "google_auth_oauthlib.flow",
        "googleapiclient",
        "googleapiclient.discovery",
        "googleapiclient.errors",
        "googleapiclient.http",
        "cachetools",
        "cachetools.func",
        "charset_normalizer",
        "httplib2",
        "uritemplate",
        # Microsoft / OneDrive
        "msal",
        "msal.application",
        "msal.authority",
        "msal.token_cache",
        "requests",
        "requests.adapters",
        "requests.auth",
        "requests.models",
        "urllib3",
        "urllib3.util",
        "certifi",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "scipy", "IPython"],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Single-file EXE / binary ───────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="GmailMediaExtractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # hide terminal for end-users
    disable_windowed_traceback=False,
    target_arch=None,
    # icon="assets/icon.ico",   # uncomment + add icon file
)

# ── macOS .app bundle ──────────────────────────────────────────────────────
if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name="GmailMediaExtractor.app",
        icon=None,
        bundle_identifier="com.yourname.gmailmediaextractor",
        info_plist={
            "CFBundleShortVersionString": "1.0.0",
            "CFBundleVersion":            "1.0.0",
            "NSHighResolutionCapable":    True,
        },
    )

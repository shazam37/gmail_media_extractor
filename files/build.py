#!/usr/bin/env python3
"""
build.py — One-command build script
Run:  python build.py

Installs deps → ensures tkinter → runs PyInstaller → prints next steps.
"""

import subprocess
import sys
import platform
from pathlib import Path

ROOT = Path(__file__).parent


def run(cmd: str, **kwargs):
    print(f"\n▶  {cmd}")
    result = subprocess.run(cmd, shell=True, **kwargs)
    if result.returncode != 0:
        print(f"❌  Command failed (exit {result.returncode})")
        sys.exit(result.returncode)


def ensure_tkinter(os_name: str):
    """On Linux, tkinter is a system package — not installable via pip."""
    try:
        import tkinter  # noqa: F401
        print("✅  tkinter found.")
        return
    except ModuleNotFoundError:
        pass

    if os_name != "Linux":
        print("⚠️  tkinter not found. Install it for your OS and retry.")
        sys.exit(1)

    print("⚠️  tkinter not found — detecting distro…")

    distro_id = ""
    try:
        with open("/etc/os-release") as f:
            for line in f:
                if line.startswith("ID="):
                    distro_id = line.strip().split("=")[1].strip('"').lower()
                    break
    except FileNotFoundError:
        pass

    py_dot = f"{sys.version_info.major}.{sys.version_info.minor}"

    if distro_id in ("fedora", "rhel", "centos", "rocky", "almalinux"):
        cmd = "sudo dnf install -y python3-tkinter"
    elif distro_id in ("ubuntu", "debian", "linuxmint", "pop"):
        cmd = f"sudo apt-get install -y python{py_dot}-tk"
    elif distro_id in ("arch", "manjaro"):
        cmd = "sudo pacman -S --noconfirm tk"
    elif distro_id == "opensuse":
        cmd = "sudo zypper install -y python3-tk"
    else:
        print(
            f"\n❌  Could not auto-detect package manager for '{distro_id}'.\n"
            f"    Install tkinter manually then retry:\n\n"
            f"      Fedora/RHEL  →  sudo dnf install python3-tkinter\n"
            f"      Ubuntu/Debian→  sudo apt install python{py_dot}-tk\n"
            f"      Arch         →  sudo pacman -S tk\n"
        )
        sys.exit(1)

    print(f"   Running: {cmd}")
    result = subprocess.run(cmd, shell=True)
    if result.returncode != 0:
        print("❌  Auto-install failed. Run the command above manually, then retry.")
        sys.exit(1)

    try:
        # Force re-import after system install
        import importlib
        import tkinter  # noqa: F401
        importlib.import_module("tkinter")
        print("✅  tkinter installed successfully.")
    except ModuleNotFoundError:
        print(
            "⚠️  tkinter installed but not yet visible to this Python process.\n"
            "    Open a fresh terminal and run  python build.py  again."
        )
        sys.exit(1)


def main():
    os_name = platform.system()
    print(f"{'='*55}")
    print(f"  Gmail Media Extractor — Build Script")
    print(f"  Platform: {os_name} / Python {sys.version.split()[0]}")
    print(f"{'='*55}")

    # 0. Ensure tkinter is present (system package on Linux)
    ensure_tkinter(os_name)

    # 1. Install Python deps
    run(f"{sys.executable} -m pip install -r requirements.txt --quiet")

    # 2. PyInstaller
    run(f"{sys.executable} -m PyInstaller build.spec --noconfirm")

    # 3. Platform-specific next steps
    dist = ROOT / "dist"
    print(f"\n{'='*55}")
    print(f"  ✅  Binary built → {dist}")
    print(f"{'='*55}")

    if os_name == "Windows":
        exe = dist / "GmailMediaExtractor.exe"
        print(f"""
Next steps (Windows installer):
  1. Install NSIS  →  https://nsis.sourceforge.io/Download
  2. Copy credentials.json into this folder
  3. Right-click installer_windows.nsi → "Compile NSIS Script"
  4. Distribute: GmailMediaExtractor_Setup.exe

The .exe is also at: {exe}
(You can share just the .exe + credentials.json if no installer is needed)
""")

    elif os_name == "Darwin":
        app = dist / "GmailMediaExtractor.app"
        print(f"""
Next steps (macOS DMG):
  1. brew install create-dmg
  2. Copy credentials.json into this folder
  3. chmod +x create_dmg.sh && ./create_dmg.sh
  4. Distribute: GmailMediaExtractor_Installer.dmg

The .app is at: {app}
""")

    else:
        binary = dist / "GmailMediaExtractor"
        print(f"""
Linux binary: {binary}
  1. Copy credentials.json next to the binary
  2. Zip both files and distribute

To create a .deb package, consider:
  fpm -s dir -t deb -n gmail-media-extractor \\
      --version 1.0.0 \\
      dist/GmailMediaExtractor=/usr/local/bin/GmailMediaExtractor
""")


if __name__ == "__main__":
    main()
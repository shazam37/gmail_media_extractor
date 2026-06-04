; ============================================================
;  Gmail Media Extractor — Windows Installer
;  Built with NSIS (https://nsis.sourceforge.io)
;
;  Prerequisites:
;    1. Install NSIS
;    2. Run:  pyinstaller build.spec
;    3. Compile this script in NSIS
;
;  Output: GmailMediaExtractor_Setup.exe
; ============================================================

!define APP_NAME        "Gmail Media Extractor"
!define APP_VERSION     "1.0.0"
!define APP_EXE         "GmailMediaExtractor.exe"
!define INSTALL_DIR     "$PROGRAMFILES64\GmailMediaExtractor"
!define UNINSTALL_KEY   "Software\Microsoft\Windows\CurrentVersion\Uninstall\GmailMediaExtractor"

Name "${APP_NAME}"
OutFile "GmailMediaExtractor_Setup.exe"
InstallDir "${INSTALL_DIR}"
InstallDirRegKey HKLM "${UNINSTALL_KEY}" "InstallLocation"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
Unicode True

; ── Pages ─────────────────────────────────────────────────────────────────
!include "MUI2.nsh"

!define MUI_ABORTWARNING
!define MUI_ICON       "assets\icon.ico"   ; optional — add your icon
!define MUI_UNICON     "assets\icon.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE    "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

; ── Install ────────────────────────────────────────────────────────────────
Section "MainSection" SEC01
    SetOutPath "$INSTDIR"

    ; Copy the single-file executable
    File "dist\${APP_EXE}"

    ; Copy credentials (bundled by developer)
    File "credentials.json"

    ; Desktop shortcut
    CreateShortcut "$DESKTOP\${APP_NAME}.lnk" \
                   "$INSTDIR\${APP_EXE}"

    ; Start Menu shortcut
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortcut  "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" \
                    "$INSTDIR\${APP_EXE}"
    CreateShortcut  "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk" \
                    "$INSTDIR\Uninstall.exe"

    ; Write uninstall registry keys
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "DisplayName"      "${APP_NAME}"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "UninstallString"  "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "InstallLocation"  "$INSTDIR"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "DisplayVersion"   "${APP_VERSION}"
    WriteRegDWORD HKLM "${UNINSTALL_KEY}" "NoModify"         1
    WriteRegDWORD HKLM "${UNINSTALL_KEY}" "NoRepair"         1

    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

; ── Uninstall ──────────────────────────────────────────────────────────────
Section "Uninstall"
    Delete "$INSTDIR\${APP_EXE}"
    Delete "$INSTDIR\credentials.json"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir  "$INSTDIR"

    Delete "$DESKTOP\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"
    RMDir  "$SMPROGRAMS\${APP_NAME}"

    DeleteRegKey HKLM "${UNINSTALL_KEY}"
SectionEnd

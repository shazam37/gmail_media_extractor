#!/bin/bash
# create_dmg.sh — packages the macOS .app into a distributable .dmg
#
# Prerequisites:
#   brew install create-dmg
#   pyinstaller build.spec   (produces dist/GmailMediaExtractor.app)
#
# Usage:
#   chmod +x create_dmg.sh && ./create_dmg.sh

set -e

APP_NAME="GmailMediaExtractor"
APP_BUNDLE="dist/${APP_NAME}.app"
DMG_NAME="${APP_NAME}_Installer.dmg"
STAGING="dmg_staging"

echo "📦  Preparing staging folder…"
rm -rf "$STAGING"
mkdir  "$STAGING"
cp -r  "$APP_BUNDLE"      "$STAGING/"
cp     "credentials.json" "$STAGING/"          # bundle credentials with the .app
cp     "README_user.txt"  "$STAGING/README.txt" 2>/dev/null || true

echo "💿  Creating DMG…"
create-dmg \
  --volname  "${APP_NAME} Installer" \
  --volicon  "assets/icon.icns" \
  --window-pos  200 200 \
  --window-size 660 400 \
  --icon-size   100 \
  --icon        "${APP_NAME}.app" 180 170 \
  --hide-extension "${APP_NAME}.app" \
  --app-drop-link  480 170 \
  "${DMG_NAME}" \
  "${STAGING}/"

echo "✅  Done → ${DMG_NAME}"

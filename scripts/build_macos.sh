#!/usr/bin/env bash
set -euo pipefail

APP_NAME="${APP_NAME:-NTOU_Tools}"
APP_VERSION="${APP_VERSION:-1.0}"
APP_VERSION_SANITIZED="${APP_VERSION#v}"
APP_VERSION_SANITIZED="${APP_VERSION_SANITIZED#V}"
PY_FILE="${PY_FILE:-Line_chart.py}"
DIST_DIR="dist"
BUILD_DIR="build"

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_ROOT"

DATA_ARGS=()
if [[ -f "sample_excel.txt" ]]; then
  DATA_ARGS+=(--add-data "sample_excel.txt:.")
fi
if [[ -f "sample_data.json" ]]; then
  DATA_ARGS+=(--add-data "sample_data.json:.")
fi

ICON_SOURCE="messageImage_1767257219427.jpg"
ICON_DIR="build/icons"
ICON_PATH="$ICON_DIR/${APP_NAME}.icns"
if [[ -f "$ICON_SOURCE" ]]; then
  mkdir -p "$ICON_DIR"
  if [[ ! -f "$ICON_PATH" ]]; then
    ICONSET_DIR="$ICON_DIR/${APP_NAME}.iconset"
    mkdir -p "$ICONSET_DIR"
    for size in 16 32 64 128 256 512; do
      sips -z $size $size "$ICON_SOURCE" --out "$ICONSET_DIR/icon_${size}x${size}.png" >/dev/null
    done
    sips -z 32 32 "$ICON_SOURCE" --out "$ICONSET_DIR/icon_16x16@2x.png" >/dev/null
    sips -z 64 64 "$ICON_SOURCE" --out "$ICONSET_DIR/icon_32x32@2x.png" >/dev/null
    sips -z 256 256 "$ICON_SOURCE" --out "$ICONSET_DIR/icon_128x128@2x.png" >/dev/null
    sips -z 512 512 "$ICON_SOURCE" --out "$ICONSET_DIR/icon_256x256@2x.png" >/dev/null
    sips -z 1024 1024 "$ICON_SOURCE" --out "$ICONSET_DIR/icon_512x512@2x.png" >/dev/null
    iconutil -c icns "$ICONSET_DIR" -o "$ICON_PATH"
  fi
fi

python3 -m pip install --upgrade pyinstaller
ICON_ARGS=()
if [[ -f "$ICON_PATH" ]]; then
  ICON_ARGS+=(--icon "$ICON_PATH")
fi
pyinstaller --noconfirm --clean --windowed --name "$APP_NAME" "${ICON_ARGS[@]}" "${DATA_ARGS[@]}" "$PY_FILE"

PLIST_PATH="$APP_PATH/Contents/Info.plist"
if [[ -f "$PLIST_PATH" ]]; then
  /usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString $APP_VERSION_SANITIZED" "$PLIST_PATH" || true
  /usr/libexec/PlistBuddy -c "Set :CFBundleVersion $APP_VERSION_SANITIZED" "$PLIST_PATH" || true
fi

APP_PATH="$DIST_DIR/$APP_NAME.app"
DMG_PATH="$DIST_DIR/${APP_NAME}.dmg"
STAGING_DIR="$BUILD_DIR/dmg"

if [[ -n "${APPLE_SIGN_ID:-}" ]]; then
  codesign --force --deep --options runtime --sign "$APPLE_SIGN_ID" "$APP_PATH"
fi

if [[ -n "${APPLE_ID:-}" && -n "${APPLE_TEAM_ID:-}" && -n "${APPLE_APP_PASSWORD:-}" ]]; then
  ZIP_PATH="$BUILD_DIR/${APP_NAME}.zip"
  ditto -c -k --keepParent "$APP_PATH" "$ZIP_PATH"
  xcrun notarytool submit "$ZIP_PATH" --apple-id "$APPLE_ID" --team-id "$APPLE_TEAM_ID" --password "$APPLE_APP_PASSWORD" --wait
  xcrun stapler staple "$APP_PATH"
fi

rm -rf "$STAGING_DIR"
mkdir -p "$STAGING_DIR"
cp -R "$APP_PATH" "$STAGING_DIR/"
ln -s /Applications "$STAGING_DIR/Applications"
rm -f "$DMG_PATH"
hdiutil create -volname "$APP_NAME" -srcfolder "$STAGING_DIR" -ov -format UDZO "$DMG_PATH"

if [[ -n "${APPLE_SIGN_ID:-}" ]]; then
  codesign --force --sign "$APPLE_SIGN_ID" "$DMG_PATH"
fi

# Build Installers

This project ships macOS DMG and Windows installer builds. You must build each
platform on its own OS.

## macOS (DMG + notarization)

Requirements:
- Xcode Command Line Tools
- Apple Developer ID Application certificate installed in Keychain

Environment variables:
- APP_VERSION: version string (default 1.0, accepts "v1" or "1.0")
- APPLE_SIGN_ID: "Developer ID Application: Your Name (TEAMID)"
- APPLE_ID: your Apple ID email
- APPLE_TEAM_ID: your team ID
- APPLE_APP_PASSWORD: app-specific password

Run:
```
chmod +x scripts/build_macos.sh
./scripts/build_macos.sh
```

Outputs:
- dist/NTOU_Tools.app
- dist/NTOU_Tools.dmg

## Windows (EXE installer)

Requirements:
- Python + PyInstaller
- Inno Setup (iscc in PATH)
- Code signing certificate (PFX) and signtool

Environment variables:
- APP_VERSION: version string (default 1.0, accepts "v1" or "1.0")
- SIGNTOOL_PATH: full path to signtool.exe
- SIGN_CERT_PFX: path to .pfx
- SIGN_CERT_PASS: password for .pfx

Run:
```
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

Outputs:
- dist\NTOU_Tools.exe
- dist\installer\NTOU_Tools_Setup.exe

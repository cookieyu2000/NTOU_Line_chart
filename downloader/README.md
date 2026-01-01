# NTOU_Tools_Downloader

Electron-based downloader that fetches the correct installer from GitHub Releases
and opens it for the user.

## Development

```
npm install
npm start
```

## Build

```
npm run build
```

Outputs are in `downloader/dist`.

## Release Assets

This app expects these assets to exist under the latest release:
- NTOU_Tools_Setup.exe
- NTOU_Tools.dmg

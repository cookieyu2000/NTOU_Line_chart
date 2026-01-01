const { app, BrowserWindow, ipcMain, shell } = require("electron");
const path = require("path");
const fs = require("fs");
const https = require("https");

const DOWNLOADS = {
  win: "https://github.com/cookieyu2000/NTOU_Line_chart/releases/latest/download/NTOU_Tools_Setup.exe",
  mac: "https://github.com/cookieyu2000/NTOU_Line_chart/releases/latest/download/NTOU_Tools.dmg",
};

function createWindow() {
  const win = new BrowserWindow({
    width: 720,
    height: 520,
    resizable: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
    },
  });

  win.loadFile(path.join(__dirname, "index.html"));
}

function followRedirect(url, cb) {
  https
    .get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        return followRedirect(res.headers.location, cb);
      }
      cb(res);
    })
    .on("error", cb);
}

function downloadFile(url, targetPath, onProgress) {
  return new Promise((resolve, reject) => {
    followRedirect(url, (res) => {
      if (res instanceof Error) {
        return reject(res);
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error(`Download failed: ${res.statusCode}`));
      }

      const total = Number(res.headers["content-length"] || 0);
      let received = 0;
      const file = fs.createWriteStream(targetPath);

      res.on("data", (chunk) => {
        received += chunk.length;
        if (total > 0) {
          onProgress(Math.round((received / total) * 100));
        } else {
          onProgress(null);
        }
      });

      res.pipe(file);

      file.on("finish", () => {
        file.close(() => resolve(targetPath));
      });

      file.on("error", (err) => {
        fs.unlink(targetPath, () => reject(err));
      });
    });
  });
}

ipcMain.handle("get-platform", () => process.platform);

ipcMain.handle("download-installer", async (_event, platformKey) => {
  const url = DOWNLOADS[platformKey];
  if (!url) {
    throw new Error("Unsupported platform");
  }

  const fileName = platformKey === "win" ? "NTOU_Tools_Setup.exe" : "NTOU_Tools.dmg";
  const downloadDir = app.getPath("downloads");
  const targetPath = path.join(downloadDir, fileName);

  await downloadFile(url, targetPath, (progress) => {
    _event.sender.send("download-progress", progress);
  });

  await shell.openPath(targetPath);
  return targetPath;
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

const statusEl = document.querySelector("#status");
const progressEl = document.querySelector("#progress");
const winBtn = document.querySelector("#download-win");
const macBtn = document.querySelector("#download-mac");

function setStatus(text) {
  statusEl.textContent = text;
}

function setBusy(busy) {
  winBtn.disabled = busy;
  macBtn.disabled = busy;
}

function selectDefault(platform) {
  if (platform === "win32") {
    winBtn.classList.add("primary");
  } else if (platform === "darwin") {
    macBtn.classList.add("primary");
  }
}

async function startDownload(platformKey) {
  setBusy(true);
  progressEl.value = 0;
  progressEl.classList.remove("hidden");
  setStatus("Downloading... this may take a moment.");
  try {
    const path = await window.downloader.downloadInstaller(platformKey);
    setStatus(`Downloaded to: ${path}`);
  } catch (err) {
    setStatus(`Download failed: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

window.downloader.onProgress((progress) => {
  if (typeof progress === "number") {
    progressEl.value = progress;
  } else {
    progressEl.removeAttribute("value");
  }
});

window.downloader.getPlatform().then(selectDefault);

winBtn.addEventListener("click", () => startDownload("win"));
macBtn.addEventListener("click", () => startDownload("mac"));

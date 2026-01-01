# Custom Line Chart / 客製化折線圖

[English](#english) | [中文](#中文)

## English

### Overview

This is a Windows/macOS desktop tool for creating customizable line charts. It supports Excel paste, multiple series, interval bands, and image export. Designed for non-technical users.

### Download

1. Go to **Releases** on the right side of the GitHub page.
2. Download the installer for your OS:
   - Windows: `NTOU_Tools_Setup.exe`
   - macOS: `NTOU_Tools.dmg`
3. Install and launch:
   - Windows: double-click the `.exe`.
   - macOS: open the `.dmg`, drag the app into **Applications**, then open it.

### Build (Windows/macOS)

1. Install build tools:
   ```
   python -m pip install matplotlib pyinstaller pillow
   ```
2. Run the build script (auto-detects the platform and runs the correct command line):
   ```
   python build_app.py
   ```
3. Output files:
   - Windows: `dist/LineChart.exe` and `dist/LineChart-windows.zip`
   - macOS: `dist/LineChart.app` and `dist/LineChart-macos.zip`

Note: You must build on the target OS (PyInstaller does not cross-compile). The image `messageImage_1767257219427.jpg` is used for the app icon and header.

### How to Use

1. Open the app by double-clicking the `.exe` file (Windows) or `.app` (macOS).
2. Paste your Excel data into the **Excel Paste** area and click **Apply from Excel**.
3. (Optional) Set X/Y axis units, colors, or interval bands.
4. Click **Plot** to preview the chart.
5. Click **Download Image** to save the chart.

### Excel Paste Format

#### 1) X, Y Columns

```
cm-1(X)	T(Y)
3997.43665	1.0051
3996.01357	1.00517
3994.59049	1.00465
3993.16741	1.00394
3991.74432	1.00349
3990.32124	1.00318
```

#### 2) Paired Columns (Two Lines)

```
cm-1(X)	T(Y)		cm-1(X1)	T(Y1)
3997.43665	1.0051		4000	1.0297
3996.01357	1.00517		3999	1.0297
3994.59049	1.00465		3998	1.0296
3993.16741	1.00394		3997	1.0296
3991.74432	1.00349		3996	1.0295
3990.32124	1.00318		3995	1.0294
3988.89816	1.00292		3994	1.0294
```

### Tips

- Use **Auto Colors** for multiple lines.
- Use **A4 Landscape** for PPT slides.
- Exported images are PNG by default.

## 中文

### 簡介

這是 Windows/macOS 桌面版工具，可快速製作客製化折線圖。支援 Excel 貼上、多序列、區間色帶、圖片匯出，適合一般使用者。

### 下載

1. 到 GitHub 專案頁右側的 **Releases**。
2. 下載對應作業系統的安裝檔：
   - Windows：`NTOU_Tools_Setup.exe`
   - macOS：`NTOU_Tools.dmg`
3. 安裝並開啟：
   - Windows：雙擊 `.exe`。
   - macOS：開啟 `.dmg`，把 App 拖到 **Applications** 後開啟。

### 打包（Windows/macOS）

1. 安裝打包工具：
   ```
   python -m pip install matplotlib pyinstaller pillow
   ```
2. 執行打包腳本（會自動偵測平台並使用對應指令）：
   ```
   python build_app.py
   ```
3. 產出檔案：
   - Windows：`dist/LineChart.exe` 與 `dist/LineChart-windows.zip`
   - macOS：`dist/LineChart.app` 與 `dist/LineChart-macos.zip`

注意：需在目標作業系統上打包（PyInstaller 無法跨平台打包）。`messageImage_1767257219427.jpg` 會作為 App 圖示與標頭圖片。

### 使用方式

1. 直接雙擊 `.exe`（Windows）或 `.app`（macOS）開啟程式。
2. 將 Excel 數據貼到「Excel 貼上」區塊，按「從 Excel 貼上套用」。
3. （可選）設定 X/Y 軸單位、顏色或區間色帶。
4. 按「繪製」預覽圖表。
5. 按「下載圖片」儲存圖片。

### Excel 貼上格式

#### 1) X, Y 兩欄

```
cm-1(X)	T(Y)
3997.43665	1.0051
3996.01357	1.00517
3994.59049	1.00465
3993.16741	1.00394
3991.74432	1.00349
3990.32124	1.00318
```

#### 2) 成對欄位（兩條線）

```
cm-1(X)	T(Y)		cm-1(X1)	T(Y1)
3997.43665	1.0051		4000	1.0297
3996.01357	1.00517		3999	1.0297
3994.59049	1.00465		3998	1.0296
3993.16741	1.00394		3997	1.0296
3991.74432	1.00349		3996	1.0295
3990.32124	1.00318		3995	1.0294
3988.89816	1.00292		3994	1.0294
```

### 使用小技巧

- 多條線建議勾選「自動配色」。
- 製作簡報建議使用「A4 橫式」匯出。
- 匯出圖片預設為 PNG。

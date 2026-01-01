# Custom Line Chart

This is a Windows/macOS desktop tool for creating customizable line charts. It supports Excel paste, multiple series, interval bands, and image export. Designed for non-technical users.

**Chinese guide:** [README.zh-TW.md](README.zh-TW.md)

## Download

1. Go to **Releases** on the right side of the GitHub page.
2. Download the installer for your OS:
   - Windows: `NTOU_Tools_Setup.exe`
   - macOS: `NTOU_Tools.dmg`
3. Install and launch:
   - Windows: double-click the `.exe`.
   - macOS: open the `.dmg`, drag the app into **Applications**, then open it.

## How to Use

1. Open the app by double‑clicking the `.exe` file.
2. Paste your Excel data into the **Excel Paste** area and click **Apply from Excel**.
3. (Optional) Set X/Y axis units, colors, or interval bands.
4. Click **Plot** to preview the chart.
5. Click **Download Image** to save the chart.

## Excel Paste Format

### 1) X, Y Columns

```
cm-1(X)	T(Y)
3997.43665	1.0051
3996.01357	1.00517
3994.59049	1.00465
3993.16741	1.00394
3991.74432	1.00349
3990.32124	1.00318
```

### 2) Paired Columns (Two Lines)cm-1(X) T(Y) cm-1(X1) T(Y1)

3997.43665 1.0051 4000 1.0297
3996.01357 1.00517 3999 1.0297
3994.59049 1.00465 3998 1.0296
3993.16741 1.00394 3997 1.0296
3991.74432 1.00349 3996 1.0295
3990.32124 1.00318 3995 1.0294
3988.89816 1.00292 3994 1.0294cm-1\tT\t\tcm-1\tT

```
3997.43665\t1.0051\t\t4000\t1.0297
```

## Tips

- Use **Auto Colors** for multiple lines.
- Use **A4 Landscape** for PPT slides.
- Exported images are PNG by default.

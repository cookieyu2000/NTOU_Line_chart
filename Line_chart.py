import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser

import matplotlib
matplotlib.use("TkAgg")
from matplotlib import rcParams
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

rcParams["font.sans-serif"] = ["PingFang TC", "Microsoft JhengHei", "Noto Sans CJK TC", "SimHei", "Arial Unicode MS"]
rcParams["axes.unicode_minus"] = False


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


def parse_csv_numbers(text, label):
    raw = [item.strip() for item in text.split(",") if item.strip()]
    if not raw:
        raise ValueError(f"{label} 內容為空")
    try:
        return [float(item) for item in raw]
    except ValueError as exc:
        raise ValueError(f"{label} 含有非數字") from exc


def parse_csv_strings(text, label):
    items = [item.strip() for item in text.split(",") if item.strip()]
    if not items:
        raise ValueError(f"{label} 內容為空")
    return items


def parse_interval_notes(text):
    notes = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 3:
            raise ValueError("區間備註格式需為：起點,終點,備註")
        start = parts[0]
        end = parts[1]
        label = ",".join(parts[2:])
        notes.append((start, end, label))
    return notes


def parse_excel_block(text):
    rows = [row for row in text.splitlines() if row.strip()]
    if not rows:
        raise ValueError("Excel 貼上內容為空")

    def clean_row(row):
        return [cell.strip() for cell in row.split("\t")]

    def is_number(value):
        try:
            float(value)
        except ValueError:
            return False
        return True

    table = [clean_row(row) for row in rows]
    if len(table) < 2:
        raise ValueError("Excel 貼上需至少包含標題列與一列數據")

    max_cols = max(len(row) for row in table)
    for row in table:
        if len(row) < max_cols:
            row.extend([""] * (max_cols - len(row)))
    empty_cols = [idx for idx in range(max_cols) if all(row[idx] == "" for row in table)]
    if empty_cols:
        table = [[cell for idx, cell in enumerate(row) if idx not in empty_cols] for row in table]

    header = table[0]
    data_rows = table[1:]
    has_header = any(cell and not is_number(cell) for cell in header)

    x_items = []
    x_values = []
    x_unit = ""
    series_defs = []
    series_names = set()

    if has_header:
        header = [cell.strip() for cell in header if cell.strip() != ""]
        if len(header) < 2:
            raise ValueError("Excel 標題列需包含 X 與至少一個序列名稱")

        max_cols = max(len(row) for row in data_rows)
        if len(header) % 2 == 0 and max_cols >= len(header):
            valid_xy = True
            for row in data_rows:
                for idx in range(len(header)):
                    if idx >= len(row) or row[idx] == "":
                        continue
                    if not is_number(row[idx]):
                        valid_xy = False
                        break
                if not valid_xy:
                    break
            if valid_xy:
                pairs = len(header) // 2
                x_candidates = []
                for pair_idx in range(pairs):
                    x_col = pair_idx * 2
                    y_col = x_col + 1
                    x_vals = []
                    y_vals = []
                    for row in data_rows:
                        if len(row) <= y_col:
                            continue
                        x_cell = row[x_col]
                        y_cell = row[y_col]
                        if x_cell == "" or y_cell == "":
                            continue
                        if not (is_number(x_cell) and is_number(y_cell)):
                            continue
                        x_vals.append(x_cell)
                        y_vals.append(y_cell)
                    if not x_vals or not y_vals or len(x_vals) != len(y_vals):
                        continue
                    series_name = header[y_col] or f"序列 {pair_idx + 1}"
                    if series_name in series_names:
                        series_name = f"{series_name}-{pair_idx + 1}"
                    series_names.add(series_name)
                    series_defs.append((series_name, ",".join(y_vals), [str(v) for v in x_vals]))
                    x_candidates.append(x_vals)
                    if not x_unit and header[x_col] and not is_number(header[x_col]):
                        x_unit = header[x_col]

                if not series_defs:
                    raise ValueError("Excel 貼上內容缺少可用的數據列")

                first_x = x_candidates[0]
                x_values = first_x
                x_items = first_x
            else:
                x_items = header[1:]
                for row in data_rows:
                    if not row or len(row) < 2:
                        continue
                    name = row[0].strip() or "序列"
                    values = [cell.strip() for cell in row[1:] if cell.strip() != ""]
                    if not values:
                        continue
                    series_defs.append((name, ",".join(values), None))
        else:
            x_items = header[1:]
            for row in data_rows:
                if not row or len(row) < 2:
                    continue
                name = row[0].strip() or "序列"
                values = [cell.strip() for cell in row[1:] if cell.strip() != ""]
                if not values:
                    continue
                series_defs.append((name, ",".join(values), None))
    else:
        first_row = table[0]
        if len(first_row) < 2:
            raise ValueError("Excel 貼上需至少包含 X 與 Y 兩欄")
        if not all(is_number(cell) for cell in first_row[:2]):
            raise ValueError("Excel 貼上內容格式不正確")
        x_vals = []
        y_vals = []
        for row in table:
            if len(row) < 2:
                continue
            if not (is_number(row[0]) and is_number(row[1])):
                continue
            x_vals.append(row[0])
            y_vals.append(row[1])
        if not x_vals or not y_vals:
            raise ValueError("Excel 貼上內容缺少數據列")
        x_values = x_vals
        x_items = x_vals
        series_defs.append(("序列 1", ",".join(y_vals), [str(v) for v in x_vals]))

    if not series_defs:
        raise ValueError("Excel 貼上內容缺少數據列")

    return x_items, x_values, x_unit, series_defs


class SeriesRow:
    def __init__(self, parent, index, remove_callback):
        self.frame = ttk.Frame(parent)
        self.enabled_var = tk.BooleanVar(value=True)
        self.name_var = tk.StringVar(value=f"序列 {index}")
        self.values_var = tk.StringVar()
        self.x_values = None

        ttk.Checkbutton(self.frame, variable=self.enabled_var).grid(row=0, column=0, padx=4)
        ttk.Entry(self.frame, textvariable=self.name_var, width=14).grid(row=0, column=1, padx=4)
        ttk.Entry(self.frame, textvariable=self.values_var, width=40).grid(row=0, column=2, padx=4)
        ttk.Button(self.frame, text="移除", command=remove_callback).grid(row=0, column=3, padx=4)

    def grid(self, **kwargs):
        self.frame.grid(**kwargs)

    def destroy(self):
        self.frame.destroy()


class LineChartApp:
    def __init__(self, root):
        self.root = root
        root.title("客製化折線圖")
        root.configure(background="#0d0b1a")

        self.series_rows = []

        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(".", font=("PingFang TC", 12), background="#0d0b1a", foreground="#f8f7ff")
        style.configure("TFrame", background="#0d0b1a")
        style.configure("TLabel", background="#0d0b1a", foreground="#f8f7ff")
        style.configure("Header.TLabel", font=("PingFang TC", 16, "bold"), foreground="#ffffff", background="#0d0b1a")
        style.configure("Hint.TLabel", font=("PingFang TC", 10), foreground="#c7b8ff", background="#0d0b1a")
        style.configure("Card.TLabelframe", background="#17112b", borderwidth=1, relief="groove")
        style.configure("Card.TLabelframe.Label", font=("PingFang TC", 12, "bold"), foreground="#ffe3a0", background="#17112b")
        style.configure("TEntry", fieldbackground="#1c1536", foreground="#f8f7ff")
        style.configure("TCheckbutton", background="#0d0b1a", foreground="#f8f7ff")
        style.map("TCheckbutton", background=[("active", "#0d0b1a")], foreground=[("active", "#ffffff")])
        style.configure("TButton", padding=(12, 6), background="#2a1d4a", foreground="#f8f7ff")
        style.map(
            "TButton",
            background=[("active", "#3c2b6a"), ("pressed", "#23153e")],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        style.configure("Accent.TButton", font=("PingFang TC", 12, "bold"), background="#ff7ad9", foreground="#1a0f2a")
        style.map(
            "Accent.TButton",
            background=[("active", "#ff9be6"), ("pressed", "#f05ac8")],
            foreground=[("active", "#1a0f2a"), ("pressed", "#1a0f2a")],
        )

        self.app_icon = None
        self.banner_image = None
        self.load_branding()

        main = ttk.Frame(root, padding=14)
        main.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        header = ttk.Frame(main)
        header.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        title_col = 0
        if self.banner_image:
            ttk.Label(header, image=self.banner_image).grid(row=0, column=0, padx=(0, 10))
            title_col = 1
        title = ttk.Label(header, text="折線圖設定面板", style="Header.TLabel")
        title.grid(row=0, column=title_col, sticky="w")

        pane = ttk.PanedWindow(main, orient="horizontal")
        pane.grid(row=1, column=0, columnspan=2, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        config_wrap = ttk.LabelFrame(pane, text="設定", padding=6, style="Card.TLabelframe")
        plot_frame = ttk.LabelFrame(pane, text="預覽", padding=12, style="Card.TLabelframe")
        config_wrap.configure(width=560)
        plot_frame.configure(width=560)

        pane.add(config_wrap, weight=1)
        pane.add(plot_frame, weight=1)

        root.minsize(1200, 720)
        pane.sashpos(0, 560)

        config_canvas = tk.Canvas(config_wrap, highlightthickness=0, bg="#17112b")
        config_scroll = ttk.Scrollbar(config_wrap, orient="vertical", command=config_canvas.yview)
        config_canvas.configure(yscrollcommand=config_scroll.set)
        config_canvas.pack(side="left", fill="both", expand=True)
        config_scroll.pack(side="right", fill="y")

        config = ttk.Frame(config_canvas)
        config_canvas.create_window((0, 0), window=config, anchor="nw", tags="config_window")

        def _on_configure(_event):
            config_canvas.configure(scrollregion=config_canvas.bbox("all"))
            config_canvas.itemconfigure("config_window", width=config_canvas.winfo_width())

        def _on_mousewheel(event):
            config_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        config.bind("<Configure>", _on_configure)
        config_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        config.columnconfigure(1, weight=1)
        config.columnconfigure(3, weight=1)

        ttk.Label(config, text="小提示：輸入完畢後按「繪製」即可預覽與下載。", style="Hint.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 8)
        )

        ttk.Label(config, text="X 軸項目（以逗號分隔）").grid(row=1, column=0, sticky="w")
        self.x_items_var = tk.StringVar()
        ttk.Entry(config, textvariable=self.x_items_var, width=52).grid(row=1, column=1, columnspan=3, sticky="we", pady=2)
        ttk.Label(config, text="例：一月,二月,三月", style="Hint.TLabel").grid(row=2, column=1, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(config, text="X 軸單位（選填）").grid(row=3, column=0, sticky="w")
        self.x_unit_var = tk.StringVar()
        self.x_unit_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config, text="顯示單位", variable=self.x_unit_enabled_var).grid(row=3, column=1, sticky="w", pady=2)
        ttk.Entry(config, textvariable=self.x_unit_var, width=16).grid(row=3, column=2, sticky="w", pady=2)
        ttk.Label(config, text="例：cm⁻¹、mm、秒", style="Hint.TLabel").grid(row=4, column=2, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(config, text="X 軸數值（選填）").grid(row=5, column=0, sticky="w")
        self.x_values_var = tk.StringVar()
        ttk.Entry(config, textvariable=self.x_values_var, width=16).grid(row=5, column=1, sticky="w", pady=2)
        ttk.Label(config, text="可與項目數量不同；數量相同會作為 X 座標", style="Hint.TLabel").grid(
            row=6, column=1, columnspan=3, sticky="w", pady=(0, 6)
        )

        ttk.Label(config, text="Y 軸刻度間距").grid(row=7, column=0, sticky="w")
        self.interval_var = tk.StringVar()
        ttk.Entry(config, textvariable=self.interval_var, width=12).grid(row=7, column=1, sticky="w", pady=2)
        ttk.Label(config, text="例：10（留空自動）", style="Hint.TLabel").grid(row=8, column=1, sticky="w", pady=(0, 6))

        ttk.Label(config, text="Y 軸最小值").grid(row=7, column=2, sticky="e")
        self.ymin_var = tk.StringVar()
        ttk.Entry(config, textvariable=self.ymin_var, width=10).grid(row=7, column=3, sticky="w", pady=2)

        ttk.Label(config, text="Y 軸最大值").grid(row=9, column=0, sticky="w")
        self.ymax_var = tk.StringVar()
        ttk.Entry(config, textvariable=self.ymax_var, width=12).grid(row=9, column=1, sticky="w", pady=2)

        self.allow_negative_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config, text="允許負值", variable=self.allow_negative_var).grid(row=9, column=2, columnspan=2, sticky="w")
        ttk.Label(config, text="Y 軸單位（選填）").grid(row=10, column=0, sticky="w")
        self.y_unit_var = tk.StringVar()
        self.y_unit_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config, text="顯示單位", variable=self.y_unit_enabled_var).grid(row=10, column=1, sticky="w", pady=2)
        ttk.Entry(config, textvariable=self.y_unit_var, width=16).grid(row=10, column=2, sticky="w", pady=2)
        ttk.Label(config, text="例：T、％、ppm", style="Hint.TLabel").grid(row=11, column=2, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(config, text="留空代表自動取資料最大/最小值", style="Hint.TLabel").grid(row=12, column=1, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Separator(config).grid(row=13, column=0, columnspan=4, sticky="we", pady=6)
        ttk.Label(config, text="Excel 貼上（含標題列）").grid(row=14, column=0, sticky="w")
        self.excel_text = tk.Text(config, width=52, height=4)
        self.excel_text.grid(row=14, column=1, columnspan=3, sticky="we", pady=2)
        self.excel_text.configure(background="#1c1536", foreground="#f8f7ff", insertbackground="#f8f7ff", relief="solid", borderwidth=1)
        ttk.Label(config, text="支援：X,Y 兩欄｜或 X1,Y1,X2,Y2 成對欄位", style="Hint.TLabel").grid(
            row=15, column=1, columnspan=3, sticky="w", pady=(0, 6)
        )
        ttk.Button(config, text="從 Excel 貼上套用", style="Accent.TButton", command=self.apply_excel).grid(
            row=16, column=1, sticky="w", pady=(0, 6)
        )

        ttk.Separator(config).grid(row=17, column=0, columnspan=4, sticky="we", pady=6)
        ttk.Label(config, text="X 軸區間色帶（每行一條）").grid(row=18, column=0, sticky="w")
        self.notes_text = tk.Text(config, width=52, height=5)
        self.notes_text.grid(row=18, column=1, columnspan=3, sticky="we", pady=2)
        self.notes_text.configure(background="#1c1536", foreground="#f8f7ff", insertbackground="#f8f7ff", relief="solid", borderwidth=1)
        ttk.Label(config, text="格式：起點,終點,備註　例：1,3,促銷期（可填X數值/項目名稱/序號）", style="Hint.TLabel").grid(
            row=19, column=1, columnspan=3, sticky="w", pady=(0, 6)
        )

        style_panel = ttk.LabelFrame(config, text="樣式設定", padding=8, style="Card.TLabelframe")
        style_panel.grid(row=20, column=0, columnspan=4, sticky="we", pady=(6, 4))
        ttk.Label(style_panel, text="折線顏色（單色）").grid(row=0, column=0, sticky="w")
        self.line_color_var = tk.StringVar()
        self.auto_color_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(style_panel, text="自動配色", variable=self.auto_color_var).grid(row=0, column=1, sticky="w")
        ttk.Entry(style_panel, textvariable=self.line_color_var, width=14).grid(row=0, column=2, sticky="w", padx=(4, 12))
        ttk.Button(style_panel, text="選擇", command=self.pick_line_color).grid(row=0, column=3, padx=(0, 12))
        ttk.Label(style_panel, text="自動配色時可留空", style="Hint.TLabel").grid(row=0, column=4, sticky="w")

        ttk.Label(style_panel, text="圖表背景").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.chart_bg_var = tk.StringVar()
        ttk.Entry(style_panel, textvariable=self.chart_bg_var, width=14).grid(row=1, column=1, sticky="w", padx=(4, 12), pady=(6, 0))
        ttk.Button(style_panel, text="選擇", command=self.pick_bg_color).grid(row=1, column=2, padx=(0, 12), pady=(6, 0))
        ttk.Label(style_panel, text="例：#0b1220", style="Hint.TLabel").grid(row=1, column=3, sticky="w", pady=(6, 0))

        ttk.Label(style_panel, text="匯出比例").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.export_ratio_var = tk.StringVar(value="A4 橫式")
        self.export_ratio_box = ttk.Combobox(style_panel, textvariable=self.export_ratio_var, state="readonly", width=12)
        self.export_ratio_box["values"] = ("A4 橫式", "A4 直式")
        self.export_ratio_box.grid(row=2, column=1, sticky="w", pady=(6, 0))
        ttk.Label(style_panel, text="直式可自行切換", style="Hint.TLabel").grid(row=2, column=3, sticky="w", pady=(6, 0))

        series_frame = ttk.LabelFrame(config, text="資料序列", padding=8, style="Card.TLabelframe")
        series_frame.grid(row=21, column=0, columnspan=4, sticky="we", pady=(8, 4))
        header = ttk.Frame(series_frame)
        header.grid(row=0, column=0, sticky="we")
        ttk.Label(header, text="顯示", style="Hint.TLabel").grid(row=0, column=0, padx=6, sticky="w")
        ttk.Label(header, text="名稱", style="Hint.TLabel").grid(row=0, column=1, padx=6, sticky="w")
        ttk.Label(header, text="數值（逗號分隔）", style="Hint.TLabel").grid(row=0, column=2, padx=6, sticky="w")
        self.series_container = ttk.Frame(series_frame)
        self.series_container.grid(row=1, column=0, sticky="we")

        controls = ttk.Frame(series_frame)
        controls.grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Button(controls, text="新增序列", command=self.add_series).grid(row=0, column=0, padx=2)
        ttk.Button(controls, text="移除未勾選", command=self.remove_unchecked).grid(row=0, column=1, padx=2)

        actions = ttk.Frame(config)
        actions.grid(row=22, column=0, columnspan=4, sticky="e", pady=(10, 0))
        ttk.Button(actions, text="載入範例", command=self.load_sample).grid(row=0, column=0, padx=4)
        ttk.Button(actions, text="繪製", style="Accent.TButton", command=self.plot).grid(row=0, column=1, padx=4)
        ttk.Button(actions, text="清除", command=self.clear).grid(row=0, column=2, padx=4)
        ttk.Button(actions, text="重置", command=self.reset).grid(row=0, column=3, padx=4)
        ttk.Button(actions, text="下載圖片", style="Accent.TButton", command=self.save_image).grid(row=0, column=4, padx=4)

        self.figure = Figure(figsize=(6, 4), dpi=100, facecolor="#1a1333")
        self.ax = self.figure.add_subplot(111)
        self.ax.set_facecolor("#1a1333")
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
        self.canvas.get_tk_widget().configure(background="#1a1333")
        plot_frame.rowconfigure(0, weight=1)
        plot_frame.columnconfigure(0, weight=1)

        self.sample_excel_text = ""
        sample_config = {}
        config_path = resource_path("sample_data.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as handle:
                    loaded = json.load(handle)
                    if isinstance(loaded, dict):
                        sample_config = loaded
            except (OSError, json.JSONDecodeError):
                sample_config = {}

        if not sample_config:
            sample_path = resource_path("sample_excel.txt")
            if os.path.exists(sample_path):
                try:
                    with open(sample_path, "r", encoding="utf-8") as handle:
                        self.sample_excel_text = handle.read().strip("\n")
                except OSError:
                    self.sample_excel_text = ""
            if not self.sample_excel_text:
                self.sample_excel_text = (
                    "cm-1\tT\t\tcm-1\tT\n"
                    "3997.43665\t1.0051\t\t4000\t1.0297\n"
                    "3996.01357\t1.00517\t\t3999\t1.0297\n"
                    "3994.59049\t1.00465\t\t3998\t1.0296\n"
                    "3993.16741\t1.00394\t\t3997\t1.0296\n"
                    "3991.74432\t1.00349\t\t3996\t1.0295\n"
                    "3990.32124\t1.00318\t\t3995\t1.0294\n"
                    "3988.89816\t1.00292\t\t3994\t1.0294\n"
                )

        self.sample_x_items = "一月,二月,三月,四月,五月"
        self.sample_x_values = "1,2,3,4,5"
        self.sample_x_unit = "cm-1"
        self.sample_interval = "10"
        self.sample_ymin = ""
        self.sample_ymax = ""
        self.sample_y_unit = "T"
        self.sample_allow_negative = True
        self.sample_notes = "1,2,促銷期\n3,4,新品期"
        self.sample_series = [
            ("營收", "35,50,65,80,95"),
            ("毛利", "20,30,45,60,70"),
        ]
        self.sample_line_color = ""
        self.sample_chart_bg = "#ffffff"
        self.sample_export_ratio = "A4 橫式"

        if sample_config:
            excel_block = str(sample_config.get("excel_block", "")).strip()
            if excel_block:
                self.sample_excel_text = excel_block

            if "x_items" in sample_config:
                x_items_list = sample_config.get("x_items") or []
                self.sample_x_items = ",".join(str(item) for item in x_items_list)
            if "x_values" in sample_config:
                x_values_list = sample_config.get("x_values") or []
                self.sample_x_values = ",".join(str(value) for value in x_values_list)
            if "x_unit" in sample_config:
                self.sample_x_unit = str(sample_config.get("x_unit") or "")
            if "interval" in sample_config:
                interval = sample_config.get("interval", "")
                self.sample_interval = "" if interval is None else str(interval)
            if "y_min" in sample_config:
                y_min = sample_config.get("y_min", "")
                self.sample_ymin = "" if y_min is None else str(y_min)
            if "y_max" in sample_config:
                y_max = sample_config.get("y_max", "")
                self.sample_ymax = "" if y_max is None else str(y_max)
            if "y_unit" in sample_config:
                self.sample_y_unit = str(sample_config.get("y_unit") or "")
            if "allow_negative" in sample_config:
                self.sample_allow_negative = bool(sample_config.get("allow_negative"))

            notes_list = sample_config.get("notes") or []
            if notes_list:
                note_lines = []
                for note in notes_list:
                    start = note.get("start", "")
                    end = note.get("end", "")
                    label = note.get("label", "")
                    note_lines.append(f"{start},{end},{label}".strip(","))
                self.sample_notes = "\n".join(note_lines)
            else:
                self.sample_notes = ""

            series_list = sample_config.get("series") or []
            if series_list:
                series_defs = []
                for series in series_list:
                    name = series.get("name") or "序列"
                    values = series.get("values") or []
                    x_vals = series.get("x_values") or None
                    values_text = ",".join(str(value) for value in values)
                    if x_vals:
                        x_vals_text = [str(value) for value in x_vals]
                        series_defs.append((name, values_text, x_vals_text))
                    else:
                        series_defs.append((name, values_text))
                self.sample_series = series_defs

            style_config = sample_config.get("style") or {}
            if "line_color" in style_config:
                self.sample_line_color = str(style_config.get("line_color") or "")
            if "chart_bg" in style_config:
                self.sample_chart_bg = str(style_config.get("chart_bg") or self.sample_chart_bg)
            if "export_ratio" in style_config:
                self.sample_export_ratio = str(style_config.get("export_ratio") or self.sample_export_ratio)

        if self.sample_excel_text:
            try:
                x_items, x_values, x_unit, series_defs = parse_excel_block(self.sample_excel_text)
            except ValueError:
                series_defs = None
            if series_defs:
                self.sample_x_items = ""
                self.sample_x_values = ""
                if not self.sample_x_unit:
                    self.sample_x_unit = x_unit or ""
                self.sample_interval = ""
                self.sample_ymin = ""
                self.sample_ymax = ""
                if not self.sample_y_unit:
                    self.sample_y_unit = "T"
                self.sample_allow_negative = True
                self.sample_notes = ""
                self.sample_series = series_defs

        self.default_x_items = self.sample_x_items
        self.default_x_values = self.sample_x_values
        self.default_x_unit = self.sample_x_unit
        self.default_interval = self.sample_interval
        self.default_ymin = self.sample_ymin
        self.default_ymax = self.sample_ymax
        self.default_y_unit = self.sample_y_unit
        self.default_allow_negative = self.sample_allow_negative
        self.default_notes = self.sample_notes
        self.default_line_color = self.sample_line_color
        self.default_chart_bg = self.sample_chart_bg
        self.default_export_ratio = self.sample_export_ratio

        self.apply_sample_data(self.sample_series)

    def load_branding(self):
        image_path = resource_path("messageImage_1767257219427.jpg")
        if not Image or not ImageTk or not os.path.exists(image_path):
            return
        try:
            image = Image.open(image_path)
            image = image.convert("RGBA")
        except OSError:
            return

        resample = getattr(Image, "Resampling", Image).LANCZOS
        icon_image = image.copy()
        icon_image.thumbnail((256, 256), resample)
        self.app_icon = ImageTk.PhotoImage(icon_image)
        try:
            self.root.iconphoto(True, self.app_icon)
        except tk.TclError:
            pass

        banner_image = image.copy()
        banner_image.thumbnail((96, 96), resample)
        self.banner_image = ImageTk.PhotoImage(banner_image)

    def add_series(self):
        index = len(self.series_rows) + 1
        row = SeriesRow(self.series_container, index, lambda r=None: self.remove_series(row))
        row.grid(row=index - 1, column=0, sticky="w")
        self.series_rows.append(row)
        return row

    def set_series_rows(self, series_defs):
        while self.series_rows:
            self.series_rows[0].destroy()
            self.series_rows.pop(0)
        if not series_defs:
            self.add_series()
            return
        for series in series_defs:
            if len(series) == 3:
                name, values, x_values = series
            else:
                name, values = series
                x_values = None
            row = self.add_series()
            row.name_var.set(name)
            row.values_var.set(values)
            row.x_values = x_values

    def remove_series(self, row):
        if row not in self.series_rows:
            return
        self.series_rows.remove(row)
        row.destroy()
        for i, series in enumerate(self.series_rows):
            series.name_var.set(f"序列 {i + 1}")

    def remove_unchecked(self):
        kept = []
        for row in self.series_rows:
            if row.enabled_var.get():
                kept.append(row)
            else:
                row.destroy()
        self.series_rows = kept
        if not self.series_rows:
            self.add_series()

    def clear(self):
        self.x_items_var.set("")
        self.x_unit_var.set("")
        self.y_unit_var.set("")
        self.x_values_var.set("")
        self.interval_var.set("")
        self.ymin_var.set("")
        self.ymax_var.set("")
        self.allow_negative_var.set(True)
        self.notes_text.delete("1.0", tk.END)
        self.excel_text.delete("1.0", tk.END)
        for row in self.series_rows:
            row.values_var.set("")
        self.ax.clear()
        chart_bg = self.chart_bg_var.get().strip() or "#1a1333"
        self.figure.set_facecolor(chart_bg)
        self.ax.set_facecolor(chart_bg)
        self.canvas.get_tk_widget().configure(background=chart_bg)
        self.canvas.draw()

    def reset(self):
        self.x_items_var.set(self.default_x_items)
        self.x_unit_var.set(self.default_x_unit)
        self.x_unit_enabled_var.set(True)
        self.y_unit_var.set(self.default_y_unit)
        self.y_unit_enabled_var.set(True)
        self.export_ratio_var.set(self.default_export_ratio)
        self.x_values_var.set(self.default_x_values)
        self.interval_var.set(self.default_interval)
        self.ymin_var.set(self.default_ymin)
        self.ymax_var.set(self.default_ymax)
        self.allow_negative_var.set(self.default_allow_negative)
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", self.default_notes)
        self.excel_text.delete("1.0", tk.END)
        if self.sample_excel_text:
            self.excel_text.insert("1.0", self.sample_excel_text)
        self.set_series_rows(self.sample_series)
        self.line_color_var.set(self.default_line_color)
        self.auto_color_var.set(True)
        self.chart_bg_var.set(self.default_chart_bg)
        self.ax.clear()
        self.figure.set_facecolor(self.default_chart_bg)
        self.ax.set_facecolor(self.default_chart_bg)
        self.canvas.get_tk_widget().configure(background=self.default_chart_bg)
        self.canvas.draw()

    def apply_sample_data(self, series_defs):
        self.x_items_var.set(self.sample_x_items)
        self.x_unit_var.set(self.sample_x_unit)
        self.x_unit_enabled_var.set(True)
        self.y_unit_var.set(self.sample_y_unit)
        self.y_unit_enabled_var.set(True)
        self.export_ratio_var.set(self.sample_export_ratio)
        self.x_values_var.set(self.sample_x_values)
        self.interval_var.set(self.sample_interval)
        self.ymin_var.set(self.sample_ymin)
        self.ymax_var.set(self.sample_ymax)
        self.allow_negative_var.set(self.sample_allow_negative)
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", self.sample_notes)
        self.excel_text.delete("1.0", tk.END)
        if self.sample_excel_text:
            self.excel_text.insert("1.0", self.sample_excel_text)
        self.line_color_var.set(self.sample_line_color)
        self.auto_color_var.set(True)
        self.chart_bg_var.set(self.sample_chart_bg)
        self.set_series_rows(series_defs)

    def load_sample(self):
        self.apply_sample_data(self.sample_series)

    def apply_excel(self):
        try:
            x_items, x_values, x_unit, series_defs = parse_excel_block(self.excel_text.get("1.0", tk.END))
        except ValueError as exc:
            messagebox.showerror("輸入錯誤", str(exc))
            return
        self.x_items_var.set(",".join(x_items))
        if x_values:
            self.x_values_var.set(",".join(x_values))
        else:
            self.x_values_var.set("")
        if x_unit:
            self.x_unit_var.set(x_unit)
            self.x_unit_enabled_var.set(True)
        if len(series_defs) > 1:
            self.auto_color_var.set(True)
        self.set_series_rows(series_defs)

    def normalize_color(self, value, label):
        value = value.strip()
        if not value:
            return ""
        try:
            self.root.winfo_rgb(value)
        except tk.TclError as exc:
            raise ValueError(f"{label} 格式不正確") from exc
        return value

    def color_to_rgb(self, value):
        r, g, b = self.root.winfo_rgb(value)
        return r / 65535.0, g / 65535.0, b / 65535.0

    def contrast_color(self, value):
        r, g, b = self.color_to_rgb(value)
        luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
        return "#f5f5f5" if luminance < 0.5 else "#111111"

    def blend_color(self, fg, bg, alpha):
        fr, fg_c, fb = self.color_to_rgb(fg)
        br, bg_c, bb = self.color_to_rgb(bg)
        r = fr * alpha + br * (1 - alpha)
        g = fg_c * alpha + bg_c * (1 - alpha)
        b = fb * alpha + bb * (1 - alpha)
        return f"#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}"

    def pick_line_color(self):
        color = colorchooser.askcolor(title="選擇折線顏色")[1]
        if color:
            self.line_color_var.set(color)

    def pick_bg_color(self):
        color = colorchooser.askcolor(title="選擇圖表背景顏色")[1]
        if color:
            self.chart_bg_var.set(color)

    def save_image(self):
        file_path = filedialog.asksaveasfilename(
            title="儲存圖表圖片",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png"), ("JPEG image", "*.jpg;*.jpeg"), ("All files", "*.*")],
        )
        if not file_path:
            return
        original_size = self.figure.get_size_inches()
        ratio = self.export_ratio_var.get()
        if ratio == "A4 直式":
            target_size = (8.27, 11.69)
        else:
            target_size = (11.69, 8.27)
        self.figure.set_size_inches(*target_size)
        self.figure.savefig(file_path, dpi=100, bbox_inches="tight")
        self.figure.set_size_inches(*original_size)
        messagebox.showinfo("完成", f"圖片已儲存：{file_path}")

    def plot(self):
        x_items = []
        if self.x_items_var.get().strip():
            try:
                x_items = parse_csv_strings(self.x_items_var.get(), "X 軸項目")
            except ValueError as exc:
                messagebox.showerror("輸入錯誤", str(exc))
                return

        x_values = []
        if self.x_values_var.get().strip():
            try:
                x_values = parse_csv_numbers(self.x_values_var.get(), "X 軸數值")
            except ValueError as exc:
                messagebox.showerror("輸入錯誤", str(exc))
                return

        enabled_rows = [row for row in self.series_rows if row.enabled_var.get()]
        if not enabled_rows:
            messagebox.showerror("輸入錯誤", "尚未勾選任何序列")
            return

        y_values_list = []
        series_names = []
        for row in enabled_rows:
            try:
                y_values = parse_csv_numbers(row.values_var.get(), row.name_var.get() or "序列")
            except ValueError as exc:
                messagebox.showerror("輸入錯誤", str(exc))
                return
            if row.x_values:
                if len(y_values) != len(row.x_values):
                    messagebox.showerror("輸入錯誤", f"{row.name_var.get()} 數值數量需與 X 數值相同")
                    return
            else:
                if not x_items:
                    messagebox.showerror("輸入錯誤", "請輸入 X 軸項目或使用 Excel 貼上")
                    return
                if len(y_values) != len(x_items):
                    messagebox.showerror("輸入錯誤", f"{row.name_var.get()} 數值數量需與 X 軸項目相同")
                    return
            y_values_list.append(y_values)
            series_names.append(row.name_var.get())

        if not self.allow_negative_var.get():
            for y_values in y_values_list:
                if any(value < 0 for value in y_values):
                    messagebox.showerror("輸入錯誤", "已勾選不允許負值")
                    return

        try:
            interval = float(self.interval_var.get()) if self.interval_var.get().strip() else None
        except ValueError:
            messagebox.showerror("輸入錯誤", "Y 軸刻度間距需為數字")
            return

        try:
            ymin = float(self.ymin_var.get()) if self.ymin_var.get().strip() else None
            ymax = float(self.ymax_var.get()) if self.ymax_var.get().strip() else None
        except ValueError:
            messagebox.showerror("輸入錯誤", "Y 軸最大/最小值需為數字")
            return

        try:
            notes = parse_interval_notes(self.notes_text.get("1.0", tk.END))
        except ValueError as exc:
            messagebox.showerror("輸入錯誤", str(exc))
            return

        all_values = [value for series in y_values_list for value in series]
        data_min = min(all_values)
        data_max = max(all_values)
        if ymin is None:
            ymin = data_min
        if ymax is None:
            ymax = data_max
        if not self.allow_negative_var.get():
            ymin = max(0, ymin)

        try:
            line_color = self.normalize_color(self.line_color_var.get(), "折線顏色")
            chart_bg = self.normalize_color(self.chart_bg_var.get(), "圖表背景") or "#0f1217"
        except ValueError as exc:
            messagebox.showerror("輸入錯誤", str(exc))
            return

        if not x_items and enabled_rows and enabled_rows[0].x_values:
            x_items = [str(value) for value in enabled_rows[0].x_values]

        x_positions = list(range(len(x_items)))
        xtick_positions = x_positions
        xtick_labels = x_items

        use_numeric_x = False
        numeric_x_values = []
        if x_values and x_items and len(x_values) == len(x_items):
            use_numeric_x = True
            numeric_x_values = x_values
        else:
            if x_items:
                try:
                    numeric_x_values = [float(value) for value in x_items]
                    use_numeric_x = True
                except ValueError:
                    use_numeric_x = False

        self.ax.clear()
        self.figure.set_facecolor(chart_bg)
        self.ax.set_facecolor(chart_bg)
        self.canvas.get_tk_widget().configure(background=chart_bg)
        contrast = self.contrast_color(chart_bg)
        grid_color = self.blend_color(contrast, chart_bg, 0.35)

        band_spans = []
        if notes:
            band_colors = [
                "#1f77b4",
                "#ff7f0e",
                "#2ca02c",
                "#d62728",
                "#9467bd",
                "#8c564b",
            ]

            def resolve_x_boundary(value):
                value = value.strip()
                if not value:
                    raise ValueError("區間起點/終點不可為空")
                try:
                    numeric = float(value)
                except ValueError:
                    if value in x_items:
                        return x_positions[x_items.index(value)]
                    raise ValueError(f"找不到對應的 X 軸項目：{value}")
                if x_values and len(x_values) == len(x_items):
                    return numeric
                if 0 <= numeric <= len(x_positions) - 1:
                    return numeric
                if 1 <= numeric <= len(x_positions):
                    return numeric - 1
                return numeric

            try:
                for idx, (start_raw, end_raw, label) in enumerate(notes):
                    start = resolve_x_boundary(start_raw)
                    end = resolve_x_boundary(end_raw)
                    if start > end:
                        start, end = end, start
                    color = band_colors[idx % len(band_colors)]
                    band_spans.append((start, end))
                    self.ax.axvspan(start, end, facecolor=color, alpha=0.18, label=f"{label}（{start:g}~{end:g}）")
            except ValueError as exc:
                messagebox.showerror("輸入錯誤", str(exc))
                return

        series_x_all = []
        palette = ["#60a5fa", "#f59e0b", "#34d399", "#f472b6", "#a78bfa", "#f97316"]
        use_auto_color = self.auto_color_var.get()
        default_single_color = "#e11d48"
        for idx, (row, y_values, name) in enumerate(zip(enabled_rows, y_values_list, series_names)):
            if row.x_values:
                series_x = [float(value) for value in row.x_values]
                use_numeric_x = True
            elif use_numeric_x and numeric_x_values:
                series_x = numeric_x_values
            else:
                series_x = x_positions
            series_x_all.extend(series_x)
            marker = "o" if len(y_values) <= 60 else None
            if use_auto_color and len(y_values_list) > 1:
                color = palette[idx % len(palette)]
            elif len(y_values_list) == 1 and not line_color:
                color = default_single_color
            else:
                color = line_color if line_color else contrast
            self.ax.plot(series_x, y_values, marker=marker, label=name, color=color)

        if series_x_all:
            self.ax.set_xlim(min(series_x_all), max(series_x_all))
            if len(series_x_all) >= 2 and series_x_all[0] > series_x_all[-1]:
                self.ax.invert_xaxis()

        if use_numeric_x:
            self.ax.xaxis.set_major_locator(MaxNLocator(nbins=8))
        else:
            self.ax.set_xticks(xtick_positions)
            self.ax.set_xticklabels(xtick_labels)
            if len(xtick_labels) > 8:
                self.ax.tick_params(axis="x", labelrotation=45)
        x_unit = self.x_unit_var.get().strip() if self.x_unit_enabled_var.get() else ""
        if x_unit:
            self.ax.set_xlabel(x_unit, color=contrast)
        y_unit = self.y_unit_var.get().strip() if self.y_unit_enabled_var.get() else ""
        if y_unit:
            self.ax.set_ylabel(y_unit, color=contrast)

        self.ax.set_ylim(ymin, ymax)
        if interval:
            ticks = []
            current = ymin
            while current <= ymax + 1e-9:
                ticks.append(current)
                current += interval
            self.ax.set_yticks(ticks)

        self.ax.grid(True, linestyle="--", alpha=0.5, color=grid_color)
        self.ax.tick_params(colors=contrast)
        for spine in self.ax.spines.values():
            spine.set_color(grid_color)
        legend = self.ax.legend()
        if legend:
            for text in legend.get_texts():
                text.set_color(contrast)

        if band_spans and series_x_all:
            min_span = min(span[0] for span in band_spans)
            max_span = max(span[1] for span in band_spans)
            self.ax.set_xlim(min(min(series_x_all), min_span), max(max(series_x_all), max_span))

        self.canvas.draw()


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    app = LineChartApp(root)
    root.mainloop()

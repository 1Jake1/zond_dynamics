import sys
import random
import string
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib import rcParams
from pathlib import Path

from calc import DataLoader

BG = "#f0f2f5"
FG = "#1a1a2e"
ACCENT = "#2563eb"
ACCENT_HOVER = "#1d4ed8"
PANEL_BG = "#ffffff"
HEADER_BG = "#e8eaf0"
GRID_COLOR = "#d1d5db"
CHART_BG = "#fafbfc"


def _style_app():
    style = ttk.Style()
    style.theme_use("clam")

    style.configure(".", background=BG, foreground=FG, font=("Segoe UI", 10))
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=FG, font=("Segoe UI", 10))
    style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=(12, 6))
    style.map("TButton",
              background=[("active", ACCENT_HOVER), ("!active", ACCENT)],
              foreground=[("active", "#ffffff"), ("!active", "#ffffff")])

    style.configure("Toolbar.TFrame", background="#1a1a2e")
    style.configure("Toolbar.TButton", background="#2563eb", foreground="#ffffff",
                     font=("Segoe UI", 10, "bold"), padding=(14, 7))
    style.map("Toolbar.TButton",
              background=[("active", "#3b82f6"), ("!active", "#2563eb")],
              foreground=[("active", "#ffffff"), ("!active", "#ffffff")])

    style.configure("Header.TLabel", background="#1a1a2e", foreground="#ffffff",
                     font=("Segoe UI", 11, "bold"))
    style.configure("Status.TLabel", background="#e2e8f0", foreground="#475569",
                     font=("Segoe UI", 9), padding=(10, 4))
    style.configure("Info.TLabel", background=BG, foreground="#64748b",
                     font=("Segoe UI", 9))

    style.configure("TLabelframe", background=PANEL_BG, foreground=FG,
                     font=("Segoe UI", 10, "bold"), relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=PANEL_BG, foreground=ACCENT,
                     font=("Segoe UI", 10, "bold"))

    style.configure("Treeview", background=PANEL_BG, foreground=FG,
                     fieldbackground=PANEL_BG, font=("Consolas", 10), rowheight=24)
    style.configure("Treeview.Heading", background=HEADER_BG, foreground=FG,
                     font=("Segoe UI", 10, "bold"), relief="flat")
    style.map("Treeview", background=[("selected", ACCENT)],
              foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading", background=[("active", "#d1d5db")])

    style.configure("TPanedwindow", background=BG)
    style.configure("Sash", sashthickness=6, background=GRID_COLOR)

    rcParams["font.family"] = "Segoe UI"
    rcParams["axes.facecolor"] = CHART_BG
    rcParams["figure.facecolor"] = CHART_BG
    rcParams["axes.edgecolor"] = GRID_COLOR
    rcParams["axes.labelcolor"] = FG
    rcParams["xtick.color"] = "#64748b"
    rcParams["ytick.color"] = "#64748b"


class DinamikaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка данных динамики")
        self.root.geometry("1500x850")
        self.root.minsize(1000, 600)
        self.root.configure(bg=BG)

        self.loader = DataLoader()
        self._file_path = None

        _style_app()
        self._build_ui()

    def _build_ui(self):
        self._build_menu()
        self._build_toolbar()
        self._build_content()
        self._build_statusbar()

    def _build_menu(self):
        menubar = tk.Menu(self.root, font=("Segoe UI", 10))
        file_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10))
        file_menu.add_command(label="Открыть Excel...", command=self.load_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Сохранить результат...", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.root.config(menu=menubar)
        self.root.bind("<Control-o>", lambda e: self.load_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())

    def _build_toolbar(self):
        tb = ttk.Frame(self.root, style="Toolbar.TFrame")
        tb.pack(fill=tk.X, padx=0, pady=0)

        ttk.Label(tb, text="  Динамика", style="Header.TLabel").pack(side=tk.LEFT, padx=(12, 20))

        ttk.Button(tb, text="Открыть", style="Toolbar.TButton",
                   command=self.load_file).pack(side=tk.LEFT, padx=3, pady=6)
        ttk.Button(tb, text="Рассчитать", style="Toolbar.TButton",
                   command=self.calculate).pack(side=tk.LEFT, padx=3, pady=6)
        ttk.Button(tb, text="Сохранить", style="Toolbar.TButton",
                   command=self.save_file).pack(side=tk.LEFT, padx=3, pady=6)

        self.file_label = ttk.Label(tb, text="Файл не загружен", style="Header.TLabel")
        self.file_label.pack(side=tk.RIGHT, padx=15)

    def _build_content(self):
        main_paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 0))

        top_paned = ttk.PanedWindow(main_paned, orient=tk.HORIZONTAL)
        main_paned.add(top_paned, weight=1)

        left = ttk.LabelFrame(top_paned, text=" Динамика — Исходные данные ")
        top_paned.add(left, weight=3)
        self.tree_source = self._make_tree(left)

        right = ttk.LabelFrame(top_paned, text=" Тарировка — Калибровочная кривая ")
        top_paned.add(right, weight=2)
        self.tree_calib = self._make_tree(right)

        raw_chart_frame = ttk.LabelFrame(top_paned, text=" Тугрики от времени ")
        top_paned.add(raw_chart_frame, weight=3)

        self.raw_fig = Figure(figsize=(5, 3), dpi=100)
        self.raw_ax = self.raw_fig.add_subplot(111)
        self.raw_canvas = FigureCanvasTkAgg(self.raw_fig, master=raw_chart_frame)
        self.raw_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        raw_tb = ttk.Frame(raw_chart_frame)
        raw_tb.pack(fill=tk.X, padx=4, pady=(0, 4))
        self.raw_toolbar = NavigationToolbar2Tk(self.raw_canvas, raw_tb)
        self.raw_toolbar.update()

        bot_paned = ttk.PanedWindow(main_paned, orient=tk.HORIZONTAL)
        main_paned.add(bot_paned, weight=3)

        res = ttk.LabelFrame(bot_paned, text=" Результат расчёта ")
        bot_paned.add(res, weight=2)
        self.tree_result = self._make_tree(res)

        chart = ttk.LabelFrame(bot_paned, text=" График: Перемещение от времени ")
        bot_paned.add(chart, weight=3)

        self.result_fig = Figure(figsize=(7, 4), dpi=100)
        self.result_ax = self.result_fig.add_subplot(111)
        self.result_canvas = FigureCanvasTkAgg(self.result_fig, master=chart)
        self.result_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        tb_frame = ttk.Frame(chart)
        tb_frame.pack(fill=tk.X, padx=4, pady=(0, 4))
        self.result_toolbar = NavigationToolbar2Tk(self.result_canvas, tb_frame)
        self.result_toolbar.update()

    def _build_statusbar(self):
        sb = ttk.Frame(self.root, style="Status.TFrame")
        sb.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_var = tk.StringVar(value="Готово")
        ttk.Label(sb, textvariable=self.status_var, style="Status.TLabel").pack(side=tk.LEFT)

        self.stats_var = tk.StringVar(value="")
        ttk.Label(sb, textvariable=self.stats_var, style="Status.TLabel").pack(side=tk.RIGHT)

    def _make_tree(self, parent):
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        tree = ttk.Treeview(container, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        return tree

    def _populate_tree(self, tree, df, max_rows=500):
        tree.delete(*tree.get_children())
        if df is None or df.empty:
            return
        cols = list(df.columns)
        tree["columns"] = cols
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=130, minwidth=80, anchor=tk.CENTER)
        for i, (_, row) in enumerate(df.head(max_rows).iterrows()):
            vals = [str(v) if pd.notna(v) else "" for v in row]
            tag = "even" if i % 2 == 0 else "odd"
            tree.insert("", tk.END, values=vals, tags=(tag,))
        tree.tag_configure("even", background="#f8fafc")
        tree.tag_configure("odd", background="#ffffff")

    def load_file(self):
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not path:
            return
        try:
            self.status_var.set("Загрузка файла...")
            self.file_label.configure(text="Загрузка...")
            self.root.update_idletasks()

            self.loader.load_excel(path)
            self._file_path = Path(path)

            self._populate_tree(self.tree_source, self.loader.source_data, max_rows=500)
            self._populate_tree(self.tree_calib, self.loader.calib_data, max_rows=500)
            self.tree_result.delete(*self.tree_result.get_children())

            self.result_ax.clear()
            self.result_ax.text(0.5, 0.5, "Нажмите «Рассчитать»",
                                ha="center", va="center", transform=self.result_ax.transAxes,
                                fontsize=13, color="#94a3b8", style="italic")
            self.result_ax.set_axis_off()
            self.result_fig.tight_layout()
            self.result_canvas.draw()

            self._draw_raw_chart()

            src_n = len(self.loader.source_data) if self.loader.source_data is not None else 0
            cal_n = len(self.loader.calib_data) if self.loader.calib_data is not None else 0
            self.file_label.configure(text=self._file_path.name)
            self.status_var.set(f"Загружено: {self._file_path.name}")
            self.stats_var.set(f"Исходных: {src_n}  |  Калибровка: {cal_n} точек")

            self.calculate()
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))
            self.status_var.set("Ошибка загрузки")
            self.file_label.configure(text="Ошибка")

    def calculate(self):
        if self.loader.source_data is None or self.loader.calib_data is None:
            messagebox.showwarning("Внимание", "Сначала загрузите Excel-файл")
            return

        self.status_var.set("Выполнение расчёта...")
        self.root.update_idletasks()

        try:
            self.loader.calculate()
            self._populate_tree(self.tree_result, self.loader.result_df, max_rows=500)
            self._draw_result_chart()

            n = len(self.loader.result_df)
            mn = self.loader.result_df.iloc[:, 1].min()
            mx = self.loader.result_df.iloc[:, 1].max()
            self.status_var.set(f"Расчёт завершён")
            self.stats_var.set(f"Точек: {n}  |  Диапазон: {mn} — {mx} мм")
        except Exception as e:
            messagebox.showerror("Ошибка расчёта", str(e))
            self.status_var.set("Ошибка расчёта")

    def _draw_result_chart(self):
        self.result_ax.clear()
        df = self.loader.result_df
        if df is not None and not df.empty:
            self.result_ax.plot(df["Время, мсек"], df["Перемещение, мм"],
                                linewidth=0.8, color="#2196F3")
            self.result_ax.set_xlabel("Время, мсек")
            self.result_ax.set_ylabel("Перемещение, мм")
            self.result_ax.set_title("Перемещение от времени")
            self.result_ax.grid(True, alpha=0.3)
        self.result_fig.tight_layout()
        self.result_canvas.draw()

    def _draw_raw_chart(self):
        self.raw_ax.clear()
        df = self.loader.source_data
        if df is not None and not df.empty:
            self.raw_ax.plot(df["Время, мсек"], df["Тугрики"],
                             linewidth=0.6, color="#4CAF50")
            self.raw_ax.set_xlabel("Время, мсек")
            self.raw_ax.set_ylabel("Тугрики")
            self.raw_ax.set_title("Тугрики от времени")
            self.raw_ax.grid(True, alpha=0.3)
        else:
            self.raw_ax.text(0.5, 0.5, "Нет данных",
                             ha="center", va="center", transform=self.raw_ax.transAxes,
                             fontsize=12, color="#94a3b8")
        self.raw_fig.tight_layout()
        self.raw_canvas.draw()

    def save_file(self):
        if self.loader.result_df is None or self.loader.result_df.empty:
            messagebox.showwarning("Внимание", "Нет данных для сохранения. Сначала выполните расчёт.")
            return

        code = ''.join(random.choices(string.ascii_lowercase + string.digits, k=4))
        if self._file_path:
            stem = self._file_path.stem
        else:
            stem = "Результат"
        default_name = f"{stem}_результат_{code}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Сохранить результат",
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")],
            initialfile=default_name
        )
        if not path:
            return

        try:
            self.status_var.set("Сохранение...")
            self.root.update_idletasks()

            wb_out = openpyxl.Workbook()
            ws_out = wb_out.active
            ws_out.title = "Результат"

            ws_out["A1"] = "Динамика в тугриках"
            ws_out["D1"] = "Тарировка"
            ws_out["S1"] = "Динамика в мм"

            ws_out["A2"] = "Время, мсек"
            ws_out["B2"] = "Тугрики"
            ws_out["D2"] = "Перемещение, мм"
            ws_out["E2"] = "Тугрики"
            ws_out["S2"] = "Время, мсек"
            ws_out["T2"] = "Перемещение, мм"

            for i, (_, row) in enumerate(self.loader.source_data.iterrows()):
                ws_out.cell(row=i + 3, column=1, value=row["Время, мсек"])
                ws_out.cell(row=i + 3, column=2, value=row["Тугрики"])

            for i, (_, row) in enumerate(self.loader.calib_data.iterrows()):
                ws_out.cell(row=i + 3, column=4, value=row["Перемещение, мм"])
                ws_out.cell(row=i + 3, column=5, value=row["Тугрики"])

            for i, (_, row) in enumerate(self.loader.result_df.iterrows()):
                ws_out.cell(row=i + 3, column=19, value=row["Время, мсек"])
                ws_out.cell(row=i + 3, column=20, value=row["Перемещение, мм"])

            for col, w in {"A": 15, "B": 12, "D": 16, "E": 12, "S": 15, "T": 16}.items():
                ws_out.column_dimensions[col].width = w

            wb_out.save(path)
            wb_out.close()
            self.status_var.set(f"Сохранено: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))
            self.status_var.set("Ошибка сохранения")


def main():
    root = tk.Tk()
    app = DinamikaApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

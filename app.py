"""
Soylu Makine - Pres Kapasite Hesaplama Uygulaması
Arayüz: CustomTkinter  |  Hesap motoru: capacity_engine.py
"""
import json
import os
import sys
import threading
import traceback
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from capacity_engine import CapacityEngine, normalize_machine

APP_TITLE = "Soylu Makine – Pres Kapasite Hesaplama"
DEFAULT_PLAN_SHEET = "Diğer müşteriler DAHİL"
DEFAULT_OP_SHEET = "Uyum Operasyon Bazlı Saniyeler"
ALT_COLS = ["alternatif 1", "alternatif 2", "alternatif 3", "alternatif 4"]
NO_MACHINE = "(TANIMSIZ)"

COLOR_OK = "#2e9d5b"
COLOR_WARN = "#e8a33d"
COLOR_OVER = "#d64545"
COLOR_BEFORE = "#9aa5b1"
COLOR_AFTER = "#3b7dd8"
WARN_LIMIT = 85.0    # % bu değerin üstü turuncu
OVER_LIMIT = 100.0   # % bu değerin üstü kırmızı


def app_dir() -> str:
    """.exe olarak çalışırken exe'nin klasörü, script olarak çalışırken dosyanın klasörü."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


SETTINGS_FILE = os.path.join(app_dir(), "kapasite_ayarlari.json")


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("default_capacity", 450.0)
        data.setdefault("capacities", {})
        return data
    except Exception:
        return {"default_capacity": 450.0, "capacities": {}, "plan_file": "", "op_file": ""}


def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("Uyarı", f"Ayarlar kaydedilemedi:\n{e}")


def fmt_num(v, digits=1):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "-"
    if isinstance(v, (int, float, np.integer, np.floating)):
        return f"{v:,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(v)


def load_tag(pct):
    if pct is None or (isinstance(pct, float) and np.isnan(pct)):
        return ""
    if pct > OVER_LIMIT:
        return "over"
    if pct > WARN_LIMIT:
        return "warn"
    return ""


def bar_color(pct):
    tag = load_tag(pct)
    return COLOR_OVER if tag == "over" else COLOR_WARN if tag == "warn" else COLOR_OK


# ----------------------------------------------------------------------
# Tablo bileşeni (ttk.Treeview + kaydırma çubukları + başlığa tıklayınca sıralama)
# ----------------------------------------------------------------------
class DataTable(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.tree = ttk.Treeview(self, show="headings")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("over", background="#f8d4d4")
        self.tree.tag_configure("warn", background="#fbe9c8")
        self.tree.tag_configure("total", background="#e4e9f0", font=("Segoe UI", 10, "bold"))
        self._df = None
        self._total_row = None
        self._tag_func = None
        self._sort_state = {}

    def show(self, df: pd.DataFrame, tag_func=None, total_row: dict = None):
        self._df = df.reset_index(drop=True)
        self._tag_func = tag_func
        self._total_row = total_row
        cols = list(self._df.columns)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c, command=lambda c=c: self._sort(c))
            width = max(90, min(260, len(str(c)) * 9))
            anchor = "e" if pd.api.types.is_numeric_dtype(self._df[c]) else "w"
            self.tree.column(c, width=width, anchor=anchor, stretch=True)
        self._render()

    def _render(self):
        self.tree.delete(*self.tree.get_children())
        for _, row in self._df.iterrows():
            values = [fmt_num(v) for v in row.tolist()]
            tag = self._tag_func(row) if self._tag_func else ""
            self.tree.insert("", "end", values=values, tags=(tag,) if tag else ())
        if self._total_row:
            values = [fmt_num(self._total_row.get(c, "")) for c in self._df.columns]
            self.tree.insert("", "end", values=values, tags=("total",))

    def _sort(self, col):
        if self._df is None:
            return
        asc = not self._sort_state.get(col, False)
        self._sort_state = {col: asc}
        self._df = self._df.sort_values(col, ascending=asc, kind="mergesort").reset_index(drop=True)
        self._render()

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []


# ----------------------------------------------------------------------
# Grafik bileşeni
# ----------------------------------------------------------------------
class ChartPanel(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.fig = Figure(figsize=(10, 5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        toolbar_frame = ctk.CTkFrame(self, fg_color="transparent")
        toolbar_frame.pack(side="bottom", fill="x")
        NavigationToolbar2Tk(self.canvas, toolbar_frame).update()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.message("Veri yüklendikten sonra grafik burada görünecek.")

    def message(self, text):
        self.ax.clear()
        self.ax.axis("off")
        self.ax.text(0.5, 0.5, text, ha="center", va="center", fontsize=12, color="#666")
        self.canvas.draw_idle()

    def _style(self, title, ylabel, n):
        self.ax.set_title(title, fontsize=12, fontweight="bold", loc="left")
        self.ax.set_ylabel(ylabel)
        self.ax.spines["top"].set_visible(False)
        self.ax.spines["right"].set_visible(False)
        self.ax.grid(axis="y", alpha=0.25)
        self.ax.set_axisbelow(True)
        fs = 9 if n <= 20 else 7
        for lbl in self.ax.get_xticklabels():
            lbl.set_rotation(45 if n > 8 else 0)
            lbl.set_ha("right" if n > 8 else "center")
            lbl.set_fontsize(fs)
        self.fig.tight_layout()

    def load_chart(self, machines, hours, caps, pcts, title):
        self.ax.clear()
        self.ax.axis("on")
        x = np.arange(len(machines))
        colors = [bar_color(p) for p in pcts]
        self.ax.bar(x, hours, color=colors, width=0.7, zorder=2)
        self.ax.hlines(caps, x - 0.4, x + 0.4, colors="#222", linewidth=2, zorder=3, label="Kapasite")
        for xi, h, p in zip(x, hours, pcts):
            if not np.isnan(p):
                self.ax.text(xi, h, f"%{p:.0f}", ha="center", va="bottom",
                             fontsize=8 if len(x) <= 20 else 6, zorder=4)
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(machines)
        from matplotlib.patches import Patch
        from matplotlib.lines import Line2D
        self.ax.legend(handles=[
            Patch(color=COLOR_OK, label=f"≤ %{WARN_LIMIT:.0f}"),
            Patch(color=COLOR_WARN, label=f"%{WARN_LIMIT:.0f}–%{OVER_LIMIT:.0f}"),
            Patch(color=COLOR_OVER, label=f"> %{OVER_LIMIT:.0f} (darboğaz)"),
            Line2D([0], [0], color="#222", lw=2, label="Kapasite"),
        ], fontsize=8, loc="upper right", frameon=False)
        self._style(title, "Saat", len(x))
        self.canvas.draw_idle()

    def compare_chart(self, machines, before, after, caps, title):
        self.ax.clear()
        self.ax.axis("on")
        x = np.arange(len(machines))
        w = 0.38
        self.ax.bar(x - w / 2, before, w, color=COLOR_BEFORE, label="Önce", zorder=2)
        self.ax.bar(x + w / 2, after, w, color=COLOR_AFTER, label="Sonra", zorder=2)
        self.ax.hlines(caps, x - 0.45, x + 0.45, colors="#222", linewidth=2, zorder=3, label="Kapasite")
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(machines)
        self.ax.legend(fontsize=8, loc="upper right", frameon=False)
        self._style(title, "Saat", len(x))
        self.canvas.draw_idle()


# ----------------------------------------------------------------------
# Ana uygulama
# ----------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title(APP_TITLE)
        self.geometry("1350x820")
        self.minsize(1100, 680)

        self.report_callback_exception = self._show_exception
        self.settings = load_settings()
        self.engine: CapacityEngine = None
        self.summary: pd.DataFrame = None       # kapasite eklenmiş güncel özet
        self.sim_result: pd.DataFrame = None    # son simülasyon karşılaştırması
        self.month_vars = {}
        self.cap_entries = {}

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        self._build_top_bar()
        self._build_sidebar()
        self._build_tabs()
        self._set_status("Plan ve operasyon dosyalarını seçip 'Verileri Yükle'ye basın.")

        # son kullanılan dosyalar
        for key, entry, combo, default in (
            ("plan_file", self.plan_entry, self.plan_sheet, DEFAULT_PLAN_SHEET),
            ("op_file", self.op_entry, self.op_sheet, DEFAULT_OP_SHEET),
        ):
            path = self.settings.get(key, "")
            if path and os.path.exists(path):
                entry.insert(0, path)
                self._fill_sheets(path, combo, default)

    def _show_exception(self, exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            with open(os.path.join(app_dir(), "hata_kaydi.txt"), "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except Exception:
            pass
        messagebox.showerror("Beklenmeyen hata",
                             f"{exc_type.__name__}: {exc}\n\nAyrıntı hata_kaydi.txt dosyasına yazıldı.\n\n{text[-700:]}")

    # ---------------- Üst bar: dosya seçimi ----------------
    def _build_top_bar(self):
        bar = ctk.CTkFrame(self)
        bar.pack(side="top", fill="x", padx=10, pady=(10, 5))
        bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(bar, text="Plan dosyası:", width=120, anchor="w").grid(row=0, column=0, padx=8, pady=6)
        self.plan_entry = ctk.CTkEntry(bar)
        self.plan_entry.grid(row=0, column=1, sticky="ew", pady=6)
        ctk.CTkButton(bar, text="Seç…", width=70,
                      command=lambda: self._pick_file(self.plan_entry, self.plan_sheet, DEFAULT_PLAN_SHEET)
                      ).grid(row=0, column=2, padx=6)
        self.plan_sheet = ctk.CTkComboBox(bar, values=[DEFAULT_PLAN_SHEET], width=260)
        self.plan_sheet.set(DEFAULT_PLAN_SHEET)
        self.plan_sheet.grid(row=0, column=3, padx=6)

        ctk.CTkLabel(bar, text="Operasyon dosyası:", width=120, anchor="w").grid(row=1, column=0, padx=8, pady=6)
        self.op_entry = ctk.CTkEntry(bar)
        self.op_entry.grid(row=1, column=1, sticky="ew", pady=6)
        ctk.CTkButton(bar, text="Seç…", width=70,
                      command=lambda: self._pick_file(self.op_entry, self.op_sheet, DEFAULT_OP_SHEET)
                      ).grid(row=1, column=2, padx=6)
        self.op_sheet = ctk.CTkComboBox(bar, values=[DEFAULT_OP_SHEET], width=260)
        self.op_sheet.set(DEFAULT_OP_SHEET)
        self.op_sheet.grid(row=1, column=3, padx=6)

        self.load_btn = ctk.CTkButton(bar, text="Verileri Yükle", height=64, width=150,
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self._load_data)
        self.load_btn.grid(row=0, column=4, rowspan=2, padx=10, pady=6)

        self.status = ctk.CTkLabel(self, text="", anchor="w", text_color="#555")
        self.status.pack(side="bottom", fill="x", padx=14, pady=(0, 6))

    def _pick_file(self, entry, combo, default_sheet):
        path = filedialog.askopenfilename(filetypes=[("Excel dosyaları", "*.xlsx *.xlsm *.xls"), ("Tümü", "*.*")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)
            self._fill_sheets(path, combo, default_sheet)

    def _fill_sheets(self, path, combo, default_sheet):
        try:
            sheets = pd.ExcelFile(path).sheet_names
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya açılamadı:\n{e}")
            return
        combo.configure(values=sheets)
        combo.set(default_sheet if default_sheet in sheets else sheets[0])

    # ---------------- Sol panel: aylar ve aksiyonlar ----------------
    def _build_sidebar(self):
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(side="top", fill="both", expand=True, padx=10, pady=5)

        side = ctk.CTkFrame(self.body, width=220)
        side.pack(side="left", fill="y", padx=(0, 8))
        side.pack_propagate(False)

        ctk.CTkLabel(side, text="Hesaplanacak Aylar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(12, 4))
        self.month_frame = ctk.CTkScrollableFrame(side, height=260)
        self.month_frame.pack(fill="both", expand=True, padx=8)
        ctk.CTkLabel(self.month_frame, text="(veri yüklenmedi)", text_color="#888").pack()

        btns = ctk.CTkFrame(side, fg_color="transparent")
        btns.pack(fill="x", padx=8, pady=4)
        ctk.CTkButton(btns, text="Tümü", width=95, command=lambda: self._set_all_months(True)).pack(side="left")
        ctk.CTkButton(btns, text="Temizle", width=95, fg_color="gray55",
                      command=lambda: self._set_all_months(False)).pack(side="right")

        self.calc_btn = ctk.CTkButton(side, text="Hesapla", height=40, state="disabled",
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self._recalculate)
        self.calc_btn.pack(fill="x", padx=8, pady=(14, 6))
        self.export_btn = ctk.CTkButton(side, text="Excel'e Aktar", height=36, state="disabled",
                                        fg_color=COLOR_OK, hover_color="#24804a", command=self._export)
        self.export_btn.pack(fill="x", padx=8, pady=6)

        self.info_label = ctk.CTkLabel(side, text="", justify="left", anchor="w", wraplength=200)
        self.info_label.pack(fill="x", padx=10, pady=(12, 10))

    def _set_all_months(self, value):
        for v in self.month_vars.values():
            v.set(value)

    def _selected_months(self):
        return [m for m, v in self.month_vars.items() if v.get()]

    # ---------------- Sekmeler ----------------
    def _build_tabs(self):
        self.tabs = ctk.CTkTabview(self.body)
        self.tabs.pack(side="left", fill="both", expand=True)
        for name in ("Makine Yükü", "Grafik", "Simülasyon", "Kapasite", "Eşleşmeyenler"):
            self.tabs.add(name)

        # Makine yükü
        t = self.tabs.tab("Makine Yükü")
        self.load_table = DataTable(t)
        self.load_table.pack(fill="both", expand=True)

        # Grafik
        t = self.tabs.tab("Grafik")
        top = ctk.CTkFrame(t, fg_color="transparent")
        top.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(top, text="Dönem:").pack(side="left", padx=(4, 6))
        self.period_combo = ctk.CTkComboBox(top, values=["Seçili aylar toplamı"], width=220,
                                            command=lambda _: self._draw_load_chart())
        self.period_combo.set("Seçili aylar toplamı")
        self.period_combo.pack(side="left")
        self.only_over = ctk.CTkCheckBox(top, text=f"Sadece %{WARN_LIMIT:.0f} üstü", command=self._draw_load_chart)
        self.only_over.pack(side="left", padx=16)
        self.load_chart = ChartPanel(t)
        self.load_chart.pack(fill="both", expand=True)

        self._build_sim_tab()
        self._build_capacity_tab()

        # Eşleşmeyenler
        t = self.tabs.tab("Eşleşmeyenler")
        self.unmapped_label = ctk.CTkLabel(t, text="", anchor="w")
        self.unmapped_label.pack(fill="x", padx=4, pady=(0, 4))
        self.unmapped_table = DataTable(t)
        self.unmapped_table.pack(fill="both", expand=True)

    def _build_sim_tab(self):
        t = self.tabs.tab("Simülasyon")
        ctrl = ctk.CTkFrame(t)
        ctrl.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(ctrl, text="Darboğaz makine:").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        self.sim_machine = ctk.CTkComboBox(ctrl, values=[""], width=160, command=lambda _: self._update_sim_info())
        self.sim_machine.grid(row=0, column=1, padx=4)

        ctk.CTkLabel(ctrl, text="Hedef:").grid(row=0, column=2, padx=(16, 4))
        self.sim_alt = ctk.CTkComboBox(ctrl, values=ALT_COLS, width=140, command=lambda _: self._update_sim_info())
        self.sim_alt.set(ALT_COLS[0])
        self.sim_alt.grid(row=0, column=3, padx=4)

        ctk.CTkLabel(ctrl, text="Aktarım oranı:").grid(row=0, column=4, padx=(16, 4))
        self.sim_pct = ctk.CTkSlider(ctrl, from_=0, to=100, number_of_steps=20, width=200,
                                     command=lambda v: self.sim_pct_label.configure(text=f"%{v:.0f}"))
        self.sim_pct.set(30)
        self.sim_pct.grid(row=0, column=5, padx=4)
        self.sim_pct_label = ctk.CTkLabel(ctrl, text="%30", width=44)
        self.sim_pct_label.grid(row=0, column=6)

        self.sim_btn = ctk.CTkButton(ctrl, text="Simüle Et", state="disabled", command=self._run_simulation)
        self.sim_btn.grid(row=0, column=7, padx=12)

        self.sim_info = ctk.CTkLabel(ctrl, text="", anchor="w", text_color="#555")
        self.sim_info.grid(row=1, column=0, columnspan=8, sticky="w", padx=8, pady=(0, 8))

        pane = ttk.PanedWindow(t, orient="vertical")
        pane.pack(fill="both", expand=True)
        self.sim_chart = ChartPanel(pane)
        self.sim_chart.message("Makine, hedef ve oranı seçip 'Simüle Et'e basın.")
        self.sim_table = DataTable(pane)
        pane.add(self.sim_chart, weight=3)
        pane.add(self.sim_table, weight=2)

    def _build_capacity_tab(self):
        t = self.tabs.tab("Kapasite")
        top = ctk.CTkFrame(t)
        top.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(top, text="Varsayılan aylık kapasite (saat):").pack(side="left", padx=8, pady=8)
        self.default_cap_entry = ctk.CTkEntry(top, width=90)
        self.default_cap_entry.insert(0, fmt_input(self.settings["default_capacity"]))
        self.default_cap_entry.pack(side="left")
        ctk.CTkButton(top, text="Tümüne uygula", width=120, fg_color="gray55",
                      command=self._apply_default_to_all).pack(side="left", padx=8)
        ctk.CTkButton(top, text="Kaydet ve Yeniden Hesapla", command=self._save_capacities).pack(side="right", padx=8)
        ctk.CTkLabel(t, text="Her makine için ayda kullanılabilir üretim saatini girin "
                             "(ör. 22 gün × 3 vardiya × 7,5 saat × verim). Boş bırakılırsa varsayılan kullanılır.",
                     text_color="#555", anchor="w", justify="left", wraplength=900).pack(fill="x", padx=8)
        self.cap_frame = ctk.CTkScrollableFrame(t)
        self.cap_frame.pack(fill="both", expand=True, pady=6)

    # ---------------- Veri yükleme ----------------
    def _set_status(self, text):
        self.status.configure(text=text)

    def _load_data(self):
        plan, op = self.plan_entry.get().strip(), self.op_entry.get().strip()
        if not (plan and op):
            messagebox.showwarning("Eksik", "Lütfen iki dosyayı da seçin.")
            return
        self.load_btn.configure(state="disabled", text="Yükleniyor…")
        self._set_status("Excel dosyaları okunuyor, lütfen bekleyin…")
        plan_sheet, op_sheet = self.plan_sheet.get(), self.op_sheet.get()

        def work():
            try:
                eng = CapacityEngine(plan, op).load_data(plan_sheet=plan_sheet, op_sheet=op_sheet)
                self.after(0, lambda: self._on_loaded(eng, plan, op))
            except Exception as e:
                tb = traceback.format_exc()
                self.after(0, lambda: self._on_load_error(e, tb))

        threading.Thread(target=work, daemon=True).start()

    def _on_load_error(self, e, tb):
        self.load_btn.configure(state="normal", text="Verileri Yükle")
        self._set_status("Yükleme başarısız.")
        messagebox.showerror("Yükleme hatası",
                             f"{type(e).__name__}: {e}\n\nSayfa adlarını ve sütun başlıklarını kontrol edin.\n\n{tb[-800:]}")

    def _on_loaded(self, eng: CapacityEngine, plan, op):
        try:
            self._on_loaded_inner(eng, plan, op)
        except Exception as e:
            self.load_btn.configure(state="normal", text="Verileri Yükle")
            self._show_exception(type(e), e, e.__traceback__)

    def _on_loaded_inner(self, eng: CapacityEngine, plan, op):
        self.engine = eng
        self.load_btn.configure(state="normal", text="Verileri Yükle")
        self.settings["plan_file"], self.settings["op_file"] = plan, op
        save_settings(self.settings)

        # Aylar
        for w in self.month_frame.winfo_children():
            w.destroy()
        self.month_vars = {}
        for m in eng.available_months:
            n_weeks = sum(1 for v in eng.week_month_map.values() if v == m)
            var = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(self.month_frame, text=f"{m}  ({n_weeks} hf)", variable=var).pack(anchor="w", pady=3)
            self.month_vars[m] = var

        # Makineler
        machines = eng.machines()
        self.sim_machine.configure(values=machines or [""])
        if machines:
            self.sim_machine.set(machines[0])
        self._build_capacity_rows(eng.all_machines())

        # Eşleşmeyenler
        self._show_unmapped()

        for b in (self.calc_btn, self.export_btn, self.sim_btn):
            b.configure(state="normal")

        n_plan = len(eng.df_plan)
        n_unm = eng.unmapped_df[eng.ref_col].nunique() if eng.unmapped_df is not None else 0
        self.info_label.configure(text=(
            f"Plan satırı: {n_plan}\n"
            f"Operasyon satırı: {len(eng.df_op)}\n"
            f"Eşleşen iş: {len(eng.merged_df)}\n"
            f"Makine: {len(machines)}\n"
            f"Hafta: {len(eng.week_columns)}\n"
            f"Eşleşmeyen ref.: {n_unm}"))
        self._recalculate()
        self._set_status(f"Yüklendi: {os.path.basename(plan)}  +  {os.path.basename(op)}")

    def _show_unmapped(self):
        eng = self.engine
        df = eng.unmapped_df
        if df is None or df.empty:
            self.unmapped_label.configure(text="Tüm plan referansları operasyon dosyasında bulundu. ✓")
            self.unmapped_table.clear()
            return
        qty_cols = eng.week_columns + eng.backlog_columns
        view = df.copy()
        view["TOPLAM ADET"] = view[qty_cols].sum(axis=1) if qty_cols else 0
        info_cols = [c for c in view.columns if not c.startswith("COL_") and c not in qty_cols]
        view = view[info_cols]
        view = view.sort_values("TOPLAM ADET", ascending=False)
        self.unmapped_label.configure(
            text=f"{view[eng.ref_col].nunique()} referans operasyon dosyasında yok — bunların saati hesaba katılmadı.")
        self.unmapped_table.show(view)

    # ---------------- Kapasite ----------------
    def _default_cap(self) -> float:
        try:
            return float(self.default_cap_entry.get().replace(",", "."))
        except ValueError:
            return float(self.settings.get("default_capacity", 0))

    def _build_capacity_rows(self, machines):
        for w in self.cap_frame.winfo_children():
            w.destroy()
        self.cap_entries = {}
        caps = self.settings["capacities"]
        ctk.CTkLabel(self.cap_frame, text="Makine", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, sticky="w")
        ctk.CTkLabel(self.cap_frame, text="Aylık kapasite (saat)", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=10)
        for i, m in enumerate(machines, start=1):
            ctk.CTkLabel(self.cap_frame, text=m, width=140, anchor="w").grid(row=i, column=0, padx=10, pady=2, sticky="w")
            e = ctk.CTkEntry(self.cap_frame, width=110, placeholder_text=fmt_input(self._default_cap()))
            if m in caps:
                e.insert(0, fmt_input(caps[m]))
            e.grid(row=i, column=1, padx=10, pady=2)
            self.cap_entries[m] = e

    def _apply_default_to_all(self):
        val = fmt_input(self._default_cap())
        for e in self.cap_entries.values():
            e.delete(0, "end")
            e.insert(0, val)

    def _read_capacities(self) -> dict:
        caps = dict(self.settings.get("capacities", {}))
        bad = []
        for m, e in self.cap_entries.items():
            txt = e.get().strip().replace(",", ".")
            if not txt:
                caps.pop(m, None)
                continue
            try:
                caps[m] = float(txt)
            except ValueError:
                bad.append(m)
        if bad:
            messagebox.showwarning("Geçersiz değer", "Şu makinelerin kapasitesi sayı değil, atlandı:\n" + ", ".join(bad))
        return caps

    def _save_capacities(self):
        self.settings["default_capacity"] = self._default_cap()
        self.settings["capacities"] = self._read_capacities()
        save_settings(self.settings)
        self._set_status("Kapasite ayarları kaydedildi.")
        if self.engine is not None:
            self._recalculate()

    def _capacity_args(self):
        caps = self._read_capacities() if self.cap_entries else self.settings["capacities"]
        return caps, self._default_cap()

    # ---------------- Hesaplama ----------------
    def _recalculate(self):
        if self.engine is None:
            return
        months = self._selected_months()
        if not months:
            messagebox.showwarning("Ay seçilmedi", "En az bir ay seçin.")
            return
        try:
            raw = self.engine.calculate_workload(selected_months=months)
            caps, default_cap = self._capacity_args()
            self.summary = CapacityEngine.apply_capacity(raw, months, caps, default_cap)
        except Exception as e:
            messagebox.showerror("Hesaplama hatası", f"{type(e).__name__}: {e}\n\n{traceback.format_exc()[-800:]}")
            return
        self._show_load_table(months)
        self.period_combo.configure(values=["Seçili aylar toplamı"] + months)
        if self.period_combo.get() not in ["Seçili aylar toplamı"] + months:
            self.period_combo.set("Seçili aylar toplamı")
        self._draw_load_chart()
        self._update_sim_info()
        over = (self.summary["DOLULUK_%"] > OVER_LIMIT).sum()
        self._set_status(f"Hesaplandı: {', '.join(months)}  |  {len(self.summary)} makine  |  "
                         f"kapasite üstü: {over}")

    def _view_summary(self, df: pd.DataFrame, months):
        """Ekranda gösterilecek sade tablo (haftalar Excel'de)."""
        eng = self.engine
        month_cols = [m for m in months if f"{m}_Saat" in df.columns]
        view = pd.DataFrame({"Makine": df["Üretim presi"].replace("", NO_MACHINE)})
        for m in month_cols:
            view[f"{m} (saat)"] = df[f"{m}_Saat"]
            view[f"{m} %"] = df[f"{m}_Doluluk_%"]
        if eng.backlog_columns:
            view["Backlog (saat)"] = df[[f"{b}_Saat" for b in eng.backlog_columns]].sum(axis=1)
        view["Toplam (saat)"] = df["TOPLAM_SAAT"]
        view["Kapasite (saat)"] = df["TOPLAM_KAPASİTE"]
        view["Doluluk %"] = df["DOLULUK_%"]
        return view

    def _show_load_table(self, months):
        view = self._view_summary(self.summary, months)
        total = {"Makine": "TOPLAM"}
        for c in view.columns[1:]:
            if not c.endswith("%"):
                total[c] = view[c].sum()
        if total.get("Kapasite (saat)"):
            total["Doluluk %"] = total["Toplam (saat)"] / total["Kapasite (saat)"] * 100
        self.load_table.show(view, tag_func=lambda r: load_tag(r["Doluluk %"]), total_row=total)

    def _draw_load_chart(self):
        if self.summary is None or self.summary.empty:
            self.load_chart.message("Gösterilecek veri yok.")
            return
        df = self.summary.copy()
        period = self.period_combo.get()
        if period != "Seçili aylar toplamı" and f"{period}_Saat" in df.columns:
            hours = df[f"{period}_Saat"]
            caps = df["AYLIK_KAPASİTE"]
            pcts = df[f"{period}_Doluluk_%"]
            title = f"{period} – makine yükü"
        else:
            hours, caps, pcts = df["TOPLAM_SAAT"], df["TOPLAM_KAPASİTE"], df["DOLULUK_%"]
            title = "Seçili aylar toplamı – makine yükü"
        plot = pd.DataFrame({"m": df["Üretim presi"].replace("", NO_MACHINE), "h": hours, "c": caps, "p": pcts})
        if self.only_over.get():
            plot = plot[plot["p"] > WARN_LIMIT]
        plot = plot.sort_values("h", ascending=False)
        if plot.empty:
            self.load_chart.message("Filtreye uyan makine yok.")
            return
        self.load_chart.load_chart(plot["m"].tolist(), plot["h"].values, plot["c"].values,
                                   plot["p"].values.astype(float), title)

    # ---------------- Simülasyon ----------------
    def _update_sim_info(self):
        if self.engine is None:
            return
        m, alt = self.sim_machine.get(), self.sim_alt.get()
        df = self.engine.merged_df
        jobs = df[df["Üretim presi"] == normalize_machine(m)]
        if jobs.empty:
            self.sim_info.configure(text="Bu makinede iş yok.")
            return
        has_alt = jobs[(jobs[alt] != "") & jobs[alt].notna()]
        targets = sorted(t for t in has_alt[alt].unique() if t)
        self.sim_info.configure(text=(
            f"{m}: {len(jobs)} iş, bunların {len(has_alt)} tanesinin '{alt}' tanımlı "
            f"(aktarılabilir). Hedef makineler: {', '.join(targets) if targets else '—'}"))

    def _run_simulation(self):
        if self.engine is None or self.summary is None:
            return
        months = self._selected_months()
        if not months:
            messagebox.showwarning("Ay seçilmedi", "En az bir ay seçin.")
            return
        machine, alt, pct = self.sim_machine.get(), self.sim_alt.get(), float(self.sim_pct.get())
        try:
            sim_raw = self.engine.simulate_transfer(machine, pct, alt, months)
            caps, default_cap = self._capacity_args()
            after = CapacityEngine.apply_capacity(sim_raw, months, caps, default_cap)
        except Exception as e:
            messagebox.showerror("Simülasyon hatası", f"{type(e).__name__}: {e}\n\n{traceback.format_exc()[-800:]}")
            return

        before = self.summary[["Üretim presi", "TOPLAM_SAAT", "TOPLAM_KAPASİTE", "DOLULUK_%"]]
        after = after[["Üretim presi", "TOPLAM_SAAT", "TOPLAM_KAPASİTE", "DOLULUK_%"]]
        cmp_df = before.merge(after, on="Üretim presi", how="outer", suffixes=("_ÖNCE", "_SONRA"))
        cmp_df["TOPLAM_KAPASİTE"] = cmp_df["TOPLAM_KAPASİTE_ÖNCE"].fillna(cmp_df["TOPLAM_KAPASİTE_SONRA"])
        cmp_df[["TOPLAM_SAAT_ÖNCE", "TOPLAM_SAAT_SONRA"]] = cmp_df[["TOPLAM_SAAT_ÖNCE", "TOPLAM_SAAT_SONRA"]].fillna(0)
        cap = cmp_df["TOPLAM_KAPASİTE"].replace(0, np.nan)
        cmp_df["DOLULUK_%_ÖNCE"] = cmp_df["TOPLAM_SAAT_ÖNCE"] / cap * 100
        cmp_df["DOLULUK_%_SONRA"] = cmp_df["TOPLAM_SAAT_SONRA"] / cap * 100
        cmp_df["FARK"] = cmp_df["TOPLAM_SAAT_SONRA"] - cmp_df["TOPLAM_SAAT_ÖNCE"]

        view = pd.DataFrame({
            "Makine": cmp_df["Üretim presi"].replace("", NO_MACHINE),
            "Önce (saat)": cmp_df["TOPLAM_SAAT_ÖNCE"],
            "Sonra (saat)": cmp_df["TOPLAM_SAAT_SONRA"],
            "Fark (saat)": cmp_df["FARK"],
            "Kapasite (saat)": cmp_df["TOPLAM_KAPASİTE"],
            "Önce %": cmp_df["DOLULUK_%_ÖNCE"],
            "Sonra %": cmp_df["DOLULUK_%_SONRA"],
        })
        changed = view[view["Fark (saat)"].abs() > 1e-9].sort_values("Fark (saat)")
        self.sim_result = view.copy()
        self.sim_result.insert(0, "Senaryo", f"{machine} → {alt} %{pct:.0f}")

        if changed.empty:
            self.sim_chart.message(f"{machine} için '{alt}' tanımlı iş yok – değişiklik olmadı.")
            self.sim_table.show(view, tag_func=lambda r: load_tag(r["Sonra %"]))
            return
        self.sim_chart.compare_chart(changed["Makine"].tolist(), changed["Önce (saat)"].values,
                                     changed["Sonra (saat)"].values, changed["Kapasite (saat)"].values,
                                     f"{machine} → {alt}  (%{pct:.0f} aktarım)  – etkilenen makineler")
        rest = view.drop(changed.index).sort_values("Sonra (saat)", ascending=False)
        self.sim_table.show(pd.concat([changed, rest]), tag_func=lambda r: load_tag(r["Sonra %"]))
        moved = changed.loc[changed["Fark (saat)"] > 0, "Fark (saat)"].sum()
        self._set_status(f"Simülasyon: {machine} makinesinden {fmt_num(moved)} saat aktarıldı.")

    # ---------------- Excel ----------------
    def _export(self):
        if self.engine is None:
            return
        months = self._selected_months()
        if not months:
            messagebox.showwarning("Ay seçilmedi", "En az bir ay seçin.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Makine_Yuk_Raporu.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        caps, default_cap = self._capacity_args()
        extra = {"Simülasyon": self.sim_result} if self.sim_result is not None else None
        try:
            self.engine.export_to_excel(path, months, caps, default_cap, extra_sheets=extra)
        except PermissionError:
            messagebox.showerror("Hata", "Dosya yazılamadı. Excel'de açıksa kapatıp tekrar deneyin.")
            return
        except Exception as e:
            messagebox.showerror("Hata", f"{type(e).__name__}: {e}")
            return
        self._set_status(f"Rapor kaydedildi: {path}")
        if messagebox.askyesno("Tamam", "Rapor kaydedildi. Şimdi açılsın mı?"):
            try:
                os.startfile(path)  # Windows
            except Exception:
                pass


def fmt_input(v) -> str:
    """Giriş kutuları için sayı biçimi: 450.0 -> '450', 437.5 -> '437,5'."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return str(int(v)) if v.is_integer() else str(v).replace(".", ",")


if __name__ == "__main__":
    App().mainloop()

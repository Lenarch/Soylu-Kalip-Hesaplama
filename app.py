"""
Soylu Makine - Pres Kapasite Hesaplama Uygulaması
Arayüz: CustomTkinter  |  Hesap motoru: capacity_engine.py
"""
import json
import re
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
from veritabani import (create_empty_db, table_counts, list_imports, undo_import,
                        load_operations, save_operations, OP_COLUMNS)
from excel_aktar import sheet_names, parse_plan, check_conflicts, import_plan, parse_operations, OP_SHEET_DEFAULT, fold

APP_TITLE = "Soylu Makine – Pres Kapasite Hesaplama"
DEFAULT_DB_NAME = "kapasite.db"
ALT_COLS = ["alternatif 1", "alternatif 2", "alternatif 3", "alternatif 4"]
NO_MACHINE = "(TANIMSIZ)"

DEFAULT_WEEKLY_CAP = 105.0     # saat/hafta (5 gün × 3 vardiya × 7 saat)
WEEKS_PER_MONTH = 52 / 12      # eski aylık ayarları çevirmek için
ALL_PERIOD = "Seçili aylar toplamı"

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
    except Exception:
        data = {"db_file": ""}
    # Kapasite artık HAFTALIK. Eski sürümün aylık değerleri varsa ayda ~4,33 haftaya bölünerek çevrilir.
    if "default_weekly_capacity" not in data:
        old = data.get("default_capacity")
        data["default_weekly_capacity"] = round(float(old) / WEEKS_PER_MONTH, 1) if old else DEFAULT_WEEKLY_CAP
    if "weekly_capacities" not in data:
        data["weekly_capacities"] = {m: round(float(v) / WEEKS_PER_MONTH, 1)
                                     for m, v in (data.get("capacities") or {}).items()}
    return data


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


def backlog_col(eng):
    """Ekranda gösterilecek backlog sütunu: varsa 'Güncel Backlog', yoksa ilk backlog sütunu."""
    cols = list(getattr(eng, "backlog_columns", []) or [])
    if not cols:
        return None
    return next((c for c in cols if "GÜNCEL" in c.upper() or "GUNCEL" in c.upper()), cols[0])


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
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.ops_draft = pd.DataFrame(columns=OP_COLUMNS)   # Operasyonlar sekmesindeki taslak
        self.ops_dirty = False
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
        self._set_status("Veritabanını seçip 'Verileri Yükle'ye basın.")

        # son kullanılan veritabanı (yoksa exe'nin yanındaki kapasite.db)
        db = self.settings.get("db_file") or os.path.join(app_dir(), DEFAULT_DB_NAME)
        self.db_entry.insert(0, db)
        self._show_db_info()
        self._refresh_imports()
        self._ops_reload()

    def _on_close(self):
        if self.ops_dirty and not messagebox.askyesno(
                "Kaydedilmemiş değişiklik",
                "Operasyonlar sekmesinde kaydedilmemiş değişiklikler var.\n\n"
                "Kaydetmeden çıkılsın mı? (Değişiklikler iptal olur)"):
            return
        self.destroy()

    def _confirm_discard_ops(self) -> bool:
        """Taslakta kaydedilmemiş değişiklik varsa kullanıcıya sorar. True = devam edilebilir."""
        if not self.ops_dirty:
            return True
        return messagebox.askyesno("Kaydedilmemiş değişiklik",
                                   "Operasyonlar sekmesinde kaydedilmemiş değişiklikler var.\n"
                                   "Devam edilirse bu değişiklikler iptal olacak. Devam edilsin mi?")

    def _show_exception(self, exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            with open(os.path.join(app_dir(), "hata_kaydi.txt"), "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except Exception:
            pass
        messagebox.showerror("Beklenmeyen hata",
                             f"{exc_type.__name__}: {exc}\n\nAyrıntı hata_kaydi.txt dosyasına yazıldı.\n\n{text[-700:]}")

    # ---------------- Üst bar: veritabanı seçimi ----------------
    def _build_top_bar(self):
        bar = ctk.CTkFrame(self)
        bar.pack(side="top", fill="x", padx=10, pady=(10, 5))
        bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(bar, text="Veritabanı:", width=100, anchor="w").grid(row=0, column=0, padx=8, pady=(10, 2))
        self.db_entry = ctk.CTkEntry(bar)
        self.db_entry.grid(row=0, column=1, sticky="ew", pady=(10, 2))
        ctk.CTkButton(bar, text="Seç…", width=70, command=self._pick_db).grid(row=0, column=2, padx=6, pady=(10, 2))
        ctk.CTkButton(bar, text="Yeni DB oluştur", width=130, fg_color="gray55",
                      command=self._new_db).grid(row=0, column=3, padx=6, pady=(10, 2))

        self.db_info = ctk.CTkLabel(bar, text="", anchor="w", text_color="#555")
        self.db_info.grid(row=1, column=1, columnspan=3, sticky="w", pady=(0, 8))

        self.import_btn = ctk.CTkButton(bar, text="Excel'den Aktar", height=56, width=140,
                                        fg_color=COLOR_OK, hover_color="#24804a",
                                        font=ctk.CTkFont(size=13, weight="bold"), command=self._open_import)
        self.import_btn.grid(row=0, column=4, rowspan=2, padx=(10, 4), pady=8)
        self.load_btn = ctk.CTkButton(bar, text="Verileri Yükle", height=56, width=150,
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self._load_data)
        self.load_btn.grid(row=0, column=5, rowspan=2, padx=(4, 10), pady=8)

        self.status = ctk.CTkLabel(self, text="", anchor="w", text_color="#555")
        self.status.pack(side="bottom", fill="x", padx=14, pady=(0, 6))

    def _db_path(self) -> str:
        return self.db_entry.get().strip()

    def _show_db_info(self):
        path = self._db_path()
        if not path or not os.path.exists(path):
            self.db_info.configure(text="Dosya bulunamadı – 'Yeni DB oluştur' ile boş veritabanı oluşturabilirsiniz.")
            return
        try:
            c = table_counts(path)
        except Exception as e:
            self.db_info.configure(text=f"Veritabanı okunamadı: {e}")
            return
        parts = [f"{t}: {'tablo yok' if n is None else n}" for t, n in c.items()]
        self.db_info.configure(text="Kayıt sayıları →  " + "   |   ".join(parts))

    def _pick_db(self):
        if not self._confirm_discard_ops():
            return
        path = filedialog.askopenfilename(filetypes=[("SQLite veritabanı", "*.db *.sqlite *.sqlite3"), ("Tümü", "*.*")])
        if path:
            self.db_entry.delete(0, "end")
            self.db_entry.insert(0, path)
            self._show_db_info()
            self._refresh_imports()
            self._ops_reload()

    def _new_db(self):
        if not self._confirm_discard_ops():
            return
        path = filedialog.asksaveasfilename(defaultextension=".db", initialfile=DEFAULT_DB_NAME,
                                            filetypes=[("SQLite veritabanı", "*.db")])
        if not path:
            return
        existed = os.path.exists(path)
        try:
            create_empty_db(path)
        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanı oluşturulamadı:\n{e}")
            return
        self.db_entry.delete(0, "end")
        self.db_entry.insert(0, path)
        self._show_db_info()
        messagebox.showinfo("Tamam", ("Var olan veritabanına eksik tablolar eklendi, mevcut verilere dokunulmadı."
                                      if existed else "Boş veritabanı oluşturuldu.") +
                            "\n\nPlan Excel'lerini 'Excel'den Aktar' ile yükleyebilirsiniz.")
        self._refresh_imports()
        self._ops_reload()

    # ---------------- Excel'den aktarma ----------------
    def _open_import(self):
        db = self._db_path()
        if not db:
            messagebox.showwarning("Eksik", "Önce bir veritabanı seçin veya 'Yeni DB oluştur'.")
            return
        if not os.path.exists(db):
            if not messagebox.askyesno("Veritabanı yok", f"{db}\nbulunamadı. Boş veritabanı oluşturulsun mu?"):
                return
            create_empty_db(db)
            self._show_db_info()
        ImportDialog(self, db, on_done=self._after_import)

    def _after_import(self):
        self._show_db_info()
        self._refresh_imports()
        self._load_data()

    def _refresh_imports(self):
        db = self._db_path()
        if not hasattr(self, "imports_table"):
            return
        if not db or not os.path.exists(db):
            self.imports_table.clear()
            return
        try:
            rows = list_imports(db)
        except Exception as e:
            self.imports_label.configure(text=f"Geçmiş okunamadı: {e}")
            return
        if not rows:
            self.imports_label.configure(text="Henüz yükleme yapılmadı.")
            self.imports_table.clear()
            return
        df = pd.DataFrame([{
            "No": str(r["id"]), "Tarih": r["tarih"], "Dosya": r["dosya_adi"], "Sayfa": r["sayfa"] or "",
            "Ay": f"{r['ay']} {r['yil']}", "Haftalar": ", ".join("W" + h for h in (r["haftalar"] or "").split(",") if h),
            "Sipariş satırı": str(r["siparis_satir"]), "Backlog satırı": str(r["backlog_satir"]),
            "Durum": ("HESAPTA" if r["gecerli"] else
                      "geri alındı" if r["durum"] != "aktif" else "eski (yenisi var)"),
        } for r in rows])
        n_valid = sum(r["gecerli"] for r in rows)
        self.imports_label.configure(text=f"{len(rows)} yükleme, {n_valid} tanesi hesapta. "
                                          "Aynı ay için en son yükleme geçerlidir.")
        self.imports_table.show(df, tag_func=lambda r: "" if r["Durum"] == "HESAPTA" else "warn")

    def _undo_selected(self):
        db = self._db_path()
        sel = self.imports_table.tree.selection()
        if not sel:
            messagebox.showinfo("Seçim yok", "Listeden geri almak istediğiniz yüklemeyi seçin.")
            return
        vals = self.imports_table.tree.item(sel[0], "values")
        aid, ay, dosya, durum = int(vals[0]), vals[4], vals[2], vals[8]
        if durum == "geri alındı":
            messagebox.showinfo("Bilgi", "Bu yükleme zaten geri alınmış.")
            return
        if not messagebox.askyesno("Geri al", f"{dosya} – {ay} yüklemesi geri alınsın mı?\n\n"
                                              "Veri silinmez; bu yükleme hesaptan çıkar. Aynı ay için daha önceki "
                                              "bir yükleme varsa o tekrar geçerli olur."):
            return
        undo_import(db, aid)
        self._show_db_info()
        self._refresh_imports()
        self._load_data()

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
        for name in ("Makine Yükü", "Grafik", "Simülasyon", "Kapasite", "Operasyonlar", "Eşleşmeyenler", "Aktarımlar"):
            self.tabs.add(name)

        # Makine yükü
        t = self.tabs.tab("Makine Yükü")
        top = ctk.CTkFrame(t, fg_color="transparent")
        top.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(top, text="Görünüm:").pack(side="left", padx=(4, 6))
        self.view_mode = ctk.CTkSegmentedButton(top, values=["Aylık", "Haftalık"],
                                                command=lambda _: self._refresh_load_table())
        self.view_mode.set("Aylık")
        self.view_mode.pack(side="left")
        self.view_hint = ctk.CTkLabel(top, text="", text_color="#555")
        self.view_hint.pack(side="left", padx=12)
        self.load_table = DataTable(t)
        self.load_table.pack(fill="both", expand=True)

        # Grafik
        t = self.tabs.tab("Grafik")
        top = ctk.CTkFrame(t, fg_color="transparent")
        top.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(top, text="Dönem:").pack(side="left", padx=(4, 6))
        self.period_combo = ctk.CTkComboBox(top, values=[ALL_PERIOD], width=240,
                                            command=lambda _: self._draw_load_chart())
        self.period_combo.set(ALL_PERIOD)
        self.period_weeks = {}   # grafikteki hafta etiketi -> hafta sütunu
        self.period_combo.pack(side="left")
        self.only_over = ctk.CTkCheckBox(top, text=f"Sadece %{WARN_LIMIT:.0f} üstü", command=self._draw_load_chart)
        self.only_over.pack(side="left", padx=16)
        self.load_chart = ChartPanel(t)
        self.load_chart.pack(fill="both", expand=True)

        self._build_sim_tab()
        self._build_capacity_tab()

        self._build_ops_tab()

        # Aktarımlar (işlem kaydı)
        t = self.tabs.tab("Aktarımlar")
        top = ctk.CTkFrame(t, fg_color="transparent")
        top.pack(fill="x", pady=(0, 4))
        self.imports_label = ctk.CTkLabel(top, text="", anchor="w")
        self.imports_label.pack(side="left", padx=4)
        ctk.CTkButton(top, text="Seçileni Geri Al", width=140, fg_color=COLOR_OVER, hover_color="#b03636",
                      command=self._undo_selected).pack(side="right", padx=4)
        ctk.CTkButton(top, text="Yenile", width=80, fg_color="gray55",
                      command=self._refresh_imports).pack(side="right", padx=4)
        self.imports_table = DataTable(t)
        self.imports_table.pack(fill="both", expand=True)

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
        ctk.CTkLabel(top, text="Varsayılan HAFTALIK kapasite (saat):").pack(side="left", padx=8, pady=8)
        self.default_cap_entry = ctk.CTkEntry(top, width=90)
        self.default_cap_entry.insert(0, fmt_input(self.settings["default_weekly_capacity"]))
        self.default_cap_entry.pack(side="left")
        ctk.CTkButton(top, text="Tümüne uygula", width=120, fg_color="gray55",
                      command=self._apply_default_to_all).pack(side="left", padx=8)
        ctk.CTkButton(top, text="Kaydet ve Yeniden Hesapla", command=self._save_capacities).pack(side="right", padx=8)
        ctk.CTkLabel(t, text="Her makine için HAFTADA kullanılabilir üretim saatini girin "
                             "(ör. 5 gün × 3 vardiya × 7,5 saat × verim). Boş bırakılırsa varsayılan kullanılır. "
                             "Ay kapasitesi = haftalık kapasite × o ayın planda kaç haftası varsa.",
                     text_color="#555", anchor="w", justify="left", wraplength=900).pack(fill="x", padx=8)
        self.cap_frame = ctk.CTkScrollableFrame(t)
        self.cap_frame.pack(fill="both", expand=True, pady=6)

    # ---------------- Veri yükleme ----------------
    def _set_status(self, text):
        self.status.configure(text=text)

    def _load_data(self):
        db = self._db_path()
        if not db:
            messagebox.showwarning("Eksik", "Lütfen bir veritabanı dosyası seçin.")
            return
        self.load_btn.configure(state="disabled", text="Yükleniyor…")
        self._set_status("Veritabanı okunuyor…")

        def work():
            try:
                eng = CapacityEngine("", "").load_from_db(db)
                self.after(0, lambda: self._on_loaded(eng, db))
            except Exception as e:
                tb = traceback.format_exc()
                self.after(0, lambda: self._on_load_error(e, tb))

        threading.Thread(target=work, daemon=True).start()

    def _on_load_error(self, e, tb):
        self.load_btn.configure(state="normal", text="Verileri Yükle")
        self._set_status("Yükleme başarısız.")
        messagebox.showerror("Yükleme hatası",
                             f"{type(e).__name__}: {e}\n\nVeritabanı dosyasını ve tabloları kontrol edin.\n\n{tb[-800:]}")

    def _on_loaded(self, eng: CapacityEngine, db):
        try:
            self._on_loaded_inner(eng, db)
        except Exception as e:
            self.load_btn.configure(state="normal", text="Verileri Yükle")
            self._show_exception(type(e), e, e.__traceback__)

    def _on_loaded_inner(self, eng: CapacityEngine, db):
        self.engine = eng
        self.load_btn.configure(state="normal", text="Verileri Yükle")
        self.settings["db_file"] = db
        save_settings(self.settings)
        self._show_db_info()

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
        if eng.available_months:
            self._recalculate()
        self._set_status(f"Yüklendi: {os.path.basename(db)}")
        if eng.warnings:
            messagebox.showwarning("Veri uyarıları", "Veriler yüklendi ama şunlara dikkat:\n\n• " +
                                   "\n• ".join(eng.warnings[:15]))

    def _show_unmapped(self):
        eng = self.engine
        df = eng.unmapped_df
        if df is None or df.empty:
            self.unmapped_label.configure(text="Tüm plan referansları operasyon dosyasında bulundu. ✓")
            self.unmapped_table.clear()
            return
        qty_cols = eng.week_columns + eng.backlog_columns
        bcol = backlog_col(eng)
        view = df.copy()
        view["TOPLAM ADET"] = view[eng.week_columns].sum(axis=1) if eng.week_columns else 0
        if bcol:   # backlog tek sefer ve eksiye düşmeden sayılır
            view["TOPLAM ADET"] += view[bcol].clip(lower=0)
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
            return float(self.settings.get("default_weekly_capacity", 0))

    def _build_capacity_rows(self, machines):
        for w in self.cap_frame.winfo_children():
            w.destroy()
        self.cap_entries = {}
        caps = self.settings["weekly_capacities"]
        ctk.CTkLabel(self.cap_frame, text="Makine", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, sticky="w")
        ctk.CTkLabel(self.cap_frame, text="Haftalık kapasite (saat)", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=10)
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
        caps = dict(self.settings.get("weekly_capacities", {}))
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
        self.settings["default_weekly_capacity"] = self._default_cap()
        self.settings["weekly_capacities"] = self._read_capacities()
        save_settings(self.settings)
        self._set_status("Kapasite ayarları kaydedildi.")
        if self.engine is not None:
            self._recalculate()

    def _capacity_args(self):
        caps = self._read_capacities() if self.cap_entries else self.settings["weekly_capacities"]
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
            self.summary = CapacityEngine.apply_capacity(raw, months, caps, default_cap,
                                                         self.engine.week_month_map)
        except Exception as e:
            messagebox.showerror("Hesaplama hatası", f"{type(e).__name__}: {e}\n\n{traceback.format_exc()[-800:]}")
            return
        self._show_load_table(months)
        self.period_weeks = {f"{w}  ({m})": w for w, m in self.engine.week_month_map.items()
                             if m in months and f"{w}_Saat" in self.summary.columns}
        periods = [ALL_PERIOD] + months + list(self.period_weeks)
        self.period_combo.configure(values=periods)
        if self.period_combo.get() not in periods:
            self.period_combo.set(ALL_PERIOD)
        self._draw_load_chart()
        self._update_sim_info()
        over = (self.summary["DOLULUK_%"] > OVER_LIMIT).sum()
        self._set_status(f"Hesaplandı: {', '.join(months)}  |  {len(self.summary)} makine  |  "
                         f"kapasite üstü: {over}")

    def _active_weeks(self, df, months):
        return [w for w, m in self.engine.week_month_map.items() if m in months and f"{w}_Saat" in df.columns]

    def _view_summary(self, df: pd.DataFrame, months, mode="Aylık"):
        """Ekranda gösterilecek tablo. Aylık: ay saat / kapasite / %. Haftalık: her hafta saat / %."""
        eng = self.engine
        month_cols = [m for m in months if f"{m}_Saat" in df.columns]
        view = pd.DataFrame({"Makine": df["Üretim presi"].replace("", NO_MACHINE)})
        view["Haftalık kap. (saat)"] = df["HAFTALIK_KAPASİTE"]
        if mode == "Haftalık":
            for w in self._active_weeks(df, months):
                view[f"{w} (saat)"] = df[f"{w}_Saat"]
                view[f"{w} %"] = df[f"{w}_Doluluk_%"]
        else:
            for m in month_cols:
                n = int(df[f"{m}_Hafta"].iloc[0]) if len(df) else 0
                view[f"{m} (saat)"] = df[f"{m}_Saat"]
                view[f"{m} kap. ({n} hf)"] = df[f"{m}_Kapasite"]
                view[f"{m} %"] = df[f"{m}_Doluluk_%"]
        bcol = backlog_col(eng)
        if bcol:
            # sadece Güncel Backlog (backlog - stok); ham Backlog ile toplanırsa aynı iş iki kez sayılır
            view["Güncel Backlog (saat)"] = df[f"{bcol}_Saat"]
        if "STOKTAN_KARSILANAN_SAAT" in df.columns:
            # fazla stokla karşılanıp toplamdan düşülen sipariş saati (bilgi)
            view["Stoktan karşılanan (saat)"] = df["STOKTAN_KARSILANAN_SAAT"]
        view["Toplam (saat)"] = df["TOPLAM_SAAT"]
        view["Kapasite (saat)"] = df["TOPLAM_KAPASİTE"]
        view["Doluluk %"] = df["DOLULUK_%"]
        return view

    def _refresh_load_table(self):
        if self.summary is not None and self.engine is not None:
            self._show_load_table(self._selected_months())

    def _show_load_table(self, months):
        mode = self.view_mode.get() or "Aylık"
        view = self._view_summary(self.summary, months, mode)
        total = {"Makine": "TOPLAM"}
        for c in view.columns[1:]:
            if not c.endswith("%"):
                total[c] = view[c].sum()
        if total.get("Kapasite (saat)"):
            total["Doluluk %"] = total["Toplam (saat)"] / total["Kapasite (saat)"] * 100
        if mode == "Haftalık":
            # satır rengi: makinenin EN DOLU haftasına göre (tek bir hafta taşsa bile görünür)
            wk = [c for c in view.columns if c.endswith(" %") and c != "Doluluk %"]
            tag = lambda r: load_tag(max([r[c] for c in wk if pd.notna(r[c])], default=np.nan))
            self.view_hint.configure(text="Satır rengi: makinenin en dolu haftası")
        else:
            tag = lambda r: load_tag(r["Doluluk %"])
            self.view_hint.configure(text="Ay kapasitesi = haftalık kapasite × o ayın hafta sayısı")
        self.load_table.show(view, tag_func=tag, total_row=total)

    def _draw_load_chart(self):
        if self.summary is None or self.summary.empty:
            self.load_chart.message("Gösterilecek veri yok.")
            return
        df = self.summary.copy()
        period = self.period_combo.get()
        week = self.period_weeks.get(period)
        if week and f"{week}_Saat" in df.columns:
            hours, caps, pcts = df[f"{week}_Saat"], df["HAFTALIK_KAPASİTE"], df[f"{week}_Doluluk_%"]
            title = f"{period} – haftalık makine yükü"
        elif period != ALL_PERIOD and f"{period}_Saat" in df.columns:
            hours = df[f"{period}_Saat"]
            caps = df[f"{period}_Kapasite"]
            pcts = df[f"{period}_Doluluk_%"]
            title = f"{period} – makine yükü ({int(df[f'{period}_Hafta'].iloc[0])} hafta)"
        else:
            hours, caps, pcts = df["TOPLAM_SAAT"], df["TOPLAM_KAPASİTE"], df["DOLULUK_%"]
            title = f"{ALL_PERIOD} – makine yükü"
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
            after = CapacityEngine.apply_capacity(sim_raw, months, caps, default_cap, self.engine.week_month_map)
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


    # ---------------- Operasyonlar (REFERANS PRES) ----------------
    def _build_ops_tab(self):
        t = self.tabs.tab("Operasyonlar")
        bar = ctk.CTkFrame(t, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 4))
        ctk.CTkButton(bar, text="Excel'den Al…", width=110, fg_color=COLOR_OK, hover_color="#24804a",
                      command=self._ops_import_excel).pack(side="left", padx=3)
        ctk.CTkButton(bar, text="Yeni Satır", width=90, command=self._ops_add).pack(side="left", padx=3)
        ctk.CTkButton(bar, text="Düzenle", width=80, command=self._ops_edit).pack(side="left", padx=3)
        ctk.CTkButton(bar, text="Sil", width=60, fg_color="gray55", command=self._ops_delete).pack(side="left", padx=3)
        ctk.CTkButton(bar, text="Tümünü Sil", width=90, fg_color=COLOR_OVER, hover_color="#b03636",
                      command=self._ops_delete_all).pack(side="left", padx=3)
        ctk.CTkLabel(bar, text="  Ara:").pack(side="left")
        self.ops_search = ctk.CTkEntry(bar, width=170, placeholder_text="ürün, pres, müşteri…")
        self.ops_search.pack(side="left", padx=3)
        self.ops_search.bind("<KeyRelease>", lambda _: self._ops_show())
        self.ops_cancel_btn = ctk.CTkButton(bar, text="Vazgeç", width=80, fg_color="gray55", state="disabled",
                                            command=self._ops_cancel)
        self.ops_cancel_btn.pack(side="right", padx=3)
        self.ops_save_btn = ctk.CTkButton(bar, text="Kaydet", width=90, state="disabled",
                                          font=ctk.CTkFont(weight="bold"), command=self._ops_save)
        self.ops_save_btn.pack(side="right", padx=3)
        self.ops_label = ctk.CTkLabel(t, text="", anchor="w")
        self.ops_label.pack(fill="x", padx=4)
        self.ops_table = DataTable(t)
        self.ops_table.pack(fill="both", expand=True)
        self.ops_table.tree.bind("<Double-1>", lambda _: self._ops_edit())
        self.ops_table.tree.bind("<Delete>", lambda _: self._ops_delete())

    def _ops_reload(self):
        """Taslağı veritabanından yeniden okur (kaydedilmemiş değişiklikler gider)."""
        db = self._db_path()
        try:
            self.ops_draft = load_operations(db) if db and os.path.exists(db) else pd.DataFrame(columns=OP_COLUMNS)
        except Exception as e:
            self.ops_draft = pd.DataFrame(columns=OP_COLUMNS)
            messagebox.showerror("Hata", f"Operasyonlar okunamadı:\n{e}")
        self.ops_draft = self.ops_draft.reset_index(drop=True)
        self._ops_set_dirty(False)

    def _ops_set_dirty(self, dirty: bool):
        self.ops_dirty = dirty
        st = "normal" if dirty else "disabled"
        self.ops_save_btn.configure(state=st)
        self.ops_cancel_btn.configure(state=st)
        self._ops_show()

    def _ops_show(self):
        df = self.ops_draft
        sn = pd.to_numeric(df["uretim_saniye"], errors="coerce").fillna(0)
        view = pd.DataFrame({
            "#": [str(i) for i in df.index],
            "MÜŞTERİ": df["musteri"], "BİTMİŞ ÜRÜN KODU": df["urun_kodu"],
            "OPERASYON STOK KODU": df["operasyon_stok_kodu"], "BASILDIĞI PRES TONAJI": df["pres_tonaji"],
            "Üretim presi": df["uretim_presi"], "alternatif 1": df["alternatif_1"],
            "alternatif 2": df["alternatif_2"], "alternatif 3": df["alternatif_3"],
            "alternatif 4": df["alternatif_4"], "ÜRETİM SANİYE": sn,
            "SAATLİK ÜRETİM ADEDİ": np.where(sn > 0, 3600 / sn.replace(0, np.nan), np.nan),
            "VARDİYALIK ÜRETİM ADEDİ ( 7,5 SAAT)": np.where(sn > 0, 3600 * 7.5 / sn.replace(0, np.nan), np.nan),
        })
        bad = (df["uretim_presi"].astype(str).str.strip() == "") | (sn <= 0)
        q = fold(self.ops_search.get()) if self.ops_search.get().strip() else ""
        if q:
            text = view.drop(columns=["#"]).astype(str).apply(lambda col: col.map(fold)).agg(" ".join, axis=1)
            view = view[text.str.contains(re.escape(q), regex=True)]
        badset = set(df.index[bad].astype(str))
        self.ops_table.show(view.reset_index(drop=True), tag_func=lambda r: "over" if r["#"] in badset else "")
        info = f"{len(df)} operasyon satırı, {df['urun_kodu'].nunique()} ürün"
        if q:
            info += f"  |  aramada {len(view)} satır"
        if bad.any():
            info += f"  |  {int(bad.sum())} satırda eksik bilgi (kırmızı: pres veya saniye boş)"
        if self.ops_dirty:
            info += "   ●  KAYDEDİLMEMİŞ DEĞİŞİKLİK VAR – 'Kaydet' demeden hesaba yansımaz"
        self.ops_label.configure(text=info, text_color=COLOR_OVER if self.ops_dirty else "#333")

    def _ops_selected(self):
        return [int(self.ops_table.tree.item(i, "values")[0]) for i in self.ops_table.tree.selection()]

    def _ops_machine_names(self):
        d = self.ops_draft
        vals = pd.unique(d[["uretim_presi", "alternatif_1", "alternatif_2", "alternatif_3", "alternatif_4"]].values.ravel())
        return sorted(v for v in vals if isinstance(v, str) and v)

    def _ops_add(self):
        OpDialog(self, None, self._ops_machine_names(), self._ops_store)

    def _ops_edit(self):
        sel = self._ops_selected()
        if len(sel) != 1:
            messagebox.showinfo("Seçim", "Düzenlemek için listeden tek bir satır seçin (veya çift tıklayın).")
            return
        OpDialog(self, self.ops_draft.loc[sel[0]].to_dict(), self._ops_machine_names(),
                 lambda rec: self._ops_store(rec, sel[0]))

    def _ops_store(self, rec, idx=None):
        if idx is None:
            self.ops_draft = pd.concat([self.ops_draft, pd.DataFrame([rec], columns=OP_COLUMNS)], ignore_index=True)
        else:
            for c in OP_COLUMNS:
                self.ops_draft.at[idx, c] = rec[c]
        self._ops_set_dirty(True)

    def _ops_delete(self):
        sel = self._ops_selected()
        if not sel:
            messagebox.showinfo("Seçim", "Silmek için listeden satır seçin (Ctrl / Shift ile birden fazla seçilebilir).")
            return
        if not messagebox.askyesno("Sil", f"{len(sel)} satır silinsin mi?\n(Kaydet'e basana kadar geri alınabilir: Vazgeç)"):
            return
        self.ops_draft = self.ops_draft.drop(index=sel).reset_index(drop=True)
        self._ops_set_dirty(True)

    def _ops_delete_all(self):
        if self.ops_draft.empty:
            return
        if not messagebox.askyesno("Tümünü sil", f"{len(self.ops_draft)} satırın TAMAMI silinsin mi?\n"
                                                 "(Kaydet'e basana kadar geri alınabilir: Vazgeç)"):
            return
        self.ops_draft = pd.DataFrame(columns=OP_COLUMNS)
        self._ops_set_dirty(True)

    def _ops_cancel(self):
        if messagebox.askyesno("Vazgeç", "Kaydedilmemiş tüm değişiklikler iptal edilsin mi?"):
            self._ops_reload()

    def _ops_save(self):
        db = self._db_path()
        if not db:
            messagebox.showwarning("Eksik", "Önce bir veritabanı seçin.")
            return
        bad = int(((self.ops_draft["uretim_presi"].astype(str).str.strip() == "") |
                   (pd.to_numeric(self.ops_draft["uretim_saniye"], errors="coerce").fillna(0) <= 0)).sum())
        msg = f"{len(self.ops_draft)} operasyon satırı veritabanına kaydedilecek (eski liste bununla değişir)."
        if bad:
            msg += f"\n\n{bad} satırda pres veya saniye boş; bu satırlar saate çevrilemez."
        if not messagebox.askyesno("Kaydet", msg + "\n\nKaydedilsin mi?"):
            return
        try:
            n = save_operations(db, self.ops_draft)
        except Exception as e:
            messagebox.showerror("Kaydedilemedi", f"Hiçbir şey değişmedi.\n\n{type(e).__name__}: {e}")
            return
        self._ops_reload()
        self._show_db_info()
        self._set_status(f"{n} operasyon satırı kaydedildi.")
        self._load_data()

    def _ops_import_excel(self):
        path = filedialog.askopenfilename(title="REFERANS PRES dosyasını seçin",
                                          filetypes=[("Excel dosyaları", "*.xlsx *.xlsm *.xls")])
        if not path:
            return
        try:
            names = sheet_names(path)
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya açılamadı:\n{e}")
            return
        sheet = OP_SHEET_DEFAULT if OP_SHEET_DEFAULT in names else None
        if sheet is None:
            for n in names:   # başlığı 'BİTMİŞ ÜRÜN KODU' olan ilk sayfa
                try:
                    parse_operations(path, n)
                    sheet = n
                    break
                except Exception:
                    continue
        if sheet is None:
            messagebox.showerror("Bulunamadı", "Dosyada 'BİTMİŞ ÜRÜN KODU' başlıklı bir operasyon sayfası bulunamadı.")
            return
        try:
            df, warns = parse_operations(path, sheet)
        except Exception as e:
            messagebox.showerror("Okunamadı", str(e))
            return
        msg = f"{os.path.basename(path)} / {sheet}\n{len(df)} operasyon satırı, {df['urun_kodu'].nunique()} ürün okundu."
        if warns:
            msg += "\n\n• " + "\n• ".join(warns)
        if not self.ops_draft.empty:
            ans = messagebox.askyesnocancel(
                "Excel'den Al", msg + f"\n\nListede şu an {len(self.ops_draft)} satır var.\n"
                "EVET = mevcut listenin yerine geçsin\nHAYIR = mevcut listenin sonuna eklensin\nİPTAL = vazgeç")
            if ans is None:
                return
            self.ops_draft = df if ans else pd.concat([self.ops_draft, df], ignore_index=True)
        else:
            if not messagebox.askyesno("Excel'den Al", msg + "\n\nListeye alınsın mı?"):
                return
            self.ops_draft = df
        self.ops_draft = self.ops_draft.reset_index(drop=True)
        self._ops_set_dirty(True)
        messagebox.showinfo("Taslağa alındı", "Satırlar listeye alındı. Kontrol edip 'Kaydet'e basınca veritabanına yazılır.")


class OpDialog(ctk.CTkToplevel):
    """Tek operasyon satırını ekleme / düzenleme penceresi (REFERANS PRES şablonundaki sütunlar)."""
    FIELDS = [("musteri", "MÜŞTERİ", False), ("urun_kodu", "BİTMİŞ ÜRÜN KODU", True),
              ("operasyon_stok_kodu", "OPERASYON STOK KODU", False), ("pres_tonaji", "BASILDIĞI PRES TONAJI", False),
              ("uretim_presi", "Üretim presi", True), ("alternatif_1", "alternatif 1", False),
              ("alternatif_2", "alternatif 2", False), ("alternatif_3", "alternatif 3", False),
              ("alternatif_4", "alternatif 4", False), ("uretim_saniye", "ÜRETİM SANİYE", True)]
    MACHINE_FIELDS = {"uretim_presi", "alternatif_1", "alternatif_2", "alternatif_3", "alternatif_4"}

    def __init__(self, master, rec, machines, on_ok):
        super().__init__(master)
        self.title("Operasyon düzenle" if rec else "Yeni operasyon")
        self.geometry("460x520")
        self.transient(master)
        self.grab_set()
        self.on_ok = on_ok
        self.machines = list(machines)
        self.widgets = {}
        f = ctk.CTkFrame(self)
        f.pack(fill="both", expand=True, padx=12, pady=12)
        for i, (key, label, req) in enumerate(self.FIELDS):
            ctk.CTkLabel(f, text=label + (" *" if req else ""), anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            if key in self.MACHINE_FIELDS:
                w = ctk.CTkComboBox(f, values=[""] + machines, width=220)
                w.set("")
            else:
                w = ctk.CTkEntry(f, width=220)
            w.grid(row=i, column=1, padx=8, pady=4)
            val = "" if rec is None else rec.get(key, "")
            if key == "uretim_saniye" and rec is not None:
                val = fmt_input(val) if val not in ("", None) and float(val or 0) > 0 else ""
            if isinstance(w, ctk.CTkComboBox):
                w.set(str(val or ""))
            elif val not in ("", None):
                w.insert(0, str(val))
            self.widgets[key] = w
        ctk.CTkLabel(f, text="* zorunlu. Pres adları listeden seçilebilir veya yazılabilir.",
                     text_color="#666").grid(row=len(self.FIELDS), column=0, columnspan=2, sticky="w", padx=8, pady=6)
        b = ctk.CTkFrame(self, fg_color="transparent")
        b.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(b, text="Tamam", command=self._ok).pack(side="right", padx=4)
        ctk.CTkButton(b, text="İptal", fg_color="gray55", command=self.destroy).pack(side="right", padx=4)

    def _ok(self):
        from capacity_engine import normalize_code, normalize_machine
        rec = {k: self.widgets[k].get().strip() for k, _, _ in self.FIELDS}
        rec["urun_kodu"] = normalize_code(rec["urun_kodu"])
        rec["operasyon_stok_kodu"] = normalize_code(rec["operasyon_stok_kodu"])
        for k in self.MACHINE_FIELDS:
            rec[k] = normalize_machine(rec[k])
        try:
            rec["uretim_saniye"] = float(rec["uretim_saniye"].replace(",", "."))
        except ValueError:
            rec["uretim_saniye"] = -1
        errors = []
        if not rec["urun_kodu"]:
            errors.append("BİTMİŞ ÜRÜN KODU boş olamaz.")
        if not rec["uretim_presi"]:
            errors.append("Üretim presi boş olamaz.")
        if rec["uretim_saniye"] <= 0:
            errors.append("ÜRETİM SANİYE 0'dan büyük bir sayı olmalı (ör. 4,5).")
        if errors:
            messagebox.showwarning("Eksik bilgi", "\n".join(errors), parent=self)
            return
        # listede olmayan pres adı: yazım hatası olabilir (ör. MK66 yerine MK066)
        import difflib
        for k in sorted(self.MACHINE_FIELDS):
            m = rec[k]
            if m and self.machines and m not in self.machines:
                near = difflib.get_close_matches(m, self.machines, n=3, cutoff=0.6)
                hint = f"\nBenzer olanlar: {', '.join(near)}" if near else ""
                if not messagebox.askyesno("Yeni pres adı",
                                           f"'{m}' listede yok.{hint}\n\nYeni bir pres olarak kaydedilsin mi?\n"
                                           "(Hayır derseniz düzeltebilirsiniz)", parent=self):
                    return
        self.on_ok(rec)
        self.destroy()


class ImportDialog(ctk.CTkToplevel):
    """Plan Excel'ini seç → oku → yılları kontrol et → veritabanına yaz."""

    def __init__(self, master, db_path, on_done):
        super().__init__(master)
        self.title("Excel'den Veritabanına Aktar")
        self.geometry("780x600")
        self.transient(master)
        self.grab_set()
        self.db_path, self.on_done, self.plan = db_path, on_done, None
        self.year_entries = []

        f = ctk.CTkFrame(self)
        f.pack(fill="x", padx=12, pady=(12, 6))
        f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(f, text="Excel dosyası:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        self.file_entry = ctk.CTkEntry(f)
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=6)
        ctk.CTkButton(f, text="Seç…", width=70, command=self._pick).grid(row=0, column=2, padx=6)
        ctk.CTkLabel(f, text="Sayfa:").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        self.sheet = ctk.CTkComboBox(f, values=[""], width=300)
        self.sheet.grid(row=1, column=1, sticky="w", pady=6)
        ctk.CTkButton(f, text="Oku", width=70, command=self._read).grid(row=1, column=2, padx=6)

        ctk.CTkLabel(self, text="Bulunan aylar (yılı kontrol edin; Excel'de yıl yazmıyorsa siz girin):",
                     anchor="w").pack(fill="x", padx=16)
        self.blocks_frame = ctk.CTkScrollableFrame(self, height=150)
        self.blocks_frame.pack(fill="x", padx=12, pady=4)
        self.info = ctk.CTkTextbox(self, height=180)
        self.info.pack(fill="both", expand=True, padx=12, pady=4)
        b = ctk.CTkFrame(self, fg_color="transparent")
        b.pack(fill="x", padx=12, pady=(4, 12))
        self.write_btn = ctk.CTkButton(b, text="Veritabanına Yaz", state="disabled", fg_color=COLOR_OK,
                                       hover_color="#24804a", command=self._write)
        self.write_btn.pack(side="right", padx=4)
        ctk.CTkButton(b, text="Kapat", fg_color="gray55", command=self.destroy).pack(side="right", padx=4)
        self._set_info("Plan Excel'ini seçip 'Oku'ya basın.")

    def _set_info(self, text):
        self.info.configure(state="normal")
        self.info.delete("1.0", "end")
        self.info.insert("1.0", text)
        self.info.configure(state="disabled")

    def _pick(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("Excel dosyaları", "*.xlsx *.xlsm *.xls")])
        if not path:
            return
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, path)
        try:
            names = sheet_names(path)
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya açılamadı:\n{e}", parent=self)
            return
        self.sheet.configure(values=names)
        self.sheet.set("Diğer müşteriler DAHİL" if "Diğer müşteriler DAHİL" in names else names[0])
        self._read()

    def _read(self):
        path, sheet = self.file_entry.get().strip(), self.sheet.get()
        if not path:
            return
        self.write_btn.configure(state="disabled")
        for w in self.blocks_frame.winfo_children():
            w.destroy()
        self.year_entries = []
        try:
            self.plan = plan = parse_plan(path, sheet)
        except Exception as e:
            self.plan = None
            self._set_info(f"OKUNAMADI\n\n{e}")
            return
        for i, blk in enumerate(plan.blocks):
            n = int((plan.orders["blok"] == i).sum())
            weeks = ", ".join(f"W{h}" for h in sorted(set(blk.weeks.values())))
            row = ctk.CTkFrame(self.blocks_frame, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=f"{blk.ay}", width=90, anchor="w",
                         font=ctk.CTkFont(weight="bold")).pack(side="left", padx=4)
            ctk.CTkLabel(row, text="Yıl:").pack(side="left")
            e = ctk.CTkEntry(row, width=70, placeholder_text="ör. 2026")
            if blk.yil:
                e.insert(0, str(blk.yil))
            e.pack(side="left", padx=4)
            note = "" if blk.yil else "  ← Excel'de yıl yok, girin"
            ctk.CTkLabel(row, text=f"Excel'de: '{blk.label}'  |  {weeks}  |  {n} sipariş satırı{note}",
                         anchor="w", text_color="#a33" if not blk.yil else "#444").pack(side="left", padx=6)
            self.year_entries.append(e)
        lines = [f"Dosya: {os.path.basename(path)}  /  Sayfa: {sheet}",
                 f"Ürün (ürün + firma): {plan.n_products}",
                 f"Sipariş satırı (adet > 0): {len(plan.orders)}",
                 f"Backlog/stok satırı: {len(plan.backlog)}  → {plan.blocks[0].ay} ayına yazılacak"]
        if plan.warnings:
            lines += ["", "UYARILAR:"] + [f"• {w}" for w in plan.warnings[:30]]
        lines += ["", "Yılları kontrol edip 'Veritabanına Yaz'a basın."]
        self._set_info("\n".join(lines))
        self.write_btn.configure(state="normal")

    def _years(self):
        years = {}
        for i, e in enumerate(self.year_entries):
            txt = e.get().strip()
            if not (txt.isdigit() and 2000 <= int(txt) <= 2100):
                raise ValueError(f"{self.plan.blocks[i].ay} için geçerli bir yıl girin (ör. 2026).")
            years[i] = int(txt)
        return years

    def _write(self):
        if self.plan is None:
            return
        try:
            years = self._years()
        except ValueError as e:
            messagebox.showwarning("Yıl eksik", str(e), parent=self)
            return
        try:
            notes = check_conflicts(self.db_path, self.plan, years)
        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanı kontrol edilemedi:\n{e}", parent=self)
            return
        summary = ", ".join(f"{b.ay} {years[i]}" for i, b in enumerate(self.plan.blocks))
        msg = f"Şu aylar veritabanına yazılacak:\n{summary}"
        if self.plan.orders.empty:
            msg += "\n\nUYARI: Dosyada hiç sipariş adedi bulunamadı (tüm haftalar boş veya 0)."
        if notes:
            msg += "\n\nDİKKAT:\n• " + "\n• ".join(notes)
        if not messagebox.askyesno("Onay", msg + "\n\nDevam edilsin mi?", parent=self):
            return
        try:
            ids = import_plan(self.db_path, self.plan, years)
        except Exception as e:
            messagebox.showerror("Yazılamadı", f"Hiçbir şey yazılmadı.\n\n{type(e).__name__}: {e}", parent=self)
            return
        messagebox.showinfo("Tamam", f"{len(ids)} ay veritabanına yazıldı ({summary}).", parent=self)
        self.destroy()
        self.on_done()


def fmt_input(v) -> str:
    """Giriş kutuları için sayı biçimi: 450.0 -> '450', 437.5 -> '437,5'."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return str(int(v)) if v.is_integer() else str(v).replace(".", ",")


if __name__ == "__main__":
    App().mainloop()

"""
Soylu Makine - Pres Kapasite Hesaplama Uygulaması
Arayüz: CustomTkinter  |  Hesap motoru: capacity_engine.py
"""
import json
from datetime import datetime
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

from capacity_engine import CapacityEngine, normalize_machine, normalize_code
from veritabani import (create_empty_db, table_counts, list_imports, undo_import, is_synced_folder, check_db,
                        is_network_path, other_sessions, register_session, unregister_session,
                        load_operations, save_operations, OP_COLUMNS,
                        load_valid_orders, save_manual_orders, AY_ADLARI)
from excel_aktar import (sheet_names, parse_plan, check_conflicts, import_plan, parse_operations, OP_SHEET_DEFAULT,
                         fold, week_month_problem, XML_PROTECTED)
from veritabani import hafta_yili, normalize_firma

APP_TITLE = "Soylu Makine – Pres Kapasite Hesaplama"
DEFAULT_DB_NAME = "kapasite.db"
ALT_COLS = ["alternatif 1", "alternatif 2", "alternatif 3", "alternatif 4"]
NO_MACHINE = "(TANIMSIZ)"

DEFAULT_WEEKLY_CAP = 105.0     # saat/hafta (5 gün × 3 vardiya × 7 saat)
WEEKS_PER_MONTH = 52 / 12      # eski aylık ayarları çevirmek için
ALL_PERIOD = "Seçili aylar toplamı"

COLOR_OK = "#2e9d5b"
COLOR_WARN_TEXT = "#b25e00"
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
ICON_FILE = "soylu_kapasite.ico"


def resource_path(name: str) -> str:
    """Exe içine gömülen dosyanın yolu (PyInstaller) ya da script klasöründeki dosya."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def set_icon(win):
    """Pencere / görev çubuğu simgesi. CustomTkinter kendi simgesini ~200 ms sonra koyduğu için biraz sonra uygulanır."""
    path = resource_path(ICON_FILE)
    if not os.path.exists(path):
        return
    def apply():
        try:
            win.iconbitmap(path)
        except Exception:
            pass
    apply()
    try:
        win.after(300, apply)
    except Exception:
        pass


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
        self.tree.tag_configure("muted", foreground="#9a9a9a")
        self.tree.tag_configure("moved", background="#dbe8fb")
        self._df = None
        self._hide = []
        self._total_row = None
        self._tag_func = None
        self._sort_state = {}

    def show(self, df: pd.DataFrame, tag_func=None, total_row: dict = None, hide=None):
        """hide: tabloda tutulan ama ekranda gösterilmeyen sütunlar (ör. iş numarası)."""
        self._df = df.reset_index(drop=True)
        self._tag_func = tag_func
        self._total_row = total_row
        self._hide = list(hide or [])
        cols = list(self._df.columns)
        self.tree.delete(*self.tree.get_children())
        # Önce "hepsini göster"e dön: eski görünümün sütun adları yeni sütunlarda yoksa Tk hata verir
        self.tree["displaycolumns"] = "#all"
        self.tree["columns"] = cols
        if self._hide:
            self.tree["displaycolumns"] = [c for c in cols if c not in self._hide]
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

    def selected_records(self) -> list:
        """Seçili satırlar (toplam satırı hariç), tablodaki verinin sözlükleri olarak."""
        if self._df is None:
            return []
        out = []
        for item in self.tree.selection():
            i = self.tree.index(item)
            if i < len(self._df):
                out.append(self._df.iloc[i].to_dict())
        return out

    def select_where(self, col, value):
        """col == value olan ilk satırı seçip görünür yapar."""
        if self._df is None or col not in self._df.columns:
            return
        hits = self._df.index[self._df[col] == value]
        if len(hits):
            item = self.tree.get_children()[int(hits[0])]
            self.tree.selection_set(item)
            self.tree.see(item)

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["displaycolumns"] = "#all"   # önce sıfırla, sonra sütunları kaldır (yoksa Tk hata verir)
        self.tree["columns"] = []
        self._df = None           # başlığa tıklanınca eski veri geri gelmesin
        self._total_row = None


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
        set_icon(self)
        self.geometry("1350x820")
        self.minsize(1100, 680)

        self.report_callback_exception = self._show_exception
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.session_db = None   # bu oturumun "açık" olarak işaretlediği veritabanı
        self.settings = load_settings()
        self.engine: CapacityEngine = None
        self.summary: pd.DataFrame = None       # kapasite eklenmiş güncel özet
        self.sim_result: pd.DataFrame = None    # son simülasyon karşılaştırması
        self.sim_key = None                      # simülasyonun yapıldığı aylar / kapasite
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

    def _on_close(self):
        if self.session_db:
            unregister_session(self.session_db)
        self.destroy()

    def _claim_session(self, db):
        """Bu veritabanını 'açık' olarak işaretle; başka biri açık tutuyorsa uyar."""
        if self.session_db and os.path.abspath(self.session_db) != os.path.abspath(db):
            unregister_session(self.session_db)
        others = other_sessions(db)
        register_session(db)
        first_time = os.path.abspath(self.session_db or "") != os.path.abspath(db)
        self.session_db = db
        if others and first_time:
            liste = "\n".join(f"  • {s.get('kullanici')} (açılış: {s.get('zaman', '').replace('T', ' ')})" for s in others)
            messagebox.showwarning(
                "Veritabanı başka yerde de açık",
                f"Bu veritabanı şu anda başka bir yerde de açık görünüyor:\n{liste}\n\n"
                "Aynı anda iki kişi yükleme / düzenleme yaparsa değişiklikler birbirini etkileyebilir. "
                "Mümkünse diğer kişiyle haberleşin; yalnızca bakıyorsanız sorun yok.\n\n"
                "(Bilgisayar kapanmadan önce program kapatılmadıysa bu uyarı yanlış olabilir; 12 saat sonra kendiliğinden kalkar.)")

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
        ctk.CTkButton(bar, text="Operasyonlar", width=120, fg_color="gray55",
                      command=self._open_ops).grid(row=0, column=4, padx=6, pady=(10, 2))
        ctk.CTkButton(bar, text="Siparişler", width=110, fg_color="gray55",
                      command=self._open_orders).grid(row=0, column=5, padx=6, pady=(10, 2))

        self.db_info = ctk.CTkLabel(bar, text="", anchor="w", text_color="#555")
        self.db_info.grid(row=1, column=1, columnspan=5, sticky="w", pady=(0, 8))

        self.import_btn = ctk.CTkButton(bar, text="Excel'den Aktar", height=56, width=140,
                                        fg_color=COLOR_OK, hover_color="#24804a",
                                        font=ctk.CTkFont(size=13, weight="bold"), command=self._open_import)
        self.import_btn.grid(row=0, column=6, rowspan=2, padx=(10, 4), pady=8)
        self.load_btn = ctk.CTkButton(bar, text="Verileri Yükle", height=56, width=150,
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self._load_data)
        self.load_btn.grid(row=0, column=7, rowspan=2, padx=(4, 10), pady=8)

        self.status = ctk.CTkLabel(self, text="", anchor="w", text_color="#555")
        self.status.pack(side="bottom", fill="x", padx=14, pady=(0, 6))

    def _open_ops(self):
        if self._db_problem():
            return
        db = self._db_path()
        if not db:
            messagebox.showwarning("Eksik", "Önce bir veritabanı seçin veya 'Yeni DB oluştur'.")
            return
        if not os.path.exists(db):
            if not messagebox.askyesno("Veritabanı yok", f"{db}\nbulunamadı. Boş veritabanı oluşturulsun mu?"):
                return
            create_empty_db(db)
            self._show_db_info()
        OpsWindow(self, db)

    def _open_orders(self):
        if self._db_problem():
            return
        db = self._db_path()
        if not db or not os.path.exists(db):
            messagebox.showwarning("Veritabanı yok", "Önce bir veritabanı seçin ve plan Excel'ini aktarın.")
            return
        OrdersWindow(self, db)

    def _db_path(self) -> str:
        return self.db_entry.get().strip()

    def _db_problem(self, show=True):
        """Seçili dosya bozuk / veritabanı değilse mesaj gösterir ve True döner."""
        problem = check_db(self._db_path()) if self._db_path() else None
        if problem and show:
            messagebox.showerror("Veritabanı sorunu", problem)
        return bool(problem)

    def _show_db_info(self):
        path = self._db_path()
        if not path or not os.path.exists(path):
            self.db_info.configure(text="Dosya bulunamadı – 'Yeni DB oluştur' ile boş veritabanı oluşturabilirsiniz.",
                                   text_color="#555")
            return
        problem = check_db(path)
        if problem:
            self.db_info.configure(text="✖ " + problem, text_color=COLOR_OVER, justify="left")
            return
        try:
            c = table_counts(path)
        except Exception as e:
            self.db_info.configure(text=f"Veritabanı okunamadı: {e}")
            return
        parts = [f"{t}: {'tablo yok' if n is None else n}" for t, n in c.items()]
        text = "Kayıt sayıları →  " + "   |   ".join(parts)
        warnings = []
        others = other_sessions(path)
        if others:
            warnings.append("⚠ Şu anda başka yerde de açık: " + ", ".join(o.get("kullanici", "?") for o in others))
        if is_network_path(path):
            warnings.append("⚠ Veritabanı AĞ klasöründe: aynı anda birden fazla kişi kullanırsa çakışma / bozulma "
                            "riski var. Tek kişi kullanacaksa yerel bir klasöre (ör. C:\\SoyluKapasite) alın.")
        elif is_synced_folder(path):
            # OneDrive vb. SQLite dosyasını senkron sırasında kilitleyebilir / çakışma kopyası oluşturabilir
            warnings.append("⚠ Veritabanı OneDrive (senkron) klasöründe: kilitlenme veya kayıt kaybı riski var. "
                            "OneDrive dışında bir klasöre (ör. C:\\SoyluKapasite) taşımanız önerilir.")
        self.db_info.configure(text="\n".join([text] + warnings),
                               text_color=COLOR_WARN_TEXT if warnings else "#555", justify="left")

    def _pick_db(self):
        path = filedialog.askopenfilename(filetypes=[("SQLite veritabanı", "*.db *.sqlite *.sqlite3"), ("Tümü", "*.*")])
        if path:
            self.db_entry.delete(0, "end")
            self.db_entry.insert(0, path)
            self._show_db_info()
            self._refresh_imports()

    def _new_db(self):
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

    # ---------------- Excel'den aktarma ----------------
    def _open_import(self):
        if self._db_problem():
            return
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
            "Tür": "üstüne eklendi" if r.get("mod") == "ekle" else "yerine geçti",
            "Kullanıcı": (r.get("kullanici") or "") + (f"  (geri alan: {r['geri_alan']}, {r.get('geri_alma_tarihi') or ''})"
                                                     if r.get("geri_alan") else ""),
        } for r in rows])
        n_valid = sum(r["gecerli"] for r in rows)
        self.imports_label.configure(text=f"{len(rows)} yükleme, {n_valid} tanesi hesapta. "
                                          "'Yerine geçti': o ayın öncekilerin yerine geçer. 'Üstüne eklendi': o aya eklenir.")
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
        for name in ("Makine Yükü", "Grafik", "Simülasyon", "Kapasite", "Eşleşmeyenler", "Aktarımlar"):
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

    SIM_HINT = "1) Makine ve iş seçin   2) Hedef ve miktarı girip 'Listeye Ekle'   3) Sonuç burada görünür."

    def _build_sim_tab(self):
        """Serbest simülasyon: iş iş seç, tanımlı alternatife saat / % / adet olarak aktar,
        birden çok aktarımı bir senaryoda birlikte gör."""
        t = self.tabs.tab("Simülasyon")
        self.sim_transfers = []     # [{"row", "source", "ref", "op", "target", "ratio"}]
        self.sim_jobs = None        # seçili makinenin iş tablosu (engine.job_list)
        self.sim_cmp = None         # son karşılaştırma (önce / sonra, kapasiteli)
        self.sim_scenario = None    # rapor için senaryo listesi
        self.summary_key = None

        outer = ttk.PanedWindow(t, orient="vertical")
        outer.pack(fill="both", expand=True)
        top = ttk.PanedWindow(outer, orient="horizontal")

        # ---- 1) İş seçimi ----
        left = ctk.CTkFrame(top)
        bar = ctk.CTkFrame(left, fg_color="transparent")
        bar.pack(fill="x", padx=6, pady=(6, 2))
        ctk.CTkLabel(bar, text="1) Makine:", font=ctk.CTkFont(weight="bold")).pack(side="left")
        self.sim_machine = ctk.CTkComboBox(bar, values=[""], width=140, command=lambda _: self._on_sim_machine())
        self.sim_machine.pack(side="left", padx=6)
        self.sim_only_movable = ctk.CTkCheckBox(bar, text="Sadece aktarılabilir işler",
                                                command=self._show_jobs)
        self.sim_only_movable.pack(side="left", padx=10)
        # bilgi satırı ayrı satırda: dar pencerede kesilmesin
        self.sim_info = ctk.CTkLabel(left, text="", anchor="w", justify="left", text_color="#555")
        self.sim_info.pack(fill="x", padx=8)

        # Form tablodan ÖNCE alta yerleşir: pencere küçülünce form değil tablo daralır.
        form = ctk.CTkFrame(left, fg_color="transparent")
        form.pack(side="bottom", fill="x", padx=6, pady=(2, 6))

        self.sim_jobs_table = DataTable(left)
        self.sim_jobs_table.pack(fill="both", expand=True, padx=6, pady=2)
        self.sim_jobs_table.tree.bind("<<TreeviewSelect>>", lambda _: self._on_job_select())

        ctk.CTkLabel(form, text="Hedef:").grid(row=0, column=0, sticky="w")
        self.sim_target = ctk.CTkComboBox(form, values=[""], width=120, state="disabled",
                                          command=lambda _: self._update_amount_hint())
        self.sim_target.grid(row=0, column=1, padx=4)
        ctk.CTkLabel(form, text="Miktar:").grid(row=0, column=2, padx=(10, 2))
        self.sim_amount = ctk.CTkEntry(form, width=90)
        self.sim_amount.grid(row=0, column=3, padx=2)
        self.sim_amount.bind("<KeyRelease>", lambda _: self._update_amount_hint())
        self.sim_amount.bind("<Return>", lambda _: self._add_transfer())
        self.sim_unit = ctk.CTkSegmentedButton(form, values=["Saat", "%", "Adet"],
                                               command=lambda _: self._update_amount_hint())
        self.sim_unit.set("Saat")
        self.sim_unit.grid(row=0, column=4, padx=4)
        # Düğmeler ikinci satırda: dar pencerede de sığsın.
        ctk.CTkButton(form, text="Kalanın tümü", width=100, fg_color="gray55",
                      command=self._fill_remaining).grid(row=1, column=2, columnspan=2, sticky="e",
                                                         padx=2, pady=(4, 0))
        self.sim_add_btn = ctk.CTkButton(form, text="Listeye Ekle", width=110, state="disabled",
                                         command=self._add_transfer)
        self.sim_add_btn.grid(row=1, column=4, sticky="w", padx=4, pady=(4, 0))
        self.sim_amount_hint = ctk.CTkLabel(form, text="Listeden bir iş seçin.", anchor="w",
                                            justify="left", text_color="#555")
        self.sim_amount_hint.grid(row=2, column=0, columnspan=5, sticky="w", pady=(4, 0))
        top.add(left, weight=3)

        # ---- 2) Senaryo ----
        right = ctk.CTkFrame(top)
        ctk.CTkLabel(right, text="2) Senaryo – aktarım listesi", font=ctk.CTkFont(weight="bold"),
                     anchor="w").pack(fill="x", padx=6, pady=(8, 2))
        btns = ctk.CTkFrame(right, fg_color="transparent")
        btns.pack(side="bottom", fill="x", padx=6, pady=(2, 6))
        self.sim_list_table = DataTable(right)
        self.sim_list_table.pack(fill="both", expand=True, padx=6, pady=2)
        ctk.CTkButton(btns, text="Seçileni Sil", width=110, fg_color=COLOR_OVER, hover_color="#b03636",
                      command=self._remove_transfer).pack(side="left")
        ctk.CTkButton(btns, text="Tümünü Temizle", width=120, fg_color="gray55",
                      command=self._clear_transfers).pack(side="left", padx=6)
        self.sim_btn = ctk.CTkButton(btns, text="Yeniden Hesapla", width=120, state="disabled",
                                     command=lambda: self._run_simulation(force=True))
        self.sim_btn.pack(side="right")
        top.add(right, weight=2)
        outer.add(top, weight=3)

        # ---- 3) Sonuç ----
        res = ctk.CTkFrame(outer)
        rbar = ctk.CTkFrame(res, fg_color="transparent")
        rbar.pack(fill="x", padx=6, pady=(6, 2))
        ctk.CTkLabel(rbar, text="3) Sonuç", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkLabel(rbar, text="Tablo:").pack(side="left", padx=(16, 4))
        self.sim_view = ctk.CTkSegmentedButton(rbar, values=["Toplam", "Aylık", "Haftalık"],
                                               command=lambda _: self._show_sim_table())
        self.sim_view.set("Toplam")
        self.sim_view.pack(side="left")
        self.sim_only_changed = ctk.CTkCheckBox(rbar, text="Sadece etkilenen makineler",
                                                command=self._show_sim_table)
        self.sim_only_changed.select()
        self.sim_only_changed.pack(side="left", padx=12)
        self.sim_summary = ctk.CTkLabel(res, text=self.SIM_HINT, anchor="w", justify="left", text_color="#555")
        self.sim_summary.pack(fill="x", padx=8)
        rpane = ttk.PanedWindow(res, orient="horizontal")
        rpane.pack(fill="both", expand=True, padx=4, pady=4)
        self.sim_chart = ChartPanel(rpane)
        self.sim_chart.message("Senaryoya aktarım ekleyince sonuç burada görünür.")
        self.sim_table = DataTable(rpane)
        rpane.add(self.sim_chart, weight=1)
        rpane.add(self.sim_table, weight=1)
        outer.add(res, weight=4)

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
        self.cap_dirty_label = ctk.CTkLabel(top, text="", text_color=COLOR_WARN_TEXT,
                                            font=ctk.CTkFont(weight="bold"))
        self.cap_dirty_label.pack(side="right", padx=8)
        self._entry_border = self.default_cap_entry.cget("border_color")   # normal çerçeve rengi
        self.default_cap_entry.bind("<KeyRelease>", lambda _: self._mark_cap_dirty())
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
        if self._db_problem():
            return
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
        self._claim_session(db)
        self._clear_results()
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

        has_data = bool(eng.available_months)
        for b in (self.calc_btn, self.export_btn, self.sim_btn):
            b.configure(state="normal" if has_data else "disabled")
        if not has_data:
            self.load_chart.message("Veritabanında geçerli sipariş yok – 'Excel'den Aktar' ile plan yükleyin.")
            ctk.CTkLabel(self.month_frame, text="(sipariş yok)", text_color="#888").pack()

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
        # Kutulara yazılmış ama henüz kaydedilmemiş değerler yeniden yüklemede kaybolmasın
        typed = {m: e.get() for m, e in self.cap_entries.items()}
        for w in self.cap_frame.winfo_children():
            w.destroy()
        self.cap_entries = {}
        caps = self.settings["weekly_capacities"]
        ctk.CTkLabel(self.cap_frame, text="Makine", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, sticky="w")
        ctk.CTkLabel(self.cap_frame, text="Haftalık kapasite (saat)", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=10)
        for i, m in enumerate(machines, start=1):
            ctk.CTkLabel(self.cap_frame, text=m, width=140, anchor="w").grid(row=i, column=0, padx=10, pady=2, sticky="w")
            e = ctk.CTkEntry(self.cap_frame, width=110, placeholder_text=fmt_input(self._default_cap()))
            val = typed[m] if m in typed else (fmt_input(caps[m]) if m in caps else "")
            if val:
                e.insert(0, val)
            e.grid(row=i, column=1, padx=10, pady=2)
            e.bind("<KeyRelease>", lambda _: self._mark_cap_dirty())
            self.cap_entries[m] = e
        self._mark_cap_dirty()

    @staticmethod
    def _cap_changed(txt, saved) -> bool:
        """Kutudaki yazı kayıtlı değerden farklı mı? (saved None = kayıtlı değer yok)"""
        txt = txt.strip()
        if not txt:
            return saved is not None
        try:
            return saved is None or abs(float(txt.replace(",", ".")) - float(saved)) > 1e-9
        except ValueError:
            return True

    def _mark_cap_dirty(self):
        """Kaydedilmemiş kapasite kutularını turuncu çerçeveyle göster, üstte uyarı yaz."""
        caps = self.settings.get("weekly_capacities", {})
        n = 0
        for m, e in self.cap_entries.items():
            changed = self._cap_changed(e.get(), caps.get(m))
            e.configure(border_color=COLOR_WARN if changed else self._entry_border)
            n += changed
        d = self.default_cap_entry.get()
        d_changed = bool(d.strip()) and self._cap_changed(d, self.settings.get("default_weekly_capacity"))
        self.default_cap_entry.configure(border_color=COLOR_WARN if d_changed else self._entry_border)
        n += d_changed
        self.cap_dirty_label.configure(
            text=f"● {n} kaydedilmemiş değişiklik – 'Kaydet ve Yeniden Hesapla'ya basın" if n else "")

    def _apply_default_to_all(self):
        val = fmt_input(self._default_cap())
        for e in self.cap_entries.values():
            e.delete(0, "end")
            e.insert(0, val)
        self._mark_cap_dirty()

    def _read_capacities(self, warn=True) -> dict:
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
        if bad and warn:
            messagebox.showwarning("Geçersiz değer", "Şu makinelerin kapasitesi sayı değil, atlandı:\n" + ", ".join(bad))
        return caps

    def _save_capacities(self):
        self.settings["default_weekly_capacity"] = self._default_cap()
        self.settings["weekly_capacities"] = self._read_capacities()
        save_settings(self.settings)
        self._mark_cap_dirty()
        self._set_status("Kapasite ayarları kaydedildi.")
        if self.engine is not None:
            self._recalculate()

    def _capacity_args(self, warn=True):
        caps = self._read_capacities(warn) if self.cap_entries else self.settings["weekly_capacities"]
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
        self.summary_key = self._sim_key(months)
        # Simülasyon makine listesi: en dolu makine başta (ilk açılışta en dolu makine seçilir)
        ranked = [m for m in self.summary.sort_values("DOLULUK_%", ascending=False)["Üretim presi"] if m]
        if ranked:
            current = self.sim_machine.get()
            self.sim_machine.configure(values=ranked)
            if current not in ranked or not getattr(self, "_sim_machine_chosen", False):
                self.sim_machine.set(ranked[0])
        self._update_sim_info()
        self._simulate_now(months)   # senaryo varsa yeni ay / kapasiteyle yeniden hesaplanır
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
    # Senaryo, işin PAYI (0–1) olarak tutulur; saat / adet karşılığı seçili aylara göre her seferinde yeniden hesaplanır.
    def _update_sim_info(self):
        """Makine değişti ya da yeniden hesaplandı: iş listesini ve bilgi satırını yenile."""
        if self.engine is None or self.summary is None:
            return
        months = self._selected_months()
        m = normalize_machine(self.sim_machine.get())
        try:
            self.sim_jobs = self.engine.job_list(m, months)
        except Exception as e:
            self.sim_jobs = None
            self.sim_info.configure(text=f"İşler okunamadı: {e}")
            return
        jobs = self.sim_jobs
        if jobs.empty:
            self.sim_info.configure(text="Bu makinede iş yok.")
        else:
            has_alt = [bool(CapacityEngine.job_targets(r)) for _, r in jobs.iterrows()]
            n_alt = sum(has_alt)
            # aktarılabilir = alternatifi var VE seçili aylarda saati var (0 saatlik iş aktarılamaz)
            n_move = sum(1 for a, h in zip(has_alt, jobs["TOPLAM_SAAT"]) if a and h > 0)
            row = self.summary[self.summary["Üretim presi"] == m]
            load = ""
            if len(row):
                r = row.iloc[0]
                load = (f"yük {fmt_num(r['TOPLAM_SAAT'])} sa / kapasite {fmt_num(r['TOPLAM_KAPASİTE'])} sa "
                        f"(%{fmt_num(r['DOLULUK_%'])})  |  ")
            self.sim_info.configure(text=f"{load}{len(jobs)} iş  |  alternatifi olan: {n_alt}  |  "
                                         f"seçili aylarda aktarılabilir: {n_move}")
        self._show_jobs()
        self._show_transfers()

    def _on_sim_machine(self):
        self._sim_machine_chosen = True    # kullanıcı seçti: yeniden hesapta değiştirme
        self._update_sim_info()

    def _assigned(self, row) -> float:
        return sum(t["ratio"] for t in self.sim_transfers if t["row"] == row)

    def _show_jobs(self):
        jobs = self.sim_jobs
        keep = [r.get("SATIR") for r in self.sim_jobs_table.selected_records()]
        if jobs is None or jobs.empty:
            self.sim_jobs_table.clear()
            self._on_job_select()
            return
        eng = self.engine
        view = pd.DataFrame({
            "SATIR": jobs["SATIR"],
            "Ürün": jobs[eng.ref_col],
            "Operasyon": jobs.get("OPERASYON STOK KODU", ""),
            "Firma": jobs.get("FİRMA", ""),
            "Saniye": jobs["ÜRETİM SANİYE"],
            "Adet": jobs["TOPLAM_ADET"],
            "Saat": jobs["TOPLAM_SAAT"],
            "Aktarılan %": [self._assigned(r) * 100 for r in jobs["SATIR"]],
            "Alternatifler": [", ".join(CapacityEngine.job_targets(r)) or "—" for _, r in jobs.iterrows()],
        })

        def tag(r):
            if r["Aktarılan %"] > 0:
                return "moved"
            if r["Saat"] <= 0 or r["Alternatifler"] == "—":
                return "muted"
            return ""
        if self.sim_only_movable.get():   # saati 0 olan veya alternatifi olmayan işleri gizle
            view = view[(view["Saat"] > 0) & (view["Alternatifler"] != "—")]
            if view.empty:
                self.sim_jobs_table.clear()
                self._on_job_select()
                return
        self.sim_jobs_table.show(view, tag_func=tag, hide=["SATIR"])
        if keep:
            self.sim_jobs_table.select_where("SATIR", keep[0])
        self._on_job_select()

    def _selected_job(self):
        """Seçili iş satırı (engine.job_list tablosundan) ya da None."""
        recs = self.sim_jobs_table.selected_records()
        if not recs or self.sim_jobs is None:
            return None
        hit = self.sim_jobs[self.sim_jobs["SATIR"] == recs[0]["SATIR"]]
        return hit.iloc[0] if len(hit) else None

    def _on_job_select(self):
        job = self._selected_job()
        if job is None:
            self.sim_target.configure(values=[""], state="disabled")
            self.sim_target.set("")
            self.sim_add_btn.configure(state="disabled")
            self.sim_amount_hint.configure(text="Listeden bir iş seçin.", text_color="#555")
            return
        targets = CapacityEngine.job_targets(job)
        if not targets:
            self.sim_target.configure(values=[""], state="disabled")
            self.sim_target.set("")
            self.sim_add_btn.configure(state="disabled")
            self.sim_amount_hint.configure(
                text=f"{job[self.engine.ref_col]}: REFERANS PRES'te alternatif pres tanımlı değil – aktarılamaz.",
                text_color=COLOR_WARN_TEXT)
            return
        self.sim_target.configure(values=targets, state="readonly")
        if self.sim_target.get() not in targets:
            self.sim_target.set(targets[0])
        if job["TOPLAM_SAAT"] <= 0:
            self.sim_add_btn.configure(state="disabled")
            self.sim_amount_hint.configure(
                text=(f"{job[self.engine.ref_col]}: seçili aylarda saati 0 (sipariş yok ya da fazla stoktan "
                      "karşılanıyor) – aktarılacak iş yok."), text_color=COLOR_WARN_TEXT)
            return
        self.sim_add_btn.configure(state="normal")
        self._update_amount_hint()

    def _amount_ratio(self, job):
        """Girilen miktarı işin payına (0–1) çevirir. Boş/geçersizse None."""
        txt = self.sim_amount.get().strip()
        if not txt:
            return None
        try:
            v = parse_qty(txt.replace("%", ""))
        except ValueError:
            return None
        unit = self.sim_unit.get()
        if unit == "%":
            return v / 100.0
        if unit == "Adet":
            return v / job["TOPLAM_ADET"] if job["TOPLAM_ADET"] > 0 else None
        return v / job["TOPLAM_SAAT"] if job["TOPLAM_SAAT"] > 0 else None

    def _load_text(self, machine, extra_hours=0.0):
        """'MK178 şu an %30,9' gibi kısa yük bilgisi (eklenecek saatle birlikte)."""
        row = self.summary[self.summary["Üretim presi"] == machine] if self.summary is not None else []
        if len(row):
            h, cap = row.iloc[0]["TOPLAM_SAAT"], row.iloc[0]["TOPLAM_KAPASİTE"]
        else:   # makinenin kendi işi yok (sadece alternatif olarak geçiyor)
            caps, default_cap = self._capacity_args()
            n = sum(1 for w, mm in self.engine.week_month_map.items() if mm in self._selected_months())
            h, cap = 0.0, caps.get(machine, default_cap) * n
        if not cap:
            return f"{machine}: kapasite tanımsız"
        txt = f"{machine} şu an %{fmt_num(h / cap * 100)}"
        if extra_hours:
            txt += f" → bu aktarımla ~%{fmt_num((h + extra_hours) / cap * 100)}"
        return txt

    def _update_amount_hint(self):
        job = self._selected_job()
        if job is None or job["TOPLAM_SAAT"] <= 0 or not CapacityEngine.job_targets(job):
            return
        left = max(0.0, 1.0 - self._assigned(job["SATIR"]))
        rest = (f"Bu işten kalan: %{fmt_num(left * 100)} = {fmt_num(left * job['TOPLAM_SAAT'])} saat "
                f"= {fmt_num(left * job['TOPLAM_ADET'], 0)} adet")
        ratio = self._amount_ratio(job)
        target = self.sim_target.get()
        if ratio is None:
            self.sim_amount_hint.configure(text=f"{rest}\n{self._load_text(target)}", text_color="#555")
            return
        bad = ratio <= 0 or ratio > left + 1e-9
        line = (f"= {fmt_num(ratio * job['TOPLAM_SAAT'])} saat = %{fmt_num(ratio * 100)} "
                f"= {fmt_num(ratio * job['TOPLAM_ADET'], 0)} adet")
        if bad:
            line += "   ⚠ kalandan fazla" if ratio > 0 else "   ⚠ 0'dan büyük olmalı"
        self.sim_amount_hint.configure(
            text=f"{line}   |   {rest}\n{self._load_text(target, 0 if bad else ratio * job['TOPLAM_SAAT'])}",
            text_color=COLOR_OVER if bad else "#333")

    def _fill_remaining(self):
        job = self._selected_job()
        if job is None:
            return
        left = max(0.0, 1.0 - self._assigned(job["SATIR"]))
        unit = self.sim_unit.get()
        val = left * 100 if unit == "%" else left * job["TOPLAM_ADET"] if unit == "Adet" else left * job["TOPLAM_SAAT"]
        self.sim_amount.delete(0, "end")
        self.sim_amount.insert(0, fmt_input(round(val, 0 if unit == "Adet" else 2)))
        self._update_amount_hint()

    def _add_transfer(self):
        job = self._selected_job()
        if job is None or str(self.sim_add_btn.cget("state")) == "disabled":
            return
        target = self.sim_target.get()
        if target not in CapacityEngine.job_targets(job):
            messagebox.showwarning("Hedef", "Hedef olarak işin tanımlı alternatiflerinden birini seçin.")
            return
        ratio = self._amount_ratio(job)
        if ratio is None:
            messagebox.showwarning("Miktar", f"Geçerli bir miktar yazın ({self.sim_unit.get()}).")
            return
        left = max(0.0, 1.0 - self._assigned(job["SATIR"]))
        if ratio <= 0:
            messagebox.showwarning("Miktar", "Miktar 0'dan büyük olmalı.")
            return
        if ratio > left + 1e-9:
            if left <= 1e-9:
                messagebox.showwarning("Miktar", "Bu işin tamamı zaten senaryoda aktarılmış.")
                return
            if not messagebox.askyesno("Kalandan fazla",
                                       f"Bu işten en çok %{fmt_num(left * 100)} "
                                       f"({fmt_num(left * job['TOPLAM_SAAT'])} saat) aktarılabilir.\n\n"
                                       "Kalanın tamamı aktarılsın mı?"):
                return
            ratio = left
        ref = job[self.engine.ref_col]
        same = next((t for t in self.sim_transfers if t["row"] == job["SATIR"] and t["target"] == target), None)
        if same:   # aynı iş + aynı hedef tek satırda toplanır
            same["ratio"] = min(1.0, same["ratio"] + ratio)
        else:
            self.sim_transfers.append({"row": int(job["SATIR"]), "source": job["Üretim presi"], "ref": ref,
                                       "op": job.get("OPERASYON STOK KODU", ""), "target": target,
                                       "ratio": ratio})
        self.sim_amount.delete(0, "end")
        self._show_jobs()
        self._run_simulation()

    def _transfers_view(self):
        """Senaryo listesi: payın seçili aylardaki saat / adet karşılığıyla."""
        rows = self.engine._row_hours(self._selected_months())[0] if self.sim_transfers else None
        out = []
        for i, t in enumerate(self.sim_transfers):
            h = a = 0.0
            if rows is not None and t["row"] in rows.index:
                h = rows.at[t["row"], "TOPLAM_SAAT"] * t["ratio"]
                a = rows.at[t["row"], "TOPLAM_ADET"] * t["ratio"]
            out.append({"NO": i, "Kaynak": t["source"], "Ürün": t["ref"], "Operasyon": t["op"],
                        "Hedef": t["target"], "Pay %": t["ratio"] * 100, "Adet": a, "Saat": h})
        return pd.DataFrame(out, columns=["NO", "Kaynak", "Ürün", "Operasyon", "Hedef", "Pay %", "Adet", "Saat"])

    def _show_transfers(self):
        if not self.sim_transfers or self.engine is None:
            self.sim_list_table.clear()
            return
        view = self._transfers_view()
        total = {"Kaynak": "TOPLAM", "Adet": view["Adet"].sum(), "Saat": view["Saat"].sum()}
        self.sim_list_table.show(view, total_row=total, hide=["NO"])

    def _remove_transfer(self):
        recs = [r for r in self.sim_list_table.selected_records() if "NO" in r]
        if not recs:
            messagebox.showinfo("Seçim yok", "Silmek için senaryo listesinden bir satır seçin.")
            return
        drop = {int(r["NO"]) for r in recs}
        self.sim_transfers = [t for i, t in enumerate(self.sim_transfers) if i not in drop]
        self._show_jobs()
        self._run_simulation()

    def _clear_transfers(self):
        if not self.sim_transfers:
            return
        if not messagebox.askyesno("Temizle", f"Senaryodaki {len(self.sim_transfers)} aktarım silinsin mi?"):
            return
        self.sim_transfers = []
        self._show_jobs()
        self._run_simulation()

    def _run_simulation(self, force=False):
        """Senaryoyu hesapla. 'Önce' değerleri güncel ay / kapasite seçimiyle olmalı:
        değiştiyse önce normal hesap yenilenir (o da sonunda simülasyonu çağırır)."""
        if self.engine is None:
            return
        months = self._selected_months()
        if not months:
            messagebox.showwarning("Ay seçilmedi", "En az bir ay seçin.")
            return
        if force or self.summary is None or self.summary_key != self._sim_key(months):
            self._recalculate()
            return
        self._simulate_now(months)

    def _simulate_now(self, months):
        self._show_transfers()
        if not self.sim_transfers:
            self.sim_cmp = None
            self.sim_result = None
            self.sim_scenario = None
            self.sim_chart.message("Senaryoya aktarım ekleyince sonuç burada görünür.")
            self.sim_table.clear()
            self.sim_summary.configure(text=self.SIM_HINT, text_color="#555")
            return
        try:
            caps, default_cap = self._capacity_args()
            raw = self.engine.simulate_plan(self.sim_transfers, months)
            after = CapacityEngine.apply_capacity(raw, months, caps, default_cap, self.engine.week_month_map)
        except Exception as e:
            messagebox.showerror("Simülasyon hatası", f"{type(e).__name__}: {e}\n\n{traceback.format_exc()[-800:]}")
            return

        weeks = self._active_weeks(self.summary, months)
        mcols = [m for m in months if f"{m}_Saat" in self.summary.columns]
        hcols = (["TOPLAM_SAAT"] + [f"{m}_Saat" for m in mcols] + [f"{w}_Saat" for w in weeks])
        kcols = ["TOPLAM_KAPASİTE", "HAFTALIK_KAPASİTE"] + [f"{m}_Kapasite" for m in mcols]
        b = self.summary[["Üretim presi"] + hcols + kcols]
        a = after[["Üretim presi"] + hcols + kcols]
        cmp_df = b.merge(a, on="Üretim presi", how="outer", suffixes=("_Ö", "_S"))
        for c in hcols:
            cmp_df[[f"{c}_Ö", f"{c}_S"]] = cmp_df[[f"{c}_Ö", f"{c}_S"]].fillna(0)
        for c in kcols:   # kapasite önce/sonra aynı; makine sadece bir tarafta varsa diğerinden al
            cmp_df[c] = cmp_df[f"{c}_Ö"].fillna(cmp_df[f"{c}_S"])
        cmp_df["FARK"] = cmp_df["TOPLAM_SAAT_S"] - cmp_df["TOPLAM_SAAT_Ö"]
        self.sim_cmp = {"df": cmp_df, "months": mcols, "weeks": weeks}
        self.sim_key = self._sim_key(months)

        self._show_sim_table()
        changed = cmp_df[cmp_df["FARK"].abs() > 1e-9].sort_values("FARK")
        names = changed["Üretim presi"].replace("", NO_MACHINE).tolist()
        if changed.empty:
            self.sim_chart.message("Senaryodaki aktarımların seçili aylarda saati yok – değişiklik olmadı.")
        else:
            self.sim_chart.compare_chart(names, changed["TOPLAM_SAAT_Ö"].values, changed["TOPLAM_SAAT_S"].values,
                                         changed["TOPLAM_KAPASİTE"].values,
                                         f"Senaryo ({len(self.sim_transfers)} aktarım) – etkilenen makineler")
        self._sim_summary_text(changed, weeks)

        # Excel raporu için
        self.sim_result = self._sim_view_df("Toplam", only_changed=False)
        self.sim_result.insert(0, "Aylar", ", ".join(months))
        sc = self._transfers_view().drop(columns=["NO"])
        sc.insert(0, "Aylar", ", ".join(months))
        self.sim_scenario = sc
        moved = changed.loc[changed["FARK"] > 0, "FARK"].sum()
        self._set_status(f"Simülasyon: {len(self.sim_transfers)} aktarım, {fmt_num(moved)} saat taşındı.")

    def _sim_summary_text(self, changed, weeks):
        moved = changed.loc[changed["FARK"] > 0, "FARK"].sum()
        parts = []
        for _, r in changed.iterrows():
            cap = r["TOPLAM_KAPASİTE"]
            if cap:
                parts.append(f"{r['Üretim presi'] or NO_MACHINE}: %{fmt_num(r['TOPLAM_SAAT_Ö'] / cap * 100)} → "
                             f"%{fmt_num(r['TOPLAM_SAAT_S'] / cap * 100)}")
        text = f"{len(self.sim_transfers)} aktarım, toplam {fmt_num(moved)} saat taşındı.   " + "   ".join(parts)
        # Yeni taşan haftalar: önce kapasite içindeydi, aktarımdan sonra aşıyor
        warn = []
        for _, r in changed[changed["FARK"] > 0].iterrows():
            wcap = r["HAFTALIK_KAPASİTE"]
            if not wcap:
                continue
            over = [w for w in weeks if r[f"{w}_Saat_S"] > wcap + 1e-9 >= r[f"{w}_Saat_Ö"]]
            if over:
                warn.append(f"{r['Üretim presi']}: {', '.join(over[:6])}{' …' if len(over) > 6 else ''}")
        if warn:
            text += "\n⚠ Aktarımla haftalık kapasiteyi aşmaya başlayan: " + "   ".join(warn)
        self.sim_summary.configure(text=text, text_color=COLOR_WARN_TEXT if warn else "#333")

    def _sim_view_df(self, mode, only_changed=True):
        c = self.sim_cmp
        df = c["df"]
        pct = lambda h, k: (h / k.replace(0, np.nan)) * 100
        view = pd.DataFrame({"Makine": df["Üretim presi"].replace("", NO_MACHINE)})
        if mode == "Aylık":
            for m in c["months"]:
                view[f"{m} önce"] = df[f"{m}_Saat_Ö"]
                view[f"{m} sonra"] = df[f"{m}_Saat_S"]
                view[f"{m} kap."] = df[f"{m}_Kapasite"]
                view[f"{m} önce %"] = pct(df[f"{m}_Saat_Ö"], df[f"{m}_Kapasite"])
                view[f"{m} sonra %"] = pct(df[f"{m}_Saat_S"], df[f"{m}_Kapasite"])
            view["Fark (saat)"] = df["FARK"]
        elif mode == "Haftalık":
            view["Haftalık kap."] = df["HAFTALIK_KAPASİTE"]
            for w in c["weeks"]:
                view[f"{w} önce"] = df[f"{w}_Saat_Ö"]
                view[f"{w} sonra"] = df[f"{w}_Saat_S"]
                view[f"{w} sonra %"] = pct(df[f"{w}_Saat_S"], df["HAFTALIK_KAPASİTE"])
            view["Fark (saat)"] = df["FARK"]
        else:
            view["Önce (saat)"] = df["TOPLAM_SAAT_Ö"]
            view["Sonra (saat)"] = df["TOPLAM_SAAT_S"]
            view["Fark (saat)"] = df["FARK"]
            view["Kapasite (saat)"] = df["TOPLAM_KAPASİTE"]
            view["Önce %"] = pct(df["TOPLAM_SAAT_Ö"], df["TOPLAM_KAPASİTE"])
            view["Sonra %"] = pct(df["TOPLAM_SAAT_S"], df["TOPLAM_KAPASİTE"])
        moved = view["Fark (saat)"].abs() > 1e-9
        first = view[moved].sort_values("Fark (saat)")
        rest = view[~moved]
        if not only_changed:
            sort_col = "Sonra (saat)" if "Sonra (saat)" in view.columns else "Makine"
            return pd.concat([first, rest.sort_values(sort_col, ascending=sort_col == "Makine")])
        return first

    def _show_sim_table(self):
        if not self.sim_cmp:
            return
        mode = self.sim_view.get() or "Toplam"
        view = self._sim_view_df(mode, only_changed=bool(self.sim_only_changed.get()))
        pcols = [c for c in view.columns if c.endswith("sonra %") or c == "Sonra %"]
        tag = lambda r: load_tag(max([r[c] for c in pcols if pd.notna(r[c])], default=np.nan))
        self.sim_table.show(view, tag_func=tag)

    # ---------------- Excel ----------------
    def _sim_key(self, months):
        """Simülasyonun geçerli olduğu koşullar: aylar + kapasite ayarları + yüklü veri.
        Kapasite, hesapla AYNI kaynaktan okunur (kutulara yazılmış, henüz kaydedilmemiş değerler dahil)."""
        caps, default_cap = self._capacity_args(warn=False)
        return (tuple(months), json.dumps(caps, sort_keys=True), default_cap, id(self.engine))

    def _clear_results(self):
        """Veri yeniden yüklenince eski hesap / simülasyon sonuçlarını ekrandan ve hafızadan temizle."""
        self.summary = None
        self.summary_key = None
        self.sim_result = None
        self.sim_scenario = None
        self.sim_key = None
        self.sim_cmp = None
        self.sim_jobs = None
        self.sim_transfers = []     # iş numaraları yeni veride değişebilir: senaryo sıfırlanır
        self._sim_machine_chosen = False
        self.load_table.clear()
        self.load_chart.message("Gösterilecek veri yok.")
        self.sim_table.clear()
        self.sim_jobs_table.clear()
        self.sim_list_table.clear()
        self._on_job_select()
        self.sim_summary.configure(text=self.SIM_HINT, text_color="#555")
        self.sim_chart.message("Senaryoya aktarım ekleyince sonuç burada görünür.")
        self.period_weeks = {}
        self.period_combo.configure(values=[ALL_PERIOD])
        self.period_combo.set(ALL_PERIOD)

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
        # Son simülasyon sadece şu anki aylar + kapasiteyle yapıldıysa rapora girer
        extra, note = None, ""
        if self.sim_result is not None:
            if getattr(self, "sim_key", None) == self._sim_key(months):
                extra = {"Simülasyon": self.sim_result}
                if self.sim_scenario is not None:
                    extra["Simülasyon_Senaryo"] = self.sim_scenario
            else:
                note = ("\n\nNot: son simülasyon farklı ay / kapasite ile yapıldığı için rapora eklenmedi. "
                        "Rapora girmesi için Simülasyon sekmesinde 'Yeniden Hesapla'ya basın.")
        try:
            self.engine.export_to_excel(path, months, caps, default_cap, extra_sheets=extra)
        except PermissionError:
            messagebox.showerror("Hata", "Dosya yazılamadı. Excel'de açıksa kapatıp tekrar deneyin.")
            return
        except Exception as e:
            messagebox.showerror("Hata", f"{type(e).__name__}: {e}")
            return
        self._set_status(f"Rapor kaydedildi: {path}")
        if messagebox.askyesno("Tamam", "Rapor kaydedildi." + note + "\n\nŞimdi açılsın mı?"):
            try:
                os.startfile(path)  # Windows
            except Exception:
                pass


def bring_to_front(win):
    """Yeni açılan pencere ana pencerenin arkasında kalmasın (Windows'ta CTkToplevel bilinen sorunu)."""
    try:
        win.lift()
        win.focus_force()
        win.after(200, win.lift)
        set_icon(win)
    except Exception:
        pass


def op_keys(df: pd.DataFrame) -> pd.Series:
    """Operasyon satırının kimliği: ürün + operasyon stok kodu (büyük/küçük harf farkı yok).
    Operasyon kodu boşsa ürün + pres + saniye kullanılır."""
    u = df["urun_kodu"].astype(str).str.strip().str.upper()
    o = df["operasyon_stok_kodu"].astype(str).str.strip().str.upper()
    p = df["uretim_presi"].astype(str).str.strip().str.upper()
    sn = pd.to_numeric(df["uretim_saniye"], errors="coerce").fillna(0).round(4).astype(str)
    return pd.Series(np.where(o != "", u + "|" + o, u + "||" + p + "|" + sn), index=df.index)


def duplicate_rows(df: pd.DataFrame) -> pd.Series:
    """Listede birden fazla kez geçen operasyon satırları (True = tekrar)."""
    if df.empty:
        return pd.Series(False, index=df.index)
    return op_keys(df).duplicated(keep=False)


class _Parented:
    """messagebox / filedialog pencerelerini verilen pencerenin önünde açar."""
    def __init__(self, mod, parent):
        self.mod, self.parent = mod, parent

    def __getattr__(self, name):
        f = getattr(self.mod, name)
        return lambda *a, **k: f(*a, parent=self.parent, **k)


class OpsWindow(ctk.CTkToplevel):
    """Operasyonlar (REFERANS PRES) penceresi. Değişiklikler 'Kaydet'e kadar taslakta kalır."""

    def __init__(self, app, db_path):
        super().__init__(app)
        self.app, self.db_path = app, db_path
        self.title(f"Operasyonlar – {os.path.basename(db_path)}")
        self.geometry("1300x700")
        self.transient(app)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)
        bring_to_front(self)
        self.ops_draft = pd.DataFrame(columns=OP_COLUMNS)
        self.ops_dirty = False
        t = ctk.CTkFrame(self, fg_color="transparent")
        t.pack(fill="both", expand=True, padx=10, pady=10)
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
        self._ops_reload()

    def _close(self):
        if self.ops_dirty and not _Parented(messagebox, self).askyesno(
                "Kaydedilmemiş değişiklik",
                "Kaydedilmemiş değişiklikler var.\n\nKaydetmeden kapatılsın mı? (Değişiklikler iptal olur)"):
            return
        self.grab_release()
        self.destroy()

    def _ops_reload(self):
        """Taslağı veritabanından yeniden okur (kaydedilmemiş değişiklikler gider)."""
        db = self.db_path
        try:
            self.ops_draft = load_operations(db) if db and os.path.exists(db) else pd.DataFrame(columns=OP_COLUMNS)
        except Exception as e:
            self.ops_draft = pd.DataFrame(columns=OP_COLUMNS)
            _Parented(messagebox, self).showerror("Hata", f"Operasyonlar okunamadı:\n{e}")
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
        dup = duplicate_rows(df)
        dupset = set(df.index[dup].astype(str))
        self.ops_table.show(view.reset_index(drop=True),
                            tag_func=lambda r: "over" if r["#"] in badset else ("warn" if r["#"] in dupset else ""))
        info = f"{len(df)} operasyon satırı, {df['urun_kodu'].nunique()} ürün"
        if q:
            info += f"  |  aramada {len(view)} satır"
        if bad.any():
            info += f"  |  {int(bad.sum())} satırda eksik bilgi (kırmızı: pres veya saniye boş)"
        if dup.any():
            info += f"  |  {int(dup.sum())} satır TEKRAR EDİYOR (turuncu) – saatler fazla sayılır"
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
            _Parented(messagebox, self).showinfo("Seçim", "Düzenlemek için listeden tek bir satır seçin (veya çift tıklayın).")
            return
        OpDialog(self, self.ops_draft.loc[sel[0]].to_dict(), self._ops_machine_names(),
                 lambda rec: self._ops_store(rec, sel[0]))

    def _ops_store(self, rec, idx=None):
        """Satırı taslağa yazar. Aynı ürün + operasyon zaten varsa yazmaz, hata metni döner."""
        others = self.ops_draft if idx is None else self.ops_draft.drop(index=idx)
        if not others.empty:
            key = op_keys(pd.DataFrame([rec], columns=OP_COLUMNS)).iloc[0]
            hit = others.index[op_keys(others) == key]
            if len(hit):
                return (f"Bu ürün + operasyon zaten listede var (satır #{hit[0]}): "
                        f"{rec['urun_kodu']} / {rec['operasyon_stok_kodu'] or '(operasyon kodu boş)'}.\n\n"
                        "Aynı operasyon iki kez olursa saatleri iki kez sayılır. "
                        "Mevcut satırı düzenleyin veya farklı bir operasyon kodu girin.")
        if idx is None:
            self.ops_draft = pd.concat([self.ops_draft, pd.DataFrame([rec], columns=OP_COLUMNS)], ignore_index=True)
        else:
            for c in OP_COLUMNS:
                self.ops_draft.at[idx, c] = rec[c]
        self._ops_set_dirty(True)

    def _ops_delete(self):
        sel = self._ops_selected()
        if not sel:
            _Parented(messagebox, self).showinfo("Seçim", "Silmek için listeden satır seçin (Ctrl / Shift ile birden fazla seçilebilir).")
            return
        if not _Parented(messagebox, self).askyesno("Sil", f"{len(sel)} satır silinsin mi?\n(Kaydet'e basana kadar geri alınabilir: Vazgeç)"):
            return
        self.ops_draft = self.ops_draft.drop(index=sel).reset_index(drop=True)
        self._ops_set_dirty(True)

    def _ops_delete_all(self):
        if self.ops_draft.empty:
            return
        if not _Parented(messagebox, self).askyesno("Tümünü sil", f"{len(self.ops_draft)} satırın TAMAMI silinsin mi?\n"
                                                 "(Kaydet'e basana kadar geri alınabilir: Vazgeç)"):
            return
        self.ops_draft = pd.DataFrame(columns=OP_COLUMNS)
        self._ops_set_dirty(True)

    def _ops_cancel(self):
        if _Parented(messagebox, self).askyesno("Vazgeç", "Kaydedilmemiş tüm değişiklikler iptal edilsin mi?"):
            self._ops_reload()

    def _ops_save(self):
        db = self.db_path
        if not db:
            _Parented(messagebox, self).showwarning("Eksik", "Önce bir veritabanı seçin.")
            return
        dup = duplicate_rows(self.ops_draft)
        if dup.any():
            keys = op_keys(self.ops_draft)
            extra = int(keys.duplicated(keep="first").sum())
            ornek = self.ops_draft[dup].drop_duplicates(["urun_kodu", "operasyon_stok_kodu"]).head(8)
            liste = "\n".join(f"  {r.urun_kodu} / {r.operasyon_stok_kodu or '-'}" for r in ornek.itertuples())
            ans = _Parented(messagebox, self).askyesnocancel(
                "Tekrar eden operasyonlar",
                f"Listede aynı ürün + operasyon {extra} kez fazladan geçiyor (turuncu satırlar), ör.:\n{liste}\n\n"
                "Bu haliyle kaydedilirse bu işlerin saatleri birden fazla sayılır.\n\n"
                "EVET = tekrarları silip (ilkini tutarak) kaydet\nHAYIR = olduğu gibi kaydet\nİPTAL = vazgeç")
            if ans is None:
                return
            if ans:
                self.ops_draft = self.ops_draft[~keys.duplicated(keep="first")].reset_index(drop=True)
                self._ops_show()
        bad = int(((self.ops_draft["uretim_presi"].astype(str).str.strip() == "") |
                   (pd.to_numeric(self.ops_draft["uretim_saniye"], errors="coerce").fillna(0) <= 0)).sum())
        msg = f"{len(self.ops_draft)} operasyon satırı veritabanına kaydedilecek (eski liste bununla değişir)."
        if bad:
            msg += f"\n\n{bad} satırda pres veya saniye boş; bu satırlar saate çevrilemez."
        if not _Parented(messagebox, self).askyesno("Kaydet", msg + "\n\nKaydedilsin mi?"):
            return
        try:
            n = save_operations(db, self.ops_draft)
        except Exception as e:
            _Parented(messagebox, self).showerror("Kaydedilemedi", f"Hiçbir şey değişmedi.\n\n{type(e).__name__}: {e}")
            return
        self._ops_reload()
        self.app._show_db_info()
        self.app._set_status(f"{n} operasyon satırı kaydedildi.")
        self.app._load_data()

    def _ops_import_excel(self):
        path = _Parented(filedialog, self).askopenfilename(title="REFERANS PRES dosyasını seçin",
                                          filetypes=[("Excel dosyaları", "*.xlsx *.xlsm *.xls")])
        if not path:
            return
        try:
            names = sheet_names(path)
        except Exception as e:
            _Parented(messagebox, self).showerror("Hata", f"Dosya açılamadı:\n{e}")
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
            _Parented(messagebox, self).showerror("Bulunamadı", "Dosyada 'BİTMİŞ ÜRÜN KODU' başlıklı bir operasyon sayfası bulunamadı.")
            return
        try:
            df, warns = parse_operations(path, sheet)
        except Exception as e:
            _Parented(messagebox, self).showerror("Okunamadı", str(e))
            return
        msg = f"{os.path.basename(path)} / {sheet}\n{len(df)} operasyon satırı, {df['urun_kodu'].nunique()} ürün okundu."
        if warns:
            msg += "\n\n• " + "\n• ".join(warns)
        if not self.ops_draft.empty:
            ans = _Parented(messagebox, self).askyesnocancel(
                "Excel'den Al", msg + f"\n\nListede şu an {len(self.ops_draft)} satır var.\n"
                "EVET = mevcut listenin yerine geçsin\nHAYIR = mevcut listenin sonuna eklensin\nİPTAL = vazgeç")
            if ans is None:
                return
            if ans:
                self.ops_draft = df
            else:
                existing = set(op_keys(self.ops_draft))
                new = df[~op_keys(df).isin(existing)]
                skipped = len(df) - len(new)
                self.ops_draft = pd.concat([self.ops_draft, new], ignore_index=True)
                if skipped:
                    _Parented(messagebox, self).showinfo(
                        "Zaten listede olanlar atlandı",
                        f"{skipped} satır listede zaten olduğu için (aynı ürün + operasyon) eklenmedi.\n"
                        f"Eklenen yeni satır: {len(new)}")
        else:
            if not _Parented(messagebox, self).askyesno("Excel'den Al", msg + "\n\nListeye alınsın mı?"):
                return
            self.ops_draft = df
        self.ops_draft = self.ops_draft.reset_index(drop=True)
        self._ops_set_dirty(True)
        _Parented(messagebox, self).showinfo("Taslağa alındı", "Satırlar listeye alındı. Kontrol edip 'Kaydet'e basınca veritabanına yazılır.")


def parse_qty(text: str) -> float:
    """Türkçe sayı girişi: '15.000' -> 15000, '1.234,5' -> 1234.5, '12,5' -> 12.5, '1500' -> 1500."""
    t = str(text).strip().replace(" ", "")
    if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", t):     # binlik nokta
        t = t.replace(".", "")
    return float(t.replace(",", "."))


def order_key(urun, firma, yil, hafta) -> str:
    return f"{urun}|{firma}|{int(yil)}|{int(hafta)}"


class OrdersWindow(ctk.CTkToplevel):
    """Hesaptaki siparişleri görme, yeni sipariş ekleme, adet değiştirme ve silme.
    Değişiklikler 'Kaydet'e kadar taslakta kalır; kayıt Aktarımlar'da 'Elle düzenleme' olarak görünür."""

    def __init__(self, app, db_path):
        super().__init__(app)
        self.app, self.db_path = app, db_path
        self.title(f"Siparişler – {os.path.basename(db_path)}")
        self.geometry("1150x680")
        self.transient(app)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)
        bring_to_front(self)
        self.changes = {}     # anahtar -> {urun_kodu, firma, yil, ay, hafta, adet}  (adet 0 = sil)
        try:
            self.products = set(load_operations(db_path)["urun_kodu"].map(normalize_code))
        except Exception:
            self.products = set()

        t = ctk.CTkFrame(self, fg_color="transparent")
        t.pack(fill="both", expand=True, padx=10, pady=10)
        bar = ctk.CTkFrame(t, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 4))
        ctk.CTkButton(bar, text="Yeni Sipariş", width=110, fg_color=COLOR_OK, hover_color="#24804a",
                      command=self._add).pack(side="left", padx=3)
        ctk.CTkButton(bar, text="Adet Değiştir", width=110, command=self._edit).pack(side="left", padx=3)
        ctk.CTkButton(bar, text="Sil", width=60, fg_color=COLOR_OVER, hover_color="#b03636",
                      command=self._delete).pack(side="left", padx=3)
        ctk.CTkLabel(bar, text="  Ay:").pack(side="left")
        self.month_filter = ctk.CTkComboBox(bar, values=["Tümü"], width=150, command=lambda _: self._show())
        self.month_filter.set("Tümü")
        self.month_filter.pack(side="left", padx=3)
        ctk.CTkLabel(bar, text="  Ara:").pack(side="left")
        self.search = ctk.CTkEntry(bar, width=160, placeholder_text="ürün kodu, firma…")
        self.search.pack(side="left", padx=3)
        self.search.bind("<KeyRelease>", lambda _: self._show())
        self.cancel_btn = ctk.CTkButton(bar, text="Vazgeç", width=80, fg_color="gray55", state="disabled",
                                        command=self._cancel)
        self.cancel_btn.pack(side="right", padx=3)
        self.save_btn = ctk.CTkButton(bar, text="Kaydet", width=90, state="disabled",
                                      font=ctk.CTkFont(weight="bold"), command=self._save)
        self.save_btn.pack(side="right", padx=3)
        self.label = ctk.CTkLabel(t, text="", anchor="w")
        self.label.pack(fill="x", padx=4)
        self.table = DataTable(t)
        self.table.pack(fill="both", expand=True)
        self.table.tree.bind("<Double-1>", lambda _: self._edit())
        self.table.tree.bind("<Delete>", lambda _: self._delete())
        self._reload()

    # --- veri ---
    def _reload(self):
        try:
            self.base = load_valid_orders(self.db_path)
        except Exception as e:
            _Parented(messagebox, self).showerror("Hata", f"Siparişler okunamadı:\n{e}")
            self.base = pd.DataFrame(columns=["urun_kodu", "firma", "yil", "ay", "hafta", "adet", "kaynak"])
        self.changes = {}
        self._set_dirty()

    def _set_dirty(self):
        st = "normal" if self.changes else "disabled"
        self.save_btn.configure(state=st)
        self.cancel_btn.configure(state=st)
        self._show()

    def _merged(self) -> pd.DataFrame:
        df = self.base.copy()
        df["_key"] = [order_key(*k) for k in zip(df["urun_kodu"], df["firma"], df["yil"], df["hafta"])]
        df["durum"] = ""
        df = df.set_index("_key", drop=False)
        for k, c in self.changes.items():
            if k in df.index:
                old = df.at[k, "adet"]
                df.at[k, "adet"] = c["adet"]
                df.at[k, "durum"] = "SİLİNECEK" if c["adet"] == 0 else f"DEĞİŞTİ (önce {fmt_num(old, 0)})"
            elif c["adet"] > 0:
                df.loc[k] = {**c, "_key": k, "kaynak": "yeni", "tarih": "", "aktarim_id": None,
                             "hafta_yili": None, "durum": "YENİ"}
        df["ay_no"] = df["ay"].map(lambda a: AY_ADLARI.index(a) + 1 if a in AY_ADLARI else 99)
        return df.sort_values(["yil", "ay_no", "hafta", "urun_kodu", "firma"]).reset_index(drop=True)

    def _show(self):
        df = self._merged()
        months = list(dict.fromkeys(f"{a} {y}" for a, y in zip(df["ay"], df["yil"])))
        self.month_filter.configure(values=["Tümü"] + months)
        sel = self.month_filter.get()
        if sel != "Tümü" and sel in months:
            df = df[(df["ay"] + " " + df["yil"].astype(str)) == sel]
        q = fold(self.search.get()) if self.search.get().strip() else ""
        if q:
            df = df[(df["urun_kodu"].map(fold) + " " + df["firma"].map(fold)).str.contains(re.escape(q), regex=True)]
        view = pd.DataFrame({
            "#": df["_key"], "Ay": df["ay"] + " " + df["yil"].astype(str), "Hafta": "W" + df["hafta"].astype(int).astype(str),
            "Firma": df["firma"], "Ürün kodu": df["urun_kodu"], "Adet": df["adet"].astype(float),
            "Kaynak": df["kaynak"], "Durum": df["durum"],
        })
        tag = lambda r: "over" if r["Durum"] == "SİLİNECEK" else ("warn" if r["Durum"] else "")
        self.table.show(view.reset_index(drop=True), tag_func=tag)
        bk = set(self.base_keys())
        n_new = sum(1 for k, c in self.changes.items() if c["adet"] > 0 and k not in bk)
        n_del = sum(1 for c in self.changes.values() if c["adet"] == 0)
        info = f"{len(view)} sipariş satırı gösteriliyor"
        if self.changes:
            info += (f"   ●  KAYDEDİLMEMİŞ: {n_new} yeni, {len(self.changes) - n_new - n_del} adet değişikliği, "
                     f"{n_del} silme – 'Kaydet' demeden hesaba yansımaz")
        self.label.configure(text=info, text_color=COLOR_OVER if self.changes else "#333")

    @property
    def firms(self):
        vals = set(self.base["firma"].dropna()) | {c["firma"] for c in self.changes.values()}
        return sorted(v for v in vals if v)

    def week_month_in_db(self, yil, hafta, ay):
        """Bu yılın bu haftası veritabanında başka bir aya kayıtlıysa (ay, yil) döner, değilse None."""
        df = self.base
        wy = hafta_yili(int(yil), AY_ADLARI.index(ay) + 1, int(hafta))
        m = (pd.to_numeric(df["hafta"]) == int(hafta)) & (pd.to_numeric(df["hafta_yili"], errors="coerce") == wy)
        found = df.loc[m, ["ay", "yil"]].drop_duplicates()
        found = found[~((found["ay"] == ay) & (found["yil"].astype(int) == int(yil)))]
        if found.empty:
            return None
        r = found.iloc[0]
        return r["ay"], int(r["yil"])

    def other_firms(self, rec):
        """Aynı ürün + yıl + hafta için başka firma adıyla duran siparişler [(firma, adet), ...]."""
        df = self._merged()
        m = ((df["urun_kodu"] == rec["urun_kodu"]) & (df["yil"].astype(int) == int(rec["yil"])) &
             (df["hafta"].astype(int) == int(rec["hafta"])) & (df["adet"] > 0) & (df["firma"] != rec["firma"]))
        return list(zip(df.loc[m, "firma"], df.loc[m, "adet"]))

    def base_keys(self):
        return [order_key(*k) for k in zip(self.base["urun_kodu"], self.base["firma"], self.base["yil"], self.base["hafta"])]

    def _selected(self):
        return [self.table.tree.item(i, "values")[0] for i in self.table.tree.selection()]

    def _row(self, key) -> dict:
        df = self._merged().set_index("_key")
        r = df.loc[key]
        return {"urun_kodu": r["urun_kodu"], "firma": r["firma"], "yil": int(r["yil"]), "ay": r["ay"],
                "hafta": int(r["hafta"]), "adet": float(r["adet"])}

    # --- işlemler ---
    def _add(self):
        sel = self.month_filter.get()
        default = None
        if sel != "Tümü":
            ay, yil = sel.rsplit(" ", 1)
            default = {"ay": ay, "yil": int(yil)}
        OrderDialog(self, None, self._store, default=default)

    def _edit(self):
        sel = self._selected()
        if len(sel) != 1:
            _Parented(messagebox, self).showinfo("Seçim", "Adedini değiştirmek için tek bir satır seçin (veya çift tıklayın).")
            return
        OrderDialog(self, self._row(sel[0]), lambda rec: self._store(rec, edit=True))

    def _store(self, rec, edit=False):
        """Taslağa yazar. Hata/vazgeçme varsa metin döner (pencere açık kalır)."""
        k = order_key(rec["urun_kodu"], rec["firma"], rec["yil"], rec["hafta"])
        if not edit:
            current = self._merged().set_index("_key")
            if k in current.index and current.at[k, "adet"] > 0:
                old = float(current.at[k, "adet"])
                ans = _Parented(messagebox, self).askyesnocancel(
                    "Bu sipariş zaten var",
                    f"{rec['urun_kodu']} / {rec['firma'] or '(firma yok)'} / {rec['ay']} {rec['yil']} W{rec['hafta']} "
                    f"için zaten {fmt_num(old, 0)} adet var.\n\n"
                    f"EVET = {fmt_num(rec['adet'], 0)} adet olsun (eskinin yerine)\n"
                    f"HAYIR = üstüne eklensin ({fmt_num(old + rec['adet'], 0)} adet)\nİPTAL = vazgeç")
                if ans is None:
                    return "vazgeçildi"
                if not ans:
                    rec = {**rec, "adet": old + rec["adet"]}
        base_keys = set(self.base_keys())
        if k not in base_keys and rec["adet"] == 0:
            self.changes.pop(k, None)      # kaydedilmemiş yeni satırın silinmesi
        else:
            self.changes[k] = dict(rec)
        self._set_dirty()
        return None

    def _delete(self):
        sel = self._selected()
        if not sel:
            _Parented(messagebox, self).showinfo("Seçim", "Silmek için satır seçin (Ctrl / Shift ile birden fazla).")
            return
        if not _Parented(messagebox, self).askyesno(
                "Sil", f"{len(sel)} sipariş satırı silinsin (iptal) mi?\n"
                       "Kaydet'e kadar 'Vazgeç' ile geri alınabilir; kaydettikten sonra da Aktarımlar'dan geri alınabilir."):
            return
        for k in sel:
            self._store({**self._row(k), "adet": 0.0}, edit=True)

    def _cancel(self):
        if _Parented(messagebox, self).askyesno("Vazgeç", "Kaydedilmemiş tüm değişiklikler iptal edilsin mi?"):
            self.changes = {}
            self._set_dirty()

    def _save(self):
        if not self.changes:
            return
        n_del = sum(1 for c in self.changes.values() if c["adet"] == 0)
        months = sorted({f"{c['ay']} {c['yil']}" for c in self.changes.values()})
        dlg = ctk.CTkInputDialog(title="Kaydet",
                                 text=f"{len(self.changes)} değişiklik ({n_del} silme) kaydedilecek: {', '.join(months)}\n\n"
                                      "Kısa bir açıklama yazın (ör. 'Valeo W41 iptal'), boş bırakabilirsiniz:")
        aciklama = dlg.get_input()
        if aciklama is None:
            return
        aciklama = aciklama.strip() if isinstance(aciklama, str) else ""
        try:
            save_manual_orders(self.db_path, list(self.changes.values()), aciklama)
        except Exception as e:
            _Parented(messagebox, self).showerror("Kaydedilemedi", f"Hiçbir şey değişmedi.\n\n{type(e).__name__}: {e}")
            return
        self._reload()
        self.app._show_db_info()
        self.app._refresh_imports()
        self.app._set_status("Sipariş değişiklikleri kaydedildi (Aktarımlar: 'Elle düzenleme').")
        self.app._load_data()

    def _close(self):
        if self.changes and not _Parented(messagebox, self).askyesno(
                "Kaydedilmemiş değişiklik", "Kaydedilmemiş değişiklikler var.\n\nKaydetmeden kapatılsın mı? (Değişiklikler iptal olur)"):
            return
        self.grab_release()
        self.destroy()


class OrderDialog(ctk.CTkToplevel):
    """Yeni sipariş ekleme veya adet değiştirme."""

    def __init__(self, master, rec, on_ok, default=None):
        super().__init__(master)
        self.edit = rec is not None
        self.title("Adet değiştir" if self.edit else "Yeni sipariş")
        self.geometry("420x360")
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        bring_to_front(self)
        self.on_ok, self.owner = on_ok, master
        now = datetime.now()
        rec = rec or {"urun_kodu": "", "firma": "", "yil": (default or {}).get("yil", now.year),
                      "ay": (default or {}).get("ay", AY_ADLARI[now.month - 1]), "hafta": "", "adet": ""}
        f = ctk.CTkFrame(self)
        f.pack(fill="both", expand=True, padx=12, pady=12)
        self.w = {}
        fields = [("urun_kodu", "Ürün kodu *"), ("firma", "Firma"), ("yil", "Yıl *"), ("ay", "Ay *"),
                  ("hafta", "Hafta (W) *"), ("adet", "Adet *")]
        for i, (k, label) in enumerate(fields):
            ctk.CTkLabel(f, text=label, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=5)
            if k == "ay":
                w = ctk.CTkComboBox(f, values=AY_ADLARI, width=200)
            elif k == "firma":   # var olan firmalardan seç ya da yeni yaz
                w = ctk.CTkComboBox(f, values=[""] + list(getattr(master, "firms", [])), width=200)
                w.set("")
            else:
                w = ctk.CTkEntry(f, width=200)
            w.grid(row=i, column=1, padx=8, pady=5)
            val = rec.get(k, "")
            if k == "adet" and val != "":
                val = fmt_input(val)
            if isinstance(w, ctk.CTkComboBox):
                w.set(str(val))
            elif str(val) != "":
                w.insert(0, str(val))
            if self.edit and k != "adet":
                w.configure(state="disabled")
            self.w[k] = w
        ctk.CTkLabel(f, text=("Sadece adet değiştirilebilir. Siparişi kaldırmak için 'Sil'i kullanın."
                              if self.edit else "Hafta numarası ISO takvimine göre (W41 = 5–11 Ekim 2026)."),
                     text_color="#666", wraplength=380, justify="left").grid(row=len(fields), column=0, columnspan=2,
                                                                          sticky="w", padx=8, pady=6)
        b = ctk.CTkFrame(self, fg_color="transparent")
        b.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(b, text="Tamam", command=self._ok).pack(side="right", padx=4)
        ctk.CTkButton(b, text="İptal", fg_color="gray55", command=self.destroy).pack(side="right", padx=4)

    def destroy(self):
        owner = self.owner
        super().destroy()
        try:
            owner.grab_set()
            owner.lift()
        except Exception:
            pass

    def _ok(self):
        mb = _Parented(messagebox, self)
        g = lambda k: self.w[k].get().strip()
        rec = {"urun_kodu": normalize_code(g("urun_kodu")), "firma": normalize_firma(g("firma")), "ay": g("ay").upper()}
        errors = []
        try:
            rec["yil"] = int(g("yil"))
            assert 2000 <= rec["yil"] <= 2100
        except Exception:
            errors.append("Yıl 2000–2100 arası bir sayı olmalı.")
        try:
            rec["hafta"] = int(g("hafta").upper().lstrip("W"))
            assert 1 <= rec["hafta"] <= 53
        except Exception:
            errors.append("Hafta 1–53 arası bir sayı olmalı (ör. 41 veya W41).")
        try:
            rec["adet"] = parse_qty(g("adet"))
            assert rec["adet"] >= 0 and (self.edit or rec["adet"] > 0)
        except Exception:
            errors.append("Adet 0'dan büyük bir sayı olmalı.")
        if not rec["urun_kodu"]:
            errors.append("Ürün kodu boş olamaz.")
        elif not re.search(r"\d", rec["urun_kodu"]):
            errors.append("Ürün kodunda en az bir rakam olmalı.")
        if rec["ay"] not in AY_ADLARI:
            errors.append("Ay listeden seçilmeli.")
        if errors:
            mb.showwarning("Eksik / hatalı bilgi", "\n".join(errors))
            return
        if rec["adet"] != int(rec["adet"]):
            yazilan = g("adet")
            okunan = f"{rec['adet']:g}".replace(".", ",")
            if not mb.askyesno("Kesirli adet",
                               f"'{yazilan}' yazdınız; bu {okunan} adet (kesirli) olarak okundu.\n\n"
                               "Binlik ayırıcı için NOKTA kullanın: 1.500 = bin beş yüz, 1,5 = bir buçuk.\n\n"
                               f"{okunan} adet olarak kaydedilsin mi?"):
                return
        if not self.edit:
            other = self.owner.week_month_in_db(rec["yil"], rec["hafta"], rec["ay"])
            if other:
                ans = mb.askyesnocancel(
                    "Hafta başka ayda",
                    f"W{rec['hafta']} haftası veritabanında {other[0]} {other[1]} ayında duruyor; "
                    f"siz {rec['ay']} {rec['yil']} yazdınız.\n\nFarklı aya yazılırsa hesapta yine {other[0]} {other[1]} "
                    "ayına sayılabilir.\n\n"
                    f"EVET = ayı {other[0]} {other[1]} yap\nHAYIR = {rec['ay']} {rec['yil']} olarak bırak\nİPTAL = vazgeç")
                if ans is None:
                    return
                if ans:
                    rec["ay"], rec["yil"] = other
            ay_no = AY_ADLARI.index(rec["ay"]) + 1
            problem = week_month_problem(rec["hafta"], hafta_yili(rec["yil"], ay_no, rec["hafta"]), ay_no, rec["yil"])
            if problem and not mb.askyesno("Hafta – ay uyumu", problem + "\n\nYine de eklensin mi?"):
                return
            others = self.owner.other_firms(rec)
            if others and not mb.askyesno(
                    "Başka firmanın siparişi var",
                    f"{rec['urun_kodu']} için {rec['ay']} {rec['yil']} W{rec['hafta']} haftasında şu firma(lar)ın "
                    f"siparişi zaten var:\n" + "\n".join(f"  • {f or '(firma yok)'}: {fmt_num(a, 0)} adet" for f, a in others) +
                    f"\n\n'{rec['firma'] or '(firma yok)'}' için AYRI, yeni bir sipariş olarak eklensin mi?\n"
                    "(Aynı müşteriyse 'Hayır' deyip firmayı listeden seçin.)"):
                return
            if self.owner.products and rec["urun_kodu"] not in self.owner.products:
                if not mb.askyesno("Operasyon yok",
                                   f"{rec['urun_kodu']} için Operasyonlar listesinde kayıt yok; bu siparişin saati "
                                   "hesaplanamaz ('Eşleşmeyenler'de görünür).\n\nYine de eklensin mi?"):
                    return
        err = self.on_ok(rec)
        if err and err != "vazgeçildi":
            mb.showwarning("Eklenemedi", err)
            return
        if err != "vazgeçildi":
            self.destroy()


class OpDialog(ctk.CTkToplevel):
    """Tek operasyon satırını ekleme / düzenleme penceresi (REFERANS PRES şablonundaki sütunlar)."""
    FIELDS = [("musteri", "MÜŞTERİ", False), ("urun_kodu", "BİTMİŞ ÜRÜN KODU", True),
              ("operasyon_stok_kodu", "OPERASYON STOK KODU", False), ("pres_tonaji", "BASILDIĞI PRES TONAJI", False),
              ("uretim_presi", "Üretim presi", True), ("alternatif_1", "alternatif 1", False),
              ("alternatif_2", "alternatif 2", False), ("alternatif_3", "alternatif 3", False),
              ("alternatif_4", "alternatif 4", False), ("uretim_saniye", "ÜRETİM SANİYE", True)]
    MACHINE_FIELDS = {"uretim_presi", "alternatif_1", "alternatif_2", "alternatif_3", "alternatif_4"}

    def destroy(self):
        owner = self.master
        super().destroy()
        try:   # Operasyonlar penceresi tekrar önde ve kilitli kalsın
            owner.grab_set()
            owner.lift()
        except Exception:
            pass

    def __init__(self, master, rec, machines, on_ok):
        super().__init__(master)
        self.title("Operasyon düzenle" if rec else "Yeni operasyon")
        self.geometry("460x520")
        self.transient(master)
        self.grab_set()
        self.on_ok = on_ok
        self.machines = list(machines)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        bring_to_front(self)
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
        err = self.on_ok(rec)
        if err:
            messagebox.showwarning("Tekrar eden operasyon", err, parent=self)
            return
        self.destroy()


class ImportDialog(ctk.CTkToplevel):
    """Plan Excel'ini seç → oku → yılları kontrol et → veritabanına yaz."""

    def __init__(self, master, db_path, on_done):
        super().__init__(master)
        self.title("Excel'den Veritabanına Aktar")
        bring_to_front(self)
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
        m = ctk.CTkFrame(self)
        m.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(m, text="Bu ay veritabanında zaten varsa:").pack(side="left", padx=8, pady=6)
        self.mode = ctk.StringVar(value="yerine")
        ctk.CTkRadioButton(m, text="Yerine geçsin (aynı dosyanın güncel hali)", variable=self.mode,
                           value="yerine").pack(side="left", padx=8)
        ctk.CTkRadioButton(m, text="Üstüne eklensin (başka müşterinin dosyası)", variable=self.mode,
                           value="ekle").pack(side="left", padx=8)
        self.info = ctk.CTkTextbox(self, height=160)
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
            n = int(((plan.orders["blok"] == i) & (plan.orders["adet"] > 0)).sum())
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
                 f"Sipariş satırı (adet > 0): {int((plan.orders['adet'] > 0).sum())}",
                 f"Backlog/stok satırı: {len(plan.backlog)}  → {plan.blocks[0].ay} ayına yazılacak",
                 "Zararlı Excel'e karşı XML koruması: " + ("AÇIK" if XML_PROTECTED else
                                                          "KAPALI (defusedxml kurulu değil – build_exe.bat'ı yeniden çalıştırın)")]
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
            mod = self.mode.get() or "yerine"
            notes = check_conflicts(self.db_path, self.plan, years, mod)
        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanı kontrol edilemedi:\n{e}", parent=self)
            return
        summary = ", ".join(f"{b.ay} {years[i]}" for i, b in enumerate(self.plan.blocks))
        msg = f"Şu aylar veritabanına yazılacak:\n{summary}"
        if not (self.plan.orders["adet"] > 0).any():
            msg += "\n\nUYARI: Dosyada hiç sipariş adedi bulunamadı (tüm haftalar boş veya 0)."
        if notes:
            msg += "\n\nDİKKAT:\n• " + "\n• ".join(notes)
        if not messagebox.askyesno("Onay", msg + "\n\nDevam edilsin mi?", parent=self):
            return
        try:
            ids = import_plan(self.db_path, self.plan, years, mod)
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

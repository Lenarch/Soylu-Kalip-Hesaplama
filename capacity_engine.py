import pandas as pd
import numpy as np
import re
import sqlite3
import os
from typing import List, Optional, Dict, Tuple


def normalize_code(value) -> str:
    """[DÜZELTME 1] Ürün kodunu tek tip metne çevirir: 12345.0 -> '12345', NaN -> ''."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    if s.upper() in ("NAN", "NONE"):
        return ""
    return s


def normalize_machine(value) -> str:
    """[DÜZELTME 6] Makine adını tek tip yapar: ' mk 5 ' -> 'MK5'. Hem dosyadaki sütunlar hem de parametre için kullanılır."""
    if pd.isna(value):
        return ""
    s = str(value).strip().upper().replace("MK ", "MK")
    if s in ("NAN", "NONE"):
        return ""
    return s


class CapacityEngine:
    def __init__(self, plan_file: str, op_file: str):
        self.plan_file = plan_file
        self.op_file = op_file
        self.df_plan = None
        self.df_op = None
        self.week_month_map: Dict[str, str] = {}
        self.available_months: List[str] = []
        self.week_columns: List[str] = []
        self.backlog_columns: List[str] = []
        self.unmapped_df = None
        self.merged_df = None
        self.ref_col: str = "REFERANS"
        self.warnings: List[str] = []
        self.backlog_month_map: Dict[str, List[str]] = {}
        self.extra_qty_columns: List[str] = []
        self.backlog_used_month: Optional[str] = None

    def load_data(self, plan_sheet: str = "Diğer müşteriler DAHİL", op_sheet: str = "Uyum Operasyon Bazlı Saniyeler"):
        """Excel dosyalarını okur, aylar ve haftaları dinamik haritalandırır."""
        # 1. Plan Dosyasını Oku (3. satır Aylar, 4. satır Sütunlar)
        raw_plan = pd.read_excel(self.plan_file, sheet_name=plan_sheet, header=None)

        month_row = raw_plan.iloc[2].ffill()  # Sağ hücrelere doğru ayı yay
        header_row = raw_plan.iloc[3]

        self.df_plan = raw_plan.iloc[4:].copy()
        # Aynı başlık birden fazla kez geçebilir (ör. her ayın altında "TOPLAM"):
        # tekrar edenlere .1, .2 eki verip sütun adlarını benzersiz yap
        names, seen = [], {}
        for i, col in enumerate(header_row):
            base = str(col).strip() if pd.notna(col) and str(col).strip() else f"COL_{i}"
            n = seen.get(base, 0)
            seen[base] = n + 1
            names.append(base if n == 0 else f"{base}.{n}")
        self.df_plan.columns = names

        # Referans sütununu bul
        ref_col = next((c for c in self.df_plan.columns if "REFERANS" in c.upper() or "ÜRÜN" in c.upper()), "REFERANS")
        self.ref_col = ref_col
        # [DÜZELTME 1] Kodları tek tip metne çevir
        self.df_plan[ref_col] = self.df_plan[ref_col].apply(normalize_code)

        # [DÜZELTME 5] Boş referanslı satırları ve TOPLAM satırlarını at
        ref_upper = self.df_plan[ref_col].str.upper()
        self.df_plan = self.df_plan[
            (self.df_plan[ref_col] != "") & (~ref_upper.str.contains("TOPLAM", na=False))
        ].copy()

        # Hafta (W...) ve Backlog sütunlarını keşfet
        self.week_columns = [c for c in self.df_plan.columns if re.match(r"^W\d+", c, re.IGNORECASE)]
        self.backlog_columns = [c for c in self.df_plan.columns if "BACKLOG" in c.upper()]

        # Hafta -> Ay Haritası
        self.week_month_map = {}
        for idx, col_str in enumerate(self.df_plan.columns):
            if col_str in self.week_columns:
                ay = str(month_row[idx]).strip() if pd.notna(month_row[idx]) else "BİLİNMEYEN"
                self.week_month_map[col_str] = ay

        self.available_months = list(dict.fromkeys(self.week_month_map.values()))

        # Sayısal değerleri temizle
        for col in (self.week_columns + self.backlog_columns):
            self.df_plan[col] = pd.to_numeric(self.df_plan[col], errors="coerce").fillna(0)

        # 2. Operasyon Dosyasını Oku
        self.df_op = pd.read_excel(self.op_file, sheet_name=op_sheet)
        return self._finalize()

    # ------------------------------------------------------------------
    # Veritabanından okuma (SQLite)
    # ------------------------------------------------------------------
    def load_from_db(self, db_path: str):
        """SQLite veritabanından okur; Excel'deki ile aynı tablo yapısını kurar.
        Sadece geçerli yüklemeler okunur (her yıl + ay için en son, geri alınmamış yükleme)."""
        from veritabani import connect, create_empty_db, hafta_yili
        if not os.path.exists(db_path):
            raise FileNotFoundError(f"Veritabanı bulunamadı: {db_path}")
        create_empty_db(db_path)   # eksik tablo/görünüm varsa ekler, veriye dokunmaz
        con = connect(db_path)
        try:
            ops = pd.read_sql_query("SELECT * FROM operasyonlar", con)
            orders = pd.read_sql_query("SELECT urun_kodu, firma, yil, hafta, hafta_yili, ay, adet FROM gecerli_siparisler", con)
            backlog = pd.read_sql_query(
                "SELECT urun_kodu, firma, yil, ay, backlog_adet, stok_adet FROM gecerli_backlog", con)
        finally:
            con.close()

        self.warnings = []
        self.ref_col = "REFERANS"
        AYLAR = ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
                 "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
        ay_no = lambda a: AYLAR.index(a) + 1 if a in AYLAR else 99
        label = lambda ay, yil: f"{ay} {yil}"

        # --- Siparişler: uzun tablo -> haftalık geniş tablo ---
        orders["urun_kodu"] = orders["urun_kodu"].apply(normalize_code)
        orders = orders[orders["urun_kodu"] != ""].copy()
        for c in ("yil", "hafta"):
            orders[c] = pd.to_numeric(orders[c], errors="coerce")
        bad = orders["yil"].isna() | orders["hafta"].isna()
        if bad.any():
            self.warnings.append(f"{int(bad.sum())} sipariş satırında yıl/hafta boş veya hatalı, atlandı.")
            orders = orders[~bad]
        orders["yil"] = orders["yil"].astype(int)
        orders["hafta"] = orders["hafta"].astype(int)
        orders["ay"] = orders["ay"].fillna("").astype(str).str.strip().str.upper()
        orders["adet"] = pd.to_numeric(orders["adet"], errors="coerce").fillna(0)

        # Haftanın yılı (ayın yılından farklı olabilir: 2026 ARALIK içindeki W1 -> 2027'nin haftası)
        hy = pd.to_numeric(orders["hafta_yili"], errors="coerce")
        orders["hafta_yili"] = [int(v) if pd.notna(v) else hafta_yili(y, ay_no(a), h)
                                for v, y, a, h in zip(hy, orders["yil"], orders["ay"], orders["hafta"])]

        # Hafta kimliği = (haftanın yılı, hafta). Birden fazla yıl varsa sütun adına yıl eklenir.
        multi_year = orders["hafta_yili"].nunique() > 1
        orders["hafta_kol"] = [f"W{h}/{y}" if multi_year else f"W{h}"
                               for h, y in zip(orders["hafta"], orders["hafta_yili"])]

        # Her hafta tek bir aya (ay + ayın yılı) ait olmalı
        week_month, week_key = {}, {}
        for (wy, h, kol), grp in orders.groupby(["hafta_yili", "hafta", "hafta_kol"], sort=True):
            grp = grp[grp["ay"] != ""]
            if grp.empty:
                y, ay = wy, "BİLİNMEYEN"
                self.warnings.append(f"W{h} ({wy}) haftasının ayı yazılmamış.")
            else:
                pairs = grp.groupby(["yil", "ay"]).size().sort_values(ascending=False)
                y, ay = pairs.index[0]
                if len(pairs) > 1:
                    self.warnings.append(f"W{h} ({wy} yılının haftası) birden fazla aya yazılmış "
                                         f"({', '.join(f'{a} {yy}' for yy, a in pairs.index)}); "
                                         f"'{ay} {y}' kabul edildi.")
            week_month[kol] = (int(y), ay)
            week_key[kol] = (int(wy), int(h))
        # ayları takvim sırasına, ay içinde haftaları sırasına göre diz
        week_cols = sorted(week_month, key=lambda k: (week_month[k][0], ay_no(week_month[k][1]), week_key[k]))
        self.week_month_map = {k: label(week_month[k][1], week_month[k][0]) for k in week_cols}
        months_sorted = sorted({week_month[k] for k in week_cols}, key=lambda t: (t[0], ay_no(t[1])))
        self.available_months = [label(a, y) for y, a in months_sorted]

        wide = (orders.pivot_table(index="urun_kodu", columns="hafta_kol", values="adet",
                                   aggfunc="sum", fill_value=0)
                .reindex(columns=week_cols, fill_value=0))

        # --- Backlog: ay ay ayrı sütunlar ("Backlog@EYLÜL 2026") ---
        backlog["urun_kodu"] = backlog["urun_kodu"].apply(normalize_code)
        backlog = backlog[backlog["urun_kodu"] != ""].copy()
        for c in ("backlog_adet", "stok_adet"):
            backlog[c] = pd.to_numeric(backlog[c], errors="coerce").fillna(0)
        backlog["lbl"] = [label(str(a).strip().upper(), int(y)) for a, y in zip(backlog["ay"], backlog["yil"])]
        backlog["guncel"] = backlog["backlog_adet"] - backlog["stok_adet"]
        self.backlog_month_map = {}
        bl_parts = []
        bl_labels = sorted(backlog["lbl"].unique(),
                           key=lambda l: (int(l.split()[-1]), ay_no(" ".join(l.split()[:-1]))))
        for lbl in bl_labels:
            g = backlog[backlog["lbl"] == lbl].groupby("urun_kodu")[["backlog_adet", "guncel"]].sum()
            g.columns = [f"Backlog@{lbl}", f"Güncel Backlog@{lbl}"]
            bl_parts.append(g)
            self.backlog_month_map[lbl] = list(g.columns)

        plan = wide
        for g in bl_parts:
            plan = plan.join(g, how="outer")
        plan = plan.fillna(0)
        plan.index.name = self.ref_col
        plan = plan.reset_index()
        # gösterim için standart sütunlar: en son ayın backlog'u (hesapta seçilen aya göre değişir)
        if bl_labels:
            plan["Backlog"] = plan[f"Backlog@{bl_labels[-1]}"]
            plan["Güncel Backlog"] = plan[f"Güncel Backlog@{bl_labels[-1]}"]
        else:
            plan["Backlog"] = 0.0
            plan["Güncel Backlog"] = 0.0

        # --- Firma bilgisi: ürünü isteyen firmalar ---
        firms = pd.concat([orders[["urun_kodu", "firma"]], backlog[["urun_kodu", "firma"]]], ignore_index=True)
        firms["firma"] = firms["firma"].fillna("").astype(str).str.strip()
        firms = firms[firms["firma"] != ""]
        firm = firms.groupby("urun_kodu")["firma"].agg(lambda x: ", ".join(sorted(set(x))))
        plan.insert(1, "FİRMA", plan[self.ref_col].map(firm).fillna(""))

        self.df_plan = plan
        self.week_columns = week_cols
        self.backlog_columns = ["Backlog", "Güncel Backlog"]
        self.extra_qty_columns = [c for cols in self.backlog_month_map.values() for c in cols]

        # --- Operasyonlar: Excel'deki sütun adlarına çevir ---
        self.df_op = ops.rename(columns={
            "urun_kodu": "BİTMİŞ ÜRÜN KODU", "operasyon_stok_kodu": "OPERASYON STOK KODU",
            "pres_tonaji": "BASILDIĞI PRES TONAJI", "uretim_presi": "Üretim presi",
            "alternatif_1": "alternatif 1", "alternatif_2": "alternatif 2",
            "alternatif_3": "alternatif 3", "alternatif_4": "alternatif 4",
            "uretim_saniye": "ÜRETİM SANİYE"})
        if not len(self.df_op):
            self.warnings.append("operasyonlar tablosu boş – siparişler saate çevrilemez (makine yükü 0 görünür).")
        if not week_cols:
            self.warnings.append("Veritabanında geçerli sipariş yok – önce 'Excel'den Aktar' ile plan yükleyin.")
        return self._finalize()

    # ------------------------------------------------------------------
    # Excel ve veritabanı için ortak son adım
    # ------------------------------------------------------------------
    def _finalize(self):
        ref_col = self.ref_col
        if not hasattr(self, "warnings"):
            self.warnings = []

        # Sütunları standartlaştır
        # [DÜZELTME 1] Operasyon tarafındaki kodları da aynı şekilde normalize et
        self.df_op["BİTMİŞ ÜRÜN KODU"] = self.df_op["BİTMİŞ ÜRÜN KODU"].apply(normalize_code)
        self.df_op["ÜRETİM SANİYE"] = pd.to_numeric(self.df_op["ÜRETİM SANİYE"], errors="coerce").fillna(0)

        # Makine isimlerini düzelt
        for m_col in ["Üretim presi", "alternatif 1", "alternatif 2", "alternatif 3", "alternatif 4"]:
            if m_col in self.df_op.columns:
                self.df_op[m_col] = self.df_op[m_col].apply(normalize_machine)

        # 3. Eşleşmeyenleri Tespit Et (Madde 3)
        plan_refs = set(self.df_plan[ref_col].unique())
        op_refs = set(self.df_op["BİTMİŞ ÜRÜN KODU"].unique())
        unmapped_refs = plan_refs - op_refs

        self.unmapped_df = self.df_plan[self.df_plan[ref_col].isin(unmapped_refs)].copy()

        # 4. Ana Eşleştirme (Merge)
        self.merged_df = pd.merge(
            self.df_plan,
            self.df_op[[
                "BİTMİŞ ÜRÜN KODU", "OPERASYON STOK KODU", "BASILDIĞI PRES TONAJI",
                "Üretim presi", "alternatif 1", "alternatif 2", "alternatif 3", "alternatif 4", "ÜRETİM SANİYE"
            ]],
            left_on=ref_col,
            right_on="BİTMİŞ ÜRÜN KODU",
            how="inner"
        )

        return self

    def calculate_workload(self, selected_months: Optional[List[str]] = None) -> pd.DataFrame:
        """
        Madde 1: İstenirse belirli aylar, seçilmezse tüm aylar için hesaplama yapar.
        """
        df = self.merged_df.copy()

        # Hangi haftalar hesaba katılacak?
        if selected_months:
            active_weeks = [w for w, m in self.week_month_map.items() if m in selected_months]
            active_months = selected_months
        else:
            active_weeks = self.week_columns
            active_months = self.available_months

        # Veritabanında backlog ay ay saklanır: seçili dönemin İLK ayının backlog'u alınır
        bmap = getattr(self, "backlog_month_map", {}) or {}
        chosen = None
        if bmap:
            chosen = next((m for m in active_months if m in bmap), None)
            if chosen is None:   # seçili ayların hiçbirinde backlog yoksa, öncesindeki en yakın ay
                order = self.available_months
                first = min((order.index(m) for m in active_months if m in order), default=None)
                before = [m for m in bmap if m in order and first is not None and order.index(m) <= first]
                chosen = before[-1] if before else None
            self.backlog_used_month = chosen
            for b in self.backlog_columns:
                df[b] = df[f"{b}@{chosen}"] if chosen else 0.0

        # Fazla stok (stok > backlog): Güncel Backlog 0 sayılır, artan stok AYNI ürünün siparişlerinden
        # tarih sırasıyla düşülür (backlog'un ait olduğu ayın ilk haftasından itibaren, seçili olmasa bile).
        df["STOK_MAHSUP_ADET"] = 0.0
        if "Güncel Backlog" in df.columns:
            remaining = (-df["Güncel Backlog"]).clip(lower=0)
            df["Güncel Backlog"] = df["Güncel Backlog"].clip(lower=0)
            order = self.available_months
            if chosen in order:
                start = order.index(chosen)
                weeks_from = [w for w in self.week_columns
                              if self.week_month_map.get(w) in order and order.index(self.week_month_map[w]) >= start]
            else:
                weeks_from = list(self.week_columns)
            for w in weeks_from:
                if not (remaining > 0).any():
                    break
                take = np.minimum(df[w], remaining)
                df[w] = df[w] - take
                remaining = remaining - take
                if w in active_weeks:
                    df["STOK_MAHSUP_ADET"] += take

        # Saatleri hesapla: (Adet * Saniye) / 3600
        for w in active_weeks:
            df[f"{w}_Saat"] = (df[w] * df["ÜRETİM SANİYE"]) / 3600.0
        for b in self.backlog_columns:
            df[f"{b}_Saat"] = (df[b] * df["ÜRETİM SANİYE"]) / 3600.0
        df["STOKTAN_KARSILANAN_SAAT"] = df["STOK_MAHSUP_ADET"] * df["ÜRETİM SANİYE"] / 3600.0

        # Aylık Toplamlar
        ay_saat_cols = []
        for ay in active_months:
            m_weeks = [f"{w}_Saat" for w, m in self.week_month_map.items() if m == ay and w in active_weeks]
            if m_weeks:
                col_name = f"{ay}_Saat"
                df[col_name] = df[m_weeks].sum(axis=1)
                ay_saat_cols.append(col_name)

        # Genel Toplam
        all_week_saat = [f"{w}_Saat" for w in active_weeks]
        df["TOPLAM_SAAT"] = df[all_week_saat].sum(axis=1)

        # Makine Bazında Özet
        group_cols = (ay_saat_cols + all_week_saat + [f"{b}_Saat" for b in self.backlog_columns]
                      + ["STOKTAN_KARSILANAN_SAAT", "TOPLAM_SAAT"])
        summary = df.groupby("Üretim presi")[group_cols].sum().reset_index()
        summary = summary.sort_values(by="TOPLAM_SAAT", ascending=False).reset_index(drop=True)
        return summary

    def simulate_transfer(
        self,
        overloaded_machine: str,
        transfer_pct: float = 30.0,
        target_alt_col: str = "alternatif 1",
        selected_months: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """
        Madde 2: Darboğazdaki makineden alternatif makineye iş aktarma simülasyonu.
        """
        df = self.merged_df.copy()
        ratio = transfer_pct / 100.0

        # [DÜZELTME 6] Parametre olarak gelen makine adını da aynı kuralla normalize et ("mk 5" -> "MK5")
        overloaded_machine = normalize_machine(overloaded_machine)

        # Seçilen makinenin işleri
        mask = (df["Üretim presi"] == overloaded_machine) & (df[target_alt_col] != "") & (df[target_alt_col].notna())

        # İki kopya oluştur: Kalan iş ana makinede, aktarılan iş alternatif makinede
        df_normal = df[~mask].copy()

        df_stay = df[mask].copy()
        df_transfer = df[mask].copy()

        # Adetleri bölüştür
        qty_cols = self.week_columns + self.backlog_columns + list(getattr(self, "extra_qty_columns", []))
        for q in qty_cols:
            df_stay[q] = df_stay[q] * (1.0 - ratio)
            df_transfer[q] = df_transfer[q] * ratio

        # Aktarılanların makinesini alternatif makine yap
        df_transfer["Üretim presi"] = df_transfer[target_alt_col]

        # Tekrar birleştir
        df_sim = pd.concat([df_normal, df_stay, df_transfer], ignore_index=True)

        # Simülasyonu hesapla
        engine_sim = CapacityEngine("", "")
        engine_sim.merged_df = df_sim
        engine_sim.week_columns = self.week_columns
        engine_sim.backlog_columns = self.backlog_columns
        engine_sim.week_month_map = self.week_month_map
        engine_sim.available_months = self.available_months
        engine_sim.backlog_month_map = getattr(self, "backlog_month_map", {})
        engine_sim.extra_qty_columns = getattr(self, "extra_qty_columns", [])

        return engine_sim.calculate_workload(selected_months=selected_months)

    # ------------------------------------------------------------------
    # Uygulama için yardımcılar
    # ------------------------------------------------------------------
    def machines(self) -> List[str]:
        """Planla eşleşen işlerdeki ana üretim preslerinin listesi."""
        if self.merged_df is None:
            return []
        return sorted(m for m in self.merged_df["Üretim presi"].unique() if m)

    def all_machines(self) -> List[str]:
        """Ana presler + alternatif olarak geçen tüm presler (kapasite girişi için)."""
        if self.merged_df is None:
            return []
        cols = [c for c in ["Üretim presi", "alternatif 1", "alternatif 2", "alternatif 3", "alternatif 4"]
                if c in self.merged_df.columns]
        vals = pd.unique(self.merged_df[cols].values.ravel())
        return sorted(v for v in vals if isinstance(v, str) and v)

    @staticmethod
    def apply_capacity(summary: pd.DataFrame, months: List[str],
                       capacities: Dict[str, float], default_cap: float,
                       week_month_map: Dict[str, str]) -> pd.DataFrame:
        """Özet tabloya kapasite ve doluluk ekler.
        capacities / default_cap: makine başına HAFTALIK kullanılabilir saat.
        Ay kapasitesi = haftalık kapasite × o ayın planda bulunan hafta sayısı."""
        df = summary.copy()
        cap = df["Üretim presi"].map(lambda m: float(capacities.get(m, default_cap) or 0)).astype(float)
        capn = cap.replace(0, np.nan)
        df["HAFTALIK_KAPASİTE"] = cap
        weeks = [w for w, m in week_month_map.items() if m in months and f"{w}_Saat" in df.columns]
        for w in weeks:
            df[f"{w}_Doluluk_%"] = df[f"{w}_Saat"] / capn * 100
        for m in months:
            if f"{m}_Saat" not in df.columns:
                continue
            n = sum(1 for w in weeks if week_month_map[w] == m)
            df[f"{m}_Hafta"] = n
            df[f"{m}_Kapasite"] = cap * n
            df[f"{m}_Doluluk_%"] = df[f"{m}_Saat"] / (cap * n).replace(0, np.nan) * 100
        df["HAFTA_SAYISI"] = len(weeks)
        df["TOPLAM_KAPASİTE"] = cap * len(weeks)
        df["DOLULUK_%"] = df["TOPLAM_SAAT"] / df["TOPLAM_KAPASİTE"].replace(0, np.nan) * 100
        return df

    def export_to_excel(self, output_file: str, selected_months: Optional[List[str]] = None,
                        capacities: Optional[Dict[str, float]] = None, default_cap: float = 0.0,
                        extra_sheets: Optional[Dict[str, pd.DataFrame]] = None):
        """Sonuçları, Eşleşmeyenleri ve Detayları çok sayfalı Excel'e basar."""
        summary = self.calculate_workload(selected_months)
        if capacities is not None:
            months = selected_months or self.available_months
            summary = self.apply_capacity(summary, months, capacities, default_cap, self.week_month_map)

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            summary.round(2).to_excel(writer, sheet_name="Makine_Yük_Özeti", index=False)
            for name, sheet_df in (extra_sheets or {}).items():
                sheet_df.round(2).to_excel(writer, sheet_name=name[:31], index=False)
            if self.unmapped_df is not None:
                self.unmapped_df.to_excel(writer, sheet_name="Eşleşmeyen_Parçalar", index=False)
            self.merged_df.round(2).to_excel(writer, sheet_name="Tüm_Operasyon_Detayları", index=False)
        return output_file

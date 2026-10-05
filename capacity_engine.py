import pandas as pd
import numpy as np
import re
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

        # Saatleri hesapla: (Adet * Saniye) / 3600
        for w in active_weeks:
            df[f"{w}_Saat"] = (df[w] * df["ÜRETİM SANİYE"]) / 3600.0

        for b in self.backlog_columns:
            df[f"{b}_Saat"] = (df[b] * df["ÜRETİM SANİYE"]) / 3600.0

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
        group_cols = ay_saat_cols + all_week_saat + [f"{b}_Saat" for b in self.backlog_columns] + ["TOPLAM_SAAT"]
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
        qty_cols = self.week_columns + self.backlog_columns
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
                       capacities: Dict[str, float], default_cap: float) -> pd.DataFrame:
        """Özet tabloya aylık kapasite ve doluluk yüzdelerini ekler."""
        df = summary.copy()
        cap = df["Üretim presi"].map(lambda m: float(capacities.get(m, default_cap) or 0)).astype(float)
        df["AYLIK_KAPASİTE"] = cap
        month_cols = [m for m in months if f"{m}_Saat" in df.columns]
        df["TOPLAM_KAPASİTE"] = cap * len(month_cols)
        for m in month_cols:
            df[f"{m}_Doluluk_%"] = np.where(cap > 0, df[f"{m}_Saat"] / cap.replace(0, np.nan) * 100, np.nan)
        tot_cap = df["TOPLAM_KAPASİTE"].replace(0, np.nan)
        df["DOLULUK_%"] = df["TOPLAM_SAAT"] / tot_cap * 100
        return df

    def export_to_excel(self, output_file: str, selected_months: Optional[List[str]] = None,
                        capacities: Optional[Dict[str, float]] = None, default_cap: float = 0.0,
                        extra_sheets: Optional[Dict[str, pd.DataFrame]] = None):
        """Sonuçları, Eşleşmeyenleri ve Detayları çok sayfalı Excel'e basar."""
        summary = self.calculate_workload(selected_months)
        if capacities is not None:
            months = selected_months or self.available_months
            summary = self.apply_capacity(summary, months, capacities, default_cap)

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            summary.round(2).to_excel(writer, sheet_name="Makine_Yük_Özeti", index=False)
            for name, sheet_df in (extra_sheets or {}).items():
                sheet_df.round(2).to_excel(writer, sheet_name=name[:31], index=False)
            if self.unmapped_df is not None:
                self.unmapped_df.to_excel(writer, sheet_name="Eşleşmeyen_Parçalar", index=False)
            self.merged_df.round(2).to_excel(writer, sheet_name="Tüm_Operasyon_Detayları", index=False)
        return output_file

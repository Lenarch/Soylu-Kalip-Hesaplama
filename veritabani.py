"""
Kapasite uygulamasının SQLite veritabanı.

Tablolar
--------
aktarimlar   : İşlem kaydı. Her Excel yüklemesi (dosya + ay) bir satır. Geri alınabilir.
siparisler   : Ürün + firma bazında haftalık sipariş adetleri (yıl + hafta + ay ile)
backlog      : Ürün + firma bazında, AY AY backlog ve stok adetleri
operasyonlar : Ürünün operasyonları, presi, alternatifleri, parça başı saniye
urunler      : (İsteğe bağlı) ürün açıklaması

Kural: Aynı yıl + ay için birden fazla yükleme varsa EN SON (geri alınmamış) yükleme geçerlidir.
Eski yüklemeler silinmez; son yükleme geri alınırsa bir öncekisi tekrar geçerli olur.
aktarim_id'si boş (NULL) satırlar elle girilmiş kabul edilir ve her zaman geçerlidir.
"""
import os
import sqlite3
from datetime import datetime

SCHEMA_VERSION = 2

SCHEMA = """
CREATE TABLE IF NOT EXISTS aktarimlar (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    tarih         TEXT NOT NULL,                 -- yükleme zamanı
    dosya_adi     TEXT NOT NULL,
    sayfa         TEXT,
    yil           INTEGER NOT NULL,
    ay            TEXT NOT NULL,                 -- AĞUSTOS, EYLÜL ...
    ay_no         INTEGER NOT NULL,              -- 1..12 (sıralama için)
    haftalar      TEXT,                          -- "34,35"
    siparis_satir INTEGER NOT NULL DEFAULT 0,
    backlog_satir INTEGER NOT NULL DEFAULT 0,
    durum         TEXT NOT NULL DEFAULT 'aktif'  -- aktif / geri alındı
        CHECK (durum IN ('aktif', 'geri alındı')),
    geri_alma_tarihi TEXT
);

CREATE TABLE IF NOT EXISTS siparisler (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    aktarim_id  INTEGER REFERENCES aktarimlar(id),   -- boşsa elle girilmiş
    urun_kodu   TEXT    NOT NULL,
    firma       TEXT    NOT NULL DEFAULT '',
    yil         INTEGER NOT NULL,                  -- AYIN yılı (ARALIK 2026 -> 2026)
    hafta       INTEGER NOT NULL CHECK (hafta BETWEEN 1 AND 53),
    hafta_yili  INTEGER,                           -- HAFTANIN yılı (2026 ARALIK içindeki W1 -> 2027)
    ay          TEXT    NOT NULL,
    adet        REAL    NOT NULL DEFAULT 0 CHECK (adet >= 0),
    UNIQUE (aktarim_id, urun_kodu, firma, yil, hafta)
);
CREATE INDEX IF NOT EXISTS ix_siparisler_aktarim ON siparisler(aktarim_id);
CREATE INDEX IF NOT EXISTS ix_siparisler_urun ON siparisler(urun_kodu);

CREATE TABLE IF NOT EXISTS backlog (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    aktarim_id    INTEGER REFERENCES aktarimlar(id),
    urun_kodu     TEXT NOT NULL,
    firma         TEXT NOT NULL DEFAULT '',
    yil           INTEGER NOT NULL,
    ay            TEXT NOT NULL,
    backlog_adet  REAL NOT NULL DEFAULT 0,
    stok_adet     REAL NOT NULL DEFAULT 0,
    UNIQUE (aktarim_id, urun_kodu, firma, yil, ay)
);
CREATE INDEX IF NOT EXISTS ix_backlog_aktarim ON backlog(aktarim_id);

CREATE TABLE IF NOT EXISTS operasyonlar (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    musteri              TEXT,                 -- MÜŞTERİ
    urun_kodu            TEXT NOT NULL,        -- BİTMİŞ ÜRÜN KODU
    operasyon_stok_kodu  TEXT,                 -- OPERASYON STOK KODU
    pres_tonaji          TEXT,                 -- BASILDIĞI PRES TONAJI (ör. 250 TON, HIZLI PRES)
    uretim_presi         TEXT NOT NULL,
    alternatif_1         TEXT,
    alternatif_2         TEXT,
    alternatif_3         TEXT,
    alternatif_4         TEXT,
    uretim_saniye        REAL NOT NULL CHECK (uretim_saniye >= 0)
);
CREATE INDEX IF NOT EXISTS ix_operasyonlar_urun ON operasyonlar(urun_kodu);

CREATE TABLE IF NOT EXISTS urunler (
    urun_kodu   TEXT PRIMARY KEY,
    aciklama    TEXT
);

-- Hesapta kullanılan (geçerli) yüklemeler: her yıl + ay için en son aktif yükleme
CREATE VIEW IF NOT EXISTS gecerli_aktarimlar AS
SELECT a.* FROM aktarimlar a
WHERE a.durum = 'aktif'
  AND a.id = (SELECT MAX(b.id) FROM aktarimlar b
              WHERE b.durum = 'aktif' AND b.yil = a.yil AND b.ay = a.ay);

CREATE VIEW IF NOT EXISTS gecerli_siparisler AS
SELECT s.* FROM siparisler s
WHERE s.aktarim_id IS NULL OR s.aktarim_id IN (SELECT id FROM gecerli_aktarimlar);

CREATE VIEW IF NOT EXISTS gecerli_backlog AS
SELECT k.* FROM backlog k
WHERE k.aktarim_id IS NULL OR k.aktarim_id IN (SELECT id FROM gecerli_aktarimlar);
"""


def hafta_yili(yil: int, ay_no: int, hafta: int) -> int:
    """Haftanın ait olduğu yıl. Aralık'taki W1, W2 ... bir sonraki yılın; Ocak'taki W52, W53 bir önceki yılın haftasıdır."""
    if ay_no == 12 and hafta < 10:
        return yil + 1
    if ay_no == 1 and hafta > 40:
        return yil - 1
    return yil


HAFTA_YILI_SQL = ("CASE WHEN ay='ARALIK' AND hafta<10 THEN yil+1 "
                  "WHEN ay='OCAK' AND hafta>40 THEN yil-1 ELSE yil END")


def connect(path: str) -> sqlite3.Connection:
    con = sqlite3.connect(path)
    con.execute("PRAGMA foreign_keys = ON")
    return con


def _columns(con, table):
    return [r[1] for r in con.execute(f"PRAGMA table_info({table})")]


def _migrate(con):
    """İlk sürümün (yıl/ay/işlem kaydı olmayan) BOŞ tablolarını yeni yapıya çevirir."""
    old = []
    if "siparisler" in _tables(con) and "aktarim_id" not in _columns(con, "siparisler"):
        old.append("siparisler")
    if "backlog" in _tables(con) and "aktarim_id" not in _columns(con, "backlog"):
        old.append("backlog")
    for t in old:
        n = con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        if n:
            raise RuntimeError(f"'{t}' tablosu eski yapıda ve içinde {n} kayıt var. "
                               "Önce bu kayıtları yedekleyin; otomatik dönüştürme yapılmadı.")
        con.execute(f"DROP TABLE {t}")


def _tables(con):
    return {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}


def create_empty_db(path: str) -> str:
    """Tabloları oluşturur (var olan veriye dokunmaz; ilk sürümün boş tablolarını günceller)."""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    con = connect(path)
    try:
        _migrate(con)
        con.executescript(SCHEMA)
        # önceki sürümde 'hafta_yili' yoktu: sütunu ekle ve var olan satırlar için hesapla
        if "musteri" not in _columns(con, "operasyonlar"):
            con.execute("ALTER TABLE operasyonlar ADD COLUMN musteri TEXT")
        if "hafta_yili" not in _columns(con, "siparisler"):
            con.execute("ALTER TABLE siparisler ADD COLUMN hafta_yili INTEGER")
        con.execute(f"UPDATE siparisler SET hafta_yili = {HAFTA_YILI_SQL} WHERE hafta_yili IS NULL")
        con.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        con.commit()
    finally:
        con.close()
    return path


def ensure_schema(path: str):
    if os.path.exists(path):
        create_empty_db(path)


def table_counts(path: str) -> dict:
    con = connect(path)
    try:
        out = {}
        for t in ("siparisler", "backlog", "operasyonlar", "aktarimlar"):
            try:
                out[t] = con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            except sqlite3.Error:
                out[t] = None
        return out
    finally:
        con.close()


# ----------------------------------------------------------------------
# İşlem kaydı (yükleme geçmişi)
# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
# Operasyonlar (REFERANS PRES karşılığı) – uygulamadaki düzenleme ekranı için
# ----------------------------------------------------------------------
OP_COLUMNS = ["musteri", "urun_kodu", "operasyon_stok_kodu", "pres_tonaji", "uretim_presi",
              "alternatif_1", "alternatif_2", "alternatif_3", "alternatif_4", "uretim_saniye"]


def load_operations(path: str):
    import pandas as pd
    create_empty_db(path)
    con = connect(path)
    try:
        df = pd.read_sql_query(f"SELECT {', '.join(OP_COLUMNS)} FROM operasyonlar ORDER BY id", con)
    finally:
        con.close()
    for c in OP_COLUMNS:
        if c != "uretim_saniye":
            df[c] = df[c].fillna("").astype(str).str.strip()
    df["uretim_saniye"] = pd.to_numeric(df["uretim_saniye"], errors="coerce").fillna(0.0)
    return df


def save_operations(path: str, df) -> int:
    """Operasyonlar tablosunu verilen listeyle DEĞİŞTİRİR (tek işlem: hata olursa eski hali kalır)."""
    create_empty_db(path)
    rows = []
    for r in df[OP_COLUMNS].itertuples(index=False):
        d = r._asdict()
        sn = float(d["uretim_saniye"]) if d["uretim_saniye"] not in ("", None) else 0.0
        rows.append(tuple((d[c] if c != "uretim_saniye" else sn) if c in ("urun_kodu", "uretim_presi", "uretim_saniye")
                          else (d[c] or None) for c in OP_COLUMNS))
    con = connect(path)
    try:
        with con:
            con.execute("DELETE FROM operasyonlar")
            con.executemany(f"INSERT INTO operasyonlar({', '.join(OP_COLUMNS)}) "
                            f"VALUES ({', '.join('?' * len(OP_COLUMNS))})", rows)
    finally:
        con.close()
    return len(rows)


def list_imports(path: str):
    """Tüm yüklemeler, en yeni en üstte. 'gecerli' = hesapta kullanılıyor mu."""
    con = connect(path)
    try:
        rows = con.execute("""
            SELECT a.id, a.tarih, a.dosya_adi, a.sayfa, a.yil, a.ay, a.haftalar,
                   a.siparis_satir, a.backlog_satir, a.durum,
                   CASE WHEN g.id IS NULL THEN 0 ELSE 1 END AS gecerli
            FROM aktarimlar a LEFT JOIN gecerli_aktarimlar g ON g.id = a.id
            ORDER BY a.id DESC""").fetchall()
        cols = ["id", "tarih", "dosya_adi", "sayfa", "yil", "ay", "haftalar",
                "siparis_satir", "backlog_satir", "durum", "gecerli"]
        return [dict(zip(cols, r)) for r in rows]
    finally:
        con.close()


def undo_import(path: str, aktarim_id: int):
    """Yüklemeyi geri alır (veri silinmez, yükleme 'geri alındı' olur; varsa bir önceki yükleme geçerli olur)."""
    con = connect(path)
    try:
        cur = con.execute("UPDATE aktarimlar SET durum='geri alındı', geri_alma_tarihi=? "
                          "WHERE id=? AND durum='aktif'",
                          (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), aktarim_id))
        con.commit()
        return cur.rowcount == 1
    finally:
        con.close()


if __name__ == "__main__":
    import sys
    p = sys.argv[1] if len(sys.argv) > 1 else "kapasite.db"
    create_empty_db(p)
    print("Oluşturuldu:", os.path.abspath(p), table_counts(p))
